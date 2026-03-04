import json
import os
import urllib.error
import urllib.parse
import urllib.request

import boto3


class LLMProvider:
    def __init__(self):
        self.provider = os.getenv("LLM_PROVIDER", "claude").strip().lower()

        self.claude_model_id = os.getenv(
            "CLAUDE_MODEL_ID",
            "global.anthropic.claude-sonnet-4-5-20250929-v1:0",
        )
        self.claude_model_version = os.getenv("CLAUDE_MODEL_VERSION", "bedrock-2023-05-31")
        self.aws_region = os.getenv("AWS_REGION", "ap-south-1")

        raw_gemini_model = os.getenv("GEMINI_MODEL", "gemini-2.0-flash")
        self.gemini_model = self._normalize_gemini_model(raw_gemini_model)
        self.gemini_api_key = os.getenv("GEMINI_API_KEY")

        self.claude_client = None
        if self.provider == "claude":
            self.claude_client = boto3.client(
                "bedrock-runtime",
                region_name=self.aws_region,
                aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
                aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
            )
        elif self.provider == "gemini":
            if not self.gemini_api_key:
                raise ValueError("GEMINI_API_KEY is required when LLM_PROVIDER=gemini")
        else:
            raise ValueError("Unsupported LLM_PROVIDER. Use 'claude' or 'gemini'.")

    def _normalize_gemini_model(self, model_name):
        model_name = (model_name or "gemini-2.0-flash").strip()
        if model_name.startswith("models/"):
            model_name = model_name.split("models/", 1)[1]
        if model_name == "gemini-1.5-pro":
            return "gemini-2.0-flash"
        return model_name

    def describe(self):
        if self.provider == "claude":
            return f"claude ({self.claude_model_id})"
        return f"gemini ({self.gemini_model})"

    def generate_text(self, prompt, max_tokens=4000, temperature=0):
        if self.provider == "claude":
            request_body = {
                "anthropic_version": self.claude_model_version,
                "max_tokens": max_tokens,
                "temperature": temperature,
                "messages": [
                    {
                        "role": "user",
                        "content": [{"type": "text", "text": prompt}],
                    }
                ],
            }

            response = self.claude_client.invoke_model(
                modelId=self.claude_model_id,
                body=json.dumps(request_body),
            )
            response_body = json.loads(response.get("body").read())
            return response_body["content"][0]["text"].strip()

        endpoint = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            f"{urllib.parse.quote(self.gemini_model, safe='')}:generateContent"
            f"?key={urllib.parse.quote(self.gemini_api_key, safe='')}"
        )
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": temperature,
                "maxOutputTokens": max_tokens,
            },
        }

        request = urllib.request.Request(
            endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                response_body = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            body = error.read().decode("utf-8", errors="ignore")
            if error.code == 404 and self.gemini_model != "gemini-2.0-flash":
                fallback_model = "gemini-2.0-flash"
                fallback_endpoint = (
                    "https://generativelanguage.googleapis.com/v1beta/models/"
                    f"{urllib.parse.quote(fallback_model, safe='')}:generateContent"
                    f"?key={urllib.parse.quote(self.gemini_api_key, safe='')}"
                )
                fallback_request = urllib.request.Request(
                    fallback_endpoint,
                    data=json.dumps(payload).encode("utf-8"),
                    headers={"Content-Type": "application/json"},
                    method="POST",
                )
                try:
                    with urllib.request.urlopen(fallback_request, timeout=120) as response:
                        response_body = json.loads(response.read().decode("utf-8"))
                    self.gemini_model = fallback_model
                except urllib.error.HTTPError:
                    raise RuntimeError(f"Gemini API error {error.code}: {body}") from error
            else:
                raise RuntimeError(f"Gemini API error {error.code}: {body}") from error

        candidates = response_body.get("candidates", [])
        if not candidates:
            raise RuntimeError(f"Gemini response missing candidates: {response_body}")

        parts = candidates[0].get("content", {}).get("parts", [])
        text = "\n".join(part.get("text", "") for part in parts if part.get("text"))
        if not text:
            raise RuntimeError(f"Gemini response missing text parts: {response_body}")

        return text.strip()

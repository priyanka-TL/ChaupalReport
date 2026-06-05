import json
import os
import time
from datetime import datetime
import urllib.error
import urllib.parse
import urllib.request

import boto3
from openai import OpenAI


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

        self.openrouter_api_key = os.getenv("OPENROUTER_API_KEY")
        self.openrouter_model = os.getenv("OPENROUTER_MODEL", "google/gemini-2.0-flash")
        self.openrouter_client = None

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
        elif self.provider == "openrouter":
            if not self.openrouter_api_key:
                raise ValueError("OPENROUTER_API_KEY is required when LLM_PROVIDER=openrouter")
            self.openrouter_client = OpenAI(
                api_key=self.openrouter_api_key,
                base_url="https://openrouter.ai/api/v1",
            )
        else:
            raise ValueError("Unsupported LLM_PROVIDER. Use 'claude', 'gemini', or 'openrouter'.")

        self.call_counter = 0
        self.log_file = os.getenv("LLM_CALL_LOG_FILE", "ai_calls.log")

    def _next_call_id(self):
        self.call_counter += 1
        return self.call_counter

    def _current_model(self):
        if self.provider == "claude":
            return self.claude_model_id
        if self.provider == "openrouter":
            return self.openrouter_model
        return self.gemini_model

    def _append_log(self, payload):
        try:
            with open(self.log_file, "a", encoding="utf-8") as file:
                file.write(json.dumps(payload, ensure_ascii=False) + "\n")
        except Exception:
            pass

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
        if self.provider == "openrouter":
            return f"openrouter ({self.openrouter_model})"
        return f"gemini ({self.gemini_model})"

    def generate_text(self, prompt, max_tokens=4000, temperature=0, thinking_budget=None):
        call_id = self._next_call_id()
        started_at = time.time()
        prompt_text = str(prompt)
        prompt_preview = prompt_text[:180].replace("\n", " ")

        print(
            f"🤖 AI Call #{call_id} started | provider={self.provider} | model={self._current_model()} | prompt_chars={len(prompt_text)}",
            flush=True,
        )
        self._append_log(
            {
                "timestamp": datetime.utcnow().isoformat() + "Z",
                "event": "start",
                "call_id": call_id,
                "provider": self.provider,
                "model": self._current_model(),
                "max_tokens": max_tokens,
                "temperature": temperature,
                "thinking_budget": thinking_budget,
                "prompt_chars": len(prompt_text),
                "prompt_preview": prompt_preview,
            }
        )

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
            output_text = response_body["content"][0]["text"].strip()
            duration = time.time() - started_at
            print(f"✅ AI Call #{call_id} completed in {duration:.2f}s", flush=True)
            self._append_log(
                {
                    "timestamp": datetime.utcnow().isoformat() + "Z",
                    "event": "success",
                    "call_id": call_id,
                    "provider": self.provider,
                    "model": self._current_model(),
                    "duration_seconds": round(duration, 3),
                    "response_chars": len(output_text),
                }
            )
            return output_text

        if self.provider == "openrouter":
            try:
                response = self.openrouter_client.chat.completions.create(
                    model=self.openrouter_model,
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=max_tokens,
                    temperature=temperature,
                    extra_headers={
                        "HTTP-Referer": "chaupal-report-pipeline",
                        "X-Title": "ChaupalReport",
                    },
                )
                output_text = response.choices[0].message.content.strip()
                duration = time.time() - started_at
                print(f"✅ AI Call #{call_id} completed in {duration:.2f}s", flush=True)
                self._append_log(
                    {
                        "timestamp": datetime.utcnow().isoformat() + "Z",
                        "event": "success",
                        "call_id": call_id,
                        "provider": self.provider,
                        "model": self._current_model(),
                        "duration_seconds": round(duration, 3),
                        "response_chars": len(output_text),
                    }
                )
                return output_text
            except Exception as error:
                duration = time.time() - started_at
                error_message = str(error)
                print(f"❌ AI Call #{call_id} failed in {duration:.2f}s: {error_message}", flush=True)
                self._append_log(
                    {
                        "timestamp": datetime.utcnow().isoformat() + "Z",
                        "event": "error",
                        "call_id": call_id,
                        "provider": self.provider,
                        "model": self._current_model(),
                        "duration_seconds": round(duration, 3),
                        "error": error_message,
                    }
                )
                raise

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
        if thinking_budget is not None:
            payload["generationConfig"]["thinkingConfig"] = {
                "thinkingBudget": int(thinking_budget)
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
            # Some Gemini models (e.g., gemini-2.5-pro) reject thinkingBudget=0 and require thinking mode.
            # Gracefully retry once without thinkingConfig for this specific error.
            if (
                error.code == 400
                and thinking_budget is not None
                and "budget 0 is invalid" in body.lower()
            ):
                payload_no_thinking = {
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {
                        "temperature": temperature,
                        "maxOutputTokens": max_tokens,
                    },
                }
                retry_request = urllib.request.Request(
                    endpoint,
                    data=json.dumps(payload_no_thinking).encode("utf-8"),
                    headers={"Content-Type": "application/json"},
                    method="POST",
                )
                print(
                    f"⚠️ AI Call #{call_id}: model rejected thinkingBudget={thinking_budget}; retrying without thinkingConfig",
                    flush=True,
                )
                try:
                    with urllib.request.urlopen(retry_request, timeout=120) as response:
                        response_body = json.loads(response.read().decode("utf-8"))
                except urllib.error.HTTPError as retry_error:
                    retry_body = retry_error.read().decode("utf-8", errors="ignore")
                    duration = time.time() - started_at
                    error_message = f"Gemini API error {retry_error.code}: {retry_body}"
                    print(f"❌ AI Call #{call_id} failed in {duration:.2f}s: {error_message}", flush=True)
                    self._append_log(
                        {
                            "timestamp": datetime.utcnow().isoformat() + "Z",
                            "event": "error",
                            "call_id": call_id,
                            "provider": self.provider,
                            "model": self._current_model(),
                            "duration_seconds": round(duration, 3),
                            "error": error_message,
                        }
                    )
                    raise RuntimeError(error_message) from retry_error
                except Exception as retry_error:
                    duration = time.time() - started_at
                    error_message = str(retry_error)
                    print(f"❌ AI Call #{call_id} failed in {duration:.2f}s: {error_message}", flush=True)
                    self._append_log(
                        {
                            "timestamp": datetime.utcnow().isoformat() + "Z",
                            "event": "error",
                            "call_id": call_id,
                            "provider": self.provider,
                            "model": self._current_model(),
                            "duration_seconds": round(duration, 3),
                            "error": error_message,
                        }
                    )
                    raise
            elif error.code == 404 and self.gemini_model != "gemini-2.0-flash":
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
                    duration = time.time() - started_at
                    error_message = f"Gemini API error {error.code}: {body}"
                    print(f"❌ AI Call #{call_id} failed in {duration:.2f}s: {error_message}", flush=True)
                    self._append_log(
                        {
                            "timestamp": datetime.utcnow().isoformat() + "Z",
                            "event": "error",
                            "call_id": call_id,
                            "provider": self.provider,
                            "model": self._current_model(),
                            "duration_seconds": round(duration, 3),
                            "error": error_message,
                        }
                    )
                    raise RuntimeError(error_message) from error
            else:
                duration = time.time() - started_at
                error_message = f"Gemini API error {error.code}: {body}"
                print(f"❌ AI Call #{call_id} failed in {duration:.2f}s: {error_message}", flush=True)
                self._append_log(
                    {
                        "timestamp": datetime.utcnow().isoformat() + "Z",
                        "event": "error",
                        "call_id": call_id,
                        "provider": self.provider,
                        "model": self._current_model(),
                        "duration_seconds": round(duration, 3),
                        "error": error_message,
                    }
                )
                raise RuntimeError(error_message) from error
        except Exception as error:
            duration = time.time() - started_at
            error_message = str(error)
            print(f"❌ AI Call #{call_id} failed in {duration:.2f}s: {error_message}", flush=True)
            self._append_log(
                {
                    "timestamp": datetime.utcnow().isoformat() + "Z",
                    "event": "error",
                    "call_id": call_id,
                    "provider": self.provider,
                    "model": self._current_model(),
                    "duration_seconds": round(duration, 3),
                    "error": error_message,
                }
            )
            raise

        candidates = response_body.get("candidates", [])
        if not candidates:
            raise RuntimeError(f"Gemini response missing candidates: {response_body}")

        parts = candidates[0].get("content", {}).get("parts", [])
        text = "\n".join(part.get("text", "") for part in parts if part.get("text"))
        if not text:
            duration = time.time() - started_at
            error_message = f"Gemini response missing text parts (possible max_tokens or thinking budget issue): {response_body}"
            print(f"❌ AI Call #{call_id} failed in {duration:.2f}s: {error_message}", flush=True)
            self._append_log(
                {
                    "timestamp": datetime.utcnow().isoformat() + "Z",
                    "event": "error",
                    "call_id": call_id,
                    "provider": self.provider,
                    "model": self._current_model(),
                    "duration_seconds": round(duration, 3),
                    "error": error_message,
                }
            )
            raise RuntimeError(error_message)

        output_text = text.strip()
        duration = time.time() - started_at
        print(f"✅ AI Call #{call_id} completed in {duration:.2f}s", flush=True)
        self._append_log(
            {
                "timestamp": datetime.utcnow().isoformat() + "Z",
                "event": "success",
                "call_id": call_id,
                "provider": self.provider,
                "model": self._current_model(),
                "duration_seconds": round(duration, 3),
                "response_chars": len(output_text),
            }
        )

        return output_text

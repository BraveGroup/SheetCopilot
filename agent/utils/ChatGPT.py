"""
Async chat-completion wrapper built on the **modern OpenAI Python SDK** (>=1.40,
tested on 2.x).

Compared with the original (which mixed ``openai.api_base`` 0.x globals with a
per-request client), this version:

* uses ``openai.AsyncOpenAI`` clients created **once** and reused, with the
  ``base_url`` / ``timeout`` / ``max_retries`` set on the client;
* works with **any OpenAI-compatible endpoint** (OpenAI, Azure-style gateways,
  vLLM / TGI / Ollama / LM Studio, etc.) -- just point ``base_url`` at your server
  and set ``api_key`` (use a dummy like ``"EMPTY"`` for local servers);
* accepts both the new config keys (``base_url``, ``api_key`` / ``api_keys``,
  ``max_tokens``) and the legacy ones (``api_base``, ``max_new_tokens``) for
  backward compatibility;
* records every query + response into a :class:`~utils.trajectory.TrajectoryLogger`
  when one is attached, capturing token usage and latency.

The public surface used by the agent is unchanged: ``await chatbot(prompt)``
returns the response text, ``chatbot.context`` is the running message list, and
``chatbot.reset_query_count()`` still exists.
"""

import asyncio
import itertools
import time

import openai
from openai import AsyncOpenAI

default_config = {
    "base_url": "https://api.openai.com/v1",
    "model_name": "gpt-4o-mini",
    "api_keys": [""],
    "max_retries": 3,
    "timeout": 60,
    "temperature": 0.0,
    "prompt_format": "gpt-chat-prompt",
    "max_total_tokens": 16384,
}


def _normalize_config(config):
    """Map new/legacy config keys to a single internal shape."""
    base_url = config.get("base_url") or config.get("api_base") or "https://api.openai.com/v1"

    keys = config.get("api_keys")
    if not keys:
        single = config.get("api_key")
        keys = [single] if single else []
    # Drop empty strings; if nothing is left use a placeholder so local,
    # auth-less OpenAI-compatible servers (vLLM/Ollama/...) still work.
    keys = [k for k in keys if k] or ["EMPTY"]

    max_tokens = config.get("max_tokens", config.get("max_new_tokens"))

    return {
        "base_url": base_url,
        "api_keys": keys,
        "model_name": config.get("model_name", default_config["model_name"]),
        "prompt_format": config.get("prompt_format", default_config["prompt_format"]),
        "max_retries": config.get("max_retries", default_config["max_retries"]),
        "timeout": config.get("timeout", default_config["timeout"]),
        "temperature": config.get("temperature", default_config["temperature"]),
        "top_p": config.get("top_p"),
        "max_tokens": max_tokens,
        "max_total_tokens": config.get("max_total_tokens", default_config["max_total_tokens"]),
        "sleep_time": config.get("sleep_time", 5),
        # Streaming is recommended for reasoning models (whose long hidden
        # reasoning can make a non-streamed response appear to hang); it also
        # lets us capture `reasoning_content` for the trajectory.
        "stream": config.get("stream", False),
    }


class ChatGPT:
    def __init__(self, config=default_config, context=None, interaction_mode=False, logger=None):
        cfg = _normalize_config(config)

        self.base_url = cfg["base_url"]
        self.model_name = cfg["model_name"]
        self.prompt_format = cfg["prompt_format"]
        self.max_retries = cfg["max_retries"]
        self.timeout = cfg["timeout"]
        self.temperature = cfg["temperature"]
        self.top_p = cfg["top_p"]
        self.max_tokens = cfg["max_tokens"]
        self.max_total_tokens = cfg["max_total_tokens"]
        self.sleep_time = cfg["sleep_time"]
        self.stream = cfg["stream"]

        self._api_keys = list(cfg["api_keys"])
        # One reused AsyncOpenAI client per key; rotate clients across calls.
        self._clients = [
            AsyncOpenAI(api_key=k, base_url=self.base_url, timeout=self.timeout, max_retries=0)
            for k in self._api_keys
        ]
        self._client_cycle = itertools.cycle(self._clients)

        self.context = context if context is not None else []
        self.interaction_mode = interaction_mode
        self.logger = logger
        # The agent sets this before each call so the trajectory records which
        # planning stage issued the query.
        self.stage = "chat"
        self._query_count = 1

        masked = ["{}***{}".format(k[:3], k[-2:]) if len(k) > 6 else "***" for k in self._api_keys]
        print("=============== Initializing the LLM ===============")
        print("=   Model      : {}".format(self.model_name))
        print("=   Base URL   : {}".format(self.base_url))
        print("=   API Keys   : {}".format(masked))
        print("=   Max Retries: {}".format(self.max_retries))
        print("=   Timeout    : {}s".format(self.timeout))
        print("=   Max Tokens : {}".format(self.max_tokens))
        print("=   Temperature: {}".format(self.temperature))
        print("======================================================")

    def reset_query_count(self):
        self._query_count = 1

    async def __call__(self, prompt) -> str:
        self.context.append({"role": "user", "content": prompt})
        return await self.__get_response__()

    async def __get_response__(self) -> str:
        last_error = None
        for i in range(self.max_retries):
            try:
                result = await self.__request__()
            except openai.APITimeoutError as e:
                last_error = e
                print("API call timed out. Retrying {}/{}...".format(i + 1, self.max_retries))
            except openai.RateLimitError as e:
                last_error = e
                print("API rate limited. Retrying {}/{}...\n{}".format(i + 1, self.max_retries, e))
            except openai.APIConnectionError as e:
                last_error = e
                print("API connection error. Retrying {}/{}...".format(i + 1, self.max_retries))
            except openai.APIError as e:
                last_error = e
                print("API error. Retrying {}/{}...\n{}".format(i + 1, self.max_retries, e))
            else:
                self._query_count += 1
                self.context.append({"role": result["role"], "content": result["content"]})
                return result["content"]

            await asyncio.sleep(self.sleep_time)

        # Record the terminal failure in the trajectory too.
        if self.logger is not None:
            self.logger.record_call(
                stage=self.stage,
                model=self.model_name,
                request_messages=self.context,
                response_content=None,
                error="API call failed after {} retries: {}".format(self.max_retries, last_error),
            )
        raise Exception("API call failed after {} retries: {}".format(self.max_retries, last_error))

    @staticmethod
    def _usage_dict(usage):
        if usage is None:
            return None
        return {
            "prompt_tokens": usage.prompt_tokens,
            "completion_tokens": usage.completion_tokens,
            "total_tokens": usage.total_tokens,
        }

    async def __request__(self) -> dict:
        client = next(self._client_cycle)

        kwargs = dict(model=self.model_name, messages=self.context, temperature=self.temperature)
        if self.max_tokens:
            kwargs["max_tokens"] = self.max_tokens
        if self.top_p is not None:
            kwargs["top_p"] = self.top_p

        t0 = time.time()
        if self.stream:
            content, reasoning, role, finish_reason, usage = await self._stream_request(client, kwargs)
        else:
            content, reasoning, role, finish_reason, usage = await self._plain_request(client, kwargs)
        latency = time.time() - t0

        # Record the full query + response (incl. any reasoning) in the trajectory.
        if self.logger is not None:
            self.logger.record_call(
                stage=self.stage,
                model=self.model_name,
                request_messages=self.context,
                response_content=content,
                reasoning_content=reasoning,
                finish_reason=finish_reason,
                role=role,
                usage=usage,
                latency_s=latency,
            )

        if usage is not None and usage["total_tokens"] > self.max_total_tokens:
            raise Exception(
                "Generated tokens ({}) exceed the limit: {}!".format(
                    usage["total_tokens"], self.max_total_tokens
                )
            )

        return {"role": role, "content": content}

    async def _plain_request(self, client, kwargs):
        response = await client.chat.completions.create(**kwargs)
        choice = response.choices[0]
        content = choice.message.content
        reasoning = getattr(choice.message, "reasoning_content", None)
        role = choice.message.role or "assistant"
        finish_reason = getattr(choice, "finish_reason", None)
        return content, reasoning, role, finish_reason, self._usage_dict(getattr(response, "usage", None))

    async def _stream_request(self, client, kwargs):
        """Stream the completion, accumulating answer + reasoning deltas.

        Reasoning models emit (hidden) ``reasoning_content`` deltas before the
        visible ``content``; streaming surfaces tokens immediately instead of
        waiting for the whole -- often very long -- reasoning to finish."""
        kwargs = dict(kwargs)
        kwargs["stream"] = True
        kwargs["stream_options"] = {"include_usage": True}

        content_parts, reasoning_parts = [], []
        role, finish_reason, usage = "assistant", None, None

        stream = await client.chat.completions.create(**kwargs)
        async for chunk in stream:
            if getattr(chunk, "usage", None) is not None:
                usage = self._usage_dict(chunk.usage)
            if not chunk.choices:
                continue
            choice = chunk.choices[0]
            delta = choice.delta
            if getattr(delta, "role", None):
                role = delta.role
            if getattr(delta, "content", None):
                content_parts.append(delta.content)
            rc = getattr(delta, "reasoning_content", None)
            if rc:
                reasoning_parts.append(rc)
            if choice.finish_reason:
                finish_reason = choice.finish_reason

        content = "".join(content_parts)
        reasoning = "".join(reasoning_parts) or None
        return content, reasoning, role, finish_reason, usage


async def test():
    chatbot = ChatGPT()
    while True:
        prompt = input("You: ")
        response = await chatbot(prompt)
        print("Bot:", response)


if __name__ == "__main__":
    asyncio.run(test())

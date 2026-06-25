"""
Structured trajectory logging for SheetCopilot agent runs.

A *trajectory* is the full, detailed trace of one task attempt: every LLM query
(the complete list of messages sent), the model's response, token usage, latency,
which planning stage issued the call, and the actions parsed/executed from each
response.  One :class:`TrajectoryLogger` accumulates these records and serialises
them to a single, human-readable JSON file -- replacing the old, redundant
``context_log_*.yaml`` dumps with one elegant artifact per attempt.

The logger is transport-agnostic: :class:`utils.ChatGPT.ChatGPT` calls
:meth:`record_call` after each request, and the agent annotates the most recent
call with the actions it extracted.  Nothing here imports ``openai``, so it can be
unit-tested without any API access.
"""

import copy
import datetime
import json
import os
import threading
import time


def _now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _copy_messages(messages):
    """Defensive deep copy of an OpenAI-style ``messages`` list."""
    try:
        return copy.deepcopy(list(messages))
    except Exception:
        # Fall back to a shallow, str-coerced copy if something is unpicklable.
        return [{"role": m.get("role"), "content": str(m.get("content"))} for m in messages]


class TrajectoryLogger:
    """Collect a structured trace of one task attempt.

    Parameters
    ----------
    meta : dict, optional
        Task-level metadata (task id, sheet name, instruction, model, etc.).
        Extra metadata can also be supplied later via :meth:`save` / :meth:`to_dict`.
    """

    def __init__(self, meta=None):
        self.meta = dict(meta or {})
        self.calls = []
        self._lock = threading.Lock()
        self._t0 = time.time()
        self.meta.setdefault("start_time", _now())

    # ------------------------------------------------------------------ #
    def record_call(
        self,
        *,
        stage,
        model,
        request_messages,
        response_content,
        reasoning_content=None,
        finish_reason=None,
        role="assistant",
        usage=None,
        latency_s=None,
        error=None,
        extra=None,
    ):
        """Append one LLM interaction to the trajectory and return the record.

        ``request_messages`` is the *complete* list of messages sent to the model
        (the query); ``response_content`` is the model's reply.  ``usage`` is the
        token-usage dict (``prompt_tokens`` / ``completion_tokens`` /
        ``total_tokens``) when available.
        """
        with self._lock:
            call = {
                "id": len(self.calls) + 1,
                "stage": stage,
                "timestamp": _now(),
                "latency_s": round(latency_s, 3) if latency_s is not None else None,
                "model": model,
                "request_messages": _copy_messages(request_messages),
                "response": {
                    "role": role,
                    "content": response_content,
                    "reasoning_content": reasoning_content,
                    "finish_reason": finish_reason,
                },
                "usage": dict(usage) if usage else None,
                "parsed_actions": None,
                "error": error,
            }
            if extra:
                call.update(extra)
            self.calls.append(call)
            return call

    def annotate_last(self, **kwargs):
        """Attach extra fields (e.g. ``parsed_actions``) to the most recent call."""
        with self._lock:
            if self.calls:
                self.calls[-1].update(kwargs)

    # ------------------------------------------------------------------ #
    def usage_summary(self):
        prompt = completion = total = 0
        for c in self.calls:
            u = c.get("usage") or {}
            prompt += u.get("prompt_tokens", 0) or 0
            completion += u.get("completion_tokens", 0) or 0
            total += u.get("total_tokens", 0) or 0
        return {
            "total_calls": len(self.calls),
            "prompt_tokens": prompt,
            "completion_tokens": completion,
            "total_tokens": total,
            "wall_time_s": round(time.time() - self._t0, 3),
        }

    def to_dict(self, **extra_meta):
        meta = dict(self.meta)
        meta.update(extra_meta)
        meta.setdefault("end_time", _now())
        meta["usage_summary"] = self.usage_summary()
        return {"meta": meta, "calls": self.calls}

    def save(self, path, **extra_meta):
        """Write the trajectory to ``path`` as pretty-printed JSON."""
        directory = os.path.dirname(os.path.abspath(path))
        os.makedirs(directory, exist_ok=True)
        data = self.to_dict(**extra_meta)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return path

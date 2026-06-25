"""Centralised logging for the Ubuntu evaluator.

Workers run in separate processes, so we use the standard ``QueueHandler`` /
``QueueListener`` pattern: every process pushes records onto a shared queue and a
single listener thread in the main process writes them to the console and to a
log file.  This keeps multi-process logs interleaved cleanly and free of torn
lines."""

import logging
import sys
from logging.handlers import QueueHandler, QueueListener

_FMT = "%(asctime)s | %(processName)-14s | %(levelname)-7s | %(message)s"
_DATEFMT = "%H:%M:%S"


def start_listener(log_path, level=logging.INFO):
    """Start a queue listener writing to ``log_path`` and stderr.

    Returns ``(queue, listener)``.  Call ``listener.stop()`` when done.
    """
    import multiprocessing

    queue = multiprocessing.Manager().Queue(-1)

    fmt = logging.Formatter(_FMT, _DATEFMT)
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(fmt)
    stream_handler = logging.StreamHandler(sys.stderr)
    stream_handler.setFormatter(fmt)

    listener = QueueListener(queue, file_handler, stream_handler, respect_handler_level=True)
    listener.start()
    return queue, listener


def configure_process(queue, level=logging.INFO):
    """Route the calling process's root logger through the shared queue."""
    root = logging.getLogger()
    root.handlers[:] = [QueueHandler(queue)]
    root.setLevel(level)

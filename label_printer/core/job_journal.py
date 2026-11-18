"""Durable, append-only batch journal. A failed write stops further submissions."""
from __future__ import annotations

import copy
import json
import os
import threading
import uuid
from datetime import datetime
from pathlib import Path


class JournalError(RuntimeError):
    """A task cannot be safely continued without its durable audit trail."""


class JobJournal:
    def __init__(self, directory: Path, snapshot: dict, batch_id: str | None = None):
        self.batch_id = batch_id or uuid.uuid4().hex
        self.snapshot = copy.deepcopy(snapshot)
        self.path = Path(directory) / f"{self.batch_id}.jsonl"
        self.summary_path = self.path.with_suffix(".json")
        self._lock = threading.Lock()
        self._sequence = 0
        self._file = None
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self._file = self.path.open("x", encoding="utf-8", newline="\n")
            self.append("batch_start", snapshot=self.snapshot)
        except (OSError, ValueError, TypeError, JournalError) as exc:
            self.close()
            raise JournalError(f"无法建立打印任务日志，已停止提交：{exc}") from exc

    def append(self, event: str, *, durable: bool = True, **data) -> None:
        with self._lock:
            if self._file is None:
                raise JournalError("任务日志已经关闭，已停止提交")
            record = {
                "schema": 1,
                "batch_id": self.batch_id,
                "sequence": self._sequence + 1,
                "time": datetime.now().astimezone().isoformat(timespec="milliseconds"),
                "event": event,
                **copy.deepcopy(data),
            }
            try:
                self._file.write(json.dumps(record, ensure_ascii=False) + "\n")
                self._file.flush()
                if durable:
                    os.fsync(self._file.fileno())
                self._sequence += 1
            except (OSError, ValueError, TypeError) as exc:
                raise JournalError(f"任务日志写入失败，已停止后续提交：{exc}") from exc

    def save_summary(self, success: bool, message: str, statuses: list[dict]) -> None:
        record = {
            "schema": 1, "batch_id": self.batch_id,
            "time": datetime.now().astimezone().isoformat(timespec="milliseconds"),
            "success": success, "message": message,
            **copy.deepcopy(self.snapshot),
            "rows": copy.deepcopy(statuses), "journal": str(self.path),
        }
        temporary = self.summary_path.with_suffix(".json.tmp")
        try:
            with temporary.open("x", encoding="utf-8") as stream:
                json.dump(record, stream, ensure_ascii=False, indent=2)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.summary_path)
        except (OSError, ValueError, TypeError) as exc:
            raise JournalError(f"任务汇总写入失败，请按增量日志核对：{exc}") from exc

    def close(self) -> None:
        with self._lock:
            stream, self._file = self._file, None
            if stream is not None:
                try:
                    stream.close()
                except OSError:
                    pass

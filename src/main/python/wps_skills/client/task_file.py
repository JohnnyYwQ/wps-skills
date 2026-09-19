"""An explicitly handed-over input file; durable Task receipts own the evidence."""

import os
from pathlib import Path
import stat


def _identity(value):
    return value.st_dev, value.st_ino, value.st_size, value.st_mtime_ns


class ConsumableTaskFile:
    def __init__(self, path):
        self.path = Path(path).absolute()
        self.report = {"path": str(self.path), "state": "retained", "reason": "not_admitted"}
        self._identity = None
        self._raw = None

    def read(self, limit):
        if not stat.S_ISREG(self.path.lstat().st_mode):
            raise ValueError("Consumable Task input must be a regular file, not a symbolic link")
        with self.path.open("rb") as stream:
            self._identity = _identity(os.fstat(stream.fileno()))
            self._raw = stream.read(limit + 1)
        if _identity(self.path.lstat()) != self._identity:
            raise ValueError("Task input changed while it was being read")
        if len(self._raw) > limit:
            raise ValueError("Task Request exceeds 8 MiB")
        return self._raw.decode("utf-8-sig")

    def admitted(self, receipt_directory):
        """Called only after the exact request and an initial receipt are on disk.

        Deletion failure retains the input; it cannot change a document outcome
        or block an already admitted request. The caller must not edit a file
        after handing it over to the Task Client.
        """
        try:
            source = self.path.resolve()
            receipt_root = receipt_directory.parent.parent.resolve()
            if source == receipt_root or receipt_root in source.parents:
                self.report["reason"] = "protected_receipt"
                return
            current = self.path.lstat()
            if not stat.S_ISREG(current.st_mode) or _identity(current) != self._identity:
                self.report["reason"] = "input_changed"
                return
            with self.path.open("rb") as stream:
                if (_identity(os.fstat(stream.fileno())) != self._identity
                        or stream.read(len(self._raw) + 1) != self._raw):
                    self.report["reason"] = "input_changed"
                    return
            self.path.unlink()
            self.report.update(state="removed", reason="durably_admitted")
        except FileNotFoundError:
            self.report.update(state="absent", reason="already_removed")
        except OSError as exc:
            self.report.update(reason="cleanup_failed", error=str(exc))

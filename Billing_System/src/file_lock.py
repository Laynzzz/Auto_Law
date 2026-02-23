"""Per-firm cross-process file locking for safe multi-PC shared access.

Uses OS-level exclusive locks: msvcrt.locking on Windows, fcntl.flock on Unix.
Lock file per firm: invoice/{FirmName}/master_{FirmName}.lock

Lock files are never deleted — only the OS lock is acquired/released.
If a process crashes, the OS automatically releases the lock when the fd closes.
"""

import json
import os
import platform
import time
from datetime import datetime
from pathlib import Path

from src.config import get_data_root


class FirmFileLock:
    """Context manager for per-firm exclusive file locking.

    Usage:
        with FirmFileLock("ABC Law"):
            # ... write to ABC Law's dataset ...

    One lock per firm — writing to ABC Law doesn't block bbc law.
    """

    def __init__(
        self,
        firm_name: str,
        timeout: float = 30.0,
        retry_interval: float = 2.0,
    ):
        self.firm_name = firm_name
        self.timeout = timeout
        self.retry_interval = retry_interval
        self._fd: int | None = None
        root = get_data_root()
        self._lock_path = root / "invoice" / firm_name / f"master_{firm_name}.lock"

    def _lock_info(self) -> bytes:
        """JSON payload identifying the current lock holder."""
        info = {
            "user": os.environ.get("USERNAME") or os.environ.get("USER", "unknown"),
            "hostname": platform.node(),
            "timestamp": datetime.now().isoformat(),
            "pid": os.getpid(),
        }
        return json.dumps(info, indent=2).encode("utf-8")

    def _read_holder_info(self) -> str:
        """Best-effort read of current lock holder info for error messages."""
        try:
            return self._lock_path.read_text(encoding="utf-8")
        except (OSError, ValueError):
            return "(unknown holder)"

    def __enter__(self):
        self._lock_path.parent.mkdir(parents=True, exist_ok=True)
        deadline = time.monotonic() + self.timeout

        while True:
            fd = None
            try:
                fd = os.open(str(self._lock_path), os.O_RDWR | os.O_CREAT)

                if os.name == "nt":
                    import msvcrt
                    os.lseek(fd, 0, os.SEEK_SET)
                    msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)

                # Lock acquired — write holder info
                os.ftruncate(fd, 0)
                os.lseek(fd, 0, os.SEEK_SET)
                os.write(fd, self._lock_info())
                self._fd = fd
                return self

            except (OSError, IOError):
                if fd is not None:
                    os.close(fd)

                if time.monotonic() >= deadline:
                    holder = self._read_holder_info()
                    raise TimeoutError(
                        f"Could not acquire lock for '{self.firm_name}' "
                        f"after {self.timeout}s.\nCurrent holder: {holder}"
                    )

                time.sleep(self.retry_interval)

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self._fd is not None:
                if os.name == "nt":
                    import msvcrt
                    try:
                        os.lseek(self._fd, 0, os.SEEK_SET)
                        msvcrt.locking(self._fd, msvcrt.LK_UNLCK, 1)
                    except OSError:
                        pass
                else:
                    import fcntl
                    try:
                        fcntl.flock(self._fd, fcntl.LOCK_UN)
                    except OSError:
                        pass
        finally:
            if self._fd is not None:
                os.close(self._fd)
                self._fd = None
        return False

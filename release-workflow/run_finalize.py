"""Launch finalize_outputs.py with the platform's release-workflow Python."""

from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys


FINALIZER = Path(__file__).with_name("finalize_outputs.py")


def command(args: list[str], platform: str, executable: str) -> list[str]:
    interpreter = ["py", "-3.14"] if platform == "nt" else [executable]
    return [*interpreter, str(FINALIZER), *args]


def main() -> int:
    completed = subprocess.run(command(sys.argv[1:], os.name, sys.executable), check=False)
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())

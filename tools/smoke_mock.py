"""Run the synthetic, offscreen UI smoke test without customer files or printing."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    output = root / ".tmp" / "ui-self-test"
    return subprocess.run(
        [sys.executable, "-m", "label_printer", "--ui-self-test", str(output)],
        cwd=root,
        check=False,
    ).returncode


if __name__ == "__main__":
    raise SystemExit(main())

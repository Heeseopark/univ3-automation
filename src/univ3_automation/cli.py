from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> None:
    script = Path(__file__).resolve().parent / "commands" / "admin_cli.py"
    subprocess.run([sys.executable, str(script)], check=False)


if __name__ == "__main__":
    main()

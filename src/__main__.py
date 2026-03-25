"""Allow `python -m src` to run the dashboard pipeline."""

from __future__ import annotations

import sys
from pathlib import Path

package_dir = Path(__file__).resolve().parent
if str(package_dir) not in sys.path:
    sys.path.insert(0, str(package_dir))

from main import run_dashboard


if __name__ == "__main__":
    run_dashboard()

"""Executable entrypoint for dashboard generation."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from process import process_current_files
from update import update_sheet


def run_dashboard() -> Path:
    """Run the dashboard pipeline and return the exported workbook path."""

    output_dir = Path("data") / "dashboards"
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y.%m.%d-%H.%M.%S")
    output_path = output_dir / f"dashboard{timestamp}.xlsx"

    debug = None
    update = False
    count_delta = False

    current_data, history_data = process_current_files(debug)
    current_data.to_excel(output_path, index=False)
    if update:
        update_sheet(current_data, count_delta, history_data)

    return output_path


if __name__ == "__main__":
    run_dashboard()

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from career_report_app import CareerReportHandler  # noqa: E402


class handler(CareerReportHandler):
    pass

"""
Main entry point for PlacementMonitor.
"""
from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from src.monitor import run_monitor


def main():
    run_monitor()


if __name__ == "__main__":
    main()

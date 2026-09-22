"""
Main entry point for PlacementMonitor.
"""
from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from src.monitor import main as monitor_main


def main():
    monitor_main(sys.argv[1:])


if __name__ == "__main__":
    main()

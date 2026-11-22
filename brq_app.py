"""Root launcher retained for existing command-line workflows."""
from pathlib import Path
import sys
APP_DIR = Path(__file__).resolve().parent / "brq app"
sys.path.insert(0, str(APP_DIR))
from main import main
if __name__ == "__main__":
    main()

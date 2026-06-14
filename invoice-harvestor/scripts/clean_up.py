from pathlib import Path
import shutil
from dataclasses import dataclass

SCRIPT_DIR = Path(__file__).resolve().parent  # Points to 'scripts/'
SKILL_ROOT = SCRIPT_DIR.parent
ASSETS_DIR = SKILL_ROOT / "assets"
TEMP_DIR = ASSETS_DIR / "temp"


@dataclass
class CleanupResult:
    """Result of cleaning up temp files."""

    success: bool


def cleanup_temp_files() -> CleanupResult:
    try:
        if TEMP_DIR.exists():
            # Deletes the temp folder and every single file inside it
            shutil.rmtree(TEMP_DIR)
            print("🧹 Cleanup complete! All temporary data files have been removed.")
            return CleanupResult(success=True)
    except Exception as e:
        print(f"Cleanup failed, error: {e}")
        return CleanupResult(success=False)

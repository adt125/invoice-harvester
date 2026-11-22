import shutil
from dataclasses import dataclass

from config import TEMP_DIR


@dataclass
class CleanupResult:
    """Result of cleaning up temp files."""

    success: bool


def cleanup_temp_files() -> CleanupResult:
    try:
        if TEMP_DIR.exists():
            # Deletes the temp folder and every single file inside it
            shutil.rmtree(TEMP_DIR)
            print("Cleanup complete! All temporary data files have been removed.")
            return CleanupResult(success=True)
        return CleanupResult(success=True)
    except Exception as e:
        print(f"Cleanup failed, error: {e}")
        return CleanupResult(success=False)

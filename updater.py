import sys
import time
import shutil
from pathlib import Path
import subprocess


def _usage() -> str:
    return (
        "Usage:\n"
        "  updater.exe <current_exe_path> <downloaded_new_exe_path>\n\n"
        "Example:\n"
        "  updater.exe C:\\Apps\\VaM.Simple.Var.Manager.exe C:\\Temp\\VSVM_1.0.1.exe\n"
    )


def _wait_for_process_release(path: Path, timeout_s: float = 25.0) -> bool:
    """
    Wait until the target exe is no longer locked by the running app.
    We'll keep trying to rename/replace; once it succeeds, lock is gone.
    """
    start = time.time()
    while time.time() - start < timeout_s:
        try:
            # Try opening for append; if locked, this often fails on Windows
            with open(path, "ab"):
                return True
        except Exception:
            time.sleep(0.25)
    return False


def _safe_rename(src: Path, dst: Path) -> None:
    if dst.exists():
        try:
            dst.unlink()
        except Exception:
            pass
    src.rename(dst)


def main() -> int:
    if len(sys.argv) < 3:
        print(_usage())
        return 2

    current_exe = Path(sys.argv[1]).resolve()
    new_exe = Path(sys.argv[2]).resolve()

    if not current_exe.exists():
        print(f"[ERROR] Current exe not found: {current_exe}")
        return 3
    if not new_exe.exists():
        print(f"[ERROR] New exe not found: {new_exe}")
        return 4

    # Wait until the main app exits (releases file lock)
    if not _wait_for_process_release(current_exe, timeout_s=30.0):
        print("[ERROR] Timed out waiting for main app to exit.")
        return 5

    folder = current_exe.parent
    backup = folder / (current_exe.name + ".bak")

    # Replace current exe with the new one
    try:
        # Backup current exe first
        try:
            if backup.exists():
                backup.unlink()
        except Exception:
            pass

        # Sometimes rename can fail briefly; retry a few times
        for _ in range(40):
            try:
                _safe_rename(current_exe, backup)
                break
            except Exception:
                time.sleep(0.25)
        else:
            print("[ERROR] Failed to backup old exe (still locked?).")
            return 6

        # Copy new exe into place
        for _ in range(40):
            try:
                shutil.copy2(new_exe, current_exe)
                break
            except Exception:
                time.sleep(0.25)
        else:
            print("[ERROR] Failed to copy new exe into place.")
            # attempt rollback
            try:
                if current_exe.exists():
                    current_exe.unlink()
            except Exception:
                pass
            try:
                if backup.exists():
                    _safe_rename(backup, current_exe)
            except Exception:
                pass
            return 7

        # Cleanup: keep backup or delete it (your choice)
        # I'd keep it so user can rollback manually if needed.
        # Try removing downloaded new exe to avoid temp clutter:
        try:
            new_exe.unlink()
        except Exception:
            pass

    except Exception as e:
        print(f"[ERROR] Update failed: {e}")
        return 8

    # Relaunch updated app
    try:
        subprocess.Popen([str(current_exe)], cwd=str(current_exe.parent))
    except Exception as e:
        print(f"[WARN] Updated but failed to relaunch: {e}")
        return 9

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

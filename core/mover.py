# core/mover.py
import shutil
from pathlib import Path

def move_unused_vars(addon_dir: Path, used_var_names: set):
    target = addon_dir / "delete candidate"
    target.mkdir(exist_ok=True)

    moved = []

    for var_file in addon_dir.glob("*.var"):
        if var_file.name not in used_var_names:
            shutil.move(str(var_file), target / var_file.name)
            moved.append(var_file.name)

    return moved

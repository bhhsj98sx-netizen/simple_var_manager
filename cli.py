from pathlib import Path
from core.resolver import collect_used_and_unused_vars
import sys
import shutil

# 1️⃣ Read AddonPackages path from CLI
addon_dir = Path(sys.argv[1])

# 2️⃣ Analyze
scene_vars, used_vars, unused_vars = collect_used_and_unused_vars(addon_dir)

print(f"Scenes: {len(scene_vars)}")
print(f"Used VARs: {len(used_vars)}")
print(f"Unused VARs: {len(unused_vars)}")

# 3️⃣ Prepare delete-candidate folder
delete_dir = addon_dir / "delete candidate"
delete_dir.mkdir(exist_ok=True)

# 4️⃣ Move unused VARs
moved = 0
for var_name in sorted(unused_vars):
    src = addon_dir / var_name
    dst = delete_dir / var_name

    if src.exists():
        print(f"Moving {moved+1}/{len(unused_vars)}: {var_name}")
        shutil.move(str(src), str(dst))
        moved += 1

print(f"\nDone. Moved {moved} unused VARs to '{delete_dir}'")

# core/scanner.py
import zipfile
import json
from pathlib import Path


def scan_var(var_path: Path):
    result = {
        "has_scene": False,
        "dependencies": set(),
        "scenes": []   # 👈 NEW
    }

    try:
        with zipfile.ZipFile(var_path, 'r') as z:
            names = z.namelist()

            # Detect scenes by folder
            scene_jsons = [
                n for n in names
                if (
                    (n.startswith("Saves/scene/") or n.startswith("Saves\\scene\\"))
                    and n.lower().endswith(".json")
                )
            ]

            if scene_jsons:
                result["has_scene"] = True

                for scene_json in scene_jsons:
                    base = scene_json.rsplit(".", 1)[0]

                    # Try jpg / png preview
                    image_name = None
                    for ext in (".jpg", ".png"):
                        candidate = base + ext
                        if candidate in names:
                            image_name = candidate
                            break

                    image_bytes = None
                    if image_name:
                        with z.open(image_name) as img:
                            image_bytes = img.read()

                    result["scenes"].append({
                        "scene_name": Path(base).name,
                        "image_bytes": image_bytes,
                        "var_name": var_path.name
                    })

            # Read meta.json
            if "meta.json" in names:
                try:
                    with z.open("meta.json") as f:
                        meta = json.load(f)

                    result["dependencies"] = set(
                        meta.get("dependencies", {}).keys()
                    )

                except (json.JSONDecodeError, UnicodeDecodeError):
                    pass

    except zipfile.BadZipFile:
        pass

    return result

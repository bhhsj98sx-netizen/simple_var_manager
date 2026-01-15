# core/scanner.py
from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any

DISABLED_SUFFIX = ".disabled"


def _open_existing_var(path: Path) -> Path:
    """
    Support both .var and .var.disabled
    """
    if path.exists():
        return path
    alt = Path(str(path) + DISABLED_SUFFIX)
    if alt.exists():
        return alt
    return path  # will fail later if truly missing


def _read_meta_json(z: zipfile.ZipFile) -> dict:
    # meta.json is always at root in VaM var
    raw = z.read("meta.json")
    try:
        return json.loads(raw.decode("utf-8", errors="ignore"))
    except Exception:
        return {}


def _extract_dependencies(meta: dict) -> list[str]:
    deps = meta.get("dependencies", [])
    if isinstance(deps, dict):
        out = []
        for k in deps.keys():
            if isinstance(k, (str, int, float)):
                out.append(str(k))
        return out
    if isinstance(deps, list):
        return [str(x) for x in deps if isinstance(x, (str, int, float))]
    if isinstance(deps, (str, int, float)):
        return [str(deps)]
    return []


def _extract_content_list(meta: dict) -> list[str]:
    """
    Tries multiple keys because some packages/tools may not use 'contentList'
    consistently, or may store it in other structures.
    Returns normalized paths with forward slashes.
    """
    # 1) common keys
    cl = meta.get("contentList")
    if cl is None:
        cl = meta.get("content_list")
    if cl is None:
        cl = meta.get("content")
    if cl is None:
        cl = meta.get("files")
    if cl is None:
        cl = meta.get("fileList")

    out: list[str] = []

    # 2) if it's a list of strings
    if isinstance(cl, list):
        for x in cl:
            if isinstance(x, str) and x.strip():
                out.append(x.replace("\\", "/").strip())
            elif isinstance(x, dict):
                # sometimes list is objects: {"path": "..."} or {"name": "..."}
                p = x.get("path") or x.get("name") or x.get("file") or ""
                if isinstance(p, str) and p.strip():
                    out.append(p.replace("\\", "/").strip())
        return out

    # 3) if it's a dict (rare)
    if isinstance(cl, dict):
        # sometimes: {"items":[...]} or {"contentList":[...]}
        for k in ("items", "list", "contentList", "files"):
            v = cl.get(k)
            if isinstance(v, list):
                for x in v:
                    if isinstance(x, str) and x.strip():
                        out.append(x.replace("\\", "/").strip())
                    elif isinstance(x, dict):
                        p = x.get("path") or x.get("name") or x.get("file") or ""
                        if isinstance(p, str) and p.strip():
                            out.append(p.replace("\\", "/").strip())
                if out:
                    return out

    return []



def _hidden_scene_names(paths: list[str]) -> set[str]:
    hidden = set()
    for p in paths:
        lp = p.replace("\\", "/").lower()
        if not lp.startswith("saves/scene/"):
            continue
        if lp.endswith(".json.hide"):
            hidden.add(Path(lp[:-5]).stem)  # strip ".hide"
        elif lp.endswith(".hide"):
            hidden.add(Path(lp[:-5]).stem)
    return hidden

def _scene_names_from_paths(paths: list[str], include_hidden: bool) -> list[str]:
    hidden = set() if include_hidden else _hidden_scene_names(paths)
    scenes = []
    for p in paths:
        lp = p.replace("\\", "/").lower()
        if not lp.startswith("saves/scene/"):
            continue
        if not lp.endswith(".json"):
            continue
        # take filename stem
        name = Path(lp).stem
        # skip default.json
        if name.lower() == "default":
            continue
        if name in hidden:
            continue
        scenes.append(name)
    # unique + stable sort
    return sorted(set(scenes), key=lambda s: s.lower())

def _has_scene_json(paths: list[str]) -> bool:
    for p in paths:
        lp = p.replace("\\", "/").lower()
        if lp.startswith("saves/scene/") and lp.endswith(".json"):
            return True
    return False


def _preview_path_for_scene(content_list: list[str], scene_name: str) -> str:
    """
    Find best preview image path in Saves/scene/ that matches scene_name.
    Common: Saves/scene/<scene>.png
    """
    scene_lower = scene_name.lower()
    best = ""
    for p in content_list:
        lp = p.replace("\\", "/").lower()
        if not lp.startswith("saves/scene/"):
            continue
        if not (lp.endswith(".png") or lp.endswith(".jpg") or lp.endswith(".jpeg")):
            continue
        if Path(lp).stem != scene_lower:
            continue
        # prefer png if multiple
        if not best:
            best = p
        else:
            if best.lower().endswith(".jpg") and lp.endswith(".png"):
                best = p
    return best


def _scene_json_path_for_scene(content_list: list[str], scene_name: str) -> str:
    """
    Find the matching Saves/scene/<scene>.json path for scene_name.
    """
    scene_lower = scene_name.lower()
    for p in content_list:
        if not isinstance(p, str):
            continue
        lp = p.lower().replace("\\", "/")
        if not lp.startswith("saves/scene/"):
            continue
        if not lp.endswith(".json"):
            continue
        if Path(lp).stem != scene_lower:
            continue
        return p
    return ""


def scan_var_meta_only(var_path: Path, include_hidden: bool = True) -> dict[str, Any]:
    """
    FAST: reads meta.json and zip namelist (no scene json read).
    Returns:
      {
        "dependencies": [...],
        "scenes": [ {"scene_name": "...", "preview_path": "Saves/scene/..png" or "", "scene_path": "Saves/scene/..json" or ""}, ... ],
        "creator": "...", (optional)
        "package_name": "...", (optional)
      }
    """
    var_path = _open_existing_var(var_path)

    out: dict[str, Any] = {"dependencies": [], "scenes": []}

    try:
        with zipfile.ZipFile(var_path, "r") as z:
            meta = _read_meta_json(z)
            content_list = _extract_content_list(meta)
            names = [] if _has_scene_json(content_list) else z.namelist()

        deps = _extract_dependencies(meta)

        # Prefer contentList, but include zip entries that are missing from it.
        paths = []
        seen = set()
        if _has_scene_json(content_list):
            for p in content_list:
                if not isinstance(p, str):
                    continue
                if p in seen:
                    continue
                seen.add(p)
                paths.append(p)
            for p in names:
                if not isinstance(p, str):
                    continue
                if p in seen:
                    continue
                seen.add(p)
                paths.append(p)
        else:
            for p in content_list + names:
                if not isinstance(p, str):
                    continue
                if p in seen:
                    continue
                seen.add(p)
                paths.append(p)

        scenes = []
        for s in _scene_names_from_paths(paths, include_hidden):
            scenes.append({
                "scene_name": s,
                "preview_path": _preview_path_for_scene(paths, s) or "",
                "scene_path": _scene_json_path_for_scene(paths, s) or "",
            })

        out["dependencies"] = deps
        out["scenes"] = scenes

        # optional metadata (nice for search later)
        # meta schema varies; keep safe
        out["creator"] = str(meta.get("creator") or meta.get("author") or "")
        out["package_name"] = str(meta.get("packageName") or meta.get("name") or "")

    except Exception:
        # corrupted zip / missing meta.json etc.
        pass

    return out


def scan_var_meta_with_previews(var_path: Path, include_hidden: bool = True) -> dict[str, Any]:
    """
    Reads meta.json AND preview image bytes in one zip open.
    Returns:
      {
        "dependencies": [...],
        "scenes": [
            {"scene_name": "...", "preview_path": "Saves/scene/..jpg", "scene_path": "Saves/scene/..json", "preview_bytes": b"..."},
            ...
        ],
        "creator": "...",
        "package_name": "..."
      }
    """
    var_path = _open_existing_var(var_path)

    out: dict[str, Any] = {"dependencies": [], "scenes": []}

    try:
        with zipfile.ZipFile(var_path, "r") as z:
            meta = _read_meta_json(z)
            content_list = _extract_content_list(meta)
            names = [] if _has_scene_json(content_list) else z.namelist()

            deps = _extract_dependencies(meta)

            paths = []
            seen = set()
            if _has_scene_json(content_list):
                for p in content_list:
                    if not isinstance(p, str):
                        continue
                    if p in seen:
                        continue
                    seen.add(p)
                    paths.append(p)
                for p in names:
                    if not isinstance(p, str):
                        continue
                    if p in seen:
                        continue
                    seen.add(p)
                    paths.append(p)
            else:
                for p in content_list + names:
                    if not isinstance(p, str):
                        continue
                    if p in seen:
                        continue
                    seen.add(p)
                    paths.append(p)

            scenes = []
            for s in _scene_names_from_paths(paths, include_hidden):
                preview_path = _preview_path_for_scene(paths, s) or ""
                scene_path = _scene_json_path_for_scene(paths, s) or ""
                preview_bytes = None
                if preview_path:
                    try:
                        preview_bytes = z.read(preview_path)
                    except Exception:
                        preview_bytes = None
                scenes.append({
                    "scene_name": s,
                    "preview_path": preview_path,
                    "scene_path": scene_path,
                    "preview_bytes": preview_bytes,
                })

            out["dependencies"] = deps
            out["scenes"] = scenes
            out["creator"] = str(meta.get("creator") or meta.get("author") or "")
            out["package_name"] = str(meta.get("packageName") or meta.get("name") or "")

    except Exception:
        pass

    return out


def read_file_from_var(var_path: Path, inner_path: str) -> bytes | None:
    """
    On-demand read (for preview images, etc).
    """
    if not inner_path:
        return None
    var_path = _open_existing_var(var_path)
    try:
        with zipfile.ZipFile(var_path, "r") as z:
            return z.read(inner_path)
    except Exception:
        return None
    
    # --- BACKWARD COMPAT WRAPPER ---
def scan_var(var_path):
    """
    Backward-compatible alias for older modules that still import scan_var.
    Returns meta-only info now.
    """
    return scan_var_meta_only(var_path)


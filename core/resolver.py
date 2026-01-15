from pathlib import Path
from core.scanner import scan_var

def is_asset_var(var_name: str) -> bool:
    lname = var_name.lower()
    return (
        "[asset]" in lname or
        "[assets]" in lname or
        ".asset" in lname  # defensive
    )


def resolve_dependency(dep_name: str, all_vars: set[str]) -> list[str]:
    """
    Resolve a VaM dependency string into actual .var filenames.

    Rules:
    - "*.latest" → match any VAR starting with base + "."
    - exact version → match exact VAR only
    """
    matches = []

    if dep_name.endswith(".var"):
        exact = dep_name
        if exact in all_vars:
            return [exact]
        return []

    if dep_name.endswith(".latest"):
        base = dep_name[:-len(".latest")]
        matches = [
            v for v in all_vars
            if v.startswith(base + ".") and v.endswith(".var")
        ]
    else:
        exact = dep_name + ".var"
        if exact in all_vars:
            matches = [exact]

    return matches


def collect_used_and_unused_vars(var_dir: Path):
    """
    Main resolver:
    - Finds scene VARs
    - Resolves dependencies recursively (VaM-correct)
    - Returns used & unused VARs
    """
    all_vars = {v.name for v in var_dir.glob("*.var")}

    used_vars = set()
    scene_vars = set()
    queue = []

    # 1️⃣ Seed from scene-containing VARs
    for var_name in all_vars:
        info = scan_var(var_dir / var_name)
        if info.get("has_scene"):
            scene_vars.add(var_name)
            used_vars.add(var_name)
            queue.extend(info.get("dependencies", []))

    # 2️⃣ Recursive dependency resolution
    while queue:
        dep = queue.pop()

        for matched_var in resolve_dependency(dep, all_vars):
            if matched_var not in used_vars:
                used_vars.add(matched_var)
                info = scan_var(var_dir / matched_var)
                queue.extend(info.get("dependencies", []))

    protected_assets = {v for v in all_vars if is_asset_var(v)}
    unused_vars = all_vars - used_vars - protected_assets
    
    return scene_vars, used_vars, unused_vars


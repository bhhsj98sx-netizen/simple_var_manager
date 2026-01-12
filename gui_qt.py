import sys
import os
import json
import subprocess
import zipfile
import time
import hashlib
import urllib.request
import re
import tempfile
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QVBoxLayout, QFileDialog, QProgressBar, QMessageBox,
    QScrollArea, QGridLayout, QLineEdit,
    QHBoxLayout, QFrame, QCheckBox, QDialog, QTextEdit, QComboBox, QTextBrowser,
    QSpacerItem, QSizePolicy, QInputDialog
)
from PySide6.QtCore import QThread, Signal, Qt, QTimer, QEvent, QSize
from PySide6.QtGui import QIcon, QDesktopServices, QPalette

from PySide6.QtCore import QUrl

from core.resolver import resolve_dependency, is_asset_var
from core.scanner import scan_var
from core.scene_card import SceneCard

def resource_path(rel_path: str) -> Path:
    """
    Get absolute path to resource, works for dev and for PyInstaller --onefile
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / rel_path
    return Path(rel_path)



# ======================
# App version + Updater (GitHub Releases)
# ======================
APP_VERSION = "1.1.0"

GITHUB_OWNER = "bhhsj98sx-netizen"
GITHUB_REPO = "simple_var_manager"
GITHUB_LATEST_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

RELEASE_ASSET_NAME = "VaM.Simple.Var.Manager.exe"
UPDATER_EXE_NAME = "updater.exe"


# ======================
# Help text (HTML)
# ======================
HELP_TEXT = """
<h2>About This App</h2>

<p>
This is a simple app made to help you <b>temporarily disable .var files you don’t need</b>,
so <b>Virt-A-Mate loads faster and runs smoother</b>.
</p>

<p>
You can choose which scenes you want to play in your current session, and the app will
<b>disable all unrelated .var files</b>, so VaM doesn’t need to read them.
</p>

<p>
Disabled files are NOT deleted.<br>
They are temporarily renamed from:<br>
<b>Something.var</b> → <b>Something.var.disabled</b>
</p>

<p>
When you close <b>VaM.exe</b>, all renamed files will automatically restore back to normal.
</p>

<hr>

<h3>How to Use</h3>

<ol>
  <li>
    <b>Select your VaM directory (Directory where VaM.exe is located)</b><br>
    (Please wait a moment while the app scans your files)
  </li>

  <li>
    <b>Tick "Enable Scene Selection"</b>
  </li>

  <li>
    <b>Select scenes</b>
    <ul>
      <li>
        <b>Choose which scenes you want to play</b><br>
        All unrelated .var files to those scenes will be temporarily disabled.
      </li>
      <li>
        <b>Select nothing</b><br>
        The app will disable only .var files that are <b>not used by any scene at all</b>.
      </li>
    </ul>
  </li>

  <li>
    <b>Click Launch VaM.exe / VaM Launcher</b><br>
    The app will temporarily disable unrelated .var files by renaming them.
  </li>

  <li>
    <b>Live mode (NEW)</b><br>
    While VaM is running, you can change selection and click <b>Update Scene Selection</b>.<br>
    Then in VaM, use its refresh / rescan packages button to update the package list.
  </li>

  <li>
    <b>No extra steps needed</b><br>
    When you close <b>VaM.exe</b>, all renamed .var files will automatically return to normal.
  </li>
</ol>

<hr>

<h3>Buttons Explanation</h3>

<ul>
  <li>
    <b>Refresh</b><br>
    Re-checks the current VaM folder for changes (new / removed / modified .var files or scenes).<br>
    Use this if you manually add, remove, or edit files while the app is open.
  </li>

  <li>
    <b>Restore VARs Manually</b><br>
    Immediately restores all previously disabled <b>.var.disabled</b> files back to normal.<br>
    Use this if:
    <ul>
      <li>VaM crashed</li>
      <li>You closed the app unexpectedly</li>
      <li>You want to cancel a lean session manually</li>
    </ul>
  </li>

  <li>
    <b>Update Scene Selection</b><br>
    Applies your current scene selection <b>while VaM is already running</b>.<br>
    This allows you to enable or disable .var files without restarting VaM.<br><br>
    After clicking this button, you may need to use VaM’s
    <b>refresh / rescan packages</b> button inside the game.
  </li>
</ul>

<hr>

<h3>Presets</h3>
<p>
You can save and load your scene selections as presets, so you don’t need to
select them again next time.
</p>

<p><b>Enjoy! ^^</b></p>

"""


# ======================
# App data dir (IMPORTANT for --onefile)
# ======================
APP_VENDOR = "LQuest"
APP_NAME = "VaM Simple VAR Manager"

def app_data_dir() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    p = base / APP_VENDOR / APP_NAME
    p.mkdir(parents=True, exist_ok=True)
    return p

def previews_dir() -> Path:
    p = app_data_dir() / "previews"
    p.mkdir(parents=True, exist_ok=True)
    return p

def cache_path() -> Path:
    return app_data_dir() / "scene_cache.json"

def config_path() -> Path:
    return app_data_dir() / "launcher_config.json"


# ======================
# Supporters JSON (remote + cache)
# ======================
SUPPORTERS_URL = "https://raw.githubusercontent.com/bhhsj98sx-netizen/varapp-supporter/refs/heads/main/supporter.json"

def supporters_cache_path() -> Path:
    return app_data_dir() / "supporters_cache.json"

def _normalize_supporters_payload(obj: dict) -> dict:
    if not isinstance(obj, dict):
        return {"updated": "unknown", "supporters": []}

    updated = obj.get("updated") or obj.get("last_updated") or obj.get("date") or "unknown"

    supporters = obj.get("supporters")
    if supporters is None:
        supporters = obj.get("supporter")
    if supporters is None:
        supporters = obj.get("names")
    if supporters is None:
        supporters = []

    norm = []
    if isinstance(supporters, list):
        for item in supporters:
            if isinstance(item, str):
                norm.append({"name": item, "tier": ""})
            elif isinstance(item, dict):
                name = item.get("name") or item.get("username") or item.get("display") or "Anonymous"
                tier = item.get("tier") or item.get("level") or ""
                norm.append({"name": str(name), "tier": str(tier)})
    elif isinstance(supporters, dict):
        if "name" in supporters:
            norm.append({"name": str(supporters.get("name", "Anonymous")), "tier": str(supporters.get("tier", ""))})
        elif "names" in supporters and isinstance(supporters["names"], list):
            for n in supporters["names"]:
                norm.append({"name": str(n), "tier": ""})

    return {"updated": str(updated), "supporters": norm}

def load_supporters_cached(max_age_hours: float = 2.0) -> dict:
    cache = supporters_cache_path()

    if cache.exists():
        try:
            age_seconds = time.time() - cache.stat().st_mtime
            if age_seconds < max_age_hours * 3600:
                obj = json.loads(cache.read_text(encoding="utf-8"))
                return _normalize_supporters_payload(obj)
        except Exception:
            pass

    try:
        req = urllib.request.Request(
            SUPPORTERS_URL,
            headers={"User-Agent": f"VSVM/{APP_VERSION}"}
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = resp.read().decode("utf-8", errors="ignore")
        obj = json.loads(data)

        cache.write_text(json.dumps(obj, indent=2), encoding="utf-8")
        return _normalize_supporters_payload(obj)
    except Exception:
        if cache.exists():
            try:
                obj = json.loads(cache.read_text(encoding="utf-8"))
                return _normalize_supporters_payload(obj)
            except Exception:
                return {"updated": "unknown", "supporters": []}
        return {"updated": "unknown", "supporters": []}


# ======================
# Config helpers (in app_data_dir)
# ======================
def load_config() -> dict:
    p = config_path()
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_config(cfg: dict):
    config_path().write_text(json.dumps(cfg, indent=2), encoding="utf-8")


# ======================
# Preview cache helpers
# ======================
def _safe_preview_key(var_name: str, scene_name: str) -> str:
    raw = f"{var_name}::{scene_name}".encode("utf-8", errors="ignore")
    return hashlib.sha1(raw).hexdigest()

def write_preview_bytes(var_name: str, scene_name: str, image_bytes: bytes | None) -> str:
    if not image_bytes:
        return ""
    key = _safe_preview_key(var_name, scene_name)
    rel = f"previews/{key}.bin"
    fp = app_data_dir() / rel
    try:
        fp.parent.mkdir(parents=True, exist_ok=True)
        fp.write_bytes(image_bytes)
        return rel
    except Exception:
        return ""

def read_preview_bytes(rel_path: str) -> bytes | None:
    if not rel_path:
        return None
    fp = app_data_dir() / rel_path
    try:
        if fp.exists():
            return fp.read_bytes()
    except Exception:
        return None
    return None

def _var_signature(p: Path) -> str:
    try:
        st = p.stat()
        return f"{st.st_size}:{int(st.st_mtime)}"
    except Exception:
        return "0:0"

def _file_signature(p: Path) -> str:
    try:
        st = p.stat()
        return f"{st.st_size}:{int(st.st_mtime)}"
    except Exception:
        return "0:0"


# ======================
# VaM-like helpers
# ======================
def _parse_var_base_and_version(var_filename: str) -> tuple[str, str]:
    name = var_filename
    if name.lower().endswith(".var"):
        name = name[:-4]
    parts = name.split(".")
    if len(parts) < 3:
        return name, ""
    version = parts[-1]
    base = ".".join(parts[:-1])
    return base, version

def _choose_latest_vars(all_var_names: list[str]) -> set[str]:
    best: dict[str, tuple[int, str, str]] = {}
    for fn in all_var_names:
        base, ver = _parse_var_base_and_version(fn)

        ver_num = -1
        try:
            ver_num = int(ver)
        except Exception:
            ver_num = -1

        current = best.get(base)
        cand = (ver_num, ver, fn)
        if current is None:
            best[base] = cand
        else:
            if cand[0] > current[0]:
                best[base] = cand
            elif cand[0] == current[0]:
                if cand[1] > current[1]:
                    best[base] = cand
                elif cand[1] == current[1]:
                    if cand[2] > current[2]:
                        best[base] = cand

    return {t[2] for t in best.values()}

def _is_hidden_sidecar(json_path: Path) -> bool:
    try:
        return (Path(str(json_path) + ".hide")).exists()
    except Exception:
        return False

def _should_count_loose_scene(relp: str) -> bool:
    rp = (relp or "").replace("\\", "/").strip()
    if not rp:
        return False

    name = Path(rp).name.lower()
    stem = Path(rp).stem

    if name == "default.json":
        return False
    if stem.isdigit():
        return False

    return True



# ======================
# Disable-by-rename helpers
# ======================
DISABLED_SUFFIX = ".disabled"

def disabled_name(original: str) -> str:
    return original + DISABLED_SUFFIX

def manifest_path_for(vam_dir: Path) -> Path:
    return vam_dir / "_vam_temp_manifest.json"


# ======================
# Loading popup
# ======================
class LoadingPopup(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Loading")
        self.setWindowFlags(
            Qt.Dialog
            | Qt.FramelessWindowHint
            | Qt.WindowStaysOnTopHint
        )
        self.setModal(False)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        self._base_text = "Loading"
        self._dots = 0

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)

        box = QFrame()
        box.setStyleSheet("""
            QFrame {
                background: rgba(20,20,20,0.92);
                border: 1px solid rgba(255,255,255,0.18);
                border-radius: 12px;
            }
        """)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(16, 14, 16, 16)
        lay.setSpacing(10)

        self.label = QLabel("Loading...")
        self.label.setStyleSheet("color: #eee; font-size: 13px; font-weight: 600;")
        self.label.setAlignment(Qt.AlignCenter)
        lay.addWidget(self.label)

        self.sub = QLabel("Please wait")
        self.sub.setStyleSheet("color: #aaa; font-size: 11px;")
        self.sub.setAlignment(Qt.AlignCenter)
        lay.addWidget(self.sub)

        self.bar = QProgressBar()
        self.bar.setRange(0, 0)
        self.bar.setMinimumHeight(16)
        self.bar.setTextVisible(False)
        self.bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid rgba(255,255,255,0.20);
                border-radius: 6px;
                background-color: rgba(255,255,255,0.08);
            }
            QProgressBar::chunk {
                background-color: #3daee9;
                border-radius: 6px;
            }
        """)
        lay.addWidget(self.bar)

        root.addWidget(box)

        self._timer = QTimer(self)
        self._timer.setInterval(180)
        self._timer.timeout.connect(self._tick)

        self.resize(320, 110)

    def _tick(self):
        self._dots = (self._dots + 1) % 4
        self.label.setText(self._base_text + ("." * self._dots))

    def start(self, text: str = "Loading", sub: str = "Please wait"):
        self._base_text = text or "Loading"
        self.sub.setText(sub or "")
        self._dots = 0
        self.label.setText(self._base_text)
        self._timer.start()

        if self.parent():
            p = self.parent().frameGeometry()
            self.move(p.center() - self.rect().center())

        self.show()
        QApplication.processEvents()

    def stop(self):
        self._timer.stop()
        self.hide()
        QApplication.processEvents()


# ======================
# Help dialog
# ======================
class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Help")
        self.resize(720, 520)

        layout = QVBoxLayout(self)
        self.text = QTextBrowser()
        self.text.setOpenExternalLinks(True)
        self.text.setHtml(HELP_TEXT)
        layout.addWidget(self.text, 1)

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)


# ======================
# Donation dialog
# ======================
class DonationDialog(QDialog):
    PATREON_URL = "https://www.patreon.com/c/LQuest"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Support VSVM")
        self.resize(520, 360)

        layout = QVBoxLayout(self)

        title = QLabel("Thanks to my supporters")
        title.setStyleSheet("font-weight: bold; font-size: 16px;")
        layout.addWidget(title)

        data = load_supporters_cached(max_age_hours=2.0)
        supporters = data.get("supporters", [])
        updated = data.get("updated", "unknown")

        if supporters:
            lines = []
            for s in supporters:
                name = (s.get("name") or "Anonymous").strip()
                tier = (s.get("tier") or "").strip()
                if tier:
                    lines.append(f"• {name}  ({tier})")
                else:
                    lines.append(f"• {name}")
            text = "\n".join(lines)
        else:
            text = "(No supporters yet)"

        self.text = QTextEdit()
        self.text.setReadOnly(True)
        self.text.setPlainText(text + f"\n\nLast updated: {updated}")
        layout.addWidget(self.text, 1)

        row = QHBoxLayout()
        row.addStretch(1)

        btn_open = QPushButton("Open Patreon")
        btn_open.clicked.connect(self.open_patreon)
        row.addWidget(btn_open)

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        row.addWidget(btn_close)

        layout.addLayout(row)

    def open_patreon(self):
        QDesktopServices.openUrl(QUrl(self.PATREON_URL))


# ======================
# Welcome popup (one-time)
# ======================
class WelcomeDialog(QDialog):
    PATREON_URL = "https://www.patreon.com/c/LQuest"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Welcome")
        self.resize(520, 260)

        layout = QVBoxLayout(self)

        html = (
            "<h3>Welcome to VaM Simple VAR Manager</h3>"
            "<p>"
            "I made this tool for free to help improve your VaM experience.<br><br>"
            "If you find it useful, you can support me on Patreon so I can keep making updates and new tools.<br><br>"
            "Enjoy! ❤️"
            "</p>"
        )

        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(False)
        self.browser.setHtml(html)
        layout.addWidget(self.browser, 1)

        row = QHBoxLayout()
        row.addStretch(1)

        btn = QPushButton("Open Patreon")
        btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(self.PATREON_URL)))
        row.addWidget(btn)

        close = QPushButton("Continue")
        close.clicked.connect(self.accept)
        row.addWidget(close)

        layout.addLayout(row)


# ======================
# Worker Thread (scan)
# ======================
class AnalyzeWorker(QThread):
    finished = Signal(object, int, dict)  # scene_entries, total_var_count, cache_obj

    def __init__(self, addon_dir: Path, saves_scene_dir: Path | None, old_cache: dict):
        super().__init__()
        self.addon_dir = addon_dir
        self.saves_scene_dir = saves_scene_dir
        self.old_cache = old_cache or {}

    def run(self):
        all_var_files = list(self.addon_dir.glob("*.var"))
        total_var_count = len(all_var_files)

        all_var_names = [p.name for p in all_var_files]
        allowed_var_names = _choose_latest_vars(all_var_names)

        show_hidden = False

        old_vars = (self.old_cache.get("vars") or {}) if isinstance(self.old_cache, dict) else {}
        old_loose = (self.old_cache.get("loose") or {}) if isinstance(self.old_cache, dict) else {}

        new_vars: dict[str, dict] = {}
        new_loose: dict[str, dict] = {}

        scene_entries: list[dict] = []

        for p in all_var_files:
            name = p.name
            if name not in allowed_var_names:
                continue

            sig = _var_signature(p)

            cached = old_vars.get(name) if isinstance(old_vars, dict) else None
            if isinstance(cached, dict) and cached.get("sig") == sig:
                scenes = cached.get("scenes", [])
                if isinstance(scenes, list):
                    for s in scenes:
                        if not isinstance(s, dict):
                            continue
                        scene_entries.append({
                            "source": "var",
                            "scene_name": s.get("scene_name", ""),
                            "var_name": name,
                            "preview_relpath": s.get("preview_relpath", ""),
                            "loose_relpath": "",
                            "is_girl_looks": None,
                        })
                new_vars[name] = cached
                continue

            info = scan_var(p)
            scenes_out = []

            if info.get("has_scene"):
                for scene in info.get("scenes", []):
                    scene_name = scene.get("scene_name", "")
                    image_bytes = scene.get("image_bytes", None)
                    rel = write_preview_bytes(name, scene_name, image_bytes)

                    scenes_out.append({
                        "scene_name": scene_name,
                        "preview_relpath": rel,
                    })

                    scene_entries.append({
                        "source": "var",
                        "scene_name": scene_name,
                        "var_name": name,
                        "preview_relpath": rel,
                        "loose_relpath": "",
                        "is_girl_looks": None,
                    })

            new_vars[name] = {"sig": sig, "scenes": scenes_out}

        if self.saves_scene_dir and self.saves_scene_dir.exists():
            for sp in self.saves_scene_dir.rglob("*.json"):
                if (not show_hidden) and _is_hidden_sidecar(sp):
                    continue

                try:
                    relp = str(sp.relative_to(self.saves_scene_dir)).replace("\\", "/")
                except Exception:
                    relp = sp.name

                if not _should_count_loose_scene(relp):
                    continue

                sig = _file_signature(sp)

                cached = old_loose.get(relp) if isinstance(old_loose, dict) else None
                if isinstance(cached, dict) and cached.get("sig") == sig:
                    scene_name = cached.get("scene_name") or relp
                    scene_entries.append({
                        "source": "loose",
                        "scene_name": scene_name,
                        "var_name": "(Saves/scene)",
                        "preview_relpath": "",
                        "loose_relpath": relp,
                        "is_girl_looks": None,
                    })
                    new_loose[relp] = cached
                    continue

                scene_name = relp
                scene_entries.append({
                    "source": "loose",
                    "scene_name": scene_name,
                    "var_name": "(Saves/scene)",
                    "preview_relpath": "",
                    "loose_relpath": relp,
                    "is_girl_looks": None,
                })
                new_loose[relp] = {"sig": sig, "scene_name": scene_name}

        cache_obj = {
            "version": 3,
            "addon_dir": str(self.addon_dir),
            "vars": new_vars,
            "loose": new_loose,
            "looks": self.old_cache.get("looks", {}) if isinstance(self.old_cache, dict) else {},
            "saved_at": time.time(),
            "only_latest_always": True,
            "show_hidden_always": False,
        }

        self.finished.emit(scene_entries, total_var_count, cache_obj)


# ======================
# ChangeCheckWorker
# ======================
class ChangeCheckWorker(QThread):
    finished = Signal(bool)  # needs_rescan?

    def __init__(self, addon_dir: Path, saves_scene_dir: Path | None, cache_obj: dict):
        super().__init__()
        self.addon_dir = addon_dir
        self.saves_scene_dir = saves_scene_dir
        self.cache_obj = cache_obj or {}

    def run(self):
        try:
            if not isinstance(self.cache_obj, dict):
                self.finished.emit(True)
                return

            cached_addon = self.cache_obj.get("addon_dir")
            if str(self.addon_dir) != str(cached_addon):
                self.finished.emit(True)
                return

            cached_vars = self.cache_obj.get("vars", {}) if isinstance(self.cache_obj.get("vars"), dict) else {}
            cached_loose = self.cache_obj.get("loose", {}) if isinstance(self.cache_obj.get("loose"), dict) else {}

            all_var_files = list(self.addon_dir.glob("*.var"))
            all_var_names = [p.name for p in all_var_files]
            allowed_now = _choose_latest_vars(all_var_names)

            if set(cached_vars.keys()) != set(allowed_now):
                self.finished.emit(True)
                return

            by_name = {p.name: p for p in all_var_files}
            for name in allowed_now:
                p = by_name.get(name)
                if not p:
                    self.finished.emit(True)
                    return
                sig_now = _var_signature(p)
                sig_cached = (cached_vars.get(name) or {}).get("sig")
                if sig_now != sig_cached:
                    self.finished.emit(True)
                    return

            show_hidden = False
            loose_now: dict[str, str] = {}
            if self.saves_scene_dir and self.saves_scene_dir.exists():
                for sp in self.saves_scene_dir.rglob("*.json"):
                    if (not show_hidden) and _is_hidden_sidecar(sp):
                        continue
                    try:
                        relp = str(sp.relative_to(self.saves_scene_dir)).replace("\\", "/")
                    except Exception:
                        relp = sp.name
                    if not _should_count_loose_scene(relp):
                        continue
                    loose_now[relp] = _file_signature(sp)

            if set(loose_now.keys()) != set(cached_loose.keys()):
                self.finished.emit(True)
                return

            for relp, sig_now in loose_now.items():
                sig_cached = (cached_loose.get(relp) or {}).get("sig")
                if sig_now != sig_cached:
                    self.finished.emit(True)
                    return

            self.finished.emit(False)
        except Exception:
            self.finished.emit(True)


# ======================
# LooksWorker
# ======================
class LooksWorker(QThread):
    finished = Signal(dict)  # looks_map { "var::scene": bool }

    def __init__(self, addon_dir: Path, scene_pairs: list[tuple[str, str]], existing: dict):
        super().__init__()
        self.addon_dir = addon_dir
        self.scene_pairs = scene_pairs
        self.existing = existing or {}

    @staticmethod
    def _key(var_name: str, scene_name: str) -> str:
        return f"{var_name}::{scene_name}"

    @staticmethod
    def _detect_looks_in_var(var_path: Path, scene_name: str) -> bool:
        wanted = f"{scene_name}.json".lower()

        # NEW: fallback to .disabled
        if not var_path.exists():
            alt = Path(str(var_path) + DISABLED_SUFFIX)
            if alt.exists():
                var_path = alt

        try:
            with zipfile.ZipFile(var_path, "r") as z:
                target = None
                for n in z.namelist():
                    ln = n.lower().replace("\\", "/")
                    if not ln.endswith(".json"):
                        continue
                    if "saves/scene/" not in ln:
                        continue
                    if Path(ln).name == wanted:
                        target = n
                        break
                if not target:
                    return False

                raw = z.read(target)
                txt = raw.decode("utf-8", errors="ignore").lower()
                return ("/female" in txt) and ("/male" not in txt)
        except Exception:
            return False


    def run(self):
        out = dict(self.existing) if isinstance(self.existing, dict) else {}
        for var_name, scene_name in self.scene_pairs:
            k = self._key(var_name, scene_name)
            if k in out:
                continue
            try:
                out[k] = self._detect_looks_in_var(self.addon_dir / var_name, scene_name)
            except Exception:
                out[k] = False
        self.finished.emit(out)


# ======================
# Dependency row widget
# ======================
class DependencyRow(QWidget):
    def __init__(self, text: str, present: bool):
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(8)

        strip = QFrame()
        strip.setFixedWidth(10)
        strip.setMinimumHeight(18)
        strip.setStyleSheet(
            "background-color: #2ecc71; border-radius: 3px;" if present
            else "background-color: #e74c3c; border-radius: 3px;"
        )
        layout.addWidget(strip)

        label = QLabel(text)
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        label.setStyleSheet("color: #ddd;")
        layout.addWidget(label, 1)

        self.setLayout(layout)


# ======================
# Main GUI
# ======================
def is_dark_mode(app: QApplication) -> bool:
    pal = app.palette()
    # value 0-255, makin kecil makin gelap
    win = pal.color(QPalette.Window).value()
    txt = pal.color(QPalette.WindowText).value()
    # kalau background gelap dan text terang -> dark mode
    return win < 128 and txt > 128

def build_btn_css(dark: bool) -> str:
    if not dark:
        # LIGHT MODE → respect system palette, only sizing + disabled clarity
        return """
        QPushButton, QComboBox {
            min-height: 26px;
            padding: 6px 14px;
            font-size: 12px;
        }

        QPushButton:disabled, QComboBox:disabled {
            color: #666;
            background-color: #f0f0f0;
            border: 1px solid #cfcfcf;
        }
        """
    else:
        # DARK MODE → custom visual styling
        return """
        QPushButton, QComboBox {
            min-height: 26px;
            padding: 6px 14px;
            font-size: 12px;

            background-color: rgba(255,255,255,0.08);
            border: 1px solid rgba(255,255,255,0.18);
            border-radius: 8px;
            color: #eee;
        }

        QPushButton:hover, QComboBox:hover {
            background-color: rgba(255,255,255,0.14);
            border: 1px solid rgba(255,255,255,0.28);
        }

        QPushButton:pressed {
            background-color: rgba(255,255,255,0.18);
        }

        QPushButton:disabled, QComboBox:disabled {
            background-color: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.08);
            color: #777;
        }
        """

def build_scrollbar_css(dark: bool) -> str:
    if dark:
        return """
        QScrollBar:vertical {
            width: 18px;
            background: rgba(255,255,255,0.05);
            margin: 4px 3px 4px 3px;
            border-radius: 9px;
        }
        QScrollBar::handle:vertical {
            background: rgba(255,255,255,0.40);
            min-height: 48px;
            border-radius: 9px;
        }
        QScrollBar::handle:vertical:hover {
            background: rgba(255,255,255,0.60);
        }
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical {
            height: 0px;
        }
        QScrollBar::add-page:vertical,
        QScrollBar::sub-page:vertical {
            background: none;
        }
        """
    else:
        return """
        QScrollBar:vertical {
            width: 18px;
            background: rgba(0,0,0,0.06);
            margin: 4px 3px 4px 3px;
            border-radius: 9px;
        }
        QScrollBar::handle:vertical {
            background: rgba(0,0,0,0.35);
            min-height: 48px;
            border-radius: 9px;
        }
        QScrollBar::handle:vertical:hover {
            background: rgba(0,0,0,0.55);
        }
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical {
            height: 0px;
        }
        QScrollBar::add-page:vertical,
        QScrollBar::sub-page:vertical {
            background: none;
        }
        """




class MainWindow(QWidget):
    def _set_dim_label(self, label: QLabel, dark_hex: str):
        if self._dark:
            label.setStyleSheet(f"color: {dark_hex};")
    def __init__(self):
        super().__init__()

        _ = app_data_dir()
        _ = previews_dir()

        self.setWindowTitle("VaM Simple VAR Manager v1.1.0")
        self.resize(1220, 780)

        self.cfg = load_config()

        self.vam_dir: Path | None = None
        self.addon_dir: Path | None = None
        self.saves_scene_dir: Path | None = None

        self.selected_scene_vars: set[str] = set()

        self.cards: list[SceneCard] = []
        self.card_by_var: dict[str, list[SceneCard]] = {}

        self.total_vars_count = 0
        self.unused_vars_count = 0
        self.scene_entries: list[dict] = []
        self.total_scene_files_found = 0

        #apply button variable
        self.session_baseline_scene_vars: set[str] = set()
        self.selection_dirty = False

        self._apply_attention_timer = QTimer(self)
        self._apply_attention_timer.setInterval(450)
        self._apply_attention_timer.timeout.connect(self._tick_apply_attention)
        self._apply_attention_on = False

                # ---- selection dirty debounce (prevents UI freeze during batch/preset load)
        self._dirty_check_timer = QTimer(self)
        self._dirty_check_timer.setSingleShot(True)
        self._dirty_check_timer.setInterval(150)
        self._dirty_check_timer.timeout.connect(self._check_selection_dirty)



        # lean session state
        self.lean_active = False
        self.vam_seen_running = False

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(1000)
        self.poll_timer.timeout.connect(self.check_vam_state)

        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.setInterval(100)
        self._resize_timer.timeout.connect(lambda: self.apply_filter(self.search.text()))

                # drag-to-scroll (touch-like)
        self._drag_scroll_active = False
        self._drag_scroll_start_pos = None
        self._drag_scroll_start_value = 0


        self._syncing_selection_ui = False
        self._batch_selection = False

        self.looks_map: dict[str, bool] = {}

        self._startup_in_progress = False
        self._refresh_in_progress = False

        self._dark = is_dark_mode(QApplication.instance())
        self._btn_css = build_btn_css(self._dark)


        self.loading = LoadingPopup(self)

        root = QHBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(12)

        # ------------------
        # LEFT
        # ------------------
        left = QWidget()
        self.left_layout = QVBoxLayout(left)
        self.left_layout.setSpacing(8)

        row_head = QHBoxLayout()
        self.title = QLabel("")
        row_head.addWidget(self.title, 1)

        self.btn_help = QPushButton("Help")
        icon_path = resource_path("icons/qmark.png")
        if icon_path.exists():
            self.btn_help.setIcon(QIcon(str(icon_path)))
            self.btn_help.setIconSize(QSize(18, 18))
        self.btn_help.setStyleSheet(self._btn_css)
        self.btn_help.clicked.connect(self.open_help)
        row_head.addWidget(self.btn_help, 0, Qt.AlignRight)

        # (You said you already moved Update beside Help earlier; keep it here.)
        self.btn_check_update = QPushButton(" Update")
        icon_path = resource_path("icons/update.png")
        if icon_path.exists():
            self.btn_check_update.setIcon(QIcon(str(icon_path)))
            self.btn_check_update.setIconSize(QSize(18, 18))
        self.btn_check_update.setToolTip("Check GitHub for a newer version and update.")
        self.btn_check_update.setStyleSheet(self._btn_css)
        self.btn_check_update.clicked.connect(self.check_update_clicked)
        row_head.addWidget(self.btn_check_update, 0, Qt.AlignRight)

        self.btn_donate = QPushButton(" Support Me")
        icon_path = resource_path("icons/donation.png")
        if icon_path.exists():
            self.btn_donate.setIcon(QIcon(str(icon_path)))
            self.btn_donate.setIconSize(QSize(18, 18))
        self.btn_donate.setStyleSheet(self._btn_css)
        self.btn_donate.clicked.connect(self.open_donation)
        row_head.addWidget(self.btn_donate, 0, Qt.AlignRight)

        self.left_layout.addLayout(row_head)

        self.status = QLabel("Starting...")
        self.left_layout.addWidget(self.status)


        self.apply_row = QHBoxLayout()
        self.apply_row.setSpacing(8)

        self.left_layout.addLayout(self.apply_row)

        self.info_label = QLabel("Total VARs: - | Unused VARs: - | Scenes: -")
        self._set_dim_label(self.info_label, "#aaa")

        self.left_layout.addWidget(self.info_label)

        self.left_layout.addItem(QSpacerItem(0, 10, QSizePolicy.Minimum, QSizePolicy.Fixed))

        self.progress = QProgressBar()
        self.progress.setMinimumHeight(20)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid rgba(255,255,255,0.25);
                border-radius: 6px;
                background-color: rgba(255,255,255,0.08);
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #3daee9;
                border-radius: 6px;
            }
        """)
        self.progress.setVisible(False)
        self.left_layout.addWidget(self.progress)

        row_controls = QHBoxLayout()
        row_controls.setSpacing(14)
        self.left_layout.addItem(QSpacerItem(0, 20, QSizePolicy.Minimum, QSizePolicy.Fixed))

        ctl_left = QWidget()
        ctl_left_lay = QVBoxLayout(ctl_left)
        ctl_left_lay.setContentsMargins(0, 0, 0, 0)
        ctl_left_lay.setSpacing(8)

        row_top = QHBoxLayout()

        self.btn_refresh = QPushButton("Refresh")
        self.btn_refresh.setToolTip("Check for changes in the current VaM folder and update if needed.")
        self.btn_refresh.setStyleSheet(self._btn_css)
        self.btn_refresh.clicked.connect(self.refresh_clicked)
        row_top.addWidget(self.btn_refresh, 1)

        self.btn_select = QPushButton("Select VaM Directory")
        self.btn_select.setStyleSheet(self._btn_css)
        self.btn_select.clicked.connect(self.select_folder)
        icon_path = resource_path("icons/folder.png")
        if icon_path.exists():
            self.btn_select.setIcon(QIcon(str(icon_path)))
            self.btn_select.setIconSize(QSize(18, 18))
        row_top.addWidget(self.btn_select, 2)

        ctl_left_lay.addLayout(row_top)

        self.btn_restore = QPushButton("Restore VARs Manually")
        self.btn_restore.setStyleSheet(self._btn_css)
        self.btn_restore.setEnabled(False)
        self.btn_restore.clicked.connect(self.restore_offloaded_vars)
        row_top.addWidget(self.btn_restore, 1)
        
        self.btn_apply_now = QPushButton("Update Scene Selection")
        self.btn_apply_now.setToolTip("Apply your current selection while VaM is running (VaM needs in-game refresh).")
        self.btn_apply_now.setStyleSheet(self._btn_css)
        self.btn_apply_now.setEnabled(False)
        self.btn_apply_now.clicked.connect(self.apply_selection_now_clicked)
        row_top.addWidget(self.btn_apply_now, 1)  

        row_launch = QHBoxLayout()

        icon_path = resource_path("icons/playbtn.png")
        icon_obj = QIcon(str(icon_path)) if icon_path.exists() else None

        self.btn_launch_vam = QPushButton("  Launch VaM.exe")
        if icon_obj:
            self.btn_launch_vam.setIcon(icon_obj)
            self.btn_launch_vam.setIconSize(QSize(13, 13))
        self.btn_launch_vam.setStyleSheet(self._btn_css)
        self.btn_launch_vam.clicked.connect(self.launch_vam_exe_lean)
        row_launch.addWidget(self.btn_launch_vam, 1)

        self.btn_launch_vd = QPushButton("  Launch VaM (Virtual Desktop)")
        if icon_obj:
            self.btn_launch_vd.setIcon(icon_obj)
            self.btn_launch_vd.setIconSize(QSize(13, 13))
        self.btn_launch_vd.setStyleSheet(self._btn_css)
        self.btn_launch_vd.clicked.connect(self.launch_vam_vd_lean)
        row_launch.addWidget(self.btn_launch_vd, 1)


        self.btn_launch_launcher = QPushButton("  Launch VaM Launcher")
        if icon_obj:
            self.btn_launch_launcher.setIcon(icon_obj)
            self.btn_launch_launcher.setIconSize(QSize(13, 13))
        self.btn_launch_launcher.setStyleSheet(self._btn_css)
        self.btn_launch_launcher.clicked.connect(self.launch_vam_launcher_lean)
        row_launch.addWidget(self.btn_launch_launcher, 1)

        
        
      

        ctl_left_lay.addLayout(row_launch)

        vline = QFrame()
        vline.setFrameShape(QFrame.VLine)
        vline.setFrameShadow(QFrame.Sunken)
        vline.setStyleSheet("color: rgba(255,255,255,0.14);")

        ctl_right = QWidget()
        ctl_right_lay = QVBoxLayout(ctl_right)
        ctl_right_lay.setContentsMargins(0, 0, 0, 0)
        ctl_right_lay.setSpacing(8)

        self.chk_select_mode = QCheckBox("Enable Scene Selection")
        self.chk_select_mode.stateChanged.connect(self.on_toggle_selection_mode)
        self.chk_select_mode.setEnabled(False)
        ctl_right_lay.addWidget(self.chk_select_mode, 0)

        row_preset = QHBoxLayout()
        self.MAX_PRESETS = 5  
        self.preset_combo = QComboBox()
        self.preset_combo.addItems([f"Scene Selection Preset {i}" for i in range(1, self.MAX_PRESETS + 1)])

        self.preset_combo.setMinimumHeight(30)
        row_preset.addWidget(self.preset_combo, 1)

        self.btn_save_preset = QPushButton("Save")
        self.btn_save_preset.setStyleSheet(self._btn_css)
        self.btn_save_preset.clicked.connect(self.save_preset)
        row_preset.addWidget(self.btn_save_preset, 0)

        self.btn_load_preset = QPushButton("Load")
        self.btn_load_preset.setStyleSheet(self._btn_css)
        self.btn_load_preset.clicked.connect(self.load_preset)
        row_preset.addWidget(self.btn_load_preset, 0)

        ctl_right_lay.addLayout(row_preset)
        ctl_right_lay.addStretch(1)

        row_controls.addWidget(ctl_left, 3)
        row_controls.addWidget(vline, 0)
        row_controls.addWidget(ctl_right, 2)

        self.left_layout.addLayout(row_controls)

        self.left_layout.addItem(QSpacerItem(0, 10, QSizePolicy.Minimum, QSizePolicy.Fixed))
        self.left_layout.addItem(QSpacerItem(0, 20, QSizePolicy.Minimum, QSizePolicy.Fixed))

        self.selection_info = QLabel("Total Scene: 0 scenes | Selected Scene: 0 scenes")
        self._set_dim_label(self.selection_info, "#aaa")
        self.left_layout.addWidget(self.selection_info)

        row_search = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search scene name / var name / creator...")
        self.search.setMinimumHeight(32)
        self.search.setEnabled(False)
        self.search.textChanged.connect(self.apply_filter)
        row_search.addWidget(self.search, 1)
        self.left_layout.addLayout(row_search)

        self.card_header = QWidget()
        self.card_header.setStyleSheet("""
            QWidget {
                background: rgba(255,255,255,0.03);
                border: 1px solid rgba(255,255,255,0.08);
                border-radius: 8px;
            }
        """)
        header_layout = QHBoxLayout(self.card_header)
        header_layout.setContentsMargins(10, 8, 10, 8)
        header_layout.setSpacing(8)

        self.btn_select_all = QPushButton("Select All (Visible)")
        self.btn_select_all.setEnabled(False)
        self.btn_select_all.clicked.connect(self.select_all_visible)
        self.btn_select_all.setStyleSheet(self._btn_css)
        header_layout.addWidget(self.btn_select_all)

        self.btn_clear_sel = QPushButton("Clear Selection")
        self.btn_clear_sel.setEnabled(False)
        self.btn_clear_sel.clicked.connect(self.clear_selection)
        self.btn_clear_sel.setStyleSheet(self._btn_css)
        header_layout.addWidget(self.btn_clear_sel)

        header_layout.addStretch(1)

        self.chk_girl_looks_only = QCheckBox('Show Girl "Looks" Only')
        self.chk_girl_looks_only.setEnabled(False)
        self.chk_girl_looks_only.stateChanged.connect(self.on_toggle_looks_only)
        header_layout.addWidget(self.chk_girl_looks_only)

        self.left_layout.addWidget(self.card_header)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet(build_scrollbar_css(self._dark))

        self.scroll.viewport().installEventFilter(self)

        self.scene_container = QWidget()
        self.scene_container_layout = QVBoxLayout(self.scene_container)
        self.scene_container_layout.setContentsMargins(10, 10, 10, 10)
        self.scene_container_layout.setSpacing(8)

        self.grid_holder = QWidget()
        self.scene_grid = QGridLayout(self.grid_holder)
        self.scene_grid.setSpacing(10)
        self.scene_grid.setContentsMargins(0, 0, 0, 0)

        self.scene_container_layout.addWidget(self.grid_holder, 1)
        self.scroll.setWidget(self.scene_container)
        self.left_layout.addWidget(self.scroll, 1)

        # ------------------
        # RIGHT
        # ------------------
        right = QWidget()
        self.right_layout = QVBoxLayout(right)
        self.right_layout.setSpacing(8)

        self.dep_title = QLabel("Dependencies")
        self.dep_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.right_layout.addWidget(self.dep_title)

        self.dep_selected = QLabel('Click a card to see dependencies.\n(Selection mode: click selects)')
        self.dep_selected.setWordWrap(True)
        self._set_dim_label(self.dep_selected, "#bbb")
        self.right_layout.addWidget(self.dep_selected)

        self.dep_summary = QLabel("")
        self._set_dim_label(self.dep_summary, "#aaa")
        self.right_layout.addWidget(self.dep_summary)

        self.dep_scroll = QScrollArea()
        self.dep_scroll.setWidgetResizable(True)

        self.dep_container = QWidget()
        self.dep_list_layout = QVBoxLayout(self.dep_container)
        self.dep_list_layout.setContentsMargins(10, 10, 10, 10)
        self.dep_list_layout.setSpacing(6)

        self.dep_scroll.setWidget(self.dep_container)
        self.right_layout.addWidget(self.dep_scroll, 1)

        right.setMinimumWidth(380)
        right.setMaximumWidth(560)

        root.addWidget(left, 3)
        root.addWidget(right, 1)

        self.setLayout(root)

        self.refresh_refresh_button()
        self.refresh_restore_button()
        self.refresh_apply_button()
        self.update_selection_ui()

        self.refresh_preset_combo_names()

        self._maybe_show_welcome_once()
        QTimer.singleShot(0, self._startup_sequence)

    # ======================
    # Startup sequence
    # ======================
    def _startup_sequence(self):
        self._startup_in_progress = True
        self.loading.start("Loading", "Preparing scenes...")
        QTimer.singleShot(50, self.auto_open_last_folder_on_startup)

    def _end_busy(self):
        self.loading.stop()
        self._startup_in_progress = False
        self._refresh_in_progress = False
        self.refresh_refresh_button()

    def refresh_preset_combo_names(self):
        current = self.preset_combo.currentIndex()
        self.preset_combo.blockSignals(True)
        try:
            self.preset_combo.clear()
            for i in range(1, self.MAX_PRESETS + 1):
                self.preset_combo.addItem(self.get_preset_name(i))
        finally:
            self.preset_combo.blockSignals(False)
        if 0 <= current < self.preset_combo.count():
            self.preset_combo.setCurrentIndex(current)

    #apply button styling

    def _schedule_check_selection_dirty(self):
        # coalesce many rapid changes into 1 expensive check + possible QSS pulse update
        if self._dirty_check_timer.isActive():
            return
        self._dirty_check_timer.start()

    def _begin_batch_selection(self):
        self._batch_selection = True
        self.scene_container.setUpdatesEnabled(False)

    def _end_batch_selection(self):
        self._batch_selection = False
        self.scene_container.setUpdatesEnabled(True)
        # do heavy UI only once
        self.apply_filter(self.search.text())
        self.update_selection_ui()
        self._check_selection_dirty()


    def _apply_btn_normal_style(self):
        self.btn_apply_now.setStyleSheet(self._btn_css)

    def _apply_btn_attention_style(self, phase: bool):
        # phase True/False to “pulse”
        if phase:
            self.btn_apply_now.setStyleSheet(self._btn_css + """
                QPushButton {
                    border: 2px solid rgba(255, 215, 64, 0.95);
                    background-color: rgba(255, 215, 64, 0.20);
                    font-weight: 700;
                }
            """)
        else:
            self.btn_apply_now.setStyleSheet(self._btn_css + """
                QPushButton {
                    border: 2px solid rgba(255, 215, 64, 0.55);
                    background-color: rgba(255, 215, 64, 0.10);
                    font-weight: 700;
                }
            """)

    def _tick_apply_attention(self):
        self._apply_attention_on = not self._apply_attention_on
        self._apply_btn_attention_style(self._apply_attention_on)

    def set_apply_attention(self, on: bool):
        # only show attention if button is enabled
        if not self.btn_apply_now.isEnabled():
            on = False

        if on:
            if not self._apply_attention_timer.isActive():
                self._apply_attention_on = False
                self._apply_attention_timer.start()
        else:
            if self._apply_attention_timer.isActive():
                self._apply_attention_timer.stop()
            self._apply_btn_normal_style()

    
    def _preset_key(self, idx: int) -> str:
        return f"preset_{idx}"

    def _preset_names_map(self) -> dict:
        self.cfg.setdefault("preset_names", {})
        if not isinstance(self.cfg["preset_names"], dict):
            self.cfg["preset_names"] = {}
        return self.cfg["preset_names"]

    def _default_preset_name(self, idx: int) -> str:
        return f"Scene Selection Preset {idx}"

    def get_preset_name(self, idx: int) -> str:
        names = self._preset_names_map()
        name = names.get(self._preset_key(idx), "")
        name = (name or "").strip()
        return name if name else self._default_preset_name(idx)

    def set_preset_name(self, idx: int, name: str):
        names = self._preset_names_map()
        clean = (name or "").strip()
        if not clean:
            clean = self._default_preset_name(idx)
        names[self._preset_key(idx)] = clean
        save_config(self.cfg)


    def all_var_names_on_disk(self) -> set[str]:
        """
        Returns ALL var names that exist in AddonPackages, including disabled ones,
        normalized to original name (Something.var).
        """
        if not self.addon_dir:
            return set()

        out: set[str] = set()

        # enabled
        for p in self.addon_dir.glob("*.var"):
            out.add(p.name)

        # disabled -> normalize back to .var
        for p in self.addon_dir.glob("*.var.disabled"):
            orig = p.name[:-len(DISABLED_SUFFIX)]
            out.add(orig)

        return out


    # ======================
    # One-time welcome
    # ======================
    def _maybe_show_welcome_once(self):
        key = "welcome_shown_v1"
        if self.cfg.get(key):
            return
        WelcomeDialog(self).exec()
        self.cfg[key] = True
        save_config(self.cfg)

    # ======================
    # Donation
    # ======================
    def open_donation(self):
        DonationDialog(self).exec()

    # ======================
    # Presets
    # ======================
    def _preset_key(self, idx: int) -> str:
        return f"preset_{idx}"

    def save_preset(self):
        if not self.is_selection_mode():
            return

        idx = self.preset_combo.currentIndex() + 1
        key = self._preset_key(idx)

        # default / current name
        current_name = self.get_preset_name(idx)

        name, ok = QInputDialog.getText(
            self,
            "Preset Name",
            f"Enter name for Preset {idx}:",
            QLineEdit.Normal,
            current_name
        )
        if not ok:
            return  # user cancelled

        name = (name or "").strip()
        if not name:
            name = self._default_preset_name(idx)

        # save selection
        self.cfg.setdefault("presets", {})
        self.cfg["presets"][key] = sorted(self.selected_scene_vars)

        # save name
        self.cfg.setdefault("preset_names", {})
        self.cfg["preset_names"][key] = name

        self.cfg["last_preset"] = idx
        save_config(self.cfg)

        # update combo display
        self.refresh_preset_combo_names()
        self.preset_combo.setCurrentIndex(idx - 1)

        QMessageBox.information(self, "Preset Saved", f'Saved selection to "{name}"')


    def load_preset(self):
        if not self.is_selection_mode():
            return
        idx = self.preset_combo.currentIndex() + 1
        key = self._preset_key(idx)

        presets = self.cfg.get("presets", {})
        items = presets.get(key, [])
        if not items:
            QMessageBox.information(self, "Empty Preset", f"Preset {idx} is empty.\nSave a selection first.")
            return

        new_set = set(items)
        old_set = set(self.selected_scene_vars)

        to_turn_off = old_set - new_set
        to_turn_on = new_set - old_set

        self._begin_batch_selection()
        try:
            # turn off removed
            for var_name in to_turn_off:
                self._set_var_selected(var_name, False)
            # turn on added
            for var_name in to_turn_on:
                self._set_var_selected(var_name, True)

            # update the internal set exactly
            self.selected_scene_vars = new_set
        finally:
            self._end_batch_selection()

        self.cfg["last_preset"] = idx
        save_config(self.cfg)

        QMessageBox.information(self, "Preset Loaded", f"Loaded Preset {idx}")


    # ======================
    # Help
    # ======================
    def open_help(self):
        HelpDialog(self).exec()

    # ======================
    # UI helpers
    # ======================

    

    
    def eventFilter(self, obj, event):
        if obj == self.scroll.viewport():
            et = event.type()

            # keep your resize behavior
            if et == QEvent.Resize:
                self._resize_timer.start()
                return False

            # --- drag-to-scroll (like phone) ---
            if et == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                # don't steal clicks from interactive widgets (cards/buttons etc.)
                child = obj.childAt(event.pos())
                if child and (isinstance(child, QPushButton) or isinstance(child, QCheckBox) or isinstance(child, QLineEdit) or isinstance(child, QComboBox)):
                    return False

                self._drag_scroll_active = True
                self._drag_scroll_start_pos = event.globalPosition().toPoint()
                self._drag_scroll_start_value = self.scroll.verticalScrollBar().value()
                obj.setCursor(Qt.ClosedHandCursor)
                event.accept()
                return True

            if et == QEvent.MouseMove and self._drag_scroll_active:
                cur = event.globalPosition().toPoint()
                dy = cur.y() - self._drag_scroll_start_pos.y()
                # invert so drag down scrolls down (natural touch feel)
                self.scroll.verticalScrollBar().setValue(self._drag_scroll_start_value - dy)
                event.accept()
                return True

            if et in (QEvent.MouseButtonRelease, QEvent.Leave) and self._drag_scroll_active:
                self._drag_scroll_active = False
                self._drag_scroll_start_pos = None
                obj.unsetCursor()
                event.accept()
                return True

        return super().eventFilter(obj, event)


    def _calc_max_cols(self) -> int:
        card_w = 220
        spacing = self.scene_grid.spacing()
        avail = max(1, self.scroll.viewport().width() - 20)
        cols = (avail + spacing) // (card_w + spacing)
        return max(1, int(cols))

    def is_selection_mode(self) -> bool:
        return self.chk_select_mode.isChecked()

    def get_vam_exe_path(self) -> Path | None:
        if not self.vam_dir:
            return None
        exe = self.vam_dir / "VaM.exe"
        return exe if exe.exists() else None

    def get_vam_updater_path(self) -> Path | None:
        if not self.vam_dir:
            return None
        return self.vam_dir / "VaM_Updater.exe"
    
    def find_running_vd_streamer_path(self) -> Path | None:
        try:
            cmd = [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-Command",
                "(Get-Process -Name 'VirtualDesktop.Streamer' -ErrorAction SilentlyContinue | "
                "Select-Object -First 1 -ExpandProperty Path)"
            ]
            out = subprocess.check_output(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
            p = out.decode("utf-8", errors="ignore").strip().strip('"')
            if not p:
                return None
            pp = Path(p)
            return pp if pp.exists() and pp.is_file() else None
        except Exception:
            return None



    def _is_process_running_windows(self, image_name: str) -> bool:
        try:
            out = subprocess.check_output(
                ["tasklist", "/FI", f"IMAGENAME eq {image_name}"],
                creationflags=subprocess.CREATE_NO_WINDOW
            ).decode(errors="ignore").lower()
            return image_name.lower() in out and "no tasks are running" not in out
        except Exception:
            return False
        
    def is_vd_streamer_running(self) -> bool:
        """
        Robust check using PowerShell Get-Process.
        Looks for VirtualDesktop.Streamer process by name (no .exe).
        """
        try:
            cmd = [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-Command",
                "if (Get-Process -Name 'VirtualDesktop.Streamer' -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"
            ]
            subprocess.check_call(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
            return True
        except Exception:
            return False

        
    def _var_enabled_path(self, var_name: str) -> Path:
        assert self.addon_dir is not None
        return self.addon_dir / var_name

    def _var_disabled_path(self, var_name: str) -> Path:
        assert self.addon_dir is not None
        return self.addon_dir / (var_name + DISABLED_SUFFIX)

    def get_var_existing_path(self, var_name: str) -> Path | None:
        """
        Return actual file path that exists on disk:
        - AddonPackages/<var_name> if exists
        - else AddonPackages/<var_name>.disabled if exists
        - else None
        """
        if not self.addon_dir:
            return None
        p1 = self._var_enabled_path(var_name)
        if p1.exists():
            return p1
        p2 = self._var_disabled_path(var_name)
        if p2.exists():
            return p2
        return None

    def list_var_state_map(self) -> dict[str, str]:
        """
        Returns mapping: { "OrigName.var": "enabled" | "disabled" }
        It normalizes both *.var and *.var.disabled to orig var name.
        """
        out: dict[str, str] = {}
        if not self.addon_dir:
            return out

        # enabled
        for p in self.addon_dir.glob("*.var"):
            out[p.name] = "enabled"

        # disabled
        for p in self.addon_dir.glob("*.var.disabled"):
            orig = p.name[:-len(DISABLED_SUFFIX)]
            # if both exist, treat enabled as truth (prefer enabled)
            if orig not in out:
                out[orig] = "disabled"

        return out

    def all_var_names_catalog(self) -> set[str]:
        """
        Catalog of VAR names we know about from scene scan UI.
        This does NOT depend on current file extension (.var vs .disabled).
        """
        return set(self.card_by_var.keys())


    def is_vam_running(self) -> bool:
        return self._is_process_running_windows("VaM.exe")

    # ======================
    # Update checker (GitHub)
    # ======================
    def _parse_version_tuple(self, v: str) -> tuple[int, int, int]:
        m = re.search(r"(\d+)\.(\d+)\.(\d+)", v or "")
        if not m:
            return (0, 0, 0)
        return (int(m.group(1)), int(m.group(2)), int(m.group(3)))

    def _is_newer(self, latest: str, current: str) -> bool:
        return self._parse_version_tuple(latest) > self._parse_version_tuple(current)

    def _http_get_json(self, url: str, timeout: int = 10) -> dict:
        req = urllib.request.Request(url, headers={"User-Agent": f"VSVM/{APP_VERSION}"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read().decode("utf-8", errors="ignore")
        return json.loads(data)

    def _download_file(self, url: str, dst: Path, timeout: int = 45):
        req = urllib.request.Request(url, headers={"User-Agent": f"VSVM/{APP_VERSION}"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            dst.write_bytes(resp.read())

    def _current_exe_path(self) -> Path:
        return Path(sys.executable).resolve()

    def _find_updater_exe(self) -> Path | None:
        p = self._current_exe_path().parent / UPDATER_EXE_NAME
        return p if p.exists() else None

    def check_update_clicked(self):
        try:
            if self._startup_in_progress or self._refresh_in_progress:
                QMessageBox.information(self, "Busy", "Please wait for current loading/refresh to finish.")
                return

            self.loading.start("Checking Update", "Contacting GitHub...")

            info = self._http_get_json(GITHUB_LATEST_API)
            tag = str(info.get("tag_name") or "").strip()
            if not tag:
                self.loading.stop()
                QMessageBox.warning(self, "Update", "Could not read latest release tag.")
                return

            latest_ver = tag.lstrip("vV")
            if not self._is_newer(latest_ver, APP_VERSION):
                self.loading.stop()
                QMessageBox.information(
                    self, "Update",
                    f"You are up to date.\n\nCurrent: {APP_VERSION}\nLatest: {latest_ver}"
                )
                return

            assets = info.get("assets", [])
            dl_url = None
            for a in assets:
                if str(a.get("name")) == RELEASE_ASSET_NAME:
                    dl_url = a.get("browser_download_url")
                    break

            self.loading.stop()

            if not dl_url:
                QMessageBox.warning(
                    self, "Update",
                    f"New version found ({latest_ver}) but asset not found.\n\n"
                    f"Expected asset name:\n{RELEASE_ASSET_NAME}"
                )
                return

            reply = QMessageBox.question(
                self, "Update Available",
                f"New version is available!\n\nCurrent: {APP_VERSION}\nLatest: {latest_ver}\n\nInstall now?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

            self._download_and_install_update(latest_ver, dl_url)

        except Exception as e:
            self.loading.stop()
            QMessageBox.warning(self, "Update", f"Update check failed:\n{e}")

    def _download_and_install_update(self, latest_ver: str, dl_url: str):
        updater = self._find_updater_exe()
        if not updater:
            QMessageBox.warning(
                self, "Updater missing",
                f"Update available ({latest_ver}) but {UPDATER_EXE_NAME} was not found.\n\n"
                f"Place {UPDATER_EXE_NAME} next to:\n{self._current_exe_path().name}"
            )
            return

        if not getattr(sys, "frozen", False):
            QMessageBox.warning(
                self, "Not supported",
                "Update is only supported when running the built .exe.\n"
                "Please run the packaged app and try again."
            )
            return

        self.loading.start("Updating", "Downloading new version...")

        try:
            tmp_dir = Path(tempfile.gettempdir()) / "VSVM_updates"
            tmp_dir.mkdir(parents=True, exist_ok=True)

            new_exe = tmp_dir / f"VSVM_{latest_ver}.exe"
            self._download_file(dl_url, new_exe)

            self.loading.stop()

            current_exe = self._current_exe_path()
            subprocess.Popen([str(updater), str(current_exe), str(new_exe)], cwd=str(updater.parent))
            QApplication.quit()

        except Exception as e:
            self.loading.stop()
            QMessageBox.critical(self, "Update failed", f"Failed to download/install:\n{e}")

    # ======================
    # Folder selection & validation
    # ======================
    def last_vam_dir(self) -> Path | None:
        p = self.cfg.get("last_vam_path")
        if not p:
            return None
        pp = Path(p)
        return pp if pp.exists() else None

    def refresh_refresh_button(self):
        enabled = (self.vam_dir is not None) or (self.last_vam_dir() is not None)
        if self._startup_in_progress or self._refresh_in_progress:
            enabled = False
        self.btn_refresh.setEnabled(enabled)
        self.btn_check_update.setEnabled(not (self._startup_in_progress or self._refresh_in_progress))

    def validate_vam_folder(self, folder: Path) -> tuple[bool, str]:
        if not folder.exists() or not folder.is_dir():
            return False, "Folder does not exist."

        vam_exe = folder / "VaM.exe"
        if not vam_exe.exists():
            return False, "VaM.exe not found in this folder.\nPlease select the VaM directory where VaM.exe is located."

        addon = folder / "AddonPackages"
        if not addon.exists() or not addon.is_dir():
            return False, "AddonPackages folder not found inside this VaM directory."

        var_count = len(list(addon.glob("*.var")))
        if var_count == 0:
            return False, "No .var files found in AddonPackages.\nMake sure this is the correct VaM folder."

        return True, ""

    def _load_scene_cache(self) -> dict:
        p = cache_path()
        if not p.exists():
            return {}
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_scene_cache(self, cache_obj: dict):
        try:
            cache_path().write_text(json.dumps(cache_obj, indent=2), encoding="utf-8")
        except Exception:
            pass

    # ======================
    # Refresh behavior
    # ======================
    def refresh_clicked(self):
        target = self.vam_dir or self.last_vam_dir()
        if not target:
            QMessageBox.information(self, "No folder", "No VaM folder selected yet.\nClick 'Select VaM Directory' first.")
            return

        ok, msg = self.validate_vam_folder(target)
        if not ok:
            QMessageBox.warning(self, "Invalid folder", msg)
            return

        if self.vam_dir is None:
            self.vam_dir = target
            self.addon_dir = target / "AddonPackages"
            self.saves_scene_dir = target / "Saves" / "scene"

        self._refresh_in_progress = True
        self.refresh_refresh_button()
        self.loading.start("Refreshing", "Checking for changes...")

        cache_obj = self._load_scene_cache()
        if self._can_use_cache_for_current_folder(cache_obj):
            if not self.cards:
                self.looks_map = cache_obj.get("looks", {}) if isinstance(cache_obj.get("looks"), dict) else {}
                self.scene_entries = self._scene_entries_from_cache(cache_obj)
                self.total_scene_files_found = len(self.scene_entries)
                try:
                    self.total_vars_count = len(list(self.addon_dir.glob("*.var"))) if self.addon_dir else 0
                except Exception:
                    self.total_vars_count = 0
                self._recompute_unused_count_from_scene_entries()

                self.status.setText(f"VaM Directory:\n{self.vam_dir}\n(Loaded from cache)")
                self.info_label.setText(
                    f"Total VARs: {self.total_vars_count} | Unused VARs: {self.unused_vars_count} | Scenes: {self.total_scene_files_found}"
                )

                self.search.setEnabled(True)
                self.chk_girl_looks_only.setEnabled(True)
                self.chk_select_mode.setEnabled(True)

                self.populate_scene_cards_from_entries()
                self.apply_filter(self.search.text())
                self.update_selection_ui()

        self._start_change_check(cache_obj)

    # ======================
    # Startup cache-first
    # ======================
    def auto_open_last_folder_on_startup(self):
        last = self.last_vam_dir()
        if not last:
            self.status.setText("Select VaM directory (where VaM.exe is)")
            self._end_busy()
            return

        ok, msg = self.validate_vam_folder(last)
        if not ok:
            self.status.setText("Select VaM directory (where VaM.exe is)")
            self._end_busy()
            return

        self.vam_dir = last
        self.addon_dir = last / "AddonPackages"
        self.saves_scene_dir = last / "Saves" / "scene"

        self.refresh_refresh_button()
        self.refresh_restore_button()
        self.refresh_apply_button()

        cache_obj = self._load_scene_cache()
        if self._can_use_cache_for_current_folder(cache_obj):
            self.looks_map = cache_obj.get("looks", {}) if isinstance(cache_obj.get("looks"), dict) else {}

            self.scene_entries = self._scene_entries_from_cache(cache_obj)
            self.total_scene_files_found = len(self.scene_entries)

            try:
                self.total_vars_count = len(list(self.addon_dir.glob("*.var"))) if self.addon_dir else 0
            except Exception:
                self.total_vars_count = 0

            self._recompute_unused_count_from_scene_entries()

            self.status.setText(f"VaM Directory:\n{self.vam_dir}\n(Loaded from cache. Checking updates...)")
            self.info_label.setText(
                f"Total VARs: {self.total_vars_count} | Unused VARs: {self.unused_vars_count} | Scenes: {self.total_scene_files_found}"
            )

            self.search.setEnabled(True)
            self.chk_girl_looks_only.setEnabled(True)
            self.chk_select_mode.setEnabled(True)

            self.populate_scene_cards_from_entries()
            self.apply_filter(self.search.text())
            self.update_selection_ui()

            self.loading.stop()
            self._start_change_check(cache_obj)
            return

        self.set_current_vam_dir(last)

    def _can_use_cache_for_current_folder(self, cache_obj: dict) -> bool:
        if not isinstance(cache_obj, dict):
            return False
        if cache_obj.get("version") != 3:
            return False
        if not self.addon_dir:
            return False
        if str(cache_obj.get("addon_dir", "")) != str(self.addon_dir):
            return False
        vars_map = cache_obj.get("vars")
        if not isinstance(vars_map, dict):
            return False
        return True

    def _scene_entries_from_cache(self, cache_obj: dict) -> list[dict]:
        out: list[dict] = []
        vars_map = cache_obj.get("vars", {})
        loose_map = cache_obj.get("loose", {})

        if isinstance(vars_map, dict):
            for var_name, vinfo in vars_map.items():
                if not isinstance(vinfo, dict):
                    continue
                scenes = vinfo.get("scenes", [])
                if not isinstance(scenes, list):
                    continue
                for s in scenes:
                    if not isinstance(s, dict):
                        continue
                    out.append({
                        "source": "var",
                        "scene_name": s.get("scene_name", ""),
                        "var_name": var_name,
                        "preview_relpath": s.get("preview_relpath", ""),
                        "loose_relpath": "",
                        "is_girl_looks": None,
                    })

        if isinstance(loose_map, dict):
            for relp, linfo in loose_map.items():
                if not isinstance(linfo, dict):
                    continue
                scene_name = linfo.get("scene_name") or relp
                out.append({
                    "source": "loose",
                    "scene_name": scene_name,
                    "var_name": "(Saves/scene)",
                    "preview_relpath": "",
                    "loose_relpath": relp,
                    "is_girl_looks": None,
                })

        out.sort(key=lambda e: (str(e.get("source", "")), str(e.get("var_name", "")), str(e.get("scene_name", "")).lower()))
        return out

    def _start_change_check(self, cache_obj: dict):
        if not self.addon_dir:
            self._end_busy()
            return
        self.change_worker = ChangeCheckWorker(self.addon_dir, self.saves_scene_dir, cache_obj)
        self.change_worker.finished.connect(self._on_change_check_done)
        self.change_worker.start()

    def _on_change_check_done(self, needs_rescan: bool):
        if not self.vam_dir or not self.addon_dir:
            self._end_busy()
            return

        if not needs_rescan:
            self.status.setText(f"VaM Directory:\n{self.vam_dir}")
            self._end_busy()
            return

        self.status.setText("Changes detected. Updating cache...")
        self.loading.start("Updating", "Scanning changes...")

        old_cache = self._load_scene_cache()
        self.looks_map = old_cache.get("looks", {}) if isinstance(old_cache, dict) else {}

        self.worker = AnalyzeWorker(self.addon_dir, self.saves_scene_dir, old_cache)
        self.worker.finished.connect(self.analysis_done)
        self.worker.start()

    def _recompute_unused_count_from_scene_entries(self):
        self.unused_vars_count = 0
        if not self.addon_dir:
            return
        try:
            all_vars = {p.name for p in self.addon_dir.glob("*.var")}
            scene_vars = {e["var_name"] for e in self.scene_entries if e.get("source") == "var" and e.get("var_name")}
            keep_all_scenes = self.compute_keep_set_for_scene_vars(scene_vars)
            unused = all_vars - keep_all_scenes
            self.unused_vars_count = len(unused)
        except Exception:
            self.unused_vars_count = 0

    # ======================
    # Folder selection
    # ======================
    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select VaM Directory (where VaM.exe is)")
        if not folder:
            return
        self.set_current_vam_dir(Path(folder))

    def set_current_vam_dir(self, folder: Path):
        ok, msg = self.validate_vam_folder(folder)
        if not ok:
            QMessageBox.warning(self, "Invalid folder", msg)
            self._end_busy()
            return

        self.vam_dir = folder
        self.addon_dir = folder / "AddonPackages"
        self.saves_scene_dir = folder / "Saves" / "scene"

        self.cfg["last_vam_path"] = str(folder)
        save_config(self.cfg)

        self.refresh_refresh_button()
        self.refresh_restore_button()
        self.refresh_apply_button()

        self.scene_entries = []
        self.total_scene_files_found = 0

        self.loading.start("Loading", "Scanning scenes...")

        self.status.setText("Scanning scenes...")
        self.info_label.setText("Total VARs: - | Unused VARs: - | Scenes: -")
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        self.clear_scene_cards()
        self.clear_dependencies_panel()

        self.search.setEnabled(False)
        self.search.setText("")
        self.chk_girl_looks_only.setEnabled(False)
        self.chk_girl_looks_only.setChecked(False)

        self.selected_scene_vars.clear()
        self.update_selection_ui()

        self.chk_select_mode.setEnabled(False)
        self.chk_select_mode.setChecked(False)

        old_cache = self._load_scene_cache()
        self.looks_map = old_cache.get("looks", {}) if isinstance(old_cache, dict) else {}

        self.worker = AnalyzeWorker(self.addon_dir, self.saves_scene_dir, old_cache)
        self.worker.finished.connect(self.analysis_done)
        self.worker.start()

    def analysis_done(self, scene_entries: list[dict], total_var_count: int, cache_obj: dict):
        self.progress.setVisible(False)

        if isinstance(cache_obj, dict):
            cache_obj["looks"] = self.looks_map
            self._save_scene_cache(cache_obj)

        self.scene_entries = scene_entries
        self.total_vars_count = total_var_count
        self.total_scene_files_found = len(self.scene_entries)

        self._recompute_unused_count_from_scene_entries()

        self.status.setText(f"VaM Directory:\n{self.vam_dir}")
        self.info_label.setText(
            f"Total VARs: {self.total_vars_count} | Unused VARs: {self.unused_vars_count} | Scenes: {self.total_scene_files_found}"
        )

        self.search.setEnabled(True)
        self.chk_girl_looks_only.setEnabled(True)
        self.chk_select_mode.setEnabled(True)

        self.populate_scene_cards_from_entries()
        self.apply_filter(self.search.text())
        self.update_selection_ui()

        self.refresh_restore_button()
        self.refresh_apply_button()

        self._end_busy()

    # ======================
    # Cards
    # ======================

    def _start_lazy_preview_loader(self):
        if hasattr(self, "_preview_timer") and self._preview_timer.isActive():
            self._preview_timer.stop()

        self._preview_queue = [
            c for c in self.cards
            if getattr(c, "_preview_rel", "") and not getattr(c, "_preview_loaded", False)
        ]

        self._preview_timer = QTimer(self)
        self._preview_timer.setInterval(1)  # cepat tapi non-blocking
        self._preview_timer.timeout.connect(self._lazy_preview_tick)
        self._preview_timer.start()


    def _lazy_preview_tick(self):
        # load max 8 preview per tick (AMAN)
        for _ in range(8):
            if not self._preview_queue:
                self._preview_timer.stop()
                return

            c = self._preview_queue.pop(0)
            rel = getattr(c, "_preview_rel", "")
            if not rel:
                continue

            img = read_preview_bytes(rel)
            if img:
                try:
                    c.set_preview_image_bytes(img)
                except Exception:
                    pass

            c._preview_loaded = True

    
    def populate_scene_cards_from_entries(self):
        self.cards = []
        self.card_by_var = {}

        entries_sorted = sorted(
            self.scene_entries,
            key=lambda e: (str(e.get("source", "")), str(e.get("var_name", "")), str(e.get("scene_name", "")).lower())
        )

        for e in entries_sorted:
            source = e.get("source", "var")
            scene_name = e.get("scene_name", "")
            var_name = e.get("var_name", "")
            rel = e.get("preview_relpath", "") or ""

            card = SceneCard(scene_name=scene_name, var_name=var_name, image_bytes=None)

            # simpan rel path untuk lazy load
            card._preview_rel = rel
            card._preview_loaded = False

            card.clicked.connect(self.on_card_clicked)

            selectable = (source == "var")
            card.set_selection_mode(self.is_selection_mode() and selectable)

            if selectable:
                card.set_checked(card.var_name in self.selected_scene_vars)
                self.card_by_var.setdefault(card.var_name, []).append(card)
            else:
                card.set_checked(False)

            k = f"{var_name}::{scene_name}"
            looks_val = self.looks_map.get(k, None)
            card.is_girl_looks = looks_val  # type: ignore[attr-defined]

            card.source = source  # type: ignore[attr-defined]
            card.loose_relpath = e.get("loose_relpath", "")  # type: ignore[attr-defined]

            self.cards.append(card)

        self.rebuild_grid(self.cards)
        self._start_lazy_preview_loader()


    def rebuild_grid(self, cards_to_show: list[SceneCard]):
        while self.scene_grid.count():
            item = self.scene_grid.takeAt(0)
            w = item.widget()
            if w:
                w.setParent(self.grid_holder)

        col = 0
        row = 0
        max_cols = self._calc_max_cols()

        for card in cards_to_show:
            self.scene_grid.addWidget(card, row, col)
            col += 1
            if col >= max_cols:
                col = 0
                row += 1

    def clear_scene_cards(self):
        self.cards = []
        self.card_by_var = {}
        while self.scene_grid.count():
            item = self.scene_grid.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def apply_filter(self, text: str):
        q = (text or "").strip().lower()
        looks_only = self.chk_girl_looks_only.isChecked()

        selected_matches: list[SceneCard] = []
        other_matches: list[SceneCard] = []

        for card in self.cards:
            scene_name = card.scene_name
            var_name = card.var_name

            is_match = (not q) or (q in scene_name.lower()) or (q in var_name.lower())
            if not is_match:
                card.setVisible(False)
                continue

            if looks_only:
                src = getattr(card, "source", "var")
                if src != "var":
                    card.setVisible(False)
                    continue

                v = getattr(card, "is_girl_looks", None)
                if v is None:
                    card.setVisible(False)
                    continue
                if not bool(v):
                    card.setVisible(False)
                    continue

            card.setVisible(True)

            if var_name in self.selected_scene_vars:
                selected_matches.append(card)
            else:
                other_matches.append(card)

        self.rebuild_grid(selected_matches + other_matches)

    # ======================
    # Looks-only
    # ======================
    def on_toggle_looks_only(self, _state: int):
        if not self.chk_girl_looks_only.isChecked():
            self.apply_filter(self.search.text())
            return

        if not self.addon_dir:
            return

        pairs: list[tuple[str, str]] = []
        for c in self.cards:
            if getattr(c, "source", "var") != "var":
                continue
            k = f"{c.var_name}::{c.scene_name}"
            if k not in self.looks_map:
                pairs.append((c.var_name, c.scene_name))

        if not pairs:
            for c in self.cards:
                if getattr(c, "source", "var") != "var":
                    continue
                k = f"{c.var_name}::{c.scene_name}"
                c.is_girl_looks = self.looks_map.get(k, False)  # type: ignore[attr-defined]
            self.apply_filter(self.search.text())
            return

        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.status.setText("Detecting Girl 'Looks' scenes...")

        self.chk_girl_looks_only.setEnabled(False)

        self.looks_worker = LooksWorker(self.addon_dir, pairs, self.looks_map)
        self.looks_worker.finished.connect(self._looks_done)
        self.looks_worker.start()

    def _looks_done(self, looks_map: dict):
        self.progress.setVisible(False)
        self.status.setText(f"VaM Directory:\n{self.vam_dir}")

        if isinstance(looks_map, dict):
            self.looks_map = looks_map

            for c in self.cards:
                if getattr(c, "source", "var") != "var":
                    continue
                k = f"{c.var_name}::{c.scene_name}"
                c.is_girl_looks = self.looks_map.get(k, False)  # type: ignore[attr-defined]

            old = self._load_scene_cache()
            if isinstance(old, dict):
                old["looks"] = self.looks_map
                self._save_scene_cache(old)

        self.chk_girl_looks_only.setEnabled(True)
        self.apply_filter(self.search.text())

    # ======================
    # Selection mode
    # ======================
    def on_toggle_selection_mode(self, _state: int):
        enabled = self.is_selection_mode()

        self.scene_container.setUpdatesEnabled(False)
        try:
            for card in self.cards:
                src = getattr(card, "source", "var")
                selectable = (src == "var")
                card.set_selection_mode(enabled and selectable)
                if selectable:
                    card.set_checked(card.var_name in self.selected_scene_vars)
                else:
                    card.set_checked(False)
        finally:
            self.scene_container.setUpdatesEnabled(True)

        self.apply_filter(self.search.text())
        self.update_selection_ui()


    def _check_selection_dirty(self):
        # Only care while VaM is running OR lean session is active
        if not (self.is_vam_running() or self.lean_active):
            self.selection_dirty = False
            self.set_apply_attention(False)
            return

        current = set(self._scene_vars_for_launch())
        dirty = (current != set(self.session_baseline_scene_vars))

        if dirty != self.selection_dirty:
            self.selection_dirty = dirty
            self.set_apply_attention(dirty)


    def _set_var_selected(self, var_name: str, selected: bool):
        if self._syncing_selection_ui:
            return

        # short-circuit if nothing changes (saves tons of work)
        already = (var_name in self.selected_scene_vars)
        if selected == already:
            return

        self._syncing_selection_ui = True
        try:
            if selected:
                self.selected_scene_vars.add(var_name)
            else:
                self.selected_scene_vars.discard(var_name)

            # update card checkmarks (cheap)
            for c in self.card_by_var.get(var_name, []):
                c.set_checked(selected)
        finally:
            self._syncing_selection_ui = False

        # IMPORTANT: during batch, don't rebuild grid / update labels / style every click
        if self._batch_selection:
            return

        self.apply_filter(self.search.text())
        self.update_selection_ui()
        self._schedule_check_selection_dirty()



    def update_selection_ui(self):
        self.selection_info.setText(
            f"Total Scene: {self.total_scene_files_found} scenes | Selected Scene: {len(self.selected_scene_vars)} scenes"
        )

        enabled_card_tools = self.is_selection_mode() and (self.addon_dir is not None) and (len(self.cards) > 0)
        self.btn_select_all.setEnabled(enabled_card_tools)
        self.btn_clear_sel.setEnabled(enabled_card_tools)

        self.btn_save_preset.setEnabled(enabled_card_tools)
        self.btn_load_preset.setEnabled(enabled_card_tools)
        self.preset_combo.setEnabled(enabled_card_tools)

    def select_all_visible(self):
        if not self.cards:
            return

        self._begin_batch_selection()
        try:
            for var_name, cards in self.card_by_var.items():
                if any(c.isVisible() for c in cards):
                    self._set_var_selected(var_name, True)
        finally:
            self._end_batch_selection()


    def clear_selection(self):
        if not self.selected_scene_vars:
            return

        self._begin_batch_selection()
        try:
            for var_name in list(self.selected_scene_vars):
                self._set_var_selected(var_name, False)
            self.selected_scene_vars.clear()
        finally:
            self._end_batch_selection()


    def on_card_clicked(self, scene_name: str, var_name: str):
        clicked_card = None
        for c in self.cards:
            if c.scene_name == scene_name and c.var_name == var_name:
                clicked_card = c
                break
        if clicked_card is None:
            return

        src = getattr(clicked_card, "source", "var")

        if src == "var":
            if not self.addon_dir:
                return
            if self.is_selection_mode():
                self._set_var_selected(var_name, not (var_name in self.selected_scene_vars))
            self.show_dependencies(scene_name, var_name)
            return

        self.show_loose_scene_info(clicked_card)

    # ======================
    # Dependencies panel
    # ======================
    def clear_dependencies_panel(self):
        self.dep_selected.setText("Click a card to see dependencies.")
        self.dep_summary.setText("")
        while self.dep_list_layout.count():
            item = self.dep_list_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

    def show_loose_scene_info(self, card: SceneCard):
        relp = getattr(card, "loose_relpath", "")
        self.dep_selected.setText(f"Selected (Saves/scene):\n{relp}")
        self.dep_summary.setText("Loose scene: dependency graph not available in this app.")
        while self.dep_list_layout.count():
            item = self.dep_list_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.dep_list_layout.addWidget(QLabel("—"))
        self.dep_list_layout.addStretch(1)

    def show_dependencies(self, scene_name: str, var_name: str):
        if not self.addon_dir:
            return

        self.dep_selected.setText(f"Selected:\n{scene_name}\n{var_name}")

        var_path = self.get_var_existing_path(var_name)
        if not var_path:
            # file bisa saja hilang benar-benar (user delete manual), jangan crash
            self.dep_summary.setText("VAR file not found on disk (enabled/disabled).")
            while self.dep_list_layout.count():
                item = self.dep_list_layout.takeAt(0)
                w = item.widget()
                if w:
                    w.deleteLater()
            self.dep_list_layout.addWidget(QLabel("(missing on disk)"))
            self.dep_list_layout.addStretch(1)
            return

        info = scan_var(var_path)

        deps = sorted(info.get("dependencies", []))

        # IMPORTANT: all vars should include disabled too
        all_vars = set(self.list_var_state_map().keys())

        while self.dep_list_layout.count():
            item = self.dep_list_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        if not deps:
            self.dep_summary.setText("No dependencies found in meta.json")
            self.dep_list_layout.addWidget(QLabel("—"))
            return

        present_count = 0
        missing_count = 0

        for dep in deps:
            matched = resolve_dependency(dep, all_vars)
            present = len(matched) > 0

            if present:
                present_count += 1
                chosen = sorted(matched)[-1]
                # show if chosen currently disabled or enabled
                state = self.list_var_state_map().get(chosen, "")
                suffix = " (disabled)" if state == "disabled" else ""
                text = f"{dep}  →  {chosen}{suffix}"
            else:
                missing_count += 1
                text = f"{dep}  →  (missing)"

            self.dep_list_layout.addWidget(DependencyRow(text=text, present=present))

        self.dep_list_layout.addStretch(1)
        self.dep_summary.setText(f"Total: {len(deps)} | Present: {present_count} | Missing: {missing_count}")


    # ======================
    # Restore / Apply / Launch (RENAME METHOD)
    # ======================
    def refresh_restore_button(self):
        if not self.vam_dir:
            self.btn_restore.setEnabled(False)
            return
        mp = manifest_path_for(self.vam_dir)
        self.btn_restore.setEnabled(mp.exists())

    def refresh_apply_button(self):
        """
        Apply button becomes useful if:
        - VaM is running (user wants live apply), OR
        - a manifest exists (lean session active), OR
        - we have a folder selected (we can start live mode even if VaM already running)
        """
        if not self.vam_dir:
            self.btn_apply_now.setEnabled(False)
            return
        mp_exists = manifest_path_for(self.vam_dir).exists()
        running = self.is_vam_running()
        self.btn_apply_now.setEnabled(running or mp_exists)

    def _read_manifest(self) -> dict:
        if not self.vam_dir:
            return {}
        mp = manifest_path_for(self.vam_dir)
        if not mp.exists():
            return {}
        try:
            return json.loads(mp.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _write_manifest(self, manifest: dict):
        if not self.vam_dir:
            return
        mp = manifest_path_for(self.vam_dir)
        try:
            mp.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        except Exception:
            pass

    def restore_offloaded_vars(self):
        if not self.vam_dir:
            QMessageBox.warning(self, "No folder", "Select VaM directory first.")
            return

        addon_dir = self.vam_dir / "AddonPackages"
        mp = manifest_path_for(self.vam_dir)

        if not mp.exists():
            QMessageBox.information(self, "Nothing to restore", "No manifest found.")
            self.refresh_restore_button()
            self.refresh_apply_button()
            return

        try:
            manifest = json.loads(mp.read_text(encoding="utf-8"))
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to read manifest:\n{e}")
            return

        renamed = manifest.get("renamed", [])
        restored = 0

        if isinstance(renamed, list):
            for item in renamed:
                if not isinstance(item, dict):
                    continue
                src_name = item.get("to")
                dst_name = item.get("from")
                if not src_name or not dst_name:
                    continue

                src = addon_dir / src_name
                dst = addon_dir / dst_name

                if src.exists() and not dst.exists():
                    try:
                        src.rename(dst)
                        restored += 1
                    except Exception:
                        pass

        try:
            mp.unlink(missing_ok=True)
        except Exception:
            pass

        self.lean_active = False
        self.vam_seen_running = False
        self.poll_timer.stop()

        self.session_baseline_scene_vars = set()
        self.selection_dirty = False
        self.set_apply_attention(False)

        QMessageBox.information(self, "Restored", f"Restored {restored} VARs back to normal in:\n{addon_dir}")
        self.refresh_restore_button()
        self.refresh_apply_button()

    def _scene_vars_for_launch(self) -> set[str]:
        if self.selected_scene_vars:
            return set(self.selected_scene_vars)
        return set(self.card_by_var.keys())

    def compute_keep_set_for_scene_vars(self, scene_vars: set[str]) -> set[str]:
        assert self.addon_dir is not None

        # ✅ IMPORTANT: dependency universe must be ALL vars on disk (enabled + disabled)
        all_vars = self.all_var_names_on_disk()

        # Protect assets/plugins by name (same behavior as before, but now sees everything)
        protected = {v for v in all_vars if is_asset_var(v)}
        protected |= {v for v in all_vars if "[plugin]" in v.lower()}

        keep = set(protected)
        queue: list[str] = []

        # Seed selected scene vars
        for scene_var in scene_vars:
            if scene_var in all_vars:
                keep.add(scene_var)
                p = self.get_var_existing_path(scene_var)
                if p:
                    info = scan_var(p)
                    queue.extend(list(info.get("dependencies", [])))

        # Resolve dependencies recursively
        while queue:
            dep = queue.pop()
            for matched_var in resolve_dependency(dep, all_vars):
                if matched_var not in keep:
                    keep.add(matched_var)
                    p = self.get_var_existing_path(matched_var)
                    if p:
                        info = scan_var(p)
                        queue.extend(list(info.get("dependencies", [])))

        return keep



    def disable_unrelated_vars_by_rename(self, keep_set: set[str]) -> dict:
        """
        Initial disable pass (used when starting a session).
        Works with both enabled and disabled state.
        Only disables currently enabled .var that are not in keep_set.
        """
        assert self.vam_dir is not None
        addon_dir = self.vam_dir / "AddonPackages"
        mp = manifest_path_for(self.vam_dir)

        renamed: list[dict] = []

        # Use state map so we don't miss vars that exist but currently disabled
        state_map = self.list_var_state_map()

        for orig_name, state in state_map.items():
            if orig_name in keep_set:
                continue

            # Only disable if currently enabled
            if state != "enabled":
                continue

            src = addon_dir / orig_name
            dst = addon_dir / (orig_name + DISABLED_SUFFIX)

            if not src.exists():
                continue
            if dst.exists():
                continue

            try:
                src.rename(dst)
                renamed.append({"from": orig_name, "to": dst.name})
            except Exception:
                continue

        manifest = {
            "addon_dir": str(addon_dir),
            "method": "rename_disabled",
            "disabled_suffix": DISABLED_SUFFIX,
            "renamed": renamed,
            "saved_at": time.time(),
        }

        mp.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        return manifest


    def _apply_keep_set_live(self, keep_set: set[str]) -> tuple[int, int]:
        """
        LIVE apply:
        - Restore anything that should now be kept (if currently disabled by our tool)
        - Disable anything not in keep_set (but only if currently enabled)
        Returns: (disabled_count, restored_count)
        """
        assert self.vam_dir is not None
        addon_dir = self.vam_dir / "AddonPackages"
        mp = manifest_path_for(self.vam_dir)

        manifest = self._read_manifest()
        renamed_list = manifest.get("renamed", [])
        if not isinstance(renamed_list, list):
            renamed_list = []

        # tool_map: orig -> disabled_filename (only what tool is responsible for)
        tool_map: dict[str, str] = {}
        for item in renamed_list:
            if not isinstance(item, dict):
                continue
            o = item.get("from")
            d = item.get("to")
            if isinstance(o, str) and isinstance(d, str) and o and d:
                tool_map[o] = d

        disabled_count = 0
        restored_count = 0

        state_map = self.list_var_state_map()

        # 1) Restore tool-disabled vars that are now in keep_set
        for orig, dis_name in list(tool_map.items()):
            if orig not in keep_set:
                continue

            src = addon_dir / dis_name
            dst = addon_dir / orig

            if src.exists() and not dst.exists():
                try:
                    src.rename(dst)
                    restored_count += 1
                    # remove from tool_map because it's no longer disabled by tool
                    tool_map.pop(orig, None)
                except Exception:
                    pass

        # Refresh state after restores
        state_map = self.list_var_state_map()

        # 2) Disable vars NOT in keep_set (only if currently enabled)
        # IMPORTANT: Use catalog so it still works when most are disabled
        catalog = self.all_var_names_catalog()
        if not catalog:
            catalog = set(state_map.keys())

        for orig in catalog:
            if orig in keep_set:
                continue

            # Already disabled by our tool?
            if orig in tool_map:
                continue

            # Only disable if currently enabled on disk
            if state_map.get(orig) != "enabled":
                continue

            src = addon_dir / orig
            dst = addon_dir / (orig + DISABLED_SUFFIX)

            if not src.exists():
                continue
            if dst.exists():
                # don't overwrite; safest do nothing
                continue

            try:
                src.rename(dst)
                tool_map[orig] = dst.name
                disabled_count += 1
            except Exception:
                pass

        # 3) Write updated manifest from tool_map (only what tool currently disables)
        new_manifest = {
            "addon_dir": str(addon_dir),
            "method": "rename_disabled",
            "disabled_suffix": DISABLED_SUFFIX,
            "renamed": [{"from": o, "to": d} for o, d in sorted(tool_map.items())],
            "saved_at": time.time(),
        }
        try:
            mp.write_text(json.dumps(new_manifest, indent=2), encoding="utf-8")
        except Exception:
            pass

        return disabled_count, restored_count


    def apply_selection_now_clicked(self):
        """
        If VaM is running, user can click Apply anytime.
        - If no manifest yet: we can start live mode now (creates manifest).
        - Then we apply current keep_set live.
        """
        if not self.vam_dir or not self.addon_dir:
            QMessageBox.warning(self, "No folder", "Select VaM directory first.")
            return

        running = self.is_vam_running()
        mp = manifest_path_for(self.vam_dir)
        mp_exists = mp.exists()

        if not running and not mp_exists:
            QMessageBox.information(
                self, "Not available",
                "Apply Selection Now is used while VaM is running (live), or when a lean session is active.\n\n"
                "Launch VaM from this app, or open VaM and try again."
            )
            self.refresh_apply_button()
            return

        if running and not mp_exists:
            reply = QMessageBox.question(
                self,
                "Start Live Lean Mode",
                "VaM is running but no lean session is active.\n\n"
                "This will start a Live Lean Mode NOW by renaming unrelated .var files.\n"
                "After applying, you may need to use VaM's in-game refresh/rescan button.\n\n"
                "Continue?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
            
            self.session_baseline_scene_vars = set(self._scene_vars_for_launch())
            self.selection_dirty = False
            self.set_apply_attention(False)

        # Apply live changes
        try:
            scene_vars = self._scene_vars_for_launch()
            keep_set = self.compute_keep_set_for_scene_vars(scene_vars)

            self.progress.setVisible(True)
            self.progress.setRange(0, 0)
            self.status.setText("Applying selection (live)...")

            # If no manifest exists yet, do initial disable to create manifest
            if not mp_exists:
                self.disable_unrelated_vars_by_rename(keep_set)

            disabled_n, restored_n = self._apply_keep_set_live(keep_set)

            # update baseline after successful apply
            self.session_baseline_scene_vars = set(self._scene_vars_for_launch())
            self.selection_dirty = False
            self.set_apply_attention(False)


            # Mark session active and start monitoring (if VaM already running, mark as seen)
            self.lean_active = True
            self.vam_seen_running = running
            self.poll_timer.start()

            self.progress.setVisible(False)
            self.status.setText(f"VaM Directory:\n{self.vam_dir}")

            QMessageBox.information(
                self,
                "Applied",
                f"Applied selection.\n\nDisabled: {disabled_n}\nRestored: {restored_n}\n\n"
                f"Tip: In VaM, use its refresh/rescan packages button if needed."
            )

        except Exception as e:
            self.progress.setVisible(False)
            self.status.setText(f"VaM Directory:\n{self.vam_dir}")
            QMessageBox.critical(self, "Apply failed", f"Failed to apply selection:\n{e}")

        self.refresh_restore_button()
        self.refresh_apply_button()

    def _start_lean_session_or_warn(self) -> bool:
        """
        For launching: we still require VaM is NOT running.
        Live mode uses Apply Selection Now instead.
        """
        if not self.vam_dir or not self.addon_dir:
            QMessageBox.warning(self, "No folder", "Select VaM directory first.")
            return False

        if self.is_vam_running():
            QMessageBox.warning(self, "VaM is running", "Close VaM.exe first, then launch from this app.\n\n(Use Apply Selection Now for live changes.)")
            return False

        if self.lean_active:
            QMessageBox.information(
                self,
                "Lean session already active",
                "Lean session already active.\n\nStart VaM.exe (desktop/VR) and I will restore after it closes.\n(Or click Restore VARs to cancel.)"
            )
            return False

        scene_vars = self._scene_vars_for_launch()
        if not scene_vars:
            QMessageBox.information(self, "No scenes", "No packaged scenes found to launch.")
            return False

        reply = QMessageBox.question(
            self,
            "Lean Launch",
            "This will TEMPORARILY disable unrelated .var files to your current scene selection"
            "By renaming extension to .disable"
            "When VaM.exe closes, it will restore them.\n\nContinue?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return False

        try:
            keep_set = self.compute_keep_set_for_scene_vars(scene_vars)
            self.disable_unrelated_vars_by_rename(keep_set)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to disable vars:\n{e}")
            self.refresh_restore_button()
            self.refresh_apply_button()
            return False

        self.lean_active = True
        self.vam_seen_running = False
        self.poll_timer.start()
        # baseline selection for this session
        self.session_baseline_scene_vars = set(self._scene_vars_for_launch())
        self.selection_dirty = False
        self.set_apply_attention(False)

        self.refresh_restore_button()
        self.refresh_apply_button()
        return True
    
        # ======================
    # Virtual Desktop Streamer helpers
    # ======================
    def get_vd_streamer_path(self) -> Path | None:
        """
        Returns saved VirtualDesktop.Streamer.exe path if valid.
        Stored in self.cfg["vd_streamer_path"].
        """
        p = self.cfg.get("vd_streamer_path", "")
        if not p:
            return None
        pp = Path(p)
        return pp if pp.exists() and pp.is_file() else None

    def ensure_vd_streamer_path(self) -> Path | None:
        """
        Priority:
        1) If VD Streamer process is running, read its real path and save it.
        2) Else: use saved config path if valid.
        3) Else: ask user to locate VirtualDesktop.Streamer.exe (optional fallback).
        """
        # 1) If running, discover from process and save
        running_path = self.find_running_vd_streamer_path()
        if running_path:
            self.cfg["vd_streamer_path"] = str(running_path)
            save_config(self.cfg)
            return running_path

        # 2) Use saved config path
        p = self.cfg.get("vd_streamer_path", "")
        if p:
            pp = Path(p)
            if pp.exists() and pp.is_file():
                return pp

        # 3) Optional fallback: ask user (you can keep or remove this)
        picked, _ = QFileDialog.getOpenFileName(
            self,
            "Select VirtualDesktop.Streamer.exe",
            "",
            "Executable (*.exe)"
        )
        if not picked:
            return None

        pp = Path(picked)
        if not pp.exists() or not pp.is_file():
            QMessageBox.warning(self, "Invalid file", "Selected file does not exist.")
            return None

        self.cfg["vd_streamer_path"] = str(pp)
        save_config(self.cfg)
        return pp


    def launch_vam_vd_lean(self):
        # 1) basic validation
        if not self.vam_dir or not self.addon_dir:
            QMessageBox.warning(self, "No folder", "Select VaM directory first.")
            return

        # 2) NEW running gate (path-based, reliable)
        vd_streamer = self.find_running_vd_streamer_path()
        if not vd_streamer:
            QMessageBox.information(
                self,
                "Virtual Desktop not running",
                "Virtual Desktop Streamer is not running.\n\n"
                "Please start Virtual Desktop Streamer first, then try again."
            )
            return

        # 3) start lean session (NOW vars can be renamed)
        if not self._start_lean_session_or_warn():
            return

        # 4) launch VaM via VD
        vam_exe = self.get_vam_exe_path()
        if not vam_exe:
            QMessageBox.critical(self, "Error", "VaM.exe not found.")
            self.restore_offloaded_vars()
            return

        subprocess.Popen(
            [str(vd_streamer), str(vam_exe)],
            cwd=str(vd_streamer.parent),
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        self.status.setText("Launched VaM via Virtual Desktop. Monitoring until it closes...")
        self.refresh_apply_button()


    def launch_vam_exe_lean(self):
        if not self._start_lean_session_or_warn():
            return

        vam_exe = self.get_vam_exe_path()
        if not vam_exe:
            QMessageBox.critical(self, "Error", "VaM.exe not found. Please re-select VaM directory.")
            self.restore_offloaded_vars()
            return

        try:
            subprocess.Popen([str(vam_exe)], cwd=str(vam_exe.parent))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to launch VaM.exe:\n{e}\n\nRestoring now...")
            self.restore_offloaded_vars()
            return

        self.status.setText("Launched VaM.exe. Monitoring until it closes...")
        self.refresh_apply_button()

    def launch_vam_launcher_lean(self):
        if not self._start_lean_session_or_warn():
            return

        updater = self.get_vam_updater_path()
        if not updater or not updater.exists():
            QMessageBox.warning(self, "Not found", "VaM_Updater.exe not found in VaM directory.\n\nRestoring now...")
            self.restore_offloaded_vars()
            return

        try:
            subprocess.Popen([str(updater)], cwd=str(updater.parent))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to launch VaM_Updater.exe:\n{e}\n\nRestoring now...")
            self.restore_offloaded_vars()
            return

        self.status.setText("Launched VaM Launcher. Waiting for VaM.exe to start (desktop/VR)...")
        self.refresh_apply_button()

    def check_vam_state(self):

        self.refresh_apply_button()

        if not self.lean_active:
          
            return

        running = self.is_vam_running()

        if running and not self.vam_seen_running:
            self.vam_seen_running = True
            self.status.setText("VaM.exe detected running. Monitoring until it closes...")
            return

        if self.vam_seen_running and not running:
            self.status.setText("VaM.exe closed. Restoring VARs...")
            self.restore_offloaded_vars()
            self.status.setText("Restore done. You can launch again.")
            return


if __name__ == "__main__":
    app = QApplication(sys.argv)

    icon_path = resource_path("icons/app.ico")
    if icon_path.exists():
        icon = QIcon(str(icon_path))
        app.setWindowIcon(icon)

    window = MainWindow()
    if icon_path.exists():
        window.setWindowIcon(QIcon(str(icon_path)))

    window.show()
    sys.exit(app.exec())

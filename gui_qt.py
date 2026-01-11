import sys
import os
import json
import subprocess
import zipfile
import time
import hashlib
import urllib.request
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QVBoxLayout, QFileDialog, QProgressBar, QMessageBox,
    QScrollArea, QGridLayout, QLineEdit,
    QHBoxLayout, QFrame, QCheckBox, QDialog, QTextEdit, QComboBox, QTextBrowser,
    QSpacerItem, QSizePolicy
)
from PySide6.QtCore import QThread, Signal, Qt, QTimer, QEvent, QSize
from PySide6.QtGui import QIcon, QDesktopServices
from PySide6.QtCore import QUrl

from core.resolver import resolve_dependency, is_asset_var
from core.scanner import scan_var
from core.scene_card import SceneCard


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
        All unrelated .var files to that scenes will be temporarily disabled.
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
    <b>No extra steps needed</b><br>
    When you close <b>VaM.exe</b>, all renamed .var files will automatically return to normal.
  </li>
</ol>

<hr>

<h3>Presets</h3>
<p>
You can save and load your scene selections as presets, so you don’t need to
select them again next time.
</p>

<hr>

<h3>Notes</h3>
<p>
This app continue to improving.
More features will be added over time, but the goal is to
<b>keep it simple and easy to use</b>.
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
            headers={"User-Agent": "VSVM/1.0"}
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

    # skip noisy/unknown files like VaM does
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
# Loading popup (animated dots + indeterminate bar)
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
        self.bar.setRange(0, 0)  # indeterminate
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
# Donation dialog (Supporters + Patreon link)
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
# Worker Thread (VaM-like by default)
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
# Lightweight change-check worker
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
# Looks detector worker (ONLY when checkbox ticked)
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
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        _ = app_data_dir()
        _ = previews_dir()

        self.setWindowTitle("VaM Simple VAR Manager")
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

        self.lean_active = False
        self.vam_seen_running = False

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(1000)
        self.poll_timer.timeout.connect(self.check_vam_state)

        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.setInterval(100)
        self._resize_timer.timeout.connect(lambda: self.apply_filter(self.search.text()))

        self._syncing_selection_ui = False
        self._batch_selection = False

        self.looks_map: dict[str, bool] = {}

        # refresh / startup flags
        self._startup_in_progress = False
        self._refresh_in_progress = False

        self._btn_css = """
            QPushButton, QComboBox {
                background-color: rgba(255,255,255,0.08);
                border: 1px solid rgba(255,255,255,0.18);
                border-radius: 8px;
                padding: 6px 10px;
                color: #eee;
                min-height: 26px;
            }
            QPushButton:hover {
                background-color: rgba(255,255,255,0.14);
                border: 1px solid rgba(255,255,255,0.28);
            }
            QPushButton:pressed {
                background-color: rgba(255,255,255,0.18);
            }
            QPushButton:disabled {
                background-color: rgba(255,255,255,0.03);
                border: 1px solid rgba(255,255,255,0.08);
                color: #777;
            }
        """

        # Loading popup
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
        icon_path = Path("icons/qmark.png")
        if icon_path.exists():
            self.btn_help.setIcon(QIcon(str(icon_path)))
            self.btn_help.setIconSize(QSize(18, 18))
        self.btn_help.setStyleSheet(self._btn_css)
        self.btn_help.clicked.connect(self.open_help)
        row_head.addWidget(self.btn_help, 0, Qt.AlignRight)

        self.btn_donate = QPushButton(" Support Me")
        icon_path = Path("icons/donation.png")
        if icon_path.exists():
            self.btn_donate.setIcon(QIcon(str(icon_path)))
            self.btn_donate.setIconSize(QSize(18, 18))
        self.btn_donate.setStyleSheet(self._btn_css)
        self.btn_donate.clicked.connect(self.open_donation)
        row_head.addWidget(self.btn_donate, 0, Qt.AlignRight)

        self.left_layout.addLayout(row_head)

        self.status = QLabel("Starting...")
        self.left_layout.addWidget(self.status)

        self.info_label = QLabel("Total VARs: - | Unused VARs: - | Scenes: -")
        self.info_label.setStyleSheet("color: #aaa;")
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

        # Renamed "Use Last Folder" -> "Refresh"
        self.btn_refresh = QPushButton("Refresh")
        self.btn_refresh.setToolTip("Check for changes in the current VaM folder and update if needed.")
        self.btn_refresh.setStyleSheet(self._btn_css)
        self.btn_refresh.clicked.connect(self.refresh_clicked)
        row_top.addWidget(self.btn_refresh, 1)

        self.btn_select = QPushButton("Select VaM Directory")
        self.btn_select.setStyleSheet(self._btn_css)
        self.btn_select.clicked.connect(self.select_folder)
        icon_path = Path("icons/folder.png")
        if icon_path.exists():
            self.btn_select.setIcon(QIcon(str(icon_path)))
            self.btn_select.setIconSize(QSize(18, 18))
        row_top.addWidget(self.btn_select, 2)

        ctl_left_lay.addLayout(row_top)

        row_launch = QHBoxLayout()

        icon_path = Path("icons/playbtn.png")
        icon_obj = QIcon(str(icon_path)) if icon_path.exists() else None

        self.btn_launch_vam = QPushButton("  Launch VaM.exe")
        if icon_obj:
            self.btn_launch_vam.setIcon(icon_obj)
            self.btn_launch_vam.setIconSize(QSize(13, 13))
        self.btn_launch_vam.setStyleSheet(self._btn_css)
        self.btn_launch_vam.clicked.connect(self.launch_vam_exe_lean)
        row_launch.addWidget(self.btn_launch_vam, 1)

        self.btn_launch_launcher = QPushButton("  Launch VaM Launcher")
        if icon_obj:
            self.btn_launch_launcher.setIcon(icon_obj)
            self.btn_launch_launcher.setIconSize(QSize(13, 13))
        self.btn_launch_launcher.setStyleSheet(self._btn_css)
        self.btn_launch_launcher.clicked.connect(self.launch_vam_launcher_lean)
        row_launch.addWidget(self.btn_launch_launcher, 1)

        self.btn_restore = QPushButton("Restore VARs")
        self.btn_restore.setStyleSheet(self._btn_css)
        self.btn_restore.setEnabled(False)
        self.btn_restore.clicked.connect(self.restore_offloaded_vars)
        row_launch.addWidget(self.btn_restore, 1)

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
        self.preset_combo = QComboBox()
        self.preset_combo.addItems([
            "Scenes Selection Preset 1",
            "Scenes Selection Preset 2",
            "Scenes Selection Preset 3",
            "Scenes Selection Preset 4",
        ])
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
        self.selection_info.setStyleSheet("color: #aaa;")
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
        self.dep_selected.setStyleSheet("color: #bbb;")
        self.right_layout.addWidget(self.dep_selected)

        self.dep_summary = QLabel("")
        self.dep_summary.setStyleSheet("color: #aaa;")
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
        self.update_selection_ui()

        self._maybe_show_welcome_once()

        # show loading popup immediately after window is shown
        QTimer.singleShot(0, self._startup_sequence)

    # ======================
    # Startup sequence with loading popup
    # ======================
    def _startup_sequence(self):
        self._startup_in_progress = True
        self.loading.start("Loading", "Preparing scenes...")
        # start auto-open after popup is visible
        QTimer.singleShot(50, self.auto_open_last_folder_on_startup)

    def _end_busy(self):
        self.loading.stop()
        self._startup_in_progress = False
        self._refresh_in_progress = False
        self.refresh_refresh_button()

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
    # Donation window
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
        self.cfg.setdefault("presets", {})
        self.cfg["presets"][key] = sorted(self.selected_scene_vars)
        self.cfg["last_preset"] = idx
        save_config(self.cfg)
        QMessageBox.information(self, "Preset Saved", f"Saved current selection to Preset {idx}")

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

        self.selected_scene_vars = set(items)

        self._batch_selection = True
        self.scene_container.setUpdatesEnabled(False)
        try:
            for var_name in self.card_by_var.keys():
                self._set_var_selected(var_name, var_name in self.selected_scene_vars)
        finally:
            self._batch_selection = False
            self.scene_container.setUpdatesEnabled(True)

        self.cfg["last_preset"] = idx
        save_config(self.cfg)

        self.apply_filter(self.search.text())
        self.update_selection_ui()
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
        if obj == self.scroll.viewport() and event.type() == QEvent.Resize:
            self._resize_timer.start()
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

    def _is_process_running_windows(self, image_name: str) -> bool:
        try:
            out = subprocess.check_output(
                ["tasklist", "/FI", f"IMAGENAME eq {image_name}"],
                creationflags=subprocess.CREATE_NO_WINDOW
            ).decode(errors="ignore").lower()
            return image_name.lower() in out and "no tasks are running" not in out
        except Exception:
            return False

    def is_vam_running(self) -> bool:
        return self._is_process_running_windows("VaM.exe")

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
        # enabled if we have a current folder OR a saved last folder
        enabled = (self.vam_dir is not None) or (self.last_vam_dir() is not None)
        # if a worker is running, disable
        if self._startup_in_progress or self._refresh_in_progress:
            enabled = False
        self.btn_refresh.setEnabled(enabled)

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
    # Refresh button behavior
    # ======================
    def refresh_clicked(self):
        # Goal: check changes and update if needed (no forced full rescan).
        target = self.vam_dir or self.last_vam_dir()
        if not target:
            QMessageBox.information(self, "No folder", "No VaM folder selected yet.\nClick 'Select VaM Directory' first.")
            return

        ok, msg = self.validate_vam_folder(target)
        if not ok:
            QMessageBox.warning(self, "Invalid folder", msg)
            return

        # If user never loaded anything yet this run, initialize paths + show cached UI if possible
        if self.vam_dir is None:
            self.vam_dir = target
            self.addon_dir = target / "AddonPackages"
            self.saves_scene_dir = target / "Saves" / "scene"

        self._refresh_in_progress = True
        self.refresh_refresh_button()
        self.loading.start("Refreshing", "Checking for changes...")

        cache_obj = self._load_scene_cache()
        if self._can_use_cache_for_current_folder(cache_obj):
            # show cached cards immediately (if UI empty)
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

        # always run change check (if cache mismatch, it will rescan)
        self._start_change_check(cache_obj)

    # ======================
    # Startup auto-load logic (cache-first)
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

            # scene card is present now -> close the startup popup (requirement #1)
            self.loading.stop()

            # Background lightweight check
            self._start_change_check(cache_obj)
            return

        # No usable cache -> do normal scan (keep loading popup until done)
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

        # Changes detected -> run AnalyzeWorker (fast due to signature reuse)
        self.status.setText("Changes detected. Updating cache...")
        # keep popup visible while rescanning
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

        self.scene_entries = []
        self.total_scene_files_found = 0

        # show loading popup for full scan
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

        # requirement #1: hide popup after process done and cards appear
        self._end_busy()

    # ======================
    # Cards
    # ======================
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

            image_bytes = read_preview_bytes(rel)

            card = SceneCard(scene_name=scene_name, var_name=var_name, image_bytes=image_bytes)
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
    # Looks-only (lazy)
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
        for card in self.cards:
            src = getattr(card, "source", "var")
            selectable = (src == "var")
            card.set_selection_mode(enabled and selectable)
            if selectable:
                card.set_checked(card.var_name in self.selected_scene_vars)
            else:
                card.set_checked(False)

        self.apply_filter(self.search.text())
        self.update_selection_ui()

    def _set_var_selected(self, var_name: str, selected: bool):
        if self._syncing_selection_ui:
            return

        self._syncing_selection_ui = True
        try:
            if selected:
                self.selected_scene_vars.add(var_name)
            else:
                self.selected_scene_vars.discard(var_name)

            for c in self.card_by_var.get(var_name, []):
                c.set_checked(selected)
        finally:
            self._syncing_selection_ui = False

        if not self._batch_selection:
            self.apply_filter(self.search.text())
            self.update_selection_ui()

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

        self._batch_selection = True
        self.scene_container.setUpdatesEnabled(False)
        try:
            for var_name, cards in self.card_by_var.items():
                if any(c.isVisible() for c in cards):
                    self._set_var_selected(var_name, True)
        finally:
            self._batch_selection = False
            self.scene_container.setUpdatesEnabled(True)

        self.apply_filter(self.search.text())
        self.update_selection_ui()

    def clear_selection(self):
        self._batch_selection = True
        self.scene_container.setUpdatesEnabled(False)
        try:
            for var_name in list(self.selected_scene_vars):
                self._set_var_selected(var_name, False)
        finally:
            self._batch_selection = False
            self.scene_container.setUpdatesEnabled(True)

        self.apply_filter(self.search.text())
        self.update_selection_ui()

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

        info = scan_var(self.addon_dir / var_name)
        deps = sorted(info.get("dependencies", []))
        all_vars = {p.name for p in self.addon_dir.glob("*.var")}

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
                text = f"{dep}  →  {chosen}"
            else:
                missing_count += 1
                text = f"{dep}  →  (missing)"

            self.dep_list_layout.addWidget(DependencyRow(text=text, present=present))

        self.dep_list_layout.addStretch(1)
        self.dep_summary.setText(f"Total: {len(deps)} | Present: {present_count} | Missing: {missing_count}")

    # ======================
    # Restore / Launch (RENAME METHOD)
    # ======================
    def refresh_restore_button(self):
        if not self.vam_dir:
            self.btn_restore.setEnabled(False)
            return
        mp = manifest_path_for(self.vam_dir)
        self.btn_restore.setEnabled(mp.exists())

    def restore_offloaded_vars(self):
        if not self.vam_dir:
            QMessageBox.warning(self, "No folder", "Select VaM directory first.")
            return

        addon_dir = self.vam_dir / "AddonPackages"
        mp = manifest_path_for(self.vam_dir)

        if not mp.exists():
            QMessageBox.information(self, "Nothing to restore", "No manifest found.")
            self.refresh_restore_button()
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

        QMessageBox.information(self, "Restored", f"Restored {restored} VARs back to normal in:\n{addon_dir}")
        self.refresh_restore_button()

    def _scene_vars_for_launch(self) -> set[str]:
        if self.selected_scene_vars:
            return set(self.selected_scene_vars)
        return set(self.card_by_var.keys())

    def compute_keep_set_for_scene_vars(self, scene_vars: set[str]) -> set[str]:
        assert self.addon_dir is not None
        all_vars = {p.name for p in self.addon_dir.glob("*.var")}

        protected = {v for v in all_vars if is_asset_var(v)}
        protected |= {v for v in all_vars if "[plugin]" in v.lower()}

        keep = set(protected)
        queue: list[str] = []

        for scene_var in scene_vars:
            if scene_var in all_vars:
                keep.add(scene_var)
                info = scan_var(self.addon_dir / scene_var)
                queue.extend(list(info.get("dependencies", [])))

        while queue:
            dep = queue.pop()
            for matched_var in resolve_dependency(dep, all_vars):
                if matched_var not in keep:
                    keep.add(matched_var)
                    info = scan_var(self.addon_dir / matched_var)
                    queue.extend(list(info.get("dependencies", [])))

        return keep

    def disable_unrelated_vars_by_rename(self, keep_set: set[str]) -> dict:
        assert self.vam_dir is not None
        addon_dir = self.vam_dir / "AddonPackages"
        mp = manifest_path_for(self.vam_dir)

        renamed: list[dict] = []
        for var_path in addon_dir.glob("*.var"):
            name = var_path.name
            if name in keep_set:
                continue

            disabled = disabled_name(name)
            disabled_path = addon_dir / disabled

            if disabled_path.exists():
                continue

            try:
                var_path.rename(disabled_path)
                renamed.append({"from": name, "to": disabled})
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

    def _start_lean_session_or_warn(self) -> bool:
        if not self.vam_dir or not self.addon_dir:
            QMessageBox.warning(self, "No folder", "Select VaM directory first.")
            return False

        if self.is_vam_running():
            QMessageBox.warning(self, "VaM is running", "Close VaM.exe first, then launch from this app.")
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
            "This will TEMPORARILY disable unrelated .var files by renaming them.\n"
            "Example: Something.var → Something.var.disabled\n"
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
            return False

        self.lean_active = True
        self.vam_seen_running = False
        self.poll_timer.start()
        self.refresh_restore_button()
        return True

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

    def check_vam_state(self):
        if not self.lean_active:
            self.poll_timer.stop()
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

    icon_path = Path("icons/app.ico")
    if icon_path.exists():
        icon = QIcon(str(icon_path))
        app.setWindowIcon(icon)

    window = MainWindow()
    if icon_path.exists():
        window.setWindowIcon(QIcon(str(icon_path)))

    window.show()
    sys.exit(app.exec())

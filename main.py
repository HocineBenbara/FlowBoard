import sys
import os
import json
import threading
import re
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, urljoin

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QScrollArea,
    QFrame, QLabel, QLineEdit, QListWidget, QPushButton, QToolButton, QGridLayout,
    QDialog, QInputDialog, QMessageBox, QListWidgetItem, QButtonGroup, QSplashScreen,
    QMenu, QGraphicsDropShadowEffect, QSizePolicy, QSplitter, QStyledItemDelegate,
    QFileIconProvider
)
from PySide6.QtCore import (
    Qt, QTimer, QUrl, QSize, Signal, QEvent, QPoint, QMimeData, QObject,
    QModelIndex, QPropertyAnimation, QEasingCurve, QRect, QFileInfo
)
from PySide6.QtGui import (
    QIcon, QKeySequence, QShortcut, QDesktopServices, QPalette, QColor,
    QDrag, QPixmap, QPainter, QFont, QLinearGradient, QBrush, QPen,
    QRadialGradient, QPainterPath
)

import webbrowser
import urllib.parse

if sys.platform == "win32":
    try:
        import pythoncom
        from win32com.shell import shell, shellcon
        HAS_WIN32 = True
    except ImportError:
        HAS_WIN32 = False
else:
    HAS_WIN32 = False

# ── Constantes partagées ──────────────────────────────────────
CARD_SIZE    = 122
CARD_STRIDE  = 134
COLS_DEFAULT = 5

CARD_STYLE_NORMAL = """
QToolButton {
    background: rgba(255,255,255,0.93);
    border-radius: 14px;
    padding: 10px 6px 6px 6px;
    margin: 3px;
    font-size: 13px;
    color: #1e1e3a;
    border: 1px solid rgba(108,92,231,0.15);
}
QToolButton:hover {
    background: rgba(255,255,255,1.0);
    border: 1px solid rgba(108,92,231,0.55);
    color: #3d2d9c;
}
QToolButton:pressed {
    background: rgba(235,230,255,1.0);
    border: 1.5px solid #6c5ce7;
}
"""
CARD_STYLE_SELECTED = """
QToolButton {
    background: rgba(235,230,255,1.0);
    border-radius: 14px;
    padding: 10px 6px 6px 6px;
    margin: 3px;
    font-size: 13px;
    color: #1e1e3a;
    border: 2px solid #6c5ce7;
}
QToolButton:hover {
    background: rgba(235,230,255,1.0);
    border: 2px solid #6c5ce7;
}
"""


# ══════════════════════════════════════════════════════════════
#  SPLASH SCREEN
# ══════════════════════════════════════════════════════════════
class QuickSplash(QSplashScreen):
    def __init__(self):
        pixmap = QPixmap(500, 240)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)

        grad = QLinearGradient(0, 0, 500, 240)
        grad.setColorAt(0.0, QColor(35, 20, 90))
        grad.setColorAt(0.6, QColor(72, 52, 200))
        grad.setColorAt(1.0, QColor(90, 75, 210))
        painter.fillRect(pixmap.rect(), QBrush(grad))

        rg = QRadialGradient(250, 120, 160)
        rg.setColorAt(0.0, QColor(150, 130, 255, 50))
        rg.setColorAt(1.0, QColor(0, 0, 0, 0))
        painter.fillRect(pixmap.rect(), QBrush(rg))

        painter.setPen(QColor(255, 255, 255))
        painter.setFont(QFont("Segoe UI", 44, QFont.Bold))
        painter.drawText(pixmap.rect().adjusted(0, -40, 0, 0), Qt.AlignCenter, "FlowBoard")

        painter.setFont(QFont("Segoe UI", 13))
        painter.setPen(QColor(210, 205, 255))
        painter.drawText(pixmap.rect().adjusted(0, 55, 0, 0), Qt.AlignCenter, "Visual Command Center")

        painter.setFont(QFont("Segoe UI", 10))
        painter.setPen(QColor(170, 160, 230))
        painter.drawText(pixmap.rect().adjusted(0, 88, 0, 0), Qt.AlignCenter, "Loading…")

        painter.end()
        super().__init__(pixmap)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.SplashScreen | Qt.FramelessWindowHint)


# ══════════════════════════════════════════════════════════════
#  ICON CACHE
# ══════════════════════════════════════════════════════════════
class IconCache:
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._cache = {}
        return cls._instance

    def get(self, key: str):
        with self._lock:
            return self._cache.get(key)

    def set(self, key: str, value):
        with self._lock:
            self._cache[key] = value


ICON_CACHE = IconCache()


# ══════════════════════════════════════════════════════════════
#  NATIVE ICON WORKER
#  Le thread ne fait QUE retourner le chemin résolu (aucun objet Qt).
#  Le QIcon est construit dans le thread UI via le signal iconReady.
# ══════════════════════════════════════════════════════════════
class NativeIconWorker(QObject):
    """Résout les chemins (.lnk sur Windows) en background, sans jamais créer de QIcon."""
    iconReady = Signal(str, str)   # (original_path, resolved_path)

    def __init__(self):
        super().__init__()
        self._pending: set[str] = set()
        self._lock = threading.Lock()

    def request(self, path: str):
        """Demande la résolution d'un chemin. Idempotent."""
        key = f"native:{path}"
        if ICON_CACHE.get(key) is not None:
            return   # déjà en cache, rien à faire
        with self._lock:
            if path in self._pending:
                return
            self._pending.add(path)
        threading.Thread(target=self._resolve, args=(path,), daemon=True).start()

    def _resolve(self, path: str):
        """Thread worker : retourne uniquement le chemin résolu, zéro Qt."""
        resolved = path
        if sys.platform == "win32" and path.lower().endswith(".lnk") and os.path.exists(path):
            resolved = _resolve_shortcut_raw(path)
            if not os.path.exists(resolved):
                resolved = path
        with self._lock:
            self._pending.discard(path)
        self.iconReady.emit(path, resolved)


def _resolve_shortcut_raw(lnk_path: str) -> str:
    """Résout un .lnk Windows sans aucun objet Qt. Appelable depuis un thread."""
    if not HAS_WIN32:
        return lnk_path
    try:
        pythoncom.CoInitialize()
        try:
            sc = pythoncom.CoCreateInstance(
                shell.CLSID_ShellLink, None,
                pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink
            )
            sc.QueryInterface(pythoncom.IID_IPersistFile).Load(lnk_path, 0)
            target, _ = sc.GetPath(shell.SLG_RAWPATH)
            return target if target and os.path.exists(target) else lnk_path
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        return lnk_path


def _build_native_icon_ui(path: str) -> QIcon:
    """
    Construit un QIcon depuis le thread UI UNIQUEMENT.
    QFileIconProvider et QFileInfo ne sont pas thread-safe.
    """
    try:
        provider = QFileIconProvider()
        icon = provider.icon(QFileInfo(path))
        if not icon.isNull():
            return icon
        if os.path.isdir(path):
            return provider.icon(QFileIconProvider.Folder)
        if path.lower().endswith((".exe", ".bat", ".cmd", ".com", ".lnk")):
            return provider.icon(QFileIconProvider.Executable)
        return provider.icon(QFileIconProvider.File)
    except Exception:
        provider = QFileIconProvider()
        return provider.icon(
            QFileIconProvider.Folder if os.path.isdir(path) else QFileIconProvider.File
        )


def _placeholder_icon(path: str) -> QIcon:
    """Icône générique immédiate (thread UI), utilisée pendant la résolution async."""
    provider = QFileIconProvider()
    return provider.icon(
        QFileIconProvider.Folder if os.path.isdir(path) else QFileIconProvider.File
    )


# ══════════════════════════════════════════════════════════════
#  ASYNC FAVICON LOADER
#  Le thread télécharge les bytes bruts.
#  QPixmap / QIcon sont créés dans le thread UI via faviconBytes.
# ══════════════════════════════════════════════════════════════
class AsyncFaviconLoader(QObject):
    faviconLoaded = Signal(str, str)          # (url, file_path) — fichier déjà sur disque
    faviconBytes  = Signal(str, str, bytes)   # (url, out_path, raw_bytes) — à sauvegarder dans UI

    def __init__(self):
        super().__init__()
        self._queue: list[str] = []
        self._running = False
        self._write_lock = threading.Lock()

    def load(self, url: str, cache_dir: Path):
        key = f"favicon:{url}"
        cached = ICON_CACHE.get(key)
        if cached:
            self.faviconLoaded.emit(url, cached)
            return
        if url not in self._queue:
            self._queue.append(url)
        if not self._running:
            self._running = True
            threading.Thread(
                target=self._process_queue, args=(cache_dir,), daemon=True
            ).start()

    def _process_queue(self, cache_dir: Path):
        while self._queue:
            url = self._queue.pop(0)
            out_path, raw = self._download_favicon(url, cache_dir)
            if out_path and raw is None:
                # Fichier déjà sur disque — émettre directement
                ICON_CACHE.set(f"favicon:{url}", out_path)
                self.faviconLoaded.emit(url, out_path)
            elif out_path and raw:
                # Bytes bruts → thread UI crée QPixmap et sauvegarde
                self.faviconBytes.emit(url, out_path, raw)
        self._running = False

    def _download_favicon(self, url: str, cache_dir: Path) -> tuple:
        """
        Télécharge les bytes bruts du favicon.
        NE crée AUCUN objet Qt (pas de QPixmap, pas de QIcon).
        Retourne (out_path, bytes|None).
        """
        try:
            import urllib.request, ssl

            parsed = urlparse(url)
            domain = (parsed.netloc or parsed.path.split('/')[0])
            domain = domain.split(':')[0].split('?')[0].split('#')[0].lower()
            if not domain:
                return ("", None)

            favicon_dir = cache_dir / "favicons"
            favicon_dir.mkdir(parents=True, exist_ok=True)
            out_path = favicon_dir / f"{domain.replace('.', '_')}.png"

            # Fichier déjà sur disque → pas besoin de re-télécharger
            if out_path.exists():
                return (str(out_path), None)

            UA = (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
            headers = {
                "User-Agent": UA,
                "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9,fr;q=0.8",
                "Connection": "keep-alive",
            }
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE

            def fetch(fetch_url: str) -> bytes | None:
                try:
                    req = urllib.request.Request(fetch_url, headers={"User-Agent": UA})
                    with urllib.request.urlopen(req, context=ctx, timeout=10) as r:
                        if r.status == 200:
                            return r.read()
                except Exception:
                    pass
                return None

            # 1. Google favicon service (meilleure qualité)
            data = fetch(f"https://www.google.com/s2/favicons?domain={domain}&sz=256")
            if data:
                return (str(out_path), data)

            # 2. Scraping de la page HTML pour les balises <link>
            try:
                req = urllib.request.Request(f"https://{domain}", headers=headers)
                with urllib.request.urlopen(req, context=ctx, timeout=10) as r:
                    html = r.read().decode("utf-8", errors="ignore")
                patterns = [
                    r'<link[^>]+rel=["\']apple-touch-icon(?:-precomposed)?["\'][^>]+href=["\']([^"\']+)["\']',
                    r'<link[^>]+sizes=["\'](?:192|180|152|144|128|96|64|32|16)x\d+["\'][^>]+href=["\']([^"\']+)["\']',
                    r'<link[^>]+href=["\']([^"\']+)["\'][^>]+sizes=["\'](?:192|180|152|144|128|96|64|32|16)x\d+["\']',
                    r'<link[^>]+rel=["\'](?:icon|shortcut icon)["\'][^>]+href=["\']([^"\']+)["\']',
                    r'<link[^>]+href=["\']([^"\']+\.ico)["\']',
                    r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']',
                ]
                for pat in patterns:
                    for icon_url in re.findall(pat, html, re.I):
                        if icon_url.startswith("//"):
                            icon_url = "https:" + icon_url
                        elif icon_url.startswith("/"):
                            icon_url = f"https://{domain}{icon_url}"
                        elif not icon_url.startswith(("http://", "https://")):
                            icon_url = urljoin(f"https://{domain}/", icon_url)
                        d = fetch(icon_url)
                        if d:
                            return (str(out_path), d)
            except Exception:
                pass

            # 3. Candidats classiques
            for candidate in [
                f"https://{domain}/favicon.ico",
                f"https://{domain}/favicon.png",
                f"https://{domain}/favicon.svg",
                f"https://{domain}/apple-touch-icon.png",
                f"https://{domain}/apple-touch-icon-precomposed.png",
                f"https://www.{domain}/favicon.ico",
                f"https://icons.duckduckgo.com/ip3/{domain}.ico",
                f"https://logo.clearbit.com/{domain}",
            ]:
                d = fetch(candidate)
                if d:
                    return (str(out_path), d)

        except Exception as e:
            print(f"⚠️ Favicon error ({url}): {e}")

        return ("", None)


# ══════════════════════════════════════════════════════════════
#  UTILS
# ══════════════════════════════════════════════════════════════
def resolve_shortcut(lnk_path: str) -> str:
    """Wrapper public — appelle la version raw (utilisable dans n'importe quel contexte)."""
    return _resolve_shortcut_raw(lnk_path)


def get_shortcut_icon(shortcut: dict, cache_dir: Path,
                      favicon_loader: "AsyncFaviconLoader",
                      native_worker: "NativeIconWorker") -> QIcon:
    """
    Retourne une icône immédiatement (placeholder si besoin).
    Le vrai chargement async met à jour la carte via les signaux.
    Appelé UNIQUEMENT depuis le thread UI.
    """
    if shortcut.get("type") == "url" and shortcut.get("url"):
        url = shortcut["url"]
        cached_path = ICON_CACHE.get(f"favicon:{url}")
        if cached_path and os.path.exists(cached_path):
            return QIcon(cached_path)
        # Lance le téléchargement async; retourne une icône générique en attendant
        favicon_loader.load(url, cache_dir)
        return QIcon.fromTheme("internet-web-browser")

    path = shortcut.get("path", "")
    if not path or not os.path.exists(path):
        return QIcon.fromTheme("text-x-generic")

    key = f"native:{path}"
    cached = ICON_CACHE.get(key)
    if cached and isinstance(cached, QIcon) and not cached.isNull():
        return cached

    # Placeholder immédiat + résolution async pour .lnk ou si pas en cache
    native_worker.request(path)
    return _placeholder_icon(path)


def _open_path(path: str, parent=None):
    """Ouvre un fichier/dossier avec gestion d'erreur propre."""
    try:
        if not os.path.exists(path):
            if parent:
                QMessageBox.warning(parent, "Not found", f"Path not found:\n{path}")
            return
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            import subprocess; subprocess.Popen(["open", path])
        else:
            import subprocess; subprocess.Popen(["xdg-open", path])
    except Exception as e:
        if parent:
            QMessageBox.warning(parent, "Error", str(e))


# ══════════════════════════════════════════════════════════════
#  RICH TOOLTIP
# ══════════════════════════════════════════════════════════════
class RichToolTip(QWidget):
    def __init__(self, parent=None):
        super().__init__(None, Qt.ToolTip | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setMaximumWidth(360)

        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self.hide)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self._frame = QFrame()
        self._frame.setObjectName("ttFrame")
        self._frame.setStyleSheet("""
            QFrame#ttFrame {
                background-color: #16162a;
                border: 1px solid #6c5ce7;
                border-radius: 12px;
            }
        """)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(28)
        shadow.setOffset(0, 6)
        shadow.setColor(QColor(0, 0, 0, 160))
        self._frame.setGraphicsEffect(shadow)

        inner = QVBoxLayout(self._frame)
        inner.setContentsMargins(16, 12, 16, 12)
        inner.setSpacing(0)

        self._note_frame = QFrame()
        self._note_frame.setStyleSheet(
            "background: rgba(108,92,231,0.18); border-radius: 7px;"
            "border-left: 3px solid #6c5ce7; padding: 0;"
        )
        note_inner = QHBoxLayout(self._note_frame)
        note_inner.setContentsMargins(10, 7, 10, 7)
        self._lbl_note = QLabel()
        self._lbl_note.setWordWrap(True)
        self._lbl_note.setStyleSheet(
            "color: #cdd6f4; font-size: 12px; background: transparent; font-style: italic;"
        )
        note_inner.addWidget(self._lbl_note)
        inner.addWidget(self._note_frame)
        root.addWidget(self._frame)

    def show_for(self, shortcut: dict, global_pos: QPoint):
        self._hide_timer.stop()
        note = shortcut.get("note", "").strip()
        if not note:
            return
        self._lbl_note.setText(note)
        self.adjustSize()
        self._place(global_pos)
        self.show()
        self.raise_()

    def hide_delayed(self, ms: int = 250):
        self._hide_timer.start(ms)

    def cancel_hide(self):
        self._hide_timer.stop()

    def _place(self, gpos: QPoint):
        screen = QApplication.primaryScreen().availableGeometry()
        x = gpos.x() + 18
        y = gpos.y() + 18
        if x + self.width() > screen.right() - 8:
            x = gpos.x() - self.width() - 10
        if y + self.height() > screen.bottom() - 8:
            y = gpos.y() - self.height() - 10
        self.move(max(screen.left() + 4, x), max(screen.top() + 4, y))


# ══════════════════════════════════════════════════════════════
#  DRAGGABLE SHORTCUT CARD
# ══════════════════════════════════════════════════════════════
class DraggableShortcutCard(QToolButton):
    doubleClicked = Signal()

    shortcutData   = None
    _drag_start    = None
    _tooltip_inst: "RichToolTip | None" = None

    _NOTE_BADGE_SIZE   = 16
    _NOTE_BADGE_MARGIN = 5

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.setAcceptDrops(False)
        self._selected = False
        self._has_note = False

    def _note_badge_rect(self) -> QRect:
        m = self._NOTE_BADGE_MARGIN
        s = self._NOTE_BADGE_SIZE
        return QRect(self.width() - s - m, self.height() - s - m, s, s)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._has_note:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.Antialiasing)
            r = self._note_badge_rect()
            painter.setBrush(QBrush(QColor(108, 92, 231, 200)))
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(r)
            painter.setFont(QFont("Segoe UI Emoji", 8))
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(r, Qt.AlignCenter, "📝")
            painter.end()

    def setSelected(self, val: bool):
        self._selected = val
        self._update_style()

    def _update_style(self):
        self.setStyleSheet(CARD_STYLE_SELECTED if self._selected else CARD_STYLE_NORMAL)

    def leaveEvent(self, event):
        tip = DraggableShortcutCard._tooltip_inst
        if tip:
            tip.hide_delayed(200)
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_start = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if (event.buttons() & Qt.LeftButton) and self._drag_start is not None:
            if (event.position().toPoint() - self._drag_start).manhattanLength() \
                    >= QApplication.startDragDistance():
                tip = DraggableShortcutCard._tooltip_inst
                if tip:
                    tip.hide()
                if self.shortcutData:
                    drag = QDrag(self)
                    mime = QMimeData()
                    mime.setData(
                        "application/x-flowboard-shortcut-card",
                        json.dumps(self.shortcutData, ensure_ascii=False).encode()
                    )
                    px = QPixmap(self.size())
                    self.render(px)
                    scaled = px.scaled(110, 110, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    combined = QPixmap(scaled.size())
                    combined.fill(Qt.transparent)
                    p = QPainter(combined)
                    p.setOpacity(0.85)
                    p.drawPixmap(0, 0, scaled)
                    p.setOpacity(1.0)
                    pen = QPen(QColor(108, 92, 231), 2)
                    p.setPen(pen)
                    p.setBrush(Qt.NoBrush)
                    p.drawRoundedRect(combined.rect().adjusted(1, 1, -1, -1), 12, 12)
                    p.end()
                    drag.setPixmap(combined)
                    drag.setHotSpot(QPoint(combined.width() // 2, combined.height() // 2))
                    drag.setMimeData(mime)
                    drag.exec(Qt.MoveAction)
                self._drag_start = None
            return

        tip = DraggableShortcutCard._tooltip_inst
        sc  = self.shortcutData
        pos = event.position().toPoint()
        if tip and sc and sc.get("note"):
            badge_rect = self._note_badge_rect().adjusted(-4, -4, 4, 4)
            if badge_rect.contains(pos):
                tip.cancel_hide()
                tip.show_for(sc, self.mapToGlobal(QPoint(self.width() + 4, pos.y())))
            else:
                tip.hide_delayed(200)
        elif tip:
            tip.hide_delayed(200)

        super().mouseMoveEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.doubleClicked.emit()
        super().mouseDoubleClickEvent(event)


# ══════════════════════════════════════════════════════════════
#  SHORTCUT GRID WIDGET
# ══════════════════════════════════════════════════════════════
class ShortcutGridWidget(QWidget):
    shortcutDropped = Signal(int, int, dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self._hover_hl: QWidget | None = None

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("application/x-flowboard-shortcut-card"):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        pos = event.position().toPoint()
        if self._hover_hl:
            self._hover_hl.deleteLater()
            self._hover_hl = None
        col = pos.x() // CARD_STRIDE
        row = pos.y() // CARD_STRIDE
        hl = QWidget(self)
        hl.setStyleSheet(
            "background: rgba(108,92,231,0.35);"
            "border: 2px solid #a29bfe; border-radius: 12px;"
        )
        hl.setGeometry(col * CARD_STRIDE + 4, row * CARD_STRIDE + 4, CARD_SIZE, CARD_SIZE)
        hl.show()
        self._hover_hl = hl
        event.acceptProposedAction()

    def dropEvent(self, event):
        if self._hover_hl:
            self._hover_hl.deleteLater()
            self._hover_hl = None
        pos = event.position().toPoint()
        col = pos.x() // CARD_STRIDE
        row = pos.y() // CARD_STRIDE
        if event.mimeData().hasFormat("application/x-flowboard-shortcut-card"):
            try:
                sc = json.loads(
                    event.mimeData().data("application/x-flowboard-shortcut-card")
                    .data().decode()
                )
                self.shortcutDropped.emit(row, col, sc)
                event.acceptProposedAction()
                return
            except Exception as e:
                print(f"Drop decode error: {e}")
        event.ignore()

    def leaveEvent(self, event):
        if self._hover_hl:
            self._hover_hl.deleteLater()
            self._hover_hl = None
        super().leaveEvent(event)


# ══════════════════════════════════════════════════════════════
#  CATEGORY LIST WIDGET
# ══════════════════════════════════════════════════════════════
class CategoryListWidget(QListWidget):
    shortcutDroppedOnCategory = Signal(str, dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.InternalMove)
        self._hovered: QListWidgetItem | None = None

    def _reset(self, item):
        if item:
            f = item.font(); f.setBold(False); item.setFont(f)
            item.setBackground(QColor(0, 0, 0, 0))

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("application/x-flowboard-shortcut-card"):
            event.acceptProposedAction()
        elif event.mimeData().hasFormat("application/x-qabstractitemmodeldatalist"):
            super().dragEnterEvent(event)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        self._reset(self._hovered)
        self._hovered = None
        item = self.itemAt(event.position().toPoint())
        if item and item.data(Qt.UserRole) != "all":
            f = item.font(); f.setBold(True); item.setFont(f)
            item.setBackground(QColor(108, 92, 231, 55))
            self._hovered = item
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        item = self.itemAt(event.position().toPoint())
        self._reset(self._hovered)
        self._hovered = None
        if event.mimeData().hasFormat("application/x-flowboard-shortcut-card"):
            if item and item.data(Qt.UserRole) != "all":
                cat_id = item.data(Qt.UserRole)
                sc = json.loads(
                    event.mimeData().data("application/x-flowboard-shortcut-card")
                    .data().decode()
                )
                self.shortcutDroppedOnCategory.emit(cat_id, sc)
                event.acceptProposedAction()
                return
        if event.mimeData().hasFormat("application/x-qabstractitemmodeldatalist"):
            super().dropEvent(event)
            self.model().rowsMoved.emit(QModelIndex(), 0, 0, QModelIndex(), 0)


# ══════════════════════════════════════════════════════════════
#  CATEGORY COUNT DELEGATE
# ══════════════════════════════════════════════════════════════
class CategoryCountDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        from PySide6.QtWidgets import QStyle
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)

        is_selected = bool(option.state & QStyle.State_Selected)
        is_hover    = bool(option.state & QStyle.State_MouseOver)
        rect = option.rect.adjusted(2, 1, -2, -1)

        if is_selected:
            bg = QColor(108, 92, 231)
        elif is_hover:
            bg = QColor(108, 92, 231, 70)
        else:
            bg = QColor(0, 0, 0, 0)
        painter.setBrush(QBrush(bg))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(rect, 8, 8)

        count = index.data(Qt.UserRole + 1)
        badge_text = str(count) if (count is not None and count > 0) else ""
        badge_w = 0
        badge_h = 17
        if badge_text:
            fm_sm = QFont("Segoe UI", 9)
            from PySide6.QtGui import QFontMetrics
            badge_w = max(QFontMetrics(fm_sm).horizontalAdvance(badge_text) + 10, 20)

        label = index.data(Qt.DisplayRole) or ""
        txt_color = QColor(255, 255, 255) if is_selected else QColor(255, 255, 255, 215)
        font_main = QFont("Segoe UI", 13)
        font_main.setBold(is_selected)
        painter.setFont(font_main)
        painter.setPen(txt_color)
        text_rect = rect.adjusted(14, 0, -(badge_w + 10) if badge_w else -8, 0)
        painter.drawText(text_rect, Qt.AlignVCenter | Qt.TextSingleLine, label)

        if badge_w > 0:
            bx = rect.right() - badge_w - 6
            by = rect.center().y() - badge_h // 2
            badge_rect = QRect(bx, by, badge_w, badge_h)
            bb = QColor(255, 255, 255, 50) if is_selected else QColor(108, 92, 231, 45)
            bt = QColor(255, 255, 255, 190) if is_selected else QColor(180, 170, 255, 155)
            painter.setFont(QFont("Segoe UI", 9))
            painter.setBrush(QBrush(bb))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(badge_rect, badge_h // 2, badge_h // 2)
            painter.setPen(bt)
            painter.drawText(badge_rect, Qt.AlignCenter, badge_text)

        painter.restore()

    def sizeHint(self, option, index):
        return QSize(option.rect.width(), 36)


# ══════════════════════════════════════════════════════════════
#  FLOWBOARD — fenêtre principale
# ══════════════════════════════════════════════════════════════
class FlowBoard(QMainWindow):

    C = {
        "primary":        "#6c5ce7",
        "primary_dark":   "#3d2d9c",
        "primary_darker": "#27196a",
        "primary_light":  "#9b8ff0",
        "bg_glass":       "rgba(255,255,255,0.10)",
        "bg_glass_hover": "rgba(255,255,255,0.18)",
        "sidebar_bg":     "rgba(30,20,80,0.88)",
        "card_bg":        "rgba(255,255,255,0.94)",
        "card_hover":     "rgba(255,255,255,1.0)",
        "dialog_bg":      "#16162a",
        "input_bg":       "#22223a",
        "input_border":   "#8b7ff5",
        "text_dim":       "rgba(255,255,255,0.60)",
        "text_mid":       "rgba(255,255,255,0.85)",
        "text_bright":    "#ffffff",
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle("FlowBoard")
        self.setGeometry(100, 100, 1260, 820)
        self.setMinimumSize(1000, 640)
        self.setAcceptDrops(True)

        self.is_processing     = False
        self.current_search    = ""
        self.search_mode       = "google"
        self._card_widgets: dict[str, DraggableShortcutCard] = {}
        self._selected_card: DraggableShortcutCard | None = None
        self._cols             = COLS_DEFAULT

        from collections import deque
        self._undo_stack: deque = deque(maxlen=20)

        # Dossiers
        self.app_dir       = Path.cwd() / "PySide6_data"
        self.app_dir.mkdir(exist_ok=True)
        self.shortcuts_dir = self.app_dir / "shortcuts"
        self.shortcuts_dir.mkdir(exist_ok=True)

        # ── Workers async ─────────────────────────────────────
        # Favicon loader : télécharge bytes en thread, QPixmap dans UI
        self.async_loader = AsyncFaviconLoader()
        self.async_loader.faviconLoaded.connect(self._on_favicon_loaded)
        self.async_loader.faviconBytes.connect(self._on_favicon_bytes)   # NOUVEAU

        # Native icon worker : résout les chemins en thread, QIcon dans UI
        self.native_worker = NativeIconWorker()
        self.native_worker.iconReady.connect(self._on_native_icon_ready)  # NOUVEAU

        # Données
        self.shortcuts  = self._load_shortcuts()
        self.shortcuts.sort(key=lambda s: s.get("position", 9999))
        self.categories = [{"id": "all", "name": "🏠 All"}] + self._load_categories()
        self.current_category = "all"

        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.timeout.connect(self._save_shortcuts)

        # Tooltip partagé
        self._tooltip = RichToolTip()
        DraggableShortcutCard._tooltip_inst = self._tooltip

        self._build_ui()
        self._bind_shortcuts()
        self.update_clock()
        QTimer.singleShot(900, self._start_favicon_loading)

    # ══════════════════════════════════════════════
    #  SLOTS ICÔNES — tous dans le thread UI
    # ══════════════════════════════════════════════

    def _on_favicon_bytes(self, url: str, out_path: str, raw: bytes):
        """
        Reçoit les bytes bruts depuis le thread worker.
        Crée QPixmap ici dans le thread UI (seul endroit légal).
        """
        px = QPixmap()
        if not px.loadFromData(raw) or px.isNull():
            return
        # Sauvegarde sur disque (thread UI — pas de race condition sur QPixmap)
        with self.async_loader._write_lock:
            if not Path(out_path).exists():
                px.save(out_path, "PNG")
        ICON_CACHE.set(f"favicon:{url}", out_path)
        self._on_favicon_loaded(url, out_path)

    def _on_favicon_loaded(self, url: str, path: str):
        """Met à jour toutes les cartes qui affichent cette URL."""
        if not path or not os.path.exists(path):
            return
        icon = QIcon(path)
        for w in self._card_widgets.values():
            if w.shortcutData and w.shortcutData.get("url") == url:
                w.setIcon(icon)

    def _on_native_icon_ready(self, original_path: str, resolved_path: str):
        """
        Reçoit le chemin résolu depuis NativeIconWorker.
        Construit le QIcon dans le thread UI — légal et sans crash.
        """
        icon = _build_native_icon_ui(resolved_path)
        key  = f"native:{original_path}"
        ICON_CACHE.set(key, icon)

        # Met à jour toutes les cartes qui utilisent ce chemin
        for w in self._card_widgets.values():
            sc = w.shortcutData
            if sc and sc.get("path") == original_path:
                w.setIcon(icon)

    # ══════════════════════════════════════════════
    #  DATA I/O
    # ══════════════════════════════════════════════
    def _load_shortcuts(self) -> list:
        unified = self.app_dir / "shortcuts.json"
        old_files = sorted(
            self.shortcuts_dir.glob("*.json"),
            key=lambda x: int(x.stem.split("_")[1]) if "_" in x.stem else 9999,
        )
        if old_files and not unified.exists():
            result = []
            for f in old_files[:500]:
                try:
                    result.append(json.loads(f.read_text("utf-8")))
                except Exception:
                    pass
            try:
                unified.write_text(json.dumps(result, indent=2, ensure_ascii=False), "utf-8")
                for f in old_files:
                    f.unlink(missing_ok=True)
                print(f"Migrated {len(old_files)} files → shortcuts.json")
            except Exception as e:
                print(f"Migration error: {e}")
            return result
        if unified.exists():
            try:
                return json.loads(unified.read_text("utf-8"))
            except Exception as e:
                print(f"Load error: {e}")
        return []

    def _load_categories(self) -> list:
        f = self.app_dir / "categories.json"
        if f.exists():
            try:
                cats = json.loads(f.read_text("utf-8"))
                cats = [c for c in cats if c.get("id") != "uncategorized"]
                for c in cats:
                    if not any(c["name"].startswith(p) for p in ("📁","🏠","⭐","🔖","📂","📦")):
                        c["name"] = f"📁 {c['name']}"
                return cats
            except Exception:
                pass
        return []

    def _save_categories(self):
        f = self.app_dir / "categories.json"
        try:
            f.write_text(
                json.dumps([c for c in self.categories if c["id"] != "all"],
                           indent=2, ensure_ascii=False),
                "utf-8"
            )
        except Exception as e:
            print(f"Save categories error: {e}")

    def _save_shortcuts(self):
        data = json.dumps(self.shortcuts, indent=2, ensure_ascii=False)
        path = self.app_dir / "shortcuts.json"
        def _write():
            try:
                path.write_text(data, "utf-8")
            except Exception as e:
                print(f"Save error: {e}")
        threading.Thread(target=_write, daemon=True).start()

    def _schedule_save(self):
        self._save_timer.start(350)

    def _push_undo(self, action: str, data: dict):
        import copy
        self._undo_stack.append({
            "action": action,
            "data": copy.deepcopy(data),
            "shortcuts_snapshot": [dict(s) for s in self.shortcuts]
        })

    def _undo(self):
        if not self._undo_stack:
            return
        state = self._undo_stack.pop()
        self.shortcuts = state["shortcuts_snapshot"]
        self._schedule_save()
        self._refresh()

    def _start_favicon_loading(self):
        """Charge les favicons par batch de 10."""
        urls = [sc["url"] for sc in self.shortcuts
                if sc.get("type") == "url" and sc.get("url")]
        def _batch(batch_urls, delay_ms):
            def _do():
                for u in batch_urls:
                    self.async_loader.load(u, self.app_dir)
            QTimer.singleShot(delay_ms, _do)
        batch_size = 10
        for i, chunk_start in enumerate(range(0, len(urls), batch_size)):
            _batch(urls[chunk_start:chunk_start + batch_size], i * 200)

    # ══════════════════════════════════════════════
    #  UI BUILD
    # ══════════════════════════════════════════════
    def _build_ui(self):
        C = self.C
        self.setStyleSheet(f"""
            QMainWindow {{
                background: qlineargradient(
                    x1:0, y1:0, x2:0.4, y2:1,
                    stop:0 {C['primary_darker']},
                    stop:0.60 {C['primary']},
                    stop:1.0 {C['primary_light']}
                );
            }}
            QFrame#sidebarFrame {{
                background-color: {C['sidebar_bg']};
                border-right: 1px solid rgba(108,92,231,0.35);
            }}
            QListWidget {{
                background: transparent;
                border: none;
                color: {C['text_mid']};
                font-size: 13px;
                padding: 4px;
            }}
            QListWidget::item {{
                border-radius: 8px;
                padding: 0px;
                margin: 2px 0;
                background: transparent;
                color: transparent;
            }}
            QListWidget::item:selected {{
                background: transparent;
                color: transparent;
            }}
            QListWidget::item:hover:!selected {{
                background: transparent;
                color: transparent;
            }}
            #searchContainer {{
                background: {C['bg_glass']};
                border: 1px solid rgba(255,255,255,0.22);
                border-radius: 22px;
            }}
            QLineEdit {{
                background: transparent;
                border: none;
                color: white;
                font-size: 14px;
                padding: 6px 2px;
                selection-background-color: {C['primary']};
            }}
            QPushButton#radio {{
                background: transparent;
                border: none;
                color: {C['text_mid']};
                font-size: 13px;
                padding: 5px 14px;
                border-radius: 12px;
            }}
            QPushButton#radio:hover {{
                background: rgba(255,255,255,0.18);
                color: white;
            }}
            QPushButton#radio:checked {{
                background: rgba(255,255,255,0.28);
                color: white;
                font-weight: bold;
            }}
            #clockBox {{
                background: rgba(30,20,80,0.55);
                border: 1px solid rgba(108,92,231,0.45);
                border-radius: 14px;
            }}
            QScrollBar:vertical {{
                background: rgba(255,255,255,0.04);
                width: 7px; border-radius: 3px;
            }}
            QScrollBar::handle:vertical {{
                background: rgba(108,92,231,0.55);
                border-radius: 3px; min-height: 28px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: rgba(108,92,231,0.85);
            }}
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {{ height: 0; }}
            QDialog {{
                background: {C['dialog_bg']};
                border-radius: 14px;
            }}
            QDialog QLabel {{
                color: {C['text_mid']};
                font-size: 12px;
                background: transparent;
            }}
            QDialog QLineEdit {{
                background: {C['input_bg']};
                border: 1px solid {C['input_border']};
                border-radius: 7px;
                color: white;
                padding: 8px 12px;
                font-size: 13px;
            }}
            QDialog QLineEdit:focus {{
                border: 1.5px solid #c0b8ff;
            }}
            QDialog QPushButton {{
                background: {C['primary']};
                color: white; border: none;
                border-radius: 7px;
                padding: 8px 20px;
                font-weight: bold;
                font-size: 13px;
            }}
            QDialog QPushButton:hover {{ background: #7d6ff0; }}
            QDialog QPushButton#btnCancel {{
                background: rgba(255,255,255,0.08);
                color: {C['text_mid']};
            }}
            QDialog QPushButton#btnCancel:hover {{
                background: rgba(255,255,255,0.16);
                color: white;
            }}
            QMenu {{
                background: #16162a;
                border: 1px solid #6c5ce7;
                border-radius: 10px;
                padding: 5px;
                color: white;
                font-size: 13px;
            }}
            QMenu::item {{ padding: 8px 18px; border-radius: 6px; }}
            QMenu::item:selected {{ background: rgba(108,92,231,0.50); }}
            QMenu::separator {{
                height: 1px;
                background: rgba(108,92,231,0.30);
                margin: 3px 8px;
            }}
        """)

        central = QWidget()
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_sidebar())
        root.addWidget(self._build_content(), 1)

        self.setCentralWidget(central)
        self._refresh()
        self._filter_cat_list()

    # ── Sidebar ──────────────────────────────────────────────
    def _build_sidebar(self) -> QFrame:
        C = self.C
        sb = QFrame()
        sb.setObjectName("sidebarFrame")
        sb.setFixedWidth(258)

        ly = QVBoxLayout(sb)
        ly.setContentsMargins(16, 20, 16, 20)
        ly.setSpacing(10)

        hdr = QHBoxLayout()
        hdr.setSpacing(8)
        lbl = QLabel("Spaces")
        lbl.setStyleSheet("color:white;font-weight:bold;font-size:15px;background:transparent;")
        btn_add = QPushButton("+")
        btn_add.setFixedSize(24, 24)
        btn_add.setToolTip("New space")
        btn_add.setStyleSheet(f"""
            QPushButton {{
                background:{C['primary']}; border-radius:12px;
                color:white; font-weight:bold; font-size:16px;
            }}
            QPushButton:hover {{ background:#7d6ff0; }}
        """)
        btn_add.clicked.connect(self._add_category)
        hdr.addWidget(lbl)
        hdr.addStretch()
        hdr.addWidget(btn_add)
        ly.addLayout(hdr)

        cat_sb = QWidget()
        cat_sb.setFixedHeight(32)
        cat_sb.setStyleSheet(
            "background:rgba(255,255,255,0.08);"
            "border:1px solid rgba(255,255,255,0.15);"
            "border-radius:16px;"
        )
        csb_ly = QHBoxLayout(cat_sb)
        csb_ly.setContentsMargins(12, 0, 12, 0)
        self.cat_search = QLineEdit()
        self.cat_search.setPlaceholderText("Filter spaces…")
        self.cat_search.setStyleSheet("background:transparent;border:none;color:white;font-size:12px;")
        self.cat_search.textChanged.connect(self._filter_cat_list)
        csb_ly.addWidget(self.cat_search, 1)
        ly.addWidget(cat_sb)

        self.cat_list = CategoryListWidget()
        self.cat_list.setItemDelegate(CategoryCountDelegate(self.cat_list))
        self.cat_list.shortcutDroppedOnCategory.connect(self._on_drop_on_category)
        self.cat_list.model().rowsMoved.connect(self._on_cats_reordered)
        self.cat_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.cat_list.customContextMenuRequested.connect(self._cat_context_menu)
        for cat in self.categories:
            item = QListWidgetItem(cat["name"])
            item.setData(Qt.UserRole, cat["id"])
            item.setData(Qt.UserRole + 1, 0)
            self.cat_list.addItem(item)
        self.cat_list.setCurrentRow(0)
        self.cat_list.itemClicked.connect(self._on_cat_clicked)
        ly.addWidget(self.cat_list, 1)

        self.lbl_count = QLabel()
        self.lbl_count.setAlignment(Qt.AlignCenter)
        self.lbl_count.setStyleSheet(
            "color:rgba(255,255,255,0.52);font-size:11px;background:transparent;"
        )
        ly.addWidget(self.lbl_count)

        return sb

    # ── Contenu principal ────────────────────────────────────
    def _build_content(self) -> QWidget:
        C = self.C
        w = QWidget()
        ly = QVBoxLayout(w)
        ly.setContentsMargins(22, 16, 22, 16)
        ly.setSpacing(12)

        hdr = QWidget()
        hdr_ly = QHBoxLayout(hdr)
        hdr_ly.setContentsMargins(0, 0, 0, 0)
        hdr_ly.setSpacing(16)

        left = QWidget()
        left_ly = QVBoxLayout(left)
        left_ly.setContentsMargins(0, 0, 0, 0)
        left_ly.setSpacing(6)

        sc_outer = QWidget()
        sc_outer.setObjectName("searchContainer")
        sc_outer.setFixedHeight(44)
        sc_outer_ly = QHBoxLayout(sc_outer)
        sc_outer_ly.setContentsMargins(14, 0, 8, 0)
        sc_outer_ly.setSpacing(8)

        self._ico_search = QLabel("🔍")
        self._ico_search.setStyleSheet(
            "color:rgba(255,255,255,0.65);font-size:15px;background:transparent;"
        )

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search on Google…")
        self.search_bar.returnPressed.connect(self._handle_search)
        self.search_bar.textChanged.connect(self._on_search_changed)
        self.search_bar.installEventFilter(self)

        self._btn_clear = QToolButton()
        self._btn_clear.setText("✕")
        self._btn_clear.setStyleSheet(
            "QToolButton{background:transparent;color:rgba(255,255,255,0.45);font-size:12px;border:none;}"
            "QToolButton:hover{color:white;}"
        )
        self._btn_clear.setVisible(False)
        self._btn_clear.clicked.connect(self._clear_search)
        self.search_bar.textChanged.connect(lambda t: self._btn_clear.setVisible(bool(t)))

        self._pill = QWidget()
        self._pill.setFixedHeight(28)
        self._pill.setStyleSheet("background:rgba(255,255,255,0.12);border-radius:14px;")
        pill_ly = QHBoxLayout(self._pill)
        pill_ly.setContentsMargins(3, 3, 3, 3)
        pill_ly.setSpacing(0)

        def pill_btn(text):
            b = QPushButton(text)
            b.setCheckable(True)
            b.setFixedHeight(22)
            b.setStyleSheet("""
                QPushButton {
                    background: transparent; border: none;
                    color: rgba(255,255,255,0.65);
                    font-size: 11px; padding: 0 10px; border-radius: 11px;
                }
                QPushButton:checked {
                    background: rgba(255,255,255,0.25);
                    color: white; font-weight: bold;
                }
                QPushButton:hover:!checked { color: white; }
            """)
            return b

        self.btn_google = pill_btn("Google")
        self.btn_google.setChecked(True)
        self.btn_fb     = pill_btn("FlowBoard")
        self.btn_google.clicked.connect(lambda: self._set_mode("google"))
        self.btn_fb.clicked.connect(lambda: self._set_mode("flowboard"))

        self.grp_radio = QButtonGroup(self)
        self.grp_radio.setExclusive(True)
        self.grp_radio.addButton(self.btn_google)
        self.grp_radio.addButton(self.btn_fb)

        pill_ly.addWidget(self.btn_google)
        pill_ly.addWidget(self.btn_fb)

        sc_outer_ly.addWidget(self._ico_search)
        sc_outer_ly.addWidget(self.search_bar, 1)
        sc_outer_ly.addWidget(self._pill)
        sc_outer_ly.addWidget(self._btn_clear)
        left_ly.addWidget(sc_outer)

        row2 = QWidget()
        row2_ly = QHBoxLayout(row2)
        row2_ly.setContentsMargins(0, 0, 0, 0)
        row2_ly.setSpacing(6)

        self.btn_add = QPushButton("＋  Add shortcut")
        self.btn_add.setFixedHeight(32)
        self.btn_add.setStyleSheet(f"""
            QPushButton {{
                background:{C['primary']};
                color:white; font-weight:bold; font-size:13px;
                padding:0 18px; border-radius:16px; border:none;
            }}
            QPushButton:hover {{ background:#7d6ff0; }}
            QPushButton:pressed {{ background:#5a4bcc; }}
        """)
        self.btn_add.clicked.connect(self._show_add_dialog)

        row2_ly.addStretch()
        row2_ly.addWidget(self.btn_add)
        left_ly.addWidget(row2)

        clock_box = QWidget()
        clock_box.setObjectName("clockBox")
        clock_box.setFixedSize(132, 52)
        cb_ly = QVBoxLayout(clock_box)
        cb_ly.setContentsMargins(0, 4, 0, 4)
        cb_ly.setSpacing(0)
        self.lbl_clock = QLabel()
        self.lbl_clock.setAlignment(Qt.AlignCenter)
        self.lbl_clock.setTextFormat(Qt.RichText)
        cb_ly.addWidget(self.lbl_clock)

        hdr_ly.addWidget(left, 1)
        hdr_ly.addWidget(clock_box, 0, Qt.AlignVCenter)
        ly.addWidget(hdr)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("background:transparent;border:none;")

        self.grid_widget = ShortcutGridWidget()
        self.grid        = QGridLayout(self.grid_widget)
        self.grid.setSpacing(12)
        self.grid.setContentsMargins(6, 6, 6, 6)
        self.grid_widget.shortcutDropped.connect(self._on_drop_in_grid)

        scroll.setWidget(self.grid_widget)
        ly.addWidget(scroll, 1)

        return w

    def _bind_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+F"), self).activated.connect(self._focus_search)
        QShortcut(QKeySequence("Escape"), self).activated.connect(self._clear_search)
        QShortcut(QKeySequence("Ctrl+N"), self).activated.connect(self._show_add_dialog)
        QShortcut(QKeySequence("Ctrl+Z"), self).activated.connect(self._undo)

    # ── Resize → colonnes dynamiques ─────────────────────────
    def resizeEvent(self, event):
        super().resizeEvent(event)
        available = self.width() - 258 - 44 - 12
        cols = max(1, available // CARD_STRIDE)
        if cols != self._cols:
            self._cols = cols
            if not hasattr(self, "_resize_timer"):
                self._resize_timer = QTimer(self)
                self._resize_timer.setSingleShot(True)
                self._resize_timer.timeout.connect(self._refresh)
            self._resize_timer.start(80)

    # ══════════════════════════════════════════════
    #  CATEGORIES
    # ══════════════════════════════════════════════
    def _filter_cat_list(self):
        q = self.cat_search.text().strip().lower()
        for i in range(self.cat_list.count()):
            item = self.cat_list.item(i)
            if item.data(Qt.UserRole) == "all":
                item.setHidden(False)
                continue
            raw = item.text()
            name = raw.split(" ", 1)[-1].lower() if " " in raw else raw.lower()
            item.setHidden(q not in name)

    def _on_cat_clicked(self, item: QListWidgetItem):
        self.current_category = item.data(Qt.UserRole)
        self._refresh()

    def _on_cats_reordered(self, *_):
        new = []
        for i in range(self.cat_list.count()):
            cid = self.cat_list.item(i).data(Qt.UserRole)
            c = next((x for x in self.categories if x["id"] == cid), None)
            if c:
                new.append(c)
        all_c = next((x for x in new if x["id"] == "all"), None)
        if all_c:
            new.remove(all_c)
            new.insert(0, {"id": "all", "name": "🏠 All"})
        self.categories = new
        self._save_categories()
        self._filter_cat_list()

    def _add_category(self):
        name, ok = QInputDialog.getText(self, "New Space", "Space name:")
        if not ok or not name.strip():
            return
        base_cid = re.sub(r"[^a-z0-9_]", "_", name.lower().strip()) or "cat"
        existing_ids = {c["id"] for c in self.categories}
        cid = base_cid
        counter = 2
        while cid in existing_ids:
            cid = f"{base_cid}_{counter}"
            counter += 1
        display = f"📁 {name.strip()}"
        self.categories.append({"id": cid, "name": display})
        self.cat_list.model().rowsMoved.disconnect(self._on_cats_reordered)
        item = QListWidgetItem(display)
        item.setData(Qt.UserRole, cid)
        item.setData(Qt.UserRole + 1, 0)
        self.cat_list.addItem(item)
        self.cat_list.model().rowsMoved.connect(self._on_cats_reordered)
        self._save_categories()
        self._filter_cat_list()

    def _rename_category(self, item, cid, current):
        clean = current.split(" ", 1)[-1] if " " in current else current
        new, ok = QInputDialog.getText(self, "Rename Space", "New name:", text=clean)
        if not ok or not new.strip():
            return
        display = f"📁 {new.strip()}"
        for c in self.categories:
            if c["id"] == cid:
                c["name"] = display
                break
        item.setText(display)
        self._save_categories()
        self._filter_cat_list()

    def _delete_category(self, item, cid, display):
        n = sum(1 for s in self.shortcuts if s.get("category") == cid)
        msg = (f"Delete space '{display}'?\n{n} shortcut(s) will appear in All only."
               if n else f"Delete space '{display}'?")
        if QMessageBox.question(self, "Confirm", msg,
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        for s in self.shortcuts:
            if s.get("category") == cid:
                s["category"] = ""
        self.cat_list.model().rowsMoved.disconnect(self._on_cats_reordered)
        self.cat_list.takeItem(self.cat_list.row(item))
        self.cat_list.model().rowsMoved.connect(self._on_cats_reordered)
        self.categories = [c for c in self.categories if c["id"] != cid]
        self._schedule_save()
        self._refresh()
        self._save_categories()
        self._filter_cat_list()

    def _cat_context_menu(self, pos):
        item = self.cat_list.itemAt(pos)
        if not item or item.data(Qt.UserRole) == "all":
            return
        cid  = item.data(Qt.UserRole)
        name = item.text()
        menu = QMenu(self)
        a_ren = menu.addAction("✏️  Rename")
        a_del = menu.addAction("🗑️  Delete")
        act = menu.exec(self.cat_list.mapToGlobal(pos))
        if act == a_ren:
            self._rename_category(item, cid, name)
        elif act == a_del:
            self._delete_category(item, cid, name)

    def _on_drop_on_category(self, cid: str, sc: dict):
        for s in self.shortcuts:
            if s.get("id") == sc.get("id"):
                self._push_undo("move_category", s)
                s["category"] = cid
                self._schedule_save()
                self._refresh()
                break

    # ══════════════════════════════════════════════
    #  GRILLE
    # ══════════════════════════════════════════════
    def _filtered(self) -> list:
        lst = (
            self.shortcuts if self.current_category == "all"
            else [s for s in self.shortcuts if s.get("category") == self.current_category]
        )
        if self.search_mode == "flowboard" and self.current_search:
            q = self.current_search
            lst = [
                s for s in lst
                if q in s.get("title", "").lower()
                or q in s.get("url", "").lower()
                or q in s.get("path", "").lower()
                or q in s.get("note", "").lower()
            ]
        return lst

    def _refresh(self):
        items = self._filtered()
        cols  = self._cols
        used  = set()

        current_positions: dict[str, tuple[int, int]] = {}
        for sid, w in self._card_widgets.items():
            idx = self.grid.indexOf(w)
            if idx != -1:
                r, c, _, _ = self.grid.getItemPosition(idx)
                current_positions[sid] = (r, c)

        for i, sc in enumerate(items):
            sid = sc["id"]
            used.add(sid)
            row, col = i // cols, i % cols

            widget = self._card_widgets.get(sid)
            if widget:
                if widget.text() != sc.get("title", "Untitled") or widget._has_note != bool(sc.get("note")):
                    widget.shortcutData = sc
                    widget.setText(sc.get("title", "Untitled"))
                    widget._has_note = bool(sc.get("note"))
                    widget.update()
                else:
                    widget.shortcutData = sc
            else:
                widget = self._make_card(sc)
                self._card_widgets[sid] = widget
                current_positions[sid] = (-1, -1)

            cur_pos = current_positions.get(sid, (-1, -1))
            if cur_pos == (-1, -1) or cur_pos != (row, col):
                self.grid.addWidget(widget, row, col)

        for sid, widget in list(self._card_widgets.items()):
            if sid not in used:
                idx = self.grid.indexOf(widget)
                if idx != -1:
                    it = self.grid.takeAt(idx)
                    if it and it.widget():
                        it.widget().deleteLater()
                del self._card_widgets[sid]

        counts: dict[str, int] = {}
        for s in self.shortcuts:
            cid = s.get("category", "") or ""
            counts[cid] = counts.get(cid, 0) + 1

        changed = False
        for i in range(self.cat_list.count()):
            list_item = self.cat_list.item(i)
            cid = list_item.data(Qt.UserRole)
            n = 0 if cid == "all" else counts.get(cid, 0)
            if list_item.data(Qt.UserRole + 1) != n:
                list_item.setData(Qt.UserRole + 1, n)
                changed = True
        if changed:
            self.cat_list.viewport().update()

        total = len(self.shortcuts)
        vis   = len(items)
        self.lbl_count.setText(
            f"Showing {vis} of {total}" if vis < total
            else f"{total} shortcut{'s' if total != 1 else ''}"
        )

    def _reorder_grid(self):
        cols = self._cols
        for i, sc in enumerate(self._filtered()):
            w = self._card_widgets.get(sc["id"])
            if w:
                idx = self.grid.indexOf(w)
                if idx != -1:
                    cr, cc, _, _ = self.grid.getItemPosition(idx)
                    tr, tc = i // cols, i % cols
                    if cr != tr or cc != tc:
                        self.grid.addWidget(w, tr, tc)

    def _on_drop_in_grid(self, row: int, col: int, sc: dict):
        tgt = row * self._cols + col
        src = next((i for i, s in enumerate(self.shortcuts) if s.get("id") == sc.get("id")), None)
        if src is None or src == tgt:
            return
        self._push_undo("reorder", sc)
        obj = self.shortcuts.pop(src)
        self.shortcuts.insert(tgt, obj)
        for i, s in enumerate(self.shortcuts):
            s["position"] = i
        self._reorder_grid()
        self._schedule_save()

    # ══════════════════════════════════════════════
    #  CARTE
    # ══════════════════════════════════════════════
    def _make_card(self, sc: dict) -> DraggableShortcutCard:
        btn = DraggableShortcutCard()
        btn.shortcutData = sc
        btn._has_note    = bool(sc.get("note"))
        btn.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        btn.setIconSize(QSize(46, 46))
        btn.setFixedSize(CARD_SIZE, CARD_SIZE)
        btn.setText(sc.get("title", "Untitled"))
        btn._update_style()

        # ── Icône : toujours dans le thread UI, placeholder immédiat ──
        btn.setIcon(get_shortcut_icon(sc, self.app_dir, self.async_loader, self.native_worker))

        if sc.get("type") == "url":
            url = sc.get("url", "")
            btn.doubleClicked.connect(lambda u=url: QDesktopServices.openUrl(QUrl(u)))
        elif sc.get("path"):
            path = sc.get("path", "")
            btn.doubleClicked.connect(lambda p=path: _open_path(p, self))

        btn.setContextMenuPolicy(Qt.CustomContextMenu)
        btn.customContextMenuRequested.connect(
            lambda pos, s=sc, b=btn: self._card_context_menu(pos, s, b)
        )
        btn.clicked.connect(lambda checked=False, b=btn: self._select_card(b))

        return btn

    def _select_card(self, btn: DraggableShortcutCard):
        if self._selected_card and self._selected_card is not btn:
            self._selected_card.setSelected(False)
        if self._selected_card is btn:
            btn.setSelected(False)
            self._selected_card = None
        else:
            btn.setSelected(True)
            self._selected_card = btn

    # ══════════════════════════════════════════════
    #  DRAG & DROP fenêtre
    # ══════════════════════════════════════════════
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        if self.is_processing:
            return
        self.is_processing = True
        try:
            for url in event.mimeData().urls():
                p = url.toLocalFile()
                if p and os.path.exists(p):
                    self._add_from_path(p)
                elif url.toString().startswith(("http://", "https://")):
                    self._add_from_url(url.toString())
            self._schedule_save()
            self._refresh()
        finally:
            self.is_processing = False
            event.acceptProposedAction()

    # ══════════════════════════════════════════════
    #  RECHERCHE
    # ══════════════════════════════════════════════
    def _set_mode(self, mode: str):
        self.search_mode = mode
        if mode == "google":
            self.search_bar.setPlaceholderText("Search on Google…")
            self._ico_search.setText("🔍")
            self.btn_google.setChecked(True)
            self.current_search = ""
            self._refresh()
        else:
            self.search_bar.setPlaceholderText("Search shortcuts (title, URL, path, note)…")
            self._ico_search.setText("🔎")
            self.btn_fb.setChecked(True)
            self._filter_shortcuts(self.search_bar.text())

    def _on_search_changed(self, text: str):
        if self.search_mode == "flowboard":
            self._filter_shortcuts(text)

    def _filter_shortcuts(self, q: str):
        self.current_search = q.strip().lower()
        self._refresh()

    def _handle_search(self):
        if self.search_mode == "google":
            q = self.search_bar.text().strip()
            if q:
                webbrowser.open_new_tab(
                    f"https://www.google.com/search?q={urllib.parse.quote_plus(q)}"
                )
                self.search_bar.clear()

    def _clear_search(self):
        self.search_bar.clear()
        if self.search_mode == "flowboard":
            self.current_search = ""
            self._refresh()

    def _focus_search(self):
        self.search_bar.setFocus()
        self.search_bar.selectAll()

    # ══════════════════════════════════════════════
    #  DIALOGS
    # ══════════════════════════════════════════════
    def _make_dialog(self, title: str):
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.setMinimumWidth(440)

        ly = QVBoxLayout(dlg)
        ly.setContentsMargins(26, 24, 26, 22)
        ly.setSpacing(14)

        lbl_h = QLabel(title)
        lbl_h.setStyleSheet("color:white;font-size:15px;font-weight:bold;background:transparent;")
        ly.addWidget(lbl_h)

        def row(label, ph="", required=False):
            rl = QVBoxLayout()
            lbl_row = QLabel(("* " if required else "") + label)
            if required:
                lbl_row.setStyleSheet("color:#c0b8ff;font-size:11px;background:transparent;")
            inp = QLineEdit()
            inp.setPlaceholderText(ph)
            rl.addWidget(lbl_row)
            rl.addWidget(inp)
            ly.addLayout(rl)
            return inp

        url_in  = row("URL or path", "https://example.com  or  C:\\path\\to\\file", required=True)
        name_in = row("Name", "Leave empty for auto-detect")
        note_in = row("Note", "Visible on hover")

        btns = QHBoxLayout()
        btns.setSpacing(10)
        btn_cancel  = QPushButton("Cancel")
        btn_cancel.setObjectName("btnCancel")
        btn_confirm = QPushButton("Add")
        btn_confirm.setObjectName("btnConfirm")
        btn_cancel.clicked.connect(dlg.reject)
        btns.addStretch()
        btns.addWidget(btn_cancel)
        btns.addWidget(btn_confirm)
        ly.addLayout(btns)

        dlg.adjustSize()
        return dlg, url_in, name_in, note_in

    def _show_add_dialog(self):
        dlg, url_in, name_in, note_in = self._make_dialog("Add shortcut")
        url_in.setText("https://")
        url_in.setCursorPosition(len("https://"))
        btn = dlg.findChild(QPushButton, "btnConfirm")
        btn.clicked.connect(lambda: self._save_shortcut(url_in.text(), name_in.text(), note_in.text(), dlg))
        dlg.exec()

    def _edit_shortcut(self, sc: dict):
        dlg, url_in, name_in, note_in = self._make_dialog("Edit shortcut")
        url_in.setText(sc.get("url", sc.get("path", "")))
        name_in.setText(sc.get("title", ""))
        note_in.setText(sc.get("note", ""))
        btn = dlg.findChild(QPushButton, "btnConfirm")
        btn.setText("Save")
        btn.clicked.connect(lambda: self._update_shortcut(sc, url_in.text(), name_in.text(), note_in.text(), dlg))
        dlg.exec()

    def _save_shortcut(self, text: str, name: str, note: str, dlg: QDialog):
        text = text.strip()
        if not text or text == "https://":
            return
        if not text.startswith(("http://", "https://")):
            if "." in text and " " not in text and not os.path.exists(text):
                text = "https://" + text
        now = datetime.now().isoformat()
        sc  = {"id": now, "created_at": now, "updated_at": now, "position": len(self.shortcuts)}
        if text.startswith(("http://", "https://")):
            auto = urlparse(text).netloc or text.split("//")[-1].split("/")[0]
            sc.update({"url": text, "title": name.strip() or auto,
                       "note": note.strip(), "type": "url",
                       "category": self.current_category})
        elif os.path.exists(text):
            is_dir = os.path.isdir(text)
            base   = os.path.basename(text)
            auto   = base.rsplit(".", 1)[0] if not is_dir and "." in base else base
            sc.update({"path": text, "title": name.strip() or auto,
                       "note": note.strip(),
                       "type": "folder" if is_dir else "file",
                       "category": self.current_category})
        else:
            QMessageBox.warning(self, "Error", f"Invalid URL or path:\n{text}")
            return
        self.shortcuts.append(sc)
        self._schedule_save()
        self._refresh()
        dlg.accept()
        if sc.get("type") == "url":
            self.async_loader.load(sc["url"], self.app_dir)
        elif sc.get("path"):
            # Déclenche la résolution async de l'icône native
            self.native_worker.request(sc["path"])

    def _update_shortcut(self, sc: dict, text: str, name: str, note: str, dlg: QDialog):
        text = text.strip()
        if not text:
            return
        if not text.startswith(("http://", "https://")):
            if "." in text and " " not in text and not os.path.exists(text):
                text = "https://" + text
        self._push_undo("edit", sc)
        if text.startswith(("http://", "https://")):
            auto = urlparse(text).netloc or text.split("//")[-1].split("/")[0]
            sc["url"] = text; sc["type"] = "url"; sc.pop("path", None)
        elif os.path.exists(text):
            auto = os.path.basename(text).rsplit(".", 1)[0]
            sc["path"] = text
            sc["type"] = "folder" if os.path.isdir(text) else "file"
            sc.pop("url", None)
        else:
            QMessageBox.warning(self, "Error", f"Invalid URL or path:\n{text}")
            return
        sc["title"]      = name.strip() or auto
        sc["note"]       = note.strip()
        sc["updated_at"] = datetime.now().isoformat()
        self._schedule_save()
        self._refresh()
        dlg.accept()
        if sc.get("type") == "url":
            self.async_loader.load(sc["url"], self.app_dir)
        elif sc.get("path"):
            self.native_worker.request(sc["path"])

    def _delete_shortcut(self, sc: dict):
        if QMessageBox.question(
            self, "Confirm",
            f"Delete '{sc.get('title', 'Untitled')}'?",
            QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            self._push_undo("delete", sc)
            self.shortcuts.remove(sc)
            self._schedule_save()
            self._refresh()

    def _move_shortcut(self, sc: dict):
        choices = [c["name"] for c in self.categories if c["id"] != "all"]
        cid     = sc.get("category", "")
        try:
            idx = next(i for i, c in enumerate(self.categories)
                       if c["id"] == cid and c["id"] != "all")
        except StopIteration:
            idx = 0
        chosen, ok = QInputDialog.getItem(self, "Move to space", "Space:", choices, idx, False)
        if ok and chosen:
            target = next(
                (c["id"] for c in self.categories if c["name"] == chosen and c["id"] != "all"),
                ""
            )
            self._push_undo("move_category", sc)
            sc["category"] = target
            self._schedule_save()
            self._refresh()

    def _card_context_menu(self, pos, sc: dict, btn: QToolButton):
        menu = QMenu(self)
        a_open  = menu.addAction("↗️  Open")
        menu.addSeparator()
        a_edit  = menu.addAction("✏️  Edit")
        a_move  = menu.addAction("➡️  Move to space…")
        menu.addSeparator()
        a_del   = menu.addAction("🗑️  Delete")
        act = menu.exec(btn.mapToGlobal(pos))
        if act == a_open:
            if sc.get("type") == "url":
                QDesktopServices.openUrl(QUrl(sc.get("url", "")))
            elif sc.get("path"):
                _open_path(sc["path"], self)
        elif act == a_edit:
            self._edit_shortcut(sc)
        elif act == a_move:
            self._move_shortcut(sc)
        elif act == a_del:
            self._delete_shortcut(sc)

    # ══════════════════════════════════════════════
    #  DRAG & DROP — ajout depuis l'extérieur
    # ══════════════════════════════════════════════
    def _add_from_path(self, path: str):
        ep = path
        if sys.platform == "win32" and path.lower().endswith(".lnk"):
            r = _resolve_shortcut_raw(path)
            if os.path.exists(r):
                ep = r
        is_dir = os.path.isdir(ep)
        base   = os.path.basename(ep)
        title  = base.rsplit(".", 1)[0] if not is_dir and "." in base else base
        now    = datetime.now().isoformat()
        self.shortcuts.append({
            "id": now, "created_at": now, "updated_at": now,
            "position": len(self.shortcuts),
            "path": ep, "title": title, "note": "",
            "type": "folder" if is_dir else "file",
            "category": self.current_category,
        })
        # Déclenche résolution icône native immédiatement
        self.native_worker.request(ep)

    def _add_from_url(self, url: str):
        url = url.strip()
        if not url.startswith(("http://", "https://")):
            if "." in url and " " not in url:
                url = "https://" + url
        title = urlparse(url).netloc or url.split("//")[-1].split("/")[0]
        now   = datetime.now().isoformat()
        self.shortcuts.append({
            "id": now, "created_at": now, "updated_at": now,
            "position": len(self.shortcuts),
            "url": url, "title": title, "note": "",
            "type": "url", "category": self.current_category,
        })
        self.async_loader.load(url, self.app_dir)

    # ══════════════════════════════════════════════
    #  HORLOGE
    # ══════════════════════════════════════════════
    def update_clock(self):
        from PySide6.QtCore import QDateTime, QLocale
        now  = QDateTime.currentDateTime()
        time = now.toString("HH:mm:ss")
        loc  = QLocale(QLocale.English, QLocale.UnitedStates)
        date = loc.toString(now, "ddd d MMM").replace(".", "").strip()
        self.lbl_clock.setText(
            f'<div style="color:white;font-size:17px;font-weight:bold;line-height:1.2;">{time}</div>'
            f'<div style="color:rgba(255,255,255,0.70);font-size:10px;">{date}</div>'
        )
        QTimer.singleShot(1000, self.update_clock)

    # ══════════════════════════════════════════════
    #  EVENT FILTER
    # ══════════════════════════════════════════════
    def eventFilter(self, obj, event):
        if obj is self.search_bar:
            sc = self.findChild(QWidget, "searchContainer")
            if sc:
                if event.type() == QEvent.FocusIn:
                    sc.setStyleSheet(
                        "#searchContainer{background:rgba(255,255,255,0.20);"
                        "border:1px solid rgba(255,255,255,0.65);border-radius:22px;}"
                    )
                elif event.type() == QEvent.FocusOut:
                    sc.setStyleSheet(
                        "#searchContainer{background:rgba(255,255,255,0.10);"
                        "border:1px solid rgba(255,255,255,0.22);border-radius:22px;}"
                    )
        return super().eventFilter(obj, event)


# ══════════════════════════════════════════════════════════════
#  LANCEMENT
# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    if (ico := Path("icon.ico")).exists():
        app.setWindowIcon(QIcon(str(ico)))

    splash = QuickSplash()
    splash.show()
    app.processEvents()

    dark = QPalette()
    dark.setColor(QPalette.Window,          QColor(22, 22, 42))
    dark.setColor(QPalette.WindowText,      QColor(255, 255, 255))
    dark.setColor(QPalette.Base,            QColor(18, 18, 36))
    dark.setColor(QPalette.AlternateBase,   QColor(35, 35, 55))
    dark.setColor(QPalette.Text,            QColor(255, 255, 255))
    dark.setColor(QPalette.Button,          QColor(45, 45, 70))
    dark.setColor(QPalette.ButtonText,      QColor(255, 255, 255))
    dark.setColor(QPalette.BrightText,      QColor(255, 80, 80))
    dark.setColor(QPalette.Link,            QColor(108, 92, 231))
    dark.setColor(QPalette.Highlight,       QColor(108, 92, 231))
    dark.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    dark.setColor(QPalette.ToolTipBase,     QColor(22, 22, 42))
    dark.setColor(QPalette.ToolTipText,     QColor(200, 200, 230))
    app.setPalette(dark)

    window = FlowBoard()
    QTimer.singleShot(450, lambda: (splash.finish(window), window.show()))
    sys.exit(app.exec())
import subprocess
import datetime
import os
import json
import urllib.parse
import sys
import webbrowser

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QFrame, QDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal


# ==========================
#   UTILS
# ==========================

def run_powershell(cmd):
    try:
        result = subprocess.run(
            ["powershell", "-NoLogo", "-NoProfile", "-Command", cmd],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore"
        )
        return result.stdout.strip()
    except:
        return ""


def make_search_url(name):
    if not name or str(name).lower() == "unknown":
        return ""
    return f"https://www.google.com/search?q={urllib.parse.quote_plus(str(name))}"


# ==========================
#   HARDWARE DETECTION
# ==========================

def get_cpu():
    cmd = "Get-CimInstance Win32_Processor | Select-Object Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors | ConvertTo-Json"
    out = run_powershell(cmd)

    try:
        data = json.loads(out) if out else []
        if isinstance(data, dict):
            data = [data]
        lst = []
        for c in data:
            lst.append({
                "name": c.get("Name", "Unknown"),
                "manufacturer": c.get("Manufacturer", "Unknown"),
                "cores": c.get("NumberOfCores", "Unknown"),
                "threads": c.get("NumberOfLogicalProcessors", "Unknown"),
                "url": make_search_url(c.get("Name"))
            })
        return lst
    except:
        return []


def get_gpu():
    cmd = "Get-CimInstance Win32_VideoController | Select-Object Name,AdapterCompatibility,DriverVersion | ConvertTo-Json"
    out = run_powershell(cmd)

    try:
        data = json.loads(out) if out else []
        if isinstance(data, dict):
            data = [data]
        lst = []
        for g in data:
            lst.append({
                "name": g.get("Name", "Unknown"),
                "vendor": g.get("AdapterCompatibility", "Unknown"),
                "driver": g.get("DriverVersion", "Unknown"),
                "url": make_search_url(g.get("Name"))
            })
        return lst
    except:
        return []


def get_ram():
    out = run_powershell("(Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory")
    try:
        return round(int(out) / (1024**3))
    except:
        return 0


def get_disks():
    cmd = "Get-CimInstance Win32_DiskDrive | Select-Object Model,MediaType,Size | ConvertTo-Json"
    out = run_powershell(cmd)

    try:
        data = json.loads(out) if out else []
        if isinstance(data, dict):
            data = [data]
        lst = []
        for d in data:
            size = 0
            try:
                size = round(int(d.get("Size", 0)) / (1024**3))
            except:
                pass

            lst.append({
                "model": d.get("Model", "Unknown"),
                "type": d.get("MediaType", "Unknown"),
                "size": size,
                "url": make_search_url(d.get("Model"))
            })
        return lst
    except:
        return []


def get_motherboard():
    cmd = "Get-CimInstance Win32_BaseBoard | Select-Object Manufacturer,Product | ConvertTo-Json"
    out = run_powershell(cmd)

    try:
        data = json.loads(out) if out else {}
        name = f"{data.get('Manufacturer','')} {data.get('Product','')}".strip()
        if not name:
            name = "Unknown"
        return {"name": name, "url": make_search_url(name)}
    except:
        return {"name": "Unknown", "url": ""}


def get_fans():
    cmd = "Get-CimInstance Win32_Fan | Select-Object Name | ConvertTo-Json"
    out = run_powershell(cmd)

    try:
        if not out:
            return []
        data = json.loads(out)
        if isinstance(data, dict):
            data = [data]
        return [x.get("Name", "Unknown") for x in data]
    except:
        return []


# ==========================
#   MONITORES COMPLETOS
# ==========================

def get_monitors():
    cmd_basic = "Get-CimInstance Win32_DesktopMonitor | Select-Object Name,PNPDeviceID,ScreenWidth,ScreenHeight | ConvertTo-Json"
    out_basic = run_powershell(cmd_basic)

    try:
        basic = json.loads(out_basic) if out_basic else []
        if isinstance(basic, dict):
            basic = [basic]
    except:
        basic = []

    cmd_friendly = r"""
    Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID |
    Select-Object InstanceName,
        @{Name='FriendlyName';Expression={ ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_]) -join '' }},
        @{Name='Manufacturer';Expression={ ($_.ManufacturerName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_]) -join '' }},
        @{Name='Serial';Expression={ ($_.SerialNumberID | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_]) -join '' }} |
    ConvertTo-Json
    """
    out_friendly = run_powershell(cmd_friendly)

    try:
        friendly = json.loads(out_friendly) if out_friendly else []
        if isinstance(friendly, dict):
            friendly = [friendly]
    except:
        friendly = []

    final = []
    for b in basic:
        pnp = (b.get("PNPDeviceID") or "").upper()
        width = b.get("ScreenWidth")
        height = b.get("ScreenHeight")

        real_name = "Unknown"
        vendor = "Unknown"
        serial = "Unknown"

        for f in friendly:
            inst = f.get("InstanceName", "").upper()
            if pnp and pnp in inst:
                real_name = f.get("FriendlyName", real_name)
                vendor = f.get("Manufacturer", vendor)
                serial = f.get("Serial", serial)
                break

        final.append({
            "name": real_name,
            "vendor": vendor,
            "serial": serial,
            "width": width,
            "height": height,
            "url": make_search_url(real_name)
        })

    return final


# ==========================
#   RED
# ==========================

def get_ip_local():
    cmd = "(Get-NetIPAddress | Where-Object {$_.AddressFamily -eq 'IPv4' -and $_.IPAddress -notlike '169.*'}).IPAddress"
    return run_powershell(cmd)


def get_ip_public():
    return run_powershell("(Invoke-RestMethod 'https://api.ipify.org')")


# ===========================================
#       GENERAR REPORTE
# ===========================================

def generate_report():
    now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        desktop = os.getcwd()
    file_path = os.path.join(desktop, f"PC_INFO_{now}.txt")

    cpu = get_cpu()
    gpu = get_gpu()
    ram = get_ram()
    disks = get_disks()
    mb = get_motherboard()
    fans = get_fans()
    monitors = get_monitors()
    ip_local = get_ip_local()
    ip_public = get_ip_public()

    with open(file_path, "w", encoding="utf-8") as f:

        f.write("======== PC INFO SCANNER (GUI) ========\n")
        f.write("=========== Hardware Report ===========\n\n")

        f.write(f"IP Local  : {ip_local}\n")
        f.write(f"IP Pública: {ip_public}\n\n")

        # CPU
        f.write("\n=== CPU ===\n")
        if cpu:
            for c in cpu:
                f.write(f"{c['name']} ({c['cores']}C/{c['threads']}T)\n")
                f.write(f"Fabricante: {c['manufacturer']}\n")
                f.write(f"URL: {c['url']}\n\n")

        # GPU
        f.write("\n=== GPU ===\n")
        for g in gpu:
            f.write(f"{g['name']} - {g['vendor']}\n")
            f.write(f"Driver: {g['driver']}\n")
            f.write(f"URL: {g['url']}\n\n")

        # RAM
        f.write("\n=== RAM ===\n")
        f.write(f"Total: {ram} GB\n\n")

        # DISKS
        f.write("\n=== ALMACENAMIENTO ===\n")
        for d in disks:
            f.write(f"Modelo : {d['model']}\n")
            f.write(f"Tipo   : {d['type']}\n")
            f.write(f"Tamaño : {d['size']} GB\n")
            f.write(f"URL    : {d['url']}\n\n")

        # Motherboard
        f.write("=== MOTHERBOARD ===\n")
        f.write(f"Modelo : {mb['name']}\n")
        f.write(f"URL    : {mb['url']}\n\n")

        # Fans
        f.write("=== VENTILADORES ===\n")
        if fans:
            for fan in fans:
                f.write(f"- {fan}\n")
        else:
            f.write("No detectados.\n")
        f.write("\n")

        # Monitors
        f.write("=== MONITORES DETECTADOS ===\n")
        for m in monitors:
            f.write(f"Monitor : {m['name']}\n")
            f.write(f"Vendor  : {m['vendor']}\n")
            f.write(f"Serial  : {m['serial']}\n")
            f.write(f"Resolución: {m['width']}x{m['height']}\n")
            f.write(f"URL: {m['url']}\n\n")

    return file_path


# ===========================================
#       THREAD
# ===========================================

class ScanThread(QThread):
    finished = pyqtSignal(str, str)
    error = pyqtSignal(str)

    def run(self):
        try:
            path = generate_report()
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            self.finished.emit(path, content)
        except Exception as e:
            self.error.emit(str(e))


# ===========================================
#         GUI
# ===========================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Ventana sin marco
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self._drag_pos = None
        self.resize(1000, 650)

        self.setStyleSheet("""
        QMainWindow { background-color: #1E293B; }
        QWidget { color: #E5E7EB; font-family: 'Segoe UI'; font-size: 10pt; }
        QLabel#TitleLabel { font-size: 16px; font-weight: 600; color: #E5E7EB; }
        QLabel#SubtitleLabel { font-size: 10px; color: #9CA3AF; }

        QFrame#TitleBar { background-color: #0F172A; border-bottom: 1px solid #111827; }

        QPushButton#TitleButton {
            background-color: transparent; border: none; color: #9CA3AF;
            padding: 4px 10px; font-size: 11px;
        }
        QPushButton#TitleButton:hover { background-color: #1F2937; color: #E5E7EB; }

        QPushButton#CloseButton {
            background-color: transparent; border: none; color: #9CA3AF;
            padding: 4px 12px; font-size: 11px;
        }
        QPushButton#CloseButton:hover { background-color: #DC2626; color: #F9FAFB; }

        QPushButton {
            background-color: #2563EB; color: white; border-radius: 6px;
            padding: 8px 20px; font-weight: 500;
        }
        QPushButton:hover { background-color: #1D4ED8; }
        QPushButton:pressed { background-color: #1E40AF; }

        QFrame#Card {
            background-color: #0F172A; border-radius: 10px;
            border: 1px solid #334155;
        }

        QTextEdit {
            background-color: #020617; border-radius: 6px;
            border: 1px solid #334155;
            padding: 8px;
            font-family: Consolas, 'Cascadia Code', monospace;
            font-size: 9pt; color: #E5E7EB;
        }
        """)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0,0,0,0)
        root.setSpacing(0)

        # ===================== BARRA SUPERIOR =====================
        title_bar = QFrame()
        title_bar.setObjectName("TitleBar")
        tb = QHBoxLayout(title_bar)
        tb.setContentsMargins(10,4,10,4)
        tb.setSpacing(6)

        title = QLabel("PCInfoScanner - Hardware Report")
        title.setObjectName("TitleLabel")
        tb.addWidget(title)
        tb.addStretch()

        btn_min = QPushButton("–")
        btn_min.setFixedWidth(32)
        btn_min.setObjectName("TitleButton")
        btn_min.clicked.connect(self.showMinimized)

        btn_close = QPushButton("×")
        btn_close.setFixedWidth(32)
        btn_close.setObjectName("CloseButton")
        btn_close.clicked.connect(self.close)

        tb.addWidget(btn_min)
        tb.addWidget(btn_close)

        root.addWidget(title_bar)
        self._title_bar = title_bar

        # ===================== CONTENIDO =====================
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(16,16,16,16)

        header = QVBoxLayout()
        h_title = QLabel("PCInfoScanner")
        h_title.setObjectName("TitleLabel")
        h_sub = QLabel("Informe detallado de hardware para Windows 10 / 11")
        h_sub.setObjectName("SubtitleLabel")
        header.addWidget(h_title)
        header.addWidget(h_sub)
        layout.addLayout(header)

        # CARD
        card = QFrame()
        card.setObjectName("Card")
        cl = QHBoxLayout(card)
        cl.setContentsMargins(14,10,14,10)

        left = QVBoxLayout()
        self.status_label = QLabel("Listo para generar el informe.")
        self.status_label.setObjectName("SubtitleLabel")
        hint = QLabel("El archivo .txt se guardará en tu Escritorio.")
        hint.setObjectName("SubtitleLabel")
        left.addWidget(self.status_label)
        left.addWidget(hint)
        left.addStretch()

        right = QVBoxLayout()
        self.scan_button = QPushButton("Generar informe")
        self.scan_button.clicked.connect(self.start_scan)
        right.addWidget(self.scan_button, alignment=Qt.AlignRight)

        cl.addLayout(left, 3)
        cl.addLayout(right, 1)

        layout.addWidget(card)

        # TEXTAREA
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        layout.addWidget(self.text_edit, stretch=1)

        # ===================== BOTÓN DE CRÉDITOS =====================
        credits_row = QHBoxLayout()
        credits_row.addStretch()
        self.credits_button = QPushButton("Créditos / Source")
        self.credits_button.setFixedWidth(150)
        self.credits_button.clicked.connect(self.show_credits)
        credits_row.addWidget(self.credits_button)
        layout.addLayout(credits_row)

        root.addWidget(content)

        self.scan_thread = None


    # ===================== POPUP OSCURO =====================
    def show_popup(self, text):
        dlg = QDialog(self)
        dlg.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        dlg.setModal(True)

        dlg.setStyleSheet("""
        QDialog {
            background-color: #0F172A;
            border: 1px solid #334155;
            border-radius: 8px;
        }
        QLabel { color: #E5E7EB; font-size: 10pt; }
        QPushButton {
            background-color: #2563EB; color: white;
            border-radius: 4px; padding: 6px 18px;
        }
        QPushButton:hover { background-color: #1D4ED8; }
        """)

        v = QVBoxLayout(dlg)
        v.setContentsMargins(18,14,18,14)

        lbl = QLabel(text)
        lbl.setWordWrap(True)
        v.addWidget(lbl)

        row = QHBoxLayout()
        row.addStretch()
        ok = QPushButton("OK")
        ok.clicked.connect(dlg.accept)
        row.addWidget(ok)

        v.addLayout(row)

        dlg.adjustSize()
        dlg.move(self.geometry().center() - dlg.rect().center())
        dlg.exec_()


    # ===================== POPUP DE CRÉDITOS =====================
    def show_credits(self):
        dlg = QDialog(self)
        dlg.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        dlg.setModal(True)

        dlg.setStyleSheet("""
        QDialog {
            background-color: #0F172A;
            border: 1px solid #334155;
            border-radius: 8px;
        }
        QLabel {
            color: #E5E7EB;
            font-size: 10pt;
        }
        QPushButton {
            background-color: #2563EB;
            color: white;
            border-radius: 4px;
            padding: 6px 18px;
        }
        QPushButton:hover { background-color: #1D4ED8; }
        """)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18,14,18,14)

        label = QLabel(
            "<b>PCInfoScanner</b><br>"
            "Desarrollado por <b>@1vcbGH</b><br><br>"
            "Código fuente disponible en:<br>"
            "<a href='https://github.com/1vcbGH/PCInfoScanner' style='color:#60A5FA;'>"
            "github.com/1vcbGH/PCInfoScanner</a>"
        )
        label.setOpenExternalLinks(True)
        label.setTextFormat(Qt.RichText)

        layout.addWidget(label)

        row = QHBoxLayout()
        row.addStretch()
        ok = QPushButton("Cerrar")
        ok.clicked.connect(dlg.accept)
        row.addWidget(ok)

        layout.addLayout(row)

        dlg.adjustSize()
        dlg.move(self.geometry().center() - dlg.rect().center())
        dlg.exec_()


    # ===================== DRAG (MOVER VENTANA) =====================
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and event.pos().y() <= self._title_bar.height():
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self._drag_pos and event.buttons() & Qt.LeftButton:
            self.move(event.globalPos() - self._drag_pos)

    def mouseReleaseEvent(self, event):
        self._drag_pos = None


    # ===================== ESCANEO =====================
    def start_scan(self):
        if self.scan_thread and self.scan_thread.isRunning():
            return

        self.status_label.setText("Generando informe...")
        self.scan_button.setEnabled(False)
        self.text_edit.clear()

        self.scan_thread = ScanThread()
        self.scan_thread.finished.connect(self.on_scan_finished)
        self.scan_thread.error.connect(self.on_scan_error)
        self.scan_thread.start()

    def on_scan_finished(self, path, content):
        self.status_label.setText(f"Informe generado en: {path}")
        self.scan_button.setEnabled(True)
        self.text_edit.setPlainText(content)
        self.show_popup(f"Informe generado en:\n{path}")

    def on_scan_error(self, err):
        self.status_label.setText("Error al generar el informe.")
        self.scan_button.setEnabled(True)
        self.show_popup(f"Error durante el escaneo:\n{err}")


# ======================================================
#   MAIN
# ======================================================

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

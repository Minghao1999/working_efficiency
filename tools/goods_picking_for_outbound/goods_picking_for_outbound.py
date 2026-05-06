import sys
import subprocess
import time
import re
import threading
import xml.etree.ElementTree as ET

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QTextEdit,
    QLabel,
    QComboBox,
    QLineEdit,
    QMessageBox
)

ADB = r"C:\Users\sun98\OneDrive\Desktop\working efficiency\tools\platform-tools\adb.exe"
class PDAWindow(QWidget):

    def __init__(self):
        super().__init__()

        self.running = False

        self.setWindowTitle("PDA Goods Picking")
        self.resize(700, 500)

        layout = QVBoxLayout()

        self.device_label = QLabel("PDA Device")
        layout.addWidget(self.device_label)

        self.device_combo = QComboBox()
        layout.addWidget(self.device_combo)

        self.refresh_btn = QPushButton("Refresh Device")
        self.refresh_btn.clicked.connect(self.load_devices)
        layout.addWidget(self.refresh_btn)

        self.container_input = QLineEdit()
        self.container_input.setText("P60009")
        layout.addWidget(self.container_input)

        self.start_btn = QPushButton("Start")
        self.start_btn.clicked.connect(self.start_program)
        layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("Stop")
        self.stop_btn.clicked.connect(self.stop_program)
        layout.addWidget(self.stop_btn)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        self.setLayout(layout)

        self.load_devices()

    def write_log(self, text):
        self.log.append(text)
        print(text)

    def adb(self, cmd):
        device = self.device_combo.currentText()

        if not device:
            return ""

        full_cmd = [ADB, "-s", device] + cmd

        result = subprocess.run(
            full_cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="ignore"
        )

        return result.stdout.strip()

    def load_devices(self):

        self.device_combo.clear()

        result = subprocess.run(
            [ADB, "devices"],
            stdout=subprocess.PIPE,
            text=True
        )

        lines = result.stdout.splitlines()

        devices = []

        for line in lines[1:]:
            if "\tdevice" in line:
                serial = line.split("\t")[0]
                devices.append(serial)

        self.device_combo.addItems(devices)

        if devices:
            self.write_log(f"✅ 找到设备: {devices}")
        else:
            self.write_log("❌ 未找到 PDA")

    def tap(self, x, y):
        self.adb(["shell", "input", "tap", str(x), str(y)])

    def input_text(self, text):
        text = str(text).replace(" ", "%s")
        self.adb(["shell", "input", "text", text])

    def press_enter(self):
        self.adb(["shell", "input", "keyevent", "66"])

    def dump_ui(self):
        self.adb(["shell", "uiautomator", "dump", "/sdcard/window.xml"])
        xml = self.adb(["shell", "cat", "/sdcard/window.xml"])
        return xml

    def parse_bounds(self, bounds):
        nums = list(map(int, re.findall(r"\d+", bounds)))
        x1, y1, x2, y2 = nums
        return (x1 + x2) // 2, (y1 + y2) // 2

    def get_nodes(self):

        xml = self.dump_ui()

        try:
            root = ET.fromstring(xml)
        except:
            return []

        nodes = []

        for node in root.iter("node"):
            nodes.append({
                "text": node.attrib.get("text", ""),
                "class": node.attrib.get("class", ""),
                "bounds": node.attrib.get("bounds", "")
            })

        return nodes

    def get_texts(self):
        return [n["text"] for n in self.get_nodes() if n["text"]]

    def current_page(self):

        texts = self.get_texts()
        joined = " ".join(texts)

        if "Scan Container Code" in joined:
            return "container"

        if "Scan Location Code" in joined:
            return "location"

        if "Scan Barcode" in joined:
            return "barcode"

        if "Enter QTY" in joined:
            return "qty"

        return "unknown"

    def find_input(self):

        for n in self.get_nodes():

            if "EditText" in n["class"]:
                return self.parse_bounds(n["bounds"])

            if n["text"] in ["Please Scan", "Please Enter"]:
                return self.parse_bounds(n["bounds"])

        return None

    def click_and_input(self, value):

        pos = self.find_input()

        if not pos:
            self.write_log("❌ 找不到输入框")
            return

        self.tap(*pos)

        time.sleep(0.2)

        self.input_text(value)

        time.sleep(0.2)

        self.press_enter()

        time.sleep(1)

    def extract_location(self):

        joined = " ".join(self.get_texts())

        m = re.search(r"A\d+-R\d+-L\d+-B\d+", joined)

        if m:
            return m.group(0)

        return None

    def extract_goods_code(self):

        joined = " ".join(self.get_texts())

        m = re.search(r"FG\d+", joined)

        if m:
            return m.group(0)

        return None

    def extract_qty(self):

        joined = " ".join(self.get_texts())

        m = re.search(r"\d+\s*/\s*(\d+)", joined)

        if m:
            return m.group(1)

        return None

    def click_confirm(self):

        for n in self.get_nodes():

            if n["text"].lower() == "confirm":

                x, y = self.parse_bounds(n["bounds"])

                self.tap(x, y)

                self.write_log("✅ 点击 Confirm")

                time.sleep(1)

    def automation_loop(self):

        self.write_log("🚀 自动化开始")

        while self.running:

            try:

                page = self.current_page()

                self.write_log(f"当前页面: {page}")

                if page == "container":

                    value = self.container_input.text()

                    self.write_log(f"输入 Container: {value}")

                    self.click_and_input(value)

                elif page == "location":

                    loc = self.extract_location()

                    if loc:
                        self.write_log(f"输入 Location: {loc}")
                        self.click_and_input(loc)

                elif page == "barcode":

                    code = self.extract_goods_code()

                    if code:
                        self.write_log(f"输入 Barcode: {code}")
                        self.click_and_input(code)

                elif page == "qty":

                    qty = self.extract_qty()

                    if qty:
                        self.write_log(f"输入 QTY: {qty}")
                        self.click_and_input(qty)
                        self.click_confirm()

                else:

                    self.write_log("⚠️ 未识别页面")

                time.sleep(1)

            except Exception as e:
                self.write_log(str(e))
                time.sleep(2)

        self.write_log("🛑 已停止")

    def start_program(self):

        if not self.device_combo.currentText():
            QMessageBox.warning(self, "错误", "未连接 PDA")
            return

        self.running = True

        thread = threading.Thread(target=self.automation_loop)
        thread.daemon = True
        thread.start()

    def stop_program(self):
        self.running = False


app = QApplication(sys.argv)

window = PDAWindow()
window.show()

sys.exit(app.exec())
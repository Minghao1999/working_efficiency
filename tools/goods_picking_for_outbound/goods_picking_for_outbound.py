import sys
import subprocess
import time
import re
import threading
import xml.etree.ElementTree as ET

from PySide6.QtCore import Signal
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
    log_signal = Signal(str)
    clear_task_order_signal = Signal()

    def __init__(self):
        super().__init__()

        self.running = False
        self.selected_device = ""
        self.task_order_no = ""
        self.container_no = ""

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

        self.task_order_label = QLabel("Picking Task Order Number")
        layout.addWidget(self.task_order_label)

        self.task_order_input = QLineEdit()
        self.task_order_input.textChanged.connect(self.set_task_order_no)
        layout.addWidget(self.task_order_input)

        self.container_label = QLabel("Container Number")
        layout.addWidget(self.container_label)

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

        self.log_signal.connect(self.append_log)
        self.clear_task_order_signal.connect(self.clear_task_order_input)
        self.load_devices()

    def append_log(self, text):
        self.log.append(text)
        print(text)

    def write_log(self, text):
        self.log_signal.emit(str(text))

    def set_task_order_no(self, text):
        self.task_order_no = text.strip()

    def clear_task_order_input(self):
        self.task_order_input.clear()

    def adb(self, cmd):
        device = self.selected_device

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
            stderr=subprocess.PIPE,
            encoding="utf-8",
            errors="ignore",
            text=True
        )

        lines = result.stdout.splitlines()

        devices = []
        unauthorized = []
        offline = []

        for line in lines[1:]:
            if "\tdevice" in line:
                serial = line.split("\t")[0]
                devices.append(serial)
            elif "\tunauthorized" in line:
                serial = line.split("\t")[0]
                unauthorized.append(serial)
            elif "\toffline" in line:
                serial = line.split("\t")[0]
                offline.append(serial)

        self.device_combo.addItems(devices)

        if devices:
            self.write_log(f"✅ 找到设备: {devices}")
        elif unauthorized:
            self.write_log(f"⚠️ PDA 未授权: {unauthorized}，请在 PDA 上点允许 USB 调试")
        elif offline:
            self.write_log(f"⚠️ PDA 离线: {offline}，请重新插拔 USB 或重启 PDA")
        else:
            self.write_log("❌ 未找到 PDA")

            if result.stderr.strip():
                self.write_log(f"ADB 错误: {result.stderr.strip()}")

    def tap(self, x, y):
        self.adb(["shell", "input", "tap", str(x), str(y)])

    def input_text(self, text):
        text = str(text).replace(" ", "%s")
        self.adb(["shell", "input", "text", text])

    def input_text_and_enter(self, text):
        text = str(text).replace(" ", "%s")
        self.adb(["shell", f"input text {text}; input keyevent 66"])

    def press_enter(self):
        self.adb(["shell", "input", "keyevent", "66"])

    def clear_focused_input(self, current_text=""):
        count = len(current_text)

        if current_text in ["", "Please Scan", "Please Enter", "Please scan/enter"]:
            return

        self.adb(["shell", "input", "keyevent", "123"])

        if count:
            self.adb([
                "shell",
                f"for i in $(seq 1 {count}); do input keyevent 67; done"
            ])

    def dump_ui(self):
        self.adb(["shell", "uiautomator", "dump", "/sdcard/window.xml"])
        xml = self.adb(["shell", "cat", "/sdcard/window.xml"])
        return xml

    def clean_xml(self, xml):
        start = xml.find("<?xml")

        if start == -1:
            start = xml.find("<hierarchy")

        if start > 0:
            xml = xml[start:]

        end = xml.rfind("</hierarchy>")

        if end != -1:
            xml = xml[:end + len("</hierarchy>")]

        return xml.strip()

    def parse_bounds(self, bounds):
        nums = list(map(int, re.findall(r"\d+", bounds)))
        x1, y1, x2, y2 = nums
        return (x1 + x2) // 2, (y1 + y2) // 2

    def get_nodes(self):

        xml = self.clean_xml(self.dump_ui())

        try:
            root = ET.fromstring(xml)
        except Exception as e:
            self.write_log(f"❌ UI XML 解析失败: {e}")
            return []

        nodes = []

        for node in root.iter("node"):
            nodes.append({
                "text": node.attrib.get("text", ""),
                "resource_id": node.attrib.get("resource-id", ""),
                "content_desc": node.attrib.get("content-desc", ""),
                "class": node.attrib.get("class", ""),
                "bounds": node.attrib.get("bounds", "")
            })

        return nodes

    def get_texts(self, nodes=None):
        texts = []

        if nodes is None:
            nodes = self.get_nodes()

        for n in nodes:
            for key in ["text", "content_desc", "resource_id"]:
                value = n.get(key, "")

                if value:
                    texts.append(value)

        return texts

    def has_any(self, text, keywords):
        text = text.lower()
        return any(keyword.lower() in text for keyword in keywords)

    def current_page(self, nodes=None):

        texts = self.get_texts(nodes)
        joined = " ".join(texts)

        if self.has_any(joined, [
            "Scan the picking task order number",
            "Scan the task order number",
            "Scan The Picking Task Order No",
            "Scan the task order No",
            "Picking Task Order No",
            "task order number",
            "task order No"
        ]):
            return "task_order"

        if self.has_any(joined, [
            "Scan Container Code",
            "Scan Container No",
            "Start with picking container"
        ]):
            return "container"

        if self.has_any(joined, [
            "Scan Barcode",
            "Scan Goods Barcode",
            "Scan Goods Code",
            "Goods Barcode",
            "barcode"
        ]):
            return "barcode"

        if self.has_any(joined, [
            "Scan Location Code",
            "Scan Location No"
        ]):
            return "location"

        if self.has_any(joined, [
            "Enter QTY",
            "Enter Qty",
            "Input QTY",
            "Input Qty",
            "Quantity",
            "qty"
        ]):
            return "qty"

        preview = " | ".join(texts[:24])

        if preview:
            self.write_log(f"页面文字: {preview}")
        else:
            self.write_log("页面文字: 空，可能当前 App 页面未暴露控件文字")

        return "unknown"

    def find_input(self, nodes=None):

        if nodes is None:
            nodes = self.get_nodes()

        for n in nodes:

            if "EditText" in n["class"]:
                return self.parse_bounds(n["bounds"]), n["text"]

            if n["text"] in ["Please Scan", "Please Enter"]:
                return self.parse_bounds(n["bounds"]), n["text"]

        return None

    def click_and_input(self, value, nodes=None):

        found = self.find_input(nodes)

        if not found:
            self.write_log("❌ 找不到输入框")
            return

        pos, current_text = found

        self.tap(*pos)

        time.sleep(0.1)

        self.clear_focused_input(current_text)

        time.sleep(0.1)

        self.input_text_and_enter(value)

        time.sleep(0.2)

    def extract_location(self, nodes=None):

        texts = self.get_texts(nodes)
        joined = " ".join(texts)

        m = re.search(r"A\d+-R\d+-L\d+-B\d+", joined)

        if m:
            return m.group(0)

        self.write_log(f"❌ 找不到 Location No，页面文字: {' | '.join(texts[:12])}")

        return None

    def extract_goods_code(self, nodes=None):

        if nodes is None:
            nodes = self.get_nodes()

        for n in nodes:
            if n.get("resource_id", "").endswith("tv_goods_fg"):
                code = re.sub(r"\s+", "", n["text"])

                if code:
                    return code

        joined = " ".join(n["text"] for n in nodes if n["text"])

        m = re.search(r"(FG\d+)\s+(\d{3,})", joined)

        if m:
            return f"{m.group(1)}{m.group(2)}"

        m = re.search(r"FG\d+", joined)

        if m:
            return m.group(0)

        self.write_log(f"❌ 找不到 Barcode，页面文字: {' | '.join(self.get_texts(nodes)[:12])}")

        return None

    def extract_qty(self, nodes=None):

        joined = " ".join(self.get_texts(nodes))

        m = re.search(r"Pick\s*QTY\.?\s*:?\s*(\d+)", joined, re.IGNORECASE)

        if m:
            return m.group(1)

        m = re.search(r"\d+\s*/\s*(\d+)", joined)

        if m:
            return m.group(1)

        return None

    def click_confirm(self, nodes=None):

        if nodes is None:
            nodes = self.get_nodes()

        for n in nodes:

            if n["text"].lower() == "confirm":

                x, y = self.parse_bounds(n["bounds"])

                self.tap(x, y)

                self.write_log("✅ 点击 Confirm")

                time.sleep(0.2)

    def automation_loop(self):

        self.write_log("🚀 自动化开始")

        while self.running:

            try:

                nodes = self.get_nodes()
                page = self.current_page(nodes)

                self.write_log(f"当前页面: {page}")

                if page == "task_order":

                    value = self.task_order_no

                    if value:
                        self.write_log(f"输入 Picking Task Order Number: {value}")
                        self.click_and_input(value, nodes)
                        self.click_confirm(nodes)
                        self.clear_task_order_signal.emit()
                    else:
                        self.write_log("⚠️ 请在 UI 输入 Picking Task Order Number")

                elif page == "container":

                    value = self.container_no

                    self.write_log(f"输入 Container: {value}")

                    self.click_and_input(value, nodes)

                elif page == "location":

                    loc = self.extract_location(nodes)

                    if loc:
                        self.write_log(f"输入 Location: {loc}")
                        self.click_and_input(loc, nodes)

                elif page == "barcode":

                    code = self.extract_goods_code(nodes)

                    if code:
                        self.write_log(f"输入 Barcode: {code}")
                        self.click_and_input(code, nodes)

                elif page == "qty":

                    qty = self.extract_qty(nodes)

                    if qty:
                        self.write_log(f"输入 QTY: {qty}")
                        self.click_and_input(qty, nodes)
                        self.click_confirm(nodes)

                else:

                    self.write_log("⚠️ 未识别页面")

                time.sleep(0.2)

            except Exception as e:
                self.write_log(str(e))
                time.sleep(2)

        self.write_log("🛑 已停止")

    def start_program(self):

        if not self.device_combo.currentText():
            QMessageBox.warning(self, "错误", "未连接 PDA")
            return

        self.selected_device = self.device_combo.currentText()
        self.task_order_no = self.task_order_input.text().strip()
        self.container_no = self.container_input.text()
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

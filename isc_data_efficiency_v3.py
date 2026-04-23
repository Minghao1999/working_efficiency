import sys
import os
import pandas as pd
import warnings
from PyQt6.QtWidgets import QScrollArea

from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QToolTip,
    QVBoxLayout,
    QWidget,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QCursor, QFont, QIcon

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.dates as mdates
from matplotlib.patches import Patch

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================
# Theme
# =========================
THEME = {
    "bg": "#F6F8FC",
    "surface": "#FFFFFF",
    "surface_alt": "#F8FAFC",
    "border": "#E5E7EB",
    "border_soft": "#EEF2F7",
    "primary": "#2563EB",
    "primary_dark": "#1D4ED8",
    "text": "#0F172A",
    "muted": "#64748B",
    "success": "#16A34A",
    "danger": "#DC2626",
    "warning": "#D97706",
    "work": "#7EC398",
    "overnight": "#8B7ABD",
    "idle": "#C6D7EB",
    "break": "#D9E6AF",
    "kpi_bg": "#EEF4FF",
}


def resource_path(relative_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def adjust_work_date(dt):
        if pd.isna(dt):
            return None
        # ✅ 凌晨时间归前一天
        if dt.hour < 6:
            return (dt - pd.Timedelta(days=1)).date()
        return dt.date()
import re

def extract_warehouse(name):
    if pd.isna(name):
        return "Unknown"
    match = re.search(r"\d+号仓", str(name))
    return match.group() if match else str(name)
# =========================
# Data logic
# =========================
def load_data(file1, file2):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    df2["employee_no"] = df2.iloc[:, 1].astype(str).str.strip()
    df2["real_name"] = df2.iloc[:, 3].astype(str).str.strip()
    df2["上班时间"] = pd.to_datetime(df2.iloc[:, 13], errors="coerce")
    df2["下班时间"] = pd.to_datetime(df2.iloc[:, 18], errors="coerce")
    df2["attendance_group"] = (
        df2.iloc[:, 26].astype(str).str.strip().replace("nan", "未分组")
    )

    mapping = dict(zip(df2["employee_no"], df2["real_name"]))
    group_mapping = dict(zip(df2["employee_no"], df2["attendance_group"]))

    df = df1.rename(
        columns={
            "employee_no": "employee_no",
            "任务单开始时间": "start",
            "任务单结束时间": "end",
            "operator": "operator",
        }
    )
    df["warehouse_name"] = df1["warehouse_name"]
    df["employee_no"] = (
        df["employee_no"].astype(str).str.strip().replace(["nan", "None", ""], pd.NA)
    )
    df["start"] = pd.to_datetime(df["start"], errors="coerce")
    df["end"] = pd.to_datetime(df["end"], errors="coerce")
    

    # ✅ 修复 ISC 跨天错误（凌晨任务归到第二天）
    def fix_isc_datetime(row):
        start = row["start"]
        end = row["end"]

        if pd.isna(start) or pd.isna(end):
            return start, end

        # 👉 如果时间在凌晨（0~6点），说明其实是“第二天”
        if start.hour < 6:
            start = start + pd.Timedelta(days=1)
            end = end + pd.Timedelta(days=1)

        return start, end

    df[["start", "end"]] = df.apply(
        lambda r: pd.Series(fix_isc_datetime(r)), axis=1
    )

    def identify_name(row):
        emp_no = row["employee_no"]
        if pd.notna(emp_no) and emp_no in mapping:
            return f"{mapping[emp_no]} ({emp_no})"
        if pd.notna(emp_no):
            return str(emp_no)
        return str(row.get("operator", "Unknown"))

    df["display_name"] = df.apply(identify_name, axis=1)
    df["type"] = "work"
    df["work_date"] = df["start"].apply(adjust_work_date)
    return df, df2, group_mapping

def find_no_operation_people(df_isc, df_iams):
    # ✅ iAMS 有打卡的人
    iams_valid = df_iams[
        df_iams["上班时间"].notna() & df_iams["下班时间"].notna()
    ].copy()

    iams_valid["employee_no"] = iams_valid["employee_no"].astype(str).str.strip()

    # ❗关键1：只按 employee_no 判断，不看 operator
    isc_valid = df_isc[
        df_isc["employee_no"].notna()
    ].copy()

    isc_valid["employee_no"] = isc_valid["employee_no"].astype(str).str.strip()

    isc_emp_set = set(isc_valid["employee_no"])

    # ✅ 差集
    no_op = iams_valid[~iams_valid["employee_no"].isin(isc_emp_set)].copy()

    return no_op[["employee_no", "real_name", "上班时间", "下班时间"]]

def load_volume_data(file3):
    try:
        df = pd.read_excel(file3)
        df.columns = df.columns.str.strip()

        # ======================
        # 1️⃣ 打包完成时间
        # ======================
        df['打包完成时间'] = pd.to_datetime(df['打包完成时间'], errors='coerce')

        # ======================
        # 2️⃣ 过滤取消订单
        # ======================
        if '是否取消' in df.columns:
            df = df[~df['是否取消'].astype(str).str.contains('是')]

        # ======================
        # 3️⃣ G列件数（第7列）
        # ======================
        df['件数'] = pd.to_numeric(df.iloc[:, 6], errors='coerce').fillna(0)

        # ======================
        # 4️⃣ 按日期分组
        # ======================
        df['日期'] = df['打包完成时间'].dt.date

        # ======================
        # 5️⃣ 汇总
        # ======================
        volume_dict = df.groupby('日期')['件数'].sum().to_dict()

        return volume_dict

    except Exception as e:
        print("Volume读取失败:", e)
        return {}



def load_break_data(file4):
    try:
        df4 = pd.read_excel(file4)
        df_clean = pd.DataFrame(
            {
                "emp_no": df4.iloc[:, 3].astype(str).str.strip(),
                "time": pd.to_datetime(df4.iloc[:, 12], errors="coerce"),
                "action": df4.iloc[:, 18].astype(str),
            }
        ).sort_values(["emp_no", "time"])

        breaks = []
        for emp, group in df_clean.groupby("emp_no"):
            out_time = None
            for _, row in group.iterrows():
                if "出仓" in row["action"]:
                    out_time = row["time"]
                elif "进仓" in row["action"] and out_time is not None:
                    duration = (row["time"] - out_time).total_seconds() / 60
                    if duration >= 20 and 11 <= out_time.hour <= 15:
                        breaks.append(
                            {
                                "emp_no": emp,
                                "start": out_time,
                                "end": row["time"],
                                "type": "break",
                            }
                        )
                    out_time = None
        return pd.DataFrame(breaks)
    except Exception:
        return pd.DataFrame()



def build_timeline(df, df2, group_mapping, df_breaks=None):
    records = []
    processed_names = set()

    for _, row2 in df2.iterrows():
        emp_no = str(row2["employee_no"])
        name = str(row2["real_name"])
        display_name = f"{name} ({emp_no})"
        group_name = str(row2["attendance_group"])
        work_start = row2["上班时间"]
        work_end = row2["下班时间"]

        if pd.isna(work_start) or pd.isna(work_end):
            continue

        day = adjust_work_date(work_start)
        processed_names.add((display_name, day))
        shift_name = str(row2.get("班次名称", ""))

        if "早" in shift_name:
            shift_tag = "morning"
        elif "晚" in shift_name:
            shift_tag = "evening"
        else:
            shift_tag = "morning" if work_start.hour < 12 else "evening"

        person_tasks = df[
            (df["employee_no"] == emp_no)
            & (df["work_date"] == day) 
            & (df["start"] < work_end) 
            & (df["end"] > work_start)
        ].copy()
        person_tasks["type"] = "work"

        if df_breaks is not None and not df_breaks.empty:
            p_breaks = df_breaks[
                (df_breaks["emp_no"] == emp_no)
                & (df_breaks["start"] < work_end)
                & (df_breaks["end"] > work_start)
            ].copy()
            combined = pd.concat(
                [person_tasks, p_breaks.rename(columns={"emp_no": "employee_no"})],
                ignore_index=True,
            ).sort_values("start")
        else:
            combined = person_tasks.sort_values("start")

        current_cursor = work_start
        for _, task in combined.iterrows():
            t_s = max(task["start"], work_start)
            t_e = min(task["end"], work_end)
            if t_s >= t_e:
                continue

            if t_s > current_cursor:
                records.append(
                    [display_name, day, current_cursor, t_s, "idle", shift_tag, False, group_name]
                )

            task_type = task["type"]
            if task_type == "work" and task["start"].date() != task["end"].date():
                task_type = "overnight"

            records.append([display_name, day, t_s, t_e, task_type, shift_tag, False, group_name])
            current_cursor = max(current_cursor, t_e)

        if current_cursor < work_end:
            records.append(
                [display_name, day, current_cursor, work_end, "idle", shift_tag, False, group_name]
            )

    d1_dates = set(df["work_date"].dropna().unique())
    for current_day in sorted(list(d1_dates)):
        day_tasks_all = df[df["work_date"] == current_day]
        for display_name, group in day_tasks_all.groupby("display_name"):
            if (display_name, current_day) in processed_names:
                continue

            emp_id = display_name.split("(")[-1].strip(")")
            sorted_group = group.sort_values("start")
            first_task_start = sorted_group["start"].min()
            auto_shift = "morning" if 6 <= first_task_start.hour < 12 else "evening"

            for _, task in sorted_group.iterrows():
                records.append(
                    [
                        display_name,
                        current_day,
                        task["start"],
                        task["end"],
                        "work",
                        auto_shift,
                        True,
                        group_mapping.get(emp_id, "Unknown"),
                    ]
                )

    timeline = pd.DataFrame(
        records,
        columns=["name", "date", "start", "end", "type", "shift", "is_absent", "group"],
    )
    if not timeline.empty:
        timeline["duration"] = (timeline["end"] - timeline["start"]).dt.total_seconds() / 3600
    else:
        timeline["duration"] = pd.Series(dtype=float)
    return timeline


# =========================
# UI helpers
# =========================
def apply_shadow(widget: QWidget, blur=24, y_offset=6, alpha=25):
    shadow = QGraphicsDropShadowEffect()
    shadow.setBlurRadius(blur)
    shadow.setXOffset(0)
    shadow.setYOffset(y_offset)
    shadow.setColor(QColor(0, 0, 0, alpha))
    widget.setGraphicsEffect(shadow)


class FilePickerCard(QFrame):
    def __init__(self, title: str, subtitle: str, button_text: str):
        super().__init__()
        self.setObjectName("uploadCard")
        self.setMinimumHeight(104)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(8)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("cardTitle")

        self.subtitle_label = QLabel(subtitle)
        self.subtitle_label.setObjectName("cardSubtitle")
        self.subtitle_label.setWordWrap(True)

        bottom = QHBoxLayout()
        bottom.setSpacing(10)

        self.button = QPushButton(button_text)
        self.button.setCursor(Qt.CursorShape.PointingHandCursor)

        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("fileNameLabel")
        self.file_label.setWordWrap(True)

        bottom.addWidget(self.button, 0)
        bottom.addWidget(self.file_label, 1)

        layout.addWidget(self.title_label)
        layout.addWidget(self.subtitle_label)
        layout.addLayout(bottom)

        apply_shadow(self, blur=18, y_offset=4, alpha=18)

    def set_file(self, path: str):
        self.file_label.setText(os.path.basename(path) if path else "No file selected")


# =========================
# Main Window
# =========================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WorkSight Pro")
        self.resize(1580, 980)

        try:
            icon_path = resource_path("favicon.ico")
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except Exception:
            pass

        self.timeline = pd.DataFrame()
        self.volume_data = {}
        self.summary = {}
        self.history_ratios = {}
        self.history_wait = {}
        self.ax_right_m = None
        self.ax_right_e = None
        self.hover_cid_m = None
        self.hover_cid_e = None
        self._kpi_visible = True

        self.init_ui()
        self.apply_styles()

    def toggle_data_sources(self):
        visible = self.upload_grid_widget.isVisible()
        self.upload_grid_widget.setVisible(not visible)

        self.ds_toggle_btn.setText("▾" if not visible else "▸")
    
    def sync_date_from_header(self, date):
        if date != self.date_combo.currentText():
            self.date_combo.setCurrentText(date)

    def sync_date_from_main(self, date):
        if date != self.header_date_combo.currentText():
            self.header_date_combo.setCurrentText(date)

    def init_ui(self):
        main = QWidget()
        self.setCentralWidget(main)

        # =============================
        # 外层布局（不滚动）
        # =============================
        outer = QVBoxLayout(main)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # =============================
        # 1️⃣ 固定 Header（不会滚动）
        # =============================
        header = QFrame()
        header.setObjectName("heroCard")

        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(24, 20, 24, 20)

        left = QVBoxLayout()
        left.setSpacing(4)

        title = QLabel("WorkSight Pro")
        title.setObjectName("heroTitle")

        subtitle = QLabel(
            "Advanced productivity analytics dashboard for ISC, iAMS, volume, and punch data"
        )
        subtitle.setObjectName("heroSubtitle")
        subtitle.setWordWrap(True)

        left.addWidget(title)
        left.addWidget(subtitle)

        right = QHBoxLayout()
        self.header_date_combo = QComboBox()
        self.header_date_combo.setObjectName("dateBadge")
        self.header_date_combo.setMinimumWidth(140)
        self.header_date_combo.setCursor(Qt.CursorShape.PointingHandCursor)

        right.addWidget(self.header_date_combo)

        header_layout.addLayout(left, 1)
        header_layout.addLayout(right)

        outer.addWidget(header)

        # =============================
        # 2️⃣ Scroll Area（下面全部可滚动）
        # =============================
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        content = QWidget()
        scroll.setWidget(content)

        outer.addWidget(scroll)

        # =============================
        # 3️⃣ 原内容 root（滚动区域）
        # =============================
        root = QVBoxLayout(content)
        root.setContentsMargins(24, 20, 24, 20)
        root.setSpacing(18)

        # =============================
        # Data Source（你的原代码）
        # =============================
        upload_wrap = QFrame()
        upload_wrap.setObjectName("sectionCard")
        apply_shadow(upload_wrap, blur=20, y_offset=6, alpha=18)

        upload_layout = QVBoxLayout(upload_wrap)
        upload_layout.setContentsMargins(20, 18, 20, 18)
        upload_layout.setSpacing(14)

        upload_top = QHBoxLayout()

        title_box = QVBoxLayout()
        self.ds_toggle_btn = QPushButton("▾")
        self.ds_toggle_btn.setObjectName("ghostButton")
        self.ds_toggle_btn.setFixedWidth(30)
        self.ds_toggle_btn.clicked.connect(self.toggle_data_sources)
        upload_title = QLabel("Data Sources")
        upload_title.setObjectName("sectionTitle")

        upload_desc = QLabel(
            "Load your source files, then generate a daily analytics view."
        )
        upload_desc.setObjectName("sectionSubtitle")
        upload_desc.setWordWrap(True)

        title_box.addWidget(upload_title)
        title_box.addWidget(upload_desc)

        self.date_combo = QComboBox()
        self.date_combo.setMinimumWidth(180)

        self.gen_btn = QPushButton("Generate Dashboard")
        self.gen_btn.setObjectName("primaryButton")
        self.gen_btn.setMinimumHeight(42)

        upload_top.addLayout(title_box, 1)
        upload_top.addWidget(self.ds_toggle_btn)
        upload_top.addWidget(QLabel("Date"))
        upload_top.addWidget(self.date_combo)
        upload_top.addWidget(self.gen_btn)

        upload_layout.addLayout(upload_top)

        self.upload_grid_widget = QWidget()
        grid = QGridLayout(self.upload_grid_widget)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(14)

        self.card1 = FilePickerCard("ISC Task Data(任务数据)", "Upload ISC data", "Upload ISC")
        self.card2 = FilePickerCard("iAMS Attendance(班次明细)", "Upload iAMS data", "Upload iAMS")
        self.card3 = FilePickerCard("Volume Data(出库件量)", "Optional", "Upload Volume")
        self.card4 = FilePickerCard("Punch Data(打卡流水)", "Optional", "Upload Punch")

        grid.addWidget(self.card1, 0, 0)
        grid.addWidget(self.card2, 0, 1)
        grid.addWidget(self.card3, 1, 0)
        grid.addWidget(self.card4, 1, 1)

        upload_layout.addWidget(self.upload_grid_widget)
        root.addWidget(upload_wrap)

        # =============================
        # KPI 区域（你原来的）
        # =============================
        self.kpi_container = QFrame()
        self.kpi_container.setObjectName("sectionCard")
        apply_shadow(self.kpi_container, blur=20, y_offset=6, alpha=18)

        outer_kpi_layout = QVBoxLayout(self.kpi_container)
        outer_kpi_layout.setContentsMargins(18, 16, 18, 18)
        # ✅ KPI header（加按钮的地方）
        kpi_header = QHBoxLayout()

        kpi_title = QLabel("KPI Dashboard")
        kpi_title.setObjectName("sectionTitle")

        self.kpi_toggle_btn = QPushButton("▾")
        self.kpi_toggle_btn.setObjectName("ghostButton")
        self.kpi_toggle_btn.setFixedWidth(30)
        self.kpi_toggle_btn.clicked.connect(self.toggle_kpi)

        kpi_header.addWidget(kpi_title)
        kpi_header.addStretch()
        kpi_header.addWidget(self.kpi_toggle_btn)

        outer_kpi_layout.addLayout(kpi_header)

        self.k_date = self.create_kpi("WAREHOUSE", "---")
        self.k_workers = self.create_kpi("TOTAL WORKERS", "0")
        self.k_total_manhours = self.create_kpi("TOTAL WORK-HOURS", "0.0h")
        self.k_efficiency_chart = self.create_efficiency_card()
        self.k_vol = self.create_kpi("TOTAL VOLUME", "0")

        self.kpi_row_widget = QWidget()
        kpi_row = QHBoxLayout(self.kpi_row_widget)
        for w in [
            self.k_date,
            self.k_workers,
            self.k_total_manhours,
            self.k_efficiency_chart,
            self.k_vol,
        ]:
            kpi_row.addWidget(w)

        outer_kpi_layout.addWidget(self.kpi_row_widget)
        root.addWidget(self.kpi_container)

        # =============================
        # Tabs
        # =============================
        self.tabs = QTabWidget()
        self.tab_m = QWidget()
        self.tab_e = QWidget()
        self.tab_fa = QWidget()
        self.tab_no_op = QWidget() 

        self.tabs.addTab(self.tab_m, "Morning Shift")
        self.tabs.addTab(self.tab_e, "Evening Shift")
        self.tabs.addTab(self.tab_fa, "First Action Analysis")
        self.tabs.addTab(self.tab_no_op, "No Operation Analysis") 

        root.addWidget(self.tabs, 1)

        self.fig_m, self.ax_m, self.can_m, self.ph_m, self.group_m = self.create_gantt_tab(self.tab_m, "morning")
        self.fig_e, self.ax_e, self.can_e, self.ph_e, self.group_e = self.create_gantt_tab(self.tab_e, "evening")

        # =============================
        # First Action Analysis UI
        # =============================
        fa_layout = QVBoxLayout(self.tab_fa)
        fa_layout.setContentsMargins(12, 12, 12, 12)

        fa_card = QFrame()
        fa_card.setObjectName("sectionCard")
        apply_shadow(fa_card, blur=20, y_offset=6, alpha=18)

        fa_inner = QVBoxLayout(fa_card)
        fa_inner.setContentsMargins(16, 16, 16, 16)
        fa_inner.setSpacing(12)

        fa_title = QLabel("First Action Analysis")
        fa_title.setObjectName("sectionTitle")

        fa_subtitle = QLabel("Wait time from punch-in to first task, with day-over-day trend.")
        fa_subtitle.setObjectName("sectionSubtitle")

        # ✅ 关键：创建 table
        self.fa_table = QTableWidget(0, 7)
        self.fa_table.setHorizontalHeaderLabels(
            ["Name", "Group", "Shift", "Punch-in", "First Task", "Wait", "Trend"]
        )

        self.fa_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.fa_table.verticalHeader().setVisible(False)
        self.fa_table.setAlternatingRowColors(True)
        self.fa_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.fa_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.fa_table.setShowGrid(False)

        fa_inner.addWidget(fa_title)
        fa_inner.addWidget(fa_subtitle)
        fa_inner.addWidget(self.fa_table)

        fa_layout.addWidget(fa_card)
        # =============================
        # No Operation Analysis UI
        # =============================
        no_op_layout = QVBoxLayout(self.tab_no_op)
        no_op_layout.setContentsMargins(12, 12, 12, 12)

        no_op_card = QFrame()
        no_op_card.setObjectName("sectionCard")
        apply_shadow(no_op_card, blur=20, y_offset=6, alpha=18)

        no_op_inner = QVBoxLayout(no_op_card)
        no_op_inner.setContentsMargins(16, 16, 16, 16)
        no_op_inner.setSpacing(12)

        title = QLabel("No System Operation Analysis")
        title.setObjectName("sectionTitle")

        subtitle = QLabel("Employees who clocked in but had no system operations.")
        subtitle.setObjectName("sectionSubtitle")

        # ✅ 表
        self.no_op_table = QTableWidget(0, 4)
        self.no_op_table.setHorizontalHeaderLabels(
            ["Employee No", "Name", "Clock In", "Clock Out"]
        )

        self.no_op_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.no_op_table.verticalHeader().setVisible(False)
        self.no_op_table.setAlternatingRowColors(True)
        self.no_op_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        no_op_inner.addWidget(title)
        no_op_inner.addWidget(subtitle)
        no_op_inner.addWidget(self.no_op_table)

        no_op_layout.addWidget(no_op_card)
        # =============================
        # signals（保持原逻辑）
        # =============================
        self.card1.button.clicked.connect(lambda: self.load_f(1))
        self.card2.button.clicked.connect(lambda: self.load_f(2))
        self.card3.button.clicked.connect(lambda: self.load_f(3))
        self.card4.button.clicked.connect(lambda: self.load_f(4))

        self.gen_btn.clicked.connect(self.generate)
        self.date_combo.currentTextChanged.connect(self.sync_date_from_main)
        self.date_combo.currentTextChanged.connect(self.refresh_all)
        self.header_date_combo.currentTextChanged.connect(self.sync_date_from_header)

        # 👉 防止底部贴死
        root.addStretch()

        # 状态栏
        status = QStatusBar()
        status.showMessage("Ready")
        self.setStatusBar(status)

    def refresh_no_op_table(self):
        self.no_op_table.setRowCount(0)

        if hasattr(self, "no_op_df") and not self.no_op_df.empty:
            for _, row in self.no_op_df.iterrows():
                r = self.no_op_table.rowCount()
                self.no_op_table.insertRow(r)

                self.no_op_table.setItem(r, 0, QTableWidgetItem(str(row["employee_no"])))
                self.no_op_table.setItem(r, 1, QTableWidgetItem(str(row["real_name"])))
                self.no_op_table.setItem(r, 2, QTableWidgetItem(str(row["上班时间"])))
                self.no_op_table.setItem(r, 3, QTableWidgetItem(str(row["下班时间"])))

    def apply_styles(self):
        self.setStyleSheet(
            f"""
            QMainWindow {{
                background: {THEME['bg']};
            }}
            QLabel {{
                color: {THEME['text']};
            }}
            #heroCard, #sectionCard {{
                background: {THEME['surface']};
                border: 1px solid {THEME['border_soft']};
                border-radius: 18px;
            }}
            #heroTitle {{
                font-size: 26px;
                font-weight: 800;
                color: {THEME['text']};
            }}
            #heroSubtitle {{
                font-size: 13px;
                color: {THEME['muted']};
            }}
            #dateBadge {{
                background: {THEME['surface_alt']};
                border: 1px solid {THEME['border']};
                border-radius: 10px;
                padding: 8px 12px;
                font-weight: 600;
                color: {THEME['muted']};
            }}
            #sectionTitle {{
                font-size: 16px;
                font-weight: 700;
                color: {THEME['text']};
            }}
            #sectionSubtitle {{
                font-size: 12px;
                color: {THEME['muted']};
            }}
            #uploadCard {{
                background: {THEME['surface_alt']};
                border: 1px solid {THEME['border']};
                border-radius: 16px;
            }}
            #cardTitle {{
                font-size: 13px;
                font-weight: 700;
            }}
            #cardSubtitle, #fileNameLabel {{
                color: {THEME['muted']};
                font-size: 12px;
            }}
            QPushButton {{
                background: {THEME['surface']};
                border: 1px solid {THEME['border']};
                border-radius: 10px;
                padding: 8px 14px;
                font-weight: 600;
                color: {THEME['text']};
            }}
            QPushButton:hover {{
                background: #F8FAFC;
                border-color: #CBD5E1;
            }}
            QPushButton:pressed {{
                background: #EEF2F7;
            }}
            QPushButton#primaryButton {{
                background: {THEME['primary']};
                color: white;
                border: none;
                padding: 10px 18px;
                font-weight: 700;
            }}
            QPushButton#primaryButton:hover {{
                background: {THEME['primary_dark']};
            }}
            QPushButton#ghostButton {{
                background: transparent;
                border: 1px solid {THEME['border']};
                color: {THEME['muted']};
                font-size: 14px;
                padding: 0;
            }}
            QPushButton#ghostButton:hover {{
                color: {THEME['text']};
                background: #F8FAFC;
            }}
            QComboBox {{
                background: {THEME['surface']};
                border: 1px solid {THEME['border']};
                border-radius: 10px;
                padding: 8px 10px;
                min-height: 20px;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 24px;
            }}
            QTabWidget::pane {{
                border: none;
                background: transparent;
                margin-top: 8px;
            }}
            QTabBar::tab {{
                background: {THEME['surface']};
                border: 1px solid {THEME['border']};
                border-radius: 10px;
                padding: 10px 18px;
                margin-right: 8px;
                color: {THEME['muted']};
                font-weight: 600;
            }}
            QTabBar::tab:selected {{
                background: {THEME['primary']};
                color: white;
                border: none;
            }}
            QTableWidget {{
                background: {THEME['surface']};
                border: 1px solid {THEME['border_soft']};
                border-radius: 14px;
                alternate-background-color: #FAFCFF;
                gridline-color: {THEME['border_soft']};
                color: {THEME['text']};
            }}
            QHeaderView::section {{
                background: #F8FAFC;
                border: none;
                border-bottom: 1px solid {THEME['border']};
                padding: 10px;
                font-weight: 700;
                color: {THEME['muted']};
            }}
            QStatusBar {{
                background: transparent;
                color: {THEME['muted']};
            }}
            """
        )

    def create_kpi(self, title: str, value: str) -> QFrame:
        frame = QFrame()
        frame.setObjectName("kpiCard")
        frame.setStyleSheet(
            f"""
            QFrame#kpiCard {{
                background: {THEME['kpi_bg']};
                border: 1px solid #DBEAFE;
                border-radius: 16px;
            }}
            """
        )
        frame.setMinimumHeight(126)
        apply_shadow(frame, blur=16, y_offset=4, alpha=14)

        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(6)

        title_label = QLabel(title)
        title_label.setStyleSheet(f"font-size: 11px; font-weight: 700; color: {THEME['muted']};")
        value_label = QLabel(value)
        value_label.setStyleSheet(f"font-size: 24px; font-weight: 800; color: {THEME['primary']};")

        layout.addWidget(title_label)
        layout.addStretch()
        layout.addWidget(value_label)
        frame.val = value_label
        return frame

    def create_efficiency_card(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("kpiCard")
        frame.setStyleSheet(
            f"""
            QFrame#kpiCard {{
                background: {THEME['kpi_bg']};
                border: 1px solid #DBEAFE;
                border-radius: 16px;
            }}
            """
        )
        frame.setMinimumHeight(126)
        frame.setMinimumWidth(240)
        apply_shadow(frame, blur=16, y_offset=4, alpha=14)

        layout = QVBoxLayout(frame)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(4)

        title = QLabel("AVG EFFICIENCY")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(f"font-size: 11px; font-weight: 700; color: {THEME['muted']};")
        layout.addWidget(title)

        self.eff_fig = Figure(figsize=(2.0, 2.0), dpi=84)
        self.eff_fig.patch.set_facecolor(THEME["kpi_bg"])
        self.eff_can = FigureCanvas(self.eff_fig)
        self.eff_can.setStyleSheet("background: transparent; border: none;")
        layout.addWidget(self.eff_can)
        return frame

    def create_gantt_tab(self, parent: QWidget, shift: str):
        root = QVBoxLayout(parent)
        root.setContentsMargins(12, 12, 12, 12)

        card = QFrame()
        card.setObjectName("sectionCard")
        apply_shadow(card, blur=20, y_offset=6, alpha=18)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 16, 16, 16)
        card.setMinimumHeight(800)
        card_layout.setSpacing(0)

        top = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel("Morning Shift Timeline" if shift == "morning" else "Evening Shift Timeline")
        title.setObjectName("sectionTitle")
        subtitle = QLabel("Efficiency-ranked work timeline with hover details and day-over-day trend.")
        subtitle.setObjectName("sectionSubtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        group_combo = QComboBox()
        group_combo.addItem("All Groups")
        group_combo.setMinimumWidth(220)
        group_combo.setCursor(Qt.CursorShape.PointingHandCursor)
        group_combo.currentTextChanged.connect(lambda: self.refresh_single_gantt(shift))

        top.addLayout(title_box, 1)
        top.addWidget(QLabel("Filter Group"))
        top.addWidget(group_combo)
        card_layout.addLayout(top)

        placeholder = QLabel("Upload data and generate to see chart")
        placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        placeholder.setStyleSheet(f"color: {THEME['muted']}; font-size: 13px; padding: 40px;")

        fig = Figure()
        fig.patch.set_alpha(0)
        canvas = FigureCanvas(fig)
        canvas.setStyleSheet("border: none; background: transparent;")
        canvas.hide()

        # ======================
        # ✅ 新增：独立滚动区域
        # ======================
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # ======================
        # scroll 内容
        # ======================
        scroll_container = QWidget()
        scroll_layout = QVBoxLayout(scroll_container)
        scroll_layout.setContentsMargins(0, 0, 0, 0)

        # ✅ 只保留一次
        scroll_layout.addWidget(canvas)

        scroll_area.setWidget(scroll_container)
        scroll_area.setMinimumHeight(500) 

        # ======================
        # 固定时间轴
        # ======================
        axis_fig = Figure(figsize=(14, 1.2), dpi=100)
        axis_canvas = FigureCanvas(axis_fig)
        axis_canvas.setStyleSheet("background: transparent; border: none;")

        # 保存引用
        if shift == "morning":
            self.axis_canvas_m = axis_canvas
            self.axis_fig_m = axis_fig
        else:
            self.axis_canvas_e = axis_canvas
            self.axis_fig_e = axis_fig

        # ======================
        # 加入布局（顺序很关键）
        # ======================
        card_layout.addWidget(scroll_area)   # ✅ 表（滚动）
        card_layout.addWidget(axis_canvas)   # ✅ 时间轴（固定）
        card_layout.addWidget(placeholder)
        root.addWidget(card)
        root.addStretch() 
        root.addSpacing(50)
        return fig, fig.add_subplot(111), canvas, placeholder, group_combo

    def toggle_kpi(self):
        self._kpi_visible = not self._kpi_visible
        self.kpi_row_widget.setVisible(self._kpi_visible)
        self.kpi_toggle_btn.setText("▾" if self._kpi_visible else "▸")

    def load_f(self, index: int):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return

        setattr(self, f"file{index}", path)
        getattr(self, f"card{index}").set_file(path)
        self.statusBar().showMessage(f"Loaded file {index}: {os.path.basename(path)}", 4000)

    def generate(self):
        if not (hasattr(self, "file1") and hasattr(self, "file2")):
            self.statusBar().showMessage("Please upload ISC and iAMS files first.", 5000)
            return

        self.gen_btn.setText("Generating...")
        self.gen_btn.setEnabled(False)
        QApplication.processEvents()

        try:
            df, df2, group_map = load_data(self.file1, self.file2)
            self.df = df
            self.df2 = df2
            # ======================
            # ✅ 提取仓库
            # ======================
            df["warehouse_short"] = df["warehouse_name"].apply(extract_warehouse)

            # 👉 当前仓库（默认取最多的）
            self.current_warehouse = df["warehouse_short"].mode()[0]
            volume_data = load_volume_data(self.file3) if hasattr(self, "file3") else {}
            break_df = load_break_data(self.file4) if hasattr(self, "file4") else None

            self.volume_data = volume_data
            self.timeline = build_timeline(df, df2, group_map, break_df)

            self.history_ratios = {}
            self.history_wait = {}

            for d, day_data in self.timeline.groupby("date"):
                d_str = d.strftime("%Y-%m-%d")
                self.history_ratios[d_str] = {}
                self.history_wait[d_str] = {}

                for name, group in day_data.groupby("name"):
                    work_duration = group[group["type"].isin(["work", "overnight"])]["duration"].sum()
                    total_duration = group["duration"].sum()
                    self.history_ratios[d_str][name] = (
                        work_duration / total_duration * 100 if total_duration > 0 else 0
                    )

                    if group["is_absent"].any():
                        continue

                    work = group[group["type"].isin(["work", "overnight"])].sort_values("start")
                    punch_in = group["start"].min()
                    if not work.empty:
                        first_task = work.iloc[0]["start"]
                        wait_seconds = (first_task - punch_in).total_seconds()
                        if "(" in name:
                            emp_no = name.split("(")[-1].replace(")", "").strip()
                        else:
                            emp_no = name
                        self.history_wait[d_str][emp_no] = wait_seconds

            dates = sorted(self.timeline["date"].astype(str).unique()) if not self.timeline.empty else []
            self.date_combo.blockSignals(True)
            self.date_combo.clear()

            dates = sorted(self.timeline["date"].astype(str).unique())

            self.date_combo.clear()
            self.date_combo.addItems(dates)

            # ✅ 在设置完日期之后再取
            selected_date = pd.to_datetime(self.date_combo.currentText()).date()

            df2_day = df2[
                df2["上班时间"].dt.date == selected_date
            ]

            self.no_op_df = find_no_operation_people(df, df2_day)
            self.header_date_combo.addItems(dates)
            self.date_combo.blockSignals(False)

            self.refresh_all()
            self.statusBar().showMessage("Dashboard generated successfully.", 5000)
        except Exception as exc:
            self.statusBar().showMessage(f"Generate failed: {exc}", 8000)
            raise
        finally:
            self.gen_btn.setText("Generate Dashboard")
            self.gen_btn.setEnabled(True)

    def refresh_all(self):
        # ✅ 每次切换日期重新计算 no operation
        selected_date = pd.to_datetime(self.date_combo.currentText()).date()

        df2_day = self.df2[
            self.df2["上班时间"].apply(adjust_work_date) == selected_date
        ]

        # ✅ 只用当天 ISC 数据
        df_day_isc = self.df[
            self.df["work_date"] == selected_date
        ]

        self.no_op_df = find_no_operation_people(df_day_isc, df2_day)
        date = self.date_combo.currentText()
        if not date or self.timeline.empty:
            return
        df_day = self.timeline[self.timeline["date"].astype(str) == date]

        self.k_date.val.setText(self.current_warehouse)
        self.k_workers.val.setText(str(df_day["name"].nunique()))

        total_man_hours = df_day["duration"].sum()
        self.k_total_manhours.val.setText(f"{total_man_hours:.1f}h")

        valid_man_hours = df_day[df_day["type"].isin(["work", "overnight"])]["duration"].sum()
        eff_ratio = (valid_man_hours / total_man_hours * 100) if total_man_hours > 0 else 0
        self.update_efficiency_donut(eff_ratio)

        volume = int(self.volume_data.get(pd.to_datetime(date).date(), 0)) if self.volume_data else 0
        self.k_vol.val.setText(str(volume))

        for combo, shift in [(self.group_m, "morning"), (self.group_e, "evening")]:
            groups = sorted(df_day[df_day["shift"] == shift]["group"].dropna().astype(str).unique())
            combo.blockSignals(True)
            current = combo.currentText()
            combo.clear()
            combo.addItem("All Groups")
            combo.addItems(groups)
            combo.setCurrentText(current if current in groups else "All Groups")
            combo.blockSignals(False)
            self.refresh_single_gantt(shift)

        self.refresh_fa_table(df_day)
        self.refresh_no_op_table()

    def update_efficiency_donut(self, ratio: float):
        self.eff_fig.clear()
        ax = self.eff_fig.add_subplot(111)
        ax.set_facecolor(THEME["kpi_bg"])
        ax.pie(
            [ratio, max(0, 100 - ratio)],
            colors=[THEME["primary"], "#D7E3FF"],
            startangle=90,
            counterclock=False,
            wedgeprops={"width": 0.34, "edgecolor": THEME["kpi_bg"]},
        )
        ax.text(
            0,
            0,
            f"{ratio:.1f}%",
            ha="center",
            va="center",
            fontsize=13,
            fontweight="bold",
            color=THEME["primary"],
        )
        ax.set_aspect("equal")
        self.eff_can.draw_idle()

    def refresh_single_gantt(self, shift: str):
        date = self.date_combo.currentText()
        if not date or self.timeline.empty:
            return

        df_shift = self.timeline[
            (self.timeline["date"].astype(str) == date) & (self.timeline["shift"] == shift)
        ].copy()

        combo = self.group_m if shift == "morning" else self.group_e
        if combo.currentText() != "All Groups":
            df_shift = df_shift[df_shift["group"] == combo.currentText()]

        for name in df_shift["name"].unique():
            group = df_shift[df_shift["name"] == name]

            work_duration = group[group["type"].isin(["work", "overnight"])]["duration"].sum()
            total_duration = group["duration"].sum()

            # ✅ 防止空
            ratio = (work_duration / total_duration * 100) if total_duration > 0 else 0

            self.summary[name] = {
                "ratio": ratio,
                "in": group["start"].min().strftime("%H:%M") if not group.empty else "--",
                "out": group["end"].max().strftime("%H:%M") if not group.empty else "--",
            }

        ax, canvas, placeholder = (
            (self.ax_m, self.can_m, self.ph_m) if shift == "morning" else (self.ax_e, self.can_e, self.ph_e)
        )

        if df_shift.empty:
            canvas.hide()
            placeholder.show()
            return

        placeholder.hide()
        canvas.show()
        self.draw_chart(ax, canvas, df_shift, shift)

    def draw_chart(self, ax, canvas, df_shift, shift: str):
        fig = canvas.figure
        fig.clear()
        fig.patch.set_alpha(0)
        fig.patch.set_edgecolor("none")
        ax = fig.add_subplot(111)
        ax.set_facecolor("none")

        current_date_str = self.date_combo.currentText()
        prev_date = pd.to_datetime(current_date_str).date() - pd.Timedelta(days=1)
        prev_date_str = prev_date.strftime("%Y-%m-%d")

        names = sorted(df_shift["name"].unique(), key=lambda n: self.summary.get(n, {}).get("ratio", 0))
        # ======================
        # ✅ 动态高度（关键）
        # ======================
        row_height = 0.4   # 每个人高度（可以调 0.3~0.6）
        fig_height = len(names) * 0.4

        fig.set_size_inches(14, fig_height)  # 宽固定，高动态
        y_map = {name: i for i, name in enumerate(names)}

        for _, row in df_shift.iterrows():
            ax.barh(
                y_map[row["name"]],
                mdates.date2num(row["end"]) - mdates.date2num(row["start"]),
                left=mdates.date2num(row["start"]),
                color=THEME.get(row["type"], "#D1D5DB"),
                height=0.62,
                edgecolor="none",
            )

        

        for i, name in enumerate(names):
            ratio = self.summary.get(name, {}).get("ratio", 0)
            prev_ratio = self.history_ratios.get(prev_date_str, {}).get(name)

            arrow = ""
            color = THEME["muted"]

            if prev_ratio is not None:
                if ratio > prev_ratio:
                    arrow = " ↑"
                    color = THEME["success"]
                elif ratio < prev_ratio:
                    arrow = " ↓"
                    color = THEME["danger"]

            # ✅ 用 ax（不是 ax_right）
            ax.text(
                1.01,
                i,
                f"{ratio:.1f}%",
                transform=ax.get_yaxis_transform(),
                va="center",
                fontsize=8,
                color=THEME["text"],
            )

            if arrow:
                ax.text(
                    1.08,
                    i,
                    arrow,
                    transform=ax.get_yaxis_transform(),
                    va="center",
                    fontsize=9,
                    color=color,
                    fontweight="bold",
                )

        

        ax.set_yticks(range(len(names)))
        ax.set_ylim(-0.5, len(names) - 0.5)
        ax.set_yticklabels(names, fontsize=8, color=THEME["text"])
        ax.tick_params(axis="x", colors=THEME["muted"], labelsize=8)
        ax.tick_params(axis="y", length=0)

        ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
        ax.grid(axis="x", color="#E2E8F0", linestyle="--", alpha=0.65)
        # ✅ 限制白班只显示到18:00
        if shift == "morning":
            start_time = pd.to_datetime(self.date_combo.currentText() + " 07:00")
            end_time = pd.to_datetime(self.date_combo.currentText() + " 18:00")
            ax.set_xlim(start_time, end_time)
        legend_items = [
            Patch(facecolor=THEME["work"], label="Work"),
            Patch(facecolor=THEME["overnight"], label="Overnight"),
            Patch(facecolor=THEME["break"], label="Break"),
            Patch(facecolor=THEME["idle"], label="Idle"),
        ]
        ax.legend(
            handles=legend_items,
            loc="upper right",
            bbox_to_anchor=(1.17, 1.0),
            frameon=False,
            fontsize=8,
        )

        for side in ["top", "right", "bottom", "left"]:
            ax.spines[side].set_visible(False)

        fig.subplots_adjust(left=0.22, right=0.88, top=0.98, bottom=0.00)        
        canvas.draw_idle()
        # ======================
        # ✅ 让内容变高 → 触发滚动
        # ======================
        row_height = 30   # 每人高度（像素）
        total_height = max(400, len(names) * row_height)

        canvas.setMinimumHeight(total_height)

        if shift == "morning":
            if self.hover_cid_m:
                canvas.mpl_disconnect(self.hover_cid_m)
            self.hover_cid_m = canvas.mpl_connect("motion_notify_event", lambda event: self.on_hover(event, names))
        else:
            if self.hover_cid_e:
                canvas.mpl_disconnect(self.hover_cid_e)
            self.hover_cid_e = canvas.mpl_connect("motion_notify_event", lambda event: self.on_hover(event, names))
        # ======================
        # ✅ 固定时间轴
        # ======================
        axis_fig = self.axis_fig_m if shift == "morning" else self.axis_fig_e
        axis_canvas = self.axis_canvas_m if shift == "morning" else self.axis_canvas_e

        axis_fig.clear()
        ax2 = axis_fig.add_subplot(111)

        # 同步时间范围
        ax2.set_xlim(ax.get_xlim())

        # 只保留 x 轴
        ax2.yaxis.set_visible(False)

        ax2.xaxis.set_major_locator(mdates.HourLocator(interval=1))
        ax2.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

        ax2.tick_params(axis="x", labelsize=9)

        # 去边框
        for side in ["top", "right", "left"]:
            ax2.spines[side].set_visible(False)

        axis_fig.subplots_adjust(left=0.22, right=0.88, top=1.0, bottom=0.4)

        axis_canvas.draw_idle()

    def on_hover(self, event, names):
        if event.inaxes is None or event.ydata is None:
            QToolTip.hideText()
            return

        for i, name in enumerate(names):
            # ✅ 判断是否在这一行范围内
            if i - 0.5 <= event.ydata <= i + 0.5:
                summary = self.summary.get(name, {
                    "in": "--",
                    "out": "--",
                    "ratio": 0
                })

                QToolTip.showText(
                    QCursor.pos(),
                    f"<b>{name}</b><br>In: {summary['in']}<br>Out: {summary['out']}<br>Eff: {summary['ratio']:.1f}%",
                )
                return

        QToolTip.hideText()

    def refresh_fa_table(self, df_day):
        current_date = self.date_combo.currentText()
        prev_date = (pd.to_datetime(current_date) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        self.fa_table.setRowCount(0)

        records = []
        for name, group in df_day.groupby("name"):
            if group["is_absent"].any():
                continue

            emp_no = name.split("(")[-1].replace(")", "").strip() if "(" in name else name
            work = group[group["type"].isin(["work", "overnight"])].sort_values("start")
            punch_in = group["start"].min()

            if not work.empty:
                first_task = work.iloc[0]["start"]
                wait_seconds = (first_task - punch_in).total_seconds()
                prev_wait = self.history_wait.get(prev_date, {}).get(emp_no)
                trend = "-"

                if prev_wait is not None:
                    if wait_seconds < prev_wait:
                        trend = "↓"
                    elif wait_seconds > prev_wait:
                        trend = "↑"
                    else:
                        trend = "="

                records.append(
                    [
                        name,
                        group["group"].iloc[0],
                        group["shift"].iloc[0],
                        punch_in.strftime("%H:%M:%S"),
                        first_task.strftime("%H:%M:%S"),
                        f"{int(wait_seconds // 60)}m {int(wait_seconds % 60)}s",
                        trend,
                        wait_seconds,
                    ]
                )

        records.sort(key=lambda x: x[7], reverse=True)

        for row_index, row in enumerate(records):
            self.fa_table.insertRow(row_index)
            for col_index in range(7):
                item = QTableWidgetItem(str(row[col_index]))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                if col_index == 5 and row[7] > 300:
                    item.setForeground(QColor(THEME["danger"]))
                if col_index == 6:
                    if row[6] == "↑":
                        item.setForeground(QColor(THEME["danger"]))
                    elif row[6] == "↓":
                        item.setForeground(QColor(THEME["success"]))

                self.fa_table.setItem(row_index, col_index, item)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    QToolTip.setFont(QFont("Segoe UI", 9))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())

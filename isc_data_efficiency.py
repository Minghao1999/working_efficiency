import sys
import os
import pandas as pd
import warnings
from PyQt6.QtWidgets import *
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QCursor, QFont, QColor
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.dates as mdates
from matplotlib.patches import Patch

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================
# 主题配置
# =========================
THEME = {
    "bg": "#F5F7FA",
    "card": "#FFFFFF",
    "primary": "#2563EB",
    "work": "#22C55E",
    "overnight": "#FACC15",
    "idle": "#D1D5DB",
    "text": "#111827",
    "danger": "#EF4444"
}

# =========================
# 数据逻辑
# =========================
def load_data(file1, file2):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    df2["employee_no"] = df2.iloc[:, 1].astype(str).str.strip()
    df2["real_name"] = df2.iloc[:, 3].astype(str).str.strip()
    df2["上班时间"] = pd.to_datetime(df2.iloc[:, 13], errors='coerce')
    df2["下班时间"] = pd.to_datetime(df2.iloc[:, 18], errors='coerce')
    df2["attendance_group"] = df2.iloc[:, 26].astype(str).str.strip().replace("nan", "未分组")
    
    mapping = dict(zip(df2["employee_no"], df2["real_name"]))
    group_mapping = dict(zip(df2["employee_no"], df2["attendance_group"]))

    df = df1.rename(columns={"employee_no": "employee_no", "任务单开始时间": "start", "任务单结束时间": "end", "operator": "operator"})
    df["employee_no"] = df["employee_no"].astype(str).str.strip().replace(["nan", "None", ""], pd.NA)
    df["start"] = pd.to_datetime(df["start"], errors='coerce')
    df["end"] = pd.to_datetime(df["end"], errors='coerce')

    def identify_name(row):
        emp_no = row["employee_no"]
        if pd.notna(emp_no) and emp_no in mapping: return f"{mapping[emp_no]} ({emp_no})"
        if pd.notna(emp_no): return str(emp_no)
        return str(row.get("operator", "Unknown"))

    df["display_name"] = df.apply(identify_name, axis=1)
    df["type"], df["date"] = "work", df["start"].dt.date
    
    return df, df2, group_mapping

def build_timeline(df, df2, group_mapping):
    records, processed_names = [], set()
    
    for _, row2 in df2.iterrows():
        emp_no = str(row2["employee_no"])
        name = str(row2["real_name"])
        display_name = f"{name} ({emp_no})"
        group_name = str(row2["attendance_group"])
        
        work_start, work_end = row2["上班时间"], row2["下班时间"]
        if pd.isna(work_start) or pd.isna(work_end): continue
        
        day = work_start.date()
        processed_names.add((display_name, day))
        shift_tag = "morning" if "早" in str(row2.iloc[7]) else "evening"
        
        person_tasks = df[(df["employee_no"] == emp_no) &
                          (df["start"] < work_end) &
                          (df["end"] > work_start)].sort_values("start")
        
        current_cursor = work_start
        for _, task in person_tasks.iterrows():
            t_s, t_e = max(task["start"], work_start), min(task["end"], work_end)
            if t_s >= t_e: continue
            t_e_disp = t_s + pd.Timedelta(minutes=1) if (t_e - t_s).total_seconds() < 60 else t_e
            t_e_disp = min(t_e_disp, work_end)
            is_overnight = task["start"].date() != task["end"].date()
            if t_s > current_cursor:
                records.append([display_name, day, current_cursor, t_s, "idle", shift_tag, False, group_name])
            records.append([display_name, day, t_s, t_e_disp, "overnight" if is_overnight else "work", shift_tag, False, group_name])
            current_cursor = max(current_cursor, t_e_disp + pd.Timedelta(seconds=1))
            
        if current_cursor < work_end:
            records.append([display_name, day, current_cursor, work_end, "idle", shift_tag, False, group_name])

    d1_dates = set(df["start"].dropna().dt.date.unique())
    for current_day in sorted(list(d1_dates)):
        day_tasks_all = df[df["start"].dt.date == current_day]
        for d_name, group in day_tasks_all.groupby("display_name"):
            if (d_name, current_day) in processed_names: continue
            emp_id_search = d_name.split('(')[-1].strip(')')
            this_group = group_mapping.get(emp_id_search, "未匹配组")
            for _, task in group.sort_values("start").iterrows():
                t_s, t_e = task["start"], task["end"]
                shift_tag = "morning" if (t_s.hour >= 6 and t_s.hour < 18) else "evening"
                t_e_disp = t_s + pd.Timedelta(minutes=1) if (t_e - t_s).total_seconds() < 60 else t_e
                records.append([d_name, current_day, t_s, t_e_disp, "work", shift_tag, True, this_group])

    tl = pd.DataFrame(records, columns=["name", "date", "start", "end", "type", "shift", "is_absent", "group"])
    tl["duration"] = (tl["end"] - tl["start"]).dt.total_seconds() / 3600
    return tl

# =========================
# 主界面
# =========================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Operations Dashboard v2.0")
        self.resize(1450, 900)
        self.ax_right_m = self.ax_right_e = None
        self.hover_cid_m = self.hover_cid_e = None
        self.timeline = pd.DataFrame()
        self.summary = {}
        self.init_ui()

    def download_fa_excel(self):
        if self.timeline.empty:
            QMessageBox.warning(self, "Warning", "No data to export")
            return

        date = self.date_combo.currentText()
        group_filter = self.group_fa.currentText()

        df_day = self.timeline[self.timeline["date"].astype(str) == date].copy()
        if group_filter != "All Groups":
            df_day = df_day[df_day["group"] == group_filter]

        export_data = []

        for name, g in df_day.groupby("name"):
            if g["is_absent"].any():
                continue

            work_tasks = g[g["type"].isin(["work", "overnight"])].sort_values("start")
            punch_in = g["start"].min()

            first_task_str = ""
            wait_str = ""

            if not work_tasks.empty:
                first_task = work_tasks.iloc[0]["start"]
                first_task_str = first_task.strftime("%H:%M:%S")

                diff = (first_task - punch_in).total_seconds()
                wait_sec = max(0, diff)

                mm, ss = divmod(int(wait_sec), 60)
                hh, mm = divmod(mm, 60)
                wait_str = f"{hh:02d}:{mm:02d}:{ss:02d}" if hh > 0 else f"{mm}m {ss}s"

            export_data.append([
                name,
                g["group"].iloc[0],
                g["shift"].iloc[0],
                punch_in.strftime("%H:%M:%S"),
                first_task_str,
                wait_str
            ])

        df_export = pd.DataFrame(export_data, columns=[
            "Name", "Group", "Shift", "Punch In", "First Task", "Wait"
        ])

        # 选择保存路径
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "first_action.xlsx", "*.xlsx")

        if file_path:
            df_export.to_excel(file_path, index=False)
            QMessageBox.information(self, "Success", "Exported successfully!")

    def init_ui(self):
        main = QWidget(); self.setCentralWidget(main); 
        layout = QVBoxLayout(main)
        top_card = QFrame(); top_card.setStyleSheet("background: white; border-radius: 12px;"); top_layout = QHBoxLayout(top_card)
        self.btn1 = QPushButton("Upload isc"); self.btn2 = QPushButton("Upload iAMS"); 
        self.gen_btn = QPushButton("Generate")
        self.gen_btn.setFixedWidth(180)
        self.gen_btn.setFixedHeight(36)
        self.gen_btn.setMinimumWidth(160)
        self.label1, self.label2 = QLabel("No file"), QLabel("No file")
        
        self.date_combo = QComboBox()
        self.date_combo.setMaxVisibleItems(15)
        self.date_combo.view().setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.date_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        
        self.gen_btn.setStyleSheet("""
            QPushButton {
                background:#2563EB;
                color:white;
                border-radius:10px;
                padding:6px;
                font-weight:bold;
            }
            QPushButton:hover {
                background:#1D4ED8;
            }
            QPushButton:pressed {
                background:#1E40AF;
            }
            QPushButton:disabled {
                background:#94A3B8;
                color:white;
            }
            """)
        self.btn1.setStyleSheet("QPushButton { background:#E2E8F0; color:#1E293B; border-radius:8px; padding:6px 14px; }")
        self.btn2.setStyleSheet("QPushButton { background:#E2E8F0; color:#1E293B; border-radius:8px; padding:6px 14px; }")
        
        self.btn1.clicked.connect(self.load_f1); self.btn2.clicked.connect(self.load_f2); self.gen_btn.clicked.connect(self.generate)
        top_layout.addWidget(self.btn1); top_layout.addWidget(self.label1); top_layout.addSpacing(15)
        top_layout.addWidget(self.btn2); top_layout.addWidget(self.label2); top_layout.addSpacing(25)
        top_layout.addWidget(QLabel("Date Selection:")); top_layout.addWidget(self.date_combo)
        top_layout.addStretch(); top_layout.addWidget(self.gen_btn)
        layout.addWidget(top_card)

        self.tabs = QTabWidget()
        self.tab_m, self.tab_e, self.tab_fa = QWidget(), QWidget(), QWidget()
        self.tabs.addTab(self.tab_m, "Morning Gantt"); self.tabs.addTab(self.tab_e, "Evening Gantt"); self.tabs.addTab(self.tab_fa, "First Action Analysis")
        layout.addWidget(self.tabs)

        self.fig_m, self.ax_m, self.can_m, self.ph_m, self.group_m = self.create_gantt_tab(self.tab_m, "morning")
        self.fig_e, self.ax_e, self.can_e, self.ph_e, self.group_e = self.create_gantt_tab(self.tab_e, "evening")
        
        fa_outer_layout = QVBoxLayout(self.tab_fa)
        fa_ctrl_layout = QHBoxLayout()
        self.group_fa = QComboBox()
        self.download_btn = QPushButton("Download Excel")
        self.download_btn.setFixedWidth(150)
        self.download_btn.clicked.connect(self.download_fa_excel)
        self.group_fa.addItem("All Groups")
        self.group_fa.setFixedWidth(250)
        self.group_fa.currentTextChanged.connect(self.refresh_fa_table)
        fa_ctrl_layout.addWidget(QLabel("Filter Analysis Group:"))
        fa_ctrl_layout.addWidget(self.group_fa)
        fa_ctrl_layout.addWidget(self.download_btn)
        fa_outer_layout.addLayout(fa_ctrl_layout)

        self.fa_table = QTableWidget(); self.fa_table.setColumnCount(6)
        self.fa_table.setHorizontalHeaderLabels(["Name", "Group", "Shift", "Punch-in", "First Task", "Wait"])
        self.fa_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        fa_outer_layout.addWidget(self.fa_table)

        self.date_combo.currentTextChanged.connect(self.refresh_all)

    def create_gantt_tab(self, parent, shift):
        layout = QVBoxLayout(parent)
        ctrl_layout = QHBoxLayout()
        group_combo = QComboBox()
        group_combo.addItem("All Groups")
        group_combo.setFixedWidth(250)
        group_combo.currentTextChanged.connect(lambda: self.refresh_single_gantt(shift))
        ctrl_layout.addWidget(QLabel(f"Filter {shift.capitalize()} Group:")); ctrl_layout.addWidget(group_combo); ctrl_layout.addStretch()
        layout.addLayout(ctrl_layout)

        ph = QLabel("📊 Upload files to start"); ph.setAlignment(Qt.AlignmentFlag.AlignCenter)
        fig = Figure(figsize=(10, 6)); ax = fig.add_subplot(111); can = FigureCanvas(fig)
        can.hide(); layout.addWidget(can); layout.addWidget(ph)
        return fig, ax, can, ph, group_combo

    def load_f1(self):
        p, _ = QFileDialog.getOpenFileName(self, "isc", "", "*.xlsx")
        if p: self.file1 = p; self.label1.setText(os.path.basename(p))
    def load_f2(self):
        p, _ = QFileDialog.getOpenFileName(self, "iAMS", "", "*.xlsx")
        if p: self.file2 = p; self.label2.setText(os.path.basename(p))

    def generate(self):
        if not hasattr(self, 'file1') or not hasattr(self, 'file2'): return
        
        self.gen_btn.setEnabled(False)
        self.gen_btn.setText("⏳ Processing...")
        QApplication.processEvents()

        try:
            df, df2, group_mapping = load_data(self.file1, self.file2)
            self.timeline = build_timeline(df, df2, group_mapping)
            
            dates = sorted(self.timeline["date"].astype(str).unique())
            self.date_combo.blockSignals(True); self.date_combo.clear(); self.date_combo.addItems(dates); self.date_combo.blockSignals(False)
            self.refresh_all()
        finally:
            self.gen_btn.setEnabled(True)
            self.gen_btn.setText("Generate")

    def refresh_all(self):
        date = self.date_combo.currentText()
        if not date or self.timeline.empty: return
        df_day_all = self.timeline[self.timeline["date"].astype(str) == date]
        
        self._update_combo(self.group_m, sorted(df_day_all[df_day_all["shift"] == "morning"]["group"].unique()))
        self._update_combo(self.group_e, sorted(df_day_all[df_day_all["shift"] == "evening"]["group"].unique()))
        self._update_combo(self.group_fa, sorted(df_day_all["group"].unique()))

        self.refresh_single_gantt("morning")
        self.refresh_single_gantt("evening")
        self.refresh_fa_table()

    def _update_combo(self, combo, items):
        current_val = combo.currentText()
        combo.blockSignals(True); combo.clear(); combo.addItem("All Groups"); combo.addItems(items)
        index = combo.findText(current_val)
        if index >= 0: combo.setCurrentIndex(index)
        else: combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def refresh_single_gantt(self, shift):
        date = self.date_combo.currentText()
        if not date or self.timeline.empty: return
        
        combo = self.group_m if shift == "morning" else self.group_e
        group_filter = combo.currentText()
        df_day = self.timeline[(self.timeline["date"].astype(str) == date) & (self.timeline["shift"] == shift)].copy()
        
        if group_filter != "All Groups":
            df_day = df_day[df_day["group"] == group_filter]

        for name, g in df_day.groupby("name"):
            work_dur = g[g["type"].isin(["work", "overnight"])]["duration"].sum()
            total_dur = g["duration"].sum()
            
            self.summary[name] = {
                "ratio": (work_dur/total_dur*100) if total_dur>0 else 0,
                "work": work_dur, "total": total_dur,
                "punch_in": g["start"].min().strftime("%H:%M:%S"),
                "punch_out": g["end"].max().strftime("%H:%M:%S")
            }

        if shift == "morning":
            self.ph_m.hide(); self.can_m.show()
            self.draw_chart(self.ax_m, self.can_m, df_day, "morning")
        else:
            self.ph_e.hide(); self.can_e.show()
            self.draw_chart(self.ax_e, self.can_e, df_day, "evening")

    def refresh_fa_table(self):
        date = self.date_combo.currentText()
        group_filter = self.group_fa.currentText()
        if not date or self.timeline.empty: return
        df_day = self.timeline[self.timeline["date"].astype(str) == date].copy()
        if group_filter != "All Groups": df_day = df_day[df_day["group"] == group_filter]

        fa_records = []
        for name, g in df_day.groupby("name"):
            if g["is_absent"].any(): continue
            work_tasks = g[g["type"].isin(["work", "overnight"])].sort_values("start")
            punch_in = g["start"].min()
            first_task_str, wait_str, wait_sec = "None", "N/A", 0
            if not work_tasks.empty:
                first_task = work_tasks.iloc[0]["start"]
                first_task_str = first_task.strftime("%H:%M:%S")
                diff = (first_task - punch_in).total_seconds()
                wait_sec = max(0, diff)
                if diff < 0: wait_str = "Early Start"
                else:
                    mm, ss = divmod(int(wait_sec), 60); hh, mm = divmod(mm, 60)
                    wait_str = f"{hh:02d}:{mm:02d}:{ss:02d}" if hh > 0 else f"{mm}m {ss}s"
            fa_records.append([name, g["group"].iloc[0], g["shift"].iloc[0], punch_in.strftime("%H:%M:%S"), first_task_str, wait_str, wait_sec])
        
        self.fa_table.setRowCount(0); fa_records.sort(key=lambda x: x[6], reverse=True)
        for i, row in enumerate(fa_records):
            self.fa_table.insertRow(i)
            for j in range(6):
                item = QTableWidgetItem(str(row[j]))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if j == 5 and row[6] > 1800: item.setForeground(QColor(THEME["danger"])); item.setFont(QFont("Arial", 9, QFont.Weight.Bold))
                self.fa_table.setItem(i, j, item)

    def draw_chart(self, ax, can, df_s, shift):
        if shift == "morning" and self.ax_right_m: self.ax_right_m.remove(); self.ax_right_m = None
        if shift == "evening" and self.ax_right_e: self.ax_right_e.remove(); self.ax_right_e = None
        ax.clear()
        if df_s.empty: can.draw(); return

        names = sorted(df_s["name"].unique(), key=lambda x: self.summary.get(x, {"ratio":0})["ratio"])
        y_map = {p: i for i, p in enumerate(names)}
        
        for _, r in df_s.iterrows():
            ax.barh(y_map[r["name"]], mdates.date2num(r["end"]) - mdates.date2num(r["start"]), 
                    left=mdates.date2num(r["start"]), color=THEME.get(r["type"], THEME["idle"]), height=0.6)

        ax_r = ax.twinx(); ax_r.set_ylim(ax.get_ylim()); ax_r.set_yticks(range(len(names)))
        ax_r.set_yticklabels([f"{self.summary.get(n, {'ratio':0})['ratio']:.1f}% " for n in names], fontsize=9)
        
        ax_r.set_ylabel("Work Efficiency %", fontsize=8, color="#64748B", fontweight='bold', labelpad=10)
        ax_r.yaxis.set_label_position("right") 

        if shift == "morning": self.ax_right_m = ax_r
        else: self.ax_right_e = ax_r
        
        ax.set_yticks(range(len(names))); ax.set_yticklabels(names, fontsize=9)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))

        for side in ['top', 'right', 'bottom', 'left']:
            ax.spines[side].set_visible(False)
            ax_r.spines[side].set_visible(False)
        ax.tick_params(axis='both', which='both', length=0)
        ax_r.tick_params(axis='both', which='both', length=0)

        legend_elements = [
            Patch(facecolor=THEME["work"], label='Work'),
            Patch(facecolor=THEME["overnight"], label='Overnight'),
            Patch(facecolor=THEME["idle"], label='Idle')
        ]
        ax.legend(handles=legend_elements, loc='upper right', bbox_to_anchor=(1.0, 1.15), 
                  ncol=3, fontsize=8, frameon=False)

        can.figure.subplots_adjust(left=0.28, right=0.85, top=0.82, bottom=0.12)
        can.draw()
        
        if shift == "morning":
            if self.hover_cid_m: can.mpl_disconnect(self.hover_cid_m)
            self.hover_cid_m = can.mpl_connect("motion_notify_event", lambda e: self.on_hover(e, ax, names))
        else:
            if self.hover_cid_e: can.mpl_disconnect(self.hover_cid_e)
            self.hover_cid_e = can.mpl_connect("motion_notify_event", lambda e: self.on_hover(e, ax, names))

    def on_hover(self, event, ax, names):
        if event.inaxes is None or event.ydata is None:
            QToolTip.hideText(); return
        y_idx = int(round(event.ydata))
        if 0 <= y_idx < len(names):
            name = names[y_idx]; s = self.summary.get(name)
            if s:
                text = (f"<b>{name}</b><br>"
                        f"Punch In: {s['punch_in']}<br>"
                        f"Punch Out: {s['punch_out']}<br>"
                        f"Efficiency: {s['ratio']:.1f}%<br>"
                        f"Work: {s['work']:.2f}h / {s['total']:.2f}h")
                QToolTip.showText(QCursor.pos(), text)

if __name__ == "__main__":
    app = QApplication(sys.argv); w = MainWindow(); w.show(); sys.exit(app.exec())
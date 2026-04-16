import sys

import os

import pandas as pd

import warnings

from PyQt6.QtWidgets import *

from PyQt6.QtCore import QTimer, Qt

from PyQt6.QtGui import QCursor

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas

from matplotlib.figure import Figure

import matplotlib.dates as mdates

from matplotlib.patches import Patch



# 忽略 openpyxl 样式警告

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")



# =========================

# 主题配置

# =========================

THEME = {

    "bg": "#F5F7FA",

    "card": "#FFFFFF",

    "primary": "#2563EB",

    "work": "#22C55E",

    "overnight": "#FACC15", #(跨天任务)

    "idle": "#D1D5DB",

    "text": "#111827"

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

   

    mapping = dict(zip(df2["employee_no"], df2["real_name"]))



    df = df1.rename(columns={"employee_no": "employee_no", "任务单开始时间": "start", "任务单结束时间": "end", "operator": "operator"})

    df["employee_no"] = df["employee_no"].astype(str).str.strip().replace(["nan", "None", ""], pd.NA)

    df["operator"] = df["operator"].astype(str).str.strip().replace(["nan", "None", ""], pd.NA)

    df["start"] = pd.to_datetime(df["start"], errors='coerce')

    df["end"] = pd.to_datetime(df["end"], errors='coerce')



    def identify_name(row):

        emp_no, oper = row["employee_no"], row["operator"]

        if pd.notna(emp_no) and emp_no in mapping: return f"{mapping[emp_no]} ({emp_no})"

        if pd.notna(emp_no): return str(emp_no)

        if pd.notna(oper): return str(oper)

        return "Unknown"



    df["display_name"] = df.apply(identify_name, axis=1)

    df["type"], df["date"] = "work", df["start"].dt.date

    return df, df2



def build_timeline(df, df2):

    records, processed_names = [], set()

   

    # --- 第一部分：处理有考勤记录的人 (严格按表2班次) ---

    for _, row2 in df2.iterrows():

        emp_no = str(row2["employee_no"])

        name = str(row2["real_name"])

        display_name = f"{name} ({emp_no})"

       

        work_start, work_end = row2["上班时间"], row2["下班时间"]

        if pd.isna(work_start) or pd.isna(work_end): continue

       

        day = work_start.date()

        processed_names.add((display_name, day))

       

        # 严格提取表2的班次字段

        shift_tag = "morning" if "早" in str(row2.iloc[7]) else "evening"

       

        person_tasks = df[(df["employee_no"] == emp_no) &

                          (df["start"] < work_end) &

                          (df["end"] > work_start)].sort_values("start")

       

        current_cursor = work_start

        for _, task in person_tasks.iterrows():

            t_s, t_e = max(task["start"], work_start), min(task["end"], work_end)

            if t_s >= t_e: continue

           

            # 强制1分钟可见逻辑

            t_e_disp = t_s + pd.Timedelta(minutes=1) if (t_e - t_s).total_seconds() < 60 else t_e

            t_e_disp = min(t_e_disp, work_end) # 有考勤者严禁超过下班边界



            is_overnight = task["start"].date() != task["end"].date()

            if t_s > current_cursor:

                records.append([display_name, day, current_cursor, t_s, "idle", shift_tag, False])

           

            records.append([display_name, day, t_s, t_e_disp, "overnight" if is_overnight else "work", shift_tag, False])

            current_cursor = max(current_cursor, t_e_disp + pd.Timedelta(seconds=1))

           

        if current_cursor < work_end:

            records.append([display_name, day, current_cursor, work_end, "idle", shift_tag, False])



    # --- 第二部分：处理【没打卡】但【有任务】的人 ---

    d1_dates = set(df["start"].dropna().dt.date.unique())

    d2_dates = set(df2["上班时间"].dropna().dt.date.unique())

    all_dates = sorted(list(d1_dates | d2_dates))

   

    for current_day in all_dates:

        # 获取当天所有任务

        day_tasks_all = df[df["start"].dt.date == current_day]

       

        for d_name, group in day_tasks_all.groupby("display_name"):

            # 如果表2有记录，说明已经在第一部分处理过了

            if (d_name, current_day) in processed_names:

                continue

           

            # 针对没打卡的人，遍历其当天所有任务

            for _, task in group.sort_values("start").iterrows():

                t_s, t_e = task["start"], task["end"]

               

                # 1. 班次逻辑：08:00 - 18:00 算早班，其余算晚班

                v_morning_start = pd.Timestamp.combine(current_day, pd.Timestamp("08:00:00").time())

                v_morning_end = pd.Timestamp.combine(current_day, pd.Timestamp("18:00:00").time())

               

                # 如果任务主要落在 8点到18点之间，算 morning

                if t_s >= v_morning_start and t_e <= v_morning_end:

                    shift_tag = "morning"

                else:

                    shift_tag = "evening"



                # 2. 跨天限制逻辑：防止任务无限拉长

                if (t_e - t_s).total_seconds() > 57600: # 超过16小时截断

                    t_e = t_s + pd.Timedelta(hours=16)



                # 3. 强制1分钟显示逻辑

                t_e_disp = t_s + pd.Timedelta(minutes=1) if (t_e - t_s).total_seconds() < 60 else t_e

               

                is_overnight = task["start"].date() != task["end"].date()

               

                records.append([

                    d_name, current_day, t_s, t_e_disp,

                    "overnight" if is_overnight else "work",

                    shift_tag,

                    True

                ])



    tl = pd.DataFrame(records, columns=["name", "date", "start", "end", "type", "shift", "is_absent"])

    tl["duration"] = (tl["end"] - tl["start"]).dt.total_seconds() / 3600

    return tl





# =========================

# UI 组件: Toast 提示

# =========================

class Toast(QLabel):

    def __init__(self, parent, text):

        super().__init__(text, parent)

        self.setStyleSheet("""

            background: rgba(30,30,30,0.85);

            color: white;

            padding: 10px 20px;

            border-radius: 12px;

            font-size: 13px;

        """)

        self.adjustSize()

        self.move(parent.width()//2 - self.width()//2, 50)

        self.show()

        QTimer.singleShot(2000, self.close)



# =========================

# 主界面

# =========================

class MainWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.setWindowTitle("Productivity Dashboard")

        self.resize(1400, 900)

        self.ax_right_m = self.ax_right_e = None

        self.hover_cid_m = self.hover_cid_e = None

        self.timeline = pd.DataFrame()

        self.summary = {}

        self.init_ui()



    def init_ui(self):

        main = QWidget(); self.setCentralWidget(main); layout = QVBoxLayout(main)

        top_card = QFrame(); top_card.setStyleSheet("background: white; border-radius: 16px;"); top_layout = QHBoxLayout(top_card)

        self.btn1 = QPushButton("Upload isc"); self.btn2 = QPushButton("Upload iAMS"); self.gen_btn = QPushButton("Generate")

        self.label1, self.label2, self.date_combo = QLabel("No file"), QLabel("No file"), QComboBox()

        for b in [self.btn1, self.btn2, self.gen_btn]:

            b.setStyleSheet("QPushButton { background:#2563EB; color:white; border-radius:10px; padding:8px 16px; }")

        self.btn1.clicked.connect(self.load_f1); self.btn2.clicked.connect(self.load_f2); self.gen_btn.clicked.connect(self.generate)

        top_layout.addWidget(self.btn1); top_layout.addWidget(self.label1); top_layout.addSpacing(20)

        top_layout.addWidget(self.btn2); top_layout.addWidget(self.label2); top_layout.addSpacing(20)

        top_layout.addWidget(self.date_combo); top_layout.addWidget(self.gen_btn); layout.addWidget(top_card)

        self.tabs = QTabWidget(); self.tab_m, self.tab_e = QWidget(), QWidget()

        self.tabs.addTab(self.tab_m, "Morning"); self.tabs.addTab(self.tab_e, "Evening"); layout.addWidget(self.tabs)

        self.fig_m, self.ax_m, self.can_m, self.ph_m = self.create_plot(self.tab_m)

        self.fig_e, self.ax_e, self.can_e, self.ph_e = self.create_plot(self.tab_e)

        self.date_combo.currentTextChanged.connect(self.refresh)



    def create_plot(self, parent):

        layout = QVBoxLayout(parent); ph = QLabel("📊\nUpload files to generate dashboard"); ph.setAlignment(Qt.AlignmentFlag.AlignCenter)

        ph.setStyleSheet("color: #9CA3AF; font-size: 16px; padding: 40px;"); layout.addWidget(ph)

        fig = Figure(figsize=(10, 6)); ax = fig.add_subplot(111); can = FigureCanvas(fig)

        can.hide(); layout.addWidget(can); return fig, ax, can, ph



    def load_f1(self):

        p, _ = QFileDialog.getOpenFileName(self, "isc", "", "*.xlsx")

        if p: self.file1 = p; self.label1.setText(os.path.basename(p)); Toast(self, "isc file uploaded")

    def load_f2(self):

        p, _ = QFileDialog.getOpenFileName(self, "iAMS", "", "*.xlsx")

        if p: self.file2 = p; self.label2.setText(os.path.basename(p)); Toast(self, "iAMS file uploaded")



    def generate(self):

        if not hasattr(self, 'file1') or not hasattr(self, 'file2'):

            QMessageBox.warning(self, "Warning", "Upload both files first")

            return

        self.gen_btn.setText("Processing..."); self.gen_btn.setEnabled(False)

        Toast(self, "Processing data...")

        QTimer.singleShot(100, self.proc_data)



    def proc_data(self):

        df, df2 = load_data(self.file1, self.file2); self.timeline = build_timeline(df, df2)

        dates = sorted(self.timeline["date"].astype(str).unique())

        self.date_combo.clear(); self.date_combo.addItems(dates)

        self.gen_btn.setText("Generate"); self.gen_btn.setEnabled(True)

        Toast(self, "Dashboard generated"); self.refresh()



    def refresh(self):

        date = self.date_combo.currentText()

        if not date: return

       

        df_day = self.timeline[self.timeline["date"].astype(str) == date].copy()

        self.summary = {}

       

        for name, g in df_day.groupby("name"):

            # 识别异常状态：只要这组数据里有一条 is_absent 为 True，整行就标记异常

            has_absent_issue = g["is_absent"].any()

           

            # 将绿色和黄色都计入工作时间

            work_mask = g["type"].isin(["work", "overnight"])

            work_duration = g[work_mask]["duration"].sum()

            total_duration = g["duration"].sum()

           

            ratio = (work_duration / total_duration * 100) if total_duration > 0 else 0

           

            self.summary[name] = {

                "start": g["start"].min(),

                "end": g["end"].max(),

                "work": work_duration,

                "idle": total_duration - work_duration,

                "total": total_duration,

                "ratio": ratio,

                "is_absent": has_absent_issue

            }

           

        self.ph_m.hide(); self.can_m.show(); self.ph_e.hide(); self.can_e.show()

        self.draw_chart(self.ax_m, self.can_m, df_day, "morning")

        self.draw_chart(self.ax_e, self.can_e, df_day, "evening")



    def draw_chart(self, ax, can, df, shift):
        # 1. 清理右侧辅助轴（如果有），防止重复叠加导致文字模糊
        if shift == "morning" and self.ax_right_m: 
            self.ax_right_m.remove()
            self.ax_right_m = None
        if shift == "evening" and self.ax_right_e: 
            self.ax_right_e.remove()
            self.ax_right_e = None
        
        ax.clear()
        
        # 2. 筛选当前班次（早班/晚班）的数据
        df_s = df[df["shift"] == shift].copy()
        if df_s.empty:
            can.draw()
            return

        # 3. 按效率（Ratio）从低到高排序
        names = sorted(df_s["name"].unique(), key=lambda x: self.summary[x]["ratio"])
        y_map = {p: i for i, p in enumerate(names)}

        # 4. 定义颜色映射
        color_map = {
            "work": THEME["work"],       # 绿色
            "overnight": THEME["overnight"], # 黄色
            "idle": THEME["idle"]        # 灰色
        }

        # 5. 绘制甘特图条形块
        for _, r in df_s.iterrows():
            duration_days = mdates.date2num(r["end"]) - mdates.date2num(r["start"])
            ax.barh(
                y_map[r["name"]], 
                duration_days, 
                left=mdates.date2num(r["start"]), 
                color=color_map.get(r["type"], THEME["idle"]), 
                height=0.6, 
                lw=0
            )

        # 6. 配置右侧效率轴
        ax_r = ax.twinx()
        ax_r.set_ylim(ax.get_ylim())
        ax_r.set_yticks(range(len(names)))
        if shift == "morning": 
            self.ax_right_m = ax_r
        else: 
            self.ax_right_e = ax_r

        # 设置右侧显示的文本标签
        ax_r.set_yticklabels([f"{self.summary[n]['ratio']:.1f}% " for n in names], fontsize=9)
        
        # 7. 配置左侧姓名轴和时间轴格式 (核心修改部分)
        ax.set_yticks(range(len(names)))
        ax.set_yticklabels(names, fontsize=9)
        
        # --- X轴整点优化逻辑 ---
        # 设置 X 轴主刻度为每小时的整点 (例如 08:00, 09:00)
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=1)) 
        # 设置显示格式为 小时:分钟
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        
        # 自动扩展 X 轴范围到最近的整点，防止进度条贴边
        data_min = df_s["start"].min()
        data_max = df_s["end"].max()
        # 向前取整点，向后延一小时取整点
        ax.set_xlim(
            mdates.date2num(data_min.replace(minute=0, second=0)),
            mdates.date2num((data_max + pd.Timedelta(hours=1)).replace(minute=0, second=0))
        )
        # -----------------------
        
        # 8. 隐藏冗余边框
        for a in [ax, ax_r]:
            for s in a.spines.values(): 
                s.set_visible(False)
            a.tick_params(length=0)

        # 9. 更新图例
        legend_elements = [
            Patch(facecolor=THEME["work"], label='Normal Work'),
            Patch(facecolor=THEME["overnight"], label='Overnight (Cross-day)'),
            Patch(facecolor=THEME["idle"], label='Idle'),
            Patch(facecolor='white', edgecolor='none', label='Right side %: Actual Work Ratio')
        ]
        ax.legend(
            handles=legend_elements, 
            loc='lower right', 
            bbox_to_anchor=(1.0, 1.02), 
            ncol=4, 
            frameon=False, 
            fontsize=9
        )

        # 10. 调整布局并刷新画布
        can.figure.subplots_adjust(left=0.22, right=0.88, top=0.9, bottom=0.1)
        can.draw()

        # 11. 重新绑定悬停事件
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

            name = names[y_idx]

            s = self.summary.get(name)

            if s:

                # 顶部姓名

                text = f"<b style='font-size:14px;'>{name}</b><br>"

               

                # 异常提醒：如果没打卡，显示红色警告

                if s.get("is_absent"):

                    text += "<span style='color: #EF4444;'>⚠️ 异常：员工当天无考勤打卡记录</span><br>"

               

                # 时间和时长详情

                text += f"Time: {s['start'].strftime('%H:%M')} - {s['end'].strftime('%H:%M')}<br>"

                text += f"<hr>Total range: {s['total']:.2f}h<br>"

                text += f"Work (Green/Yellow): {s['work']:.2f}h<br>"

                text += f"Idle (Grey): {s['idle']:.2f}h<br>"

                text += f"<b>Efficiency: {s['ratio']:.1f}%</b>"

               

                QToolTip.showText(QCursor.pos(), text)

        else:

            QToolTip.hideText()



if __name__ == "__main__":

    app = QApplication(sys.argv); w = MainWindow(); w.show(); sys.exit(app.exec())
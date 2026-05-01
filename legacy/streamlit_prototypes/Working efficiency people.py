import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Patch

# ===== 读取数据 =====
df = pd.read_excel("pickingResultsOfQuery2026-04-09.xlsx")

df['任务领取时间'] = pd.to_datetime(df['任务领取时间'], errors='coerce')
df['拣货完成时间'] = pd.to_datetime(df['拣货完成时间'], errors='coerce')

df = df.dropna(subset=['姓名', '任务领取时间', '拣货完成时间']).copy()
df['date'] = df['任务领取时间'].dt.date

all_people = sorted(df['姓名'].astype(str).unique())

# ===== 把时间转成小时数 =====
def time_to_hour(dt):
    return dt.hour + dt.minute / 60 + dt.second / 3600

# ===== 构造时间段 =====
def build_timeline(df_day):
    records = []

    if df_day.empty:
        return pd.DataFrame(records)

    df_day = df_day.sort_values(by='任务领取时间').reset_index(drop=True)

    for i in range(len(df_day)):
        start = df_day.loc[i, '任务领取时间']
        end = df_day.loc[i, '拣货完成时间']

        if pd.isna(start) or pd.isna(end):
            continue

        # 工作时间
        if end >= start:
            records.append({
                'start': start,
                'end': end,
                'type': 'work'
            })

        # 空闲时间
        if i < len(df_day) - 1:
            next_start = df_day.loc[i + 1, '任务领取时间']
            if pd.isna(next_start):
                continue

            idle_seconds = (next_start - end).total_seconds()
            if idle_seconds <= 0:
                continue

            # 午休过滤：如果空闲和 12:00-13:00 重叠 >= 20 分钟，就不记 idle
            lunch_start = end.replace(hour=12, minute=0, second=0, microsecond=0)
            lunch_end = end.replace(hour=13, minute=0, second=0, microsecond=0)

            overlap_start = max(end, lunch_start)
            overlap_end = min(next_start, lunch_end)
            overlap_seconds = max((overlap_end - overlap_start).total_seconds(), 0)

            if overlap_seconds >= 20 * 60:
                continue

            records.append({
                'start': end,
                'end': next_start,
                'type': 'idle'
            })

    return pd.DataFrame(records)

def calculate_idle_hours(timeline_df):
    total = 0
    for _, row in timeline_df.iterrows():
        if row['type'] == 'idle':
            total += (row['end'] - row['start']).total_seconds()
    return total / 3600  # 转小时

# ===== UI =====
root = tk.Tk()
root.title("Efficiency Dashboard")
root.geometry("1300x800")

selected_person = tk.StringVar()
dropdown = ttk.Combobox(root, textvariable=selected_person, state="readonly", width=30)
dropdown['values'] = all_people
dropdown.current(0)
dropdown.pack(pady=10)

fig, ax = plt.subplots(figsize=(13, 8))
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

# ===== 绘图 =====
def update_plot(event=None):
    person = selected_person.get()
    df_person = df[df['姓名'].astype(str) == person].copy()

    ax.clear()

    if df_person.empty:
        ax.set_title(f"{person} - No Data")
        canvas.draw()
        return

    dates = sorted(df_person['date'].unique())

    xmin = 24
    xmax = 0

    idle_results = []  # 存每天 idle

    for i, date in enumerate(dates):
        df_day = df_person[df_person['date'] == date].copy()
        timeline_df = build_timeline(df_day)

        if timeline_df.empty:
            idle_results.append(0)
            continue

        # ===== 计算 idle 时间 =====
        idle_hours = calculate_idle_hours(timeline_df)
        idle_results.append(idle_hours)

        for _, row in timeline_df.iterrows():
            start_hour = time_to_hour(row['start'])
            end_hour = time_to_hour(row['end'])
            width = end_hour - start_hour

            if width <= 0:
                continue

            xmin = min(xmin, start_hour)
            xmax = max(xmax, end_hour)

            color = '#2ca02c' if row['type'] == 'work' else '#ff7f0e'

            ax.barh(i, width, left=start_hour, height=0.5, color=color)

        # ===== 右侧显示 idle =====
    text_x = xmax + 0.3

    for i, idle_h in enumerate(idle_results):
        ax.text(
            text_x,
            i,
            f"{idle_h:.2f} h",
            va='center',
            fontsize=10,
            color='red'
        )

    # y轴
    ax.set_yticks(range(len(dates)))
    ax.set_yticklabels([str(d) for d in dates])

    # x轴：只显示一天内时间
    if xmin < xmax:
        ax.set_xlim(max(0, xmin - 0.2), min(24, xmax + 0.5))
    else:
        ax.set_xlim(0, 24)

   # ===== 动态时间范围 =====
    if xmin < xmax:
        start_hour = max(0, int(xmin))
        end_hour = min(24, int(xmax) + 1)
    else:
        start_hour = 0
        end_hour = 24

    ax.set_xlim(xmin - 0.1, xmax + 0.1)

    # ===== 动态刻度 =====
    tick_hours = list(range(start_hour, end_hour + 1))
    ax.set_xticks(tick_hours)
    ax.set_xticklabels([f"{h:02d}:00" for h in tick_hours])

    ax.set_xlabel("Time of Day")
    ax.set_title(f"{person}")
    ax.grid(axis='x', linestyle='--', alpha=0.3)

    # 图例
    legend_elements = [
        Patch(facecolor='#2ca02c', label='Work'),
        Patch(facecolor='#ff7f0e', label='Idle')
    ]
    ax.legend(handles=legend_elements, loc='upper right')

    fig.tight_layout()
    canvas.draw()

dropdown.bind("<<ComboboxSelected>>", update_plot)

update_plot()
root.mainloop()
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Patch

# ===== 全局变量 =====
df = pd.DataFrame()
all_dates = []

TASK_COL = '集合单号'

# ===== 时间转小时 =====
def time_to_hour(dt):
    return dt.hour + dt.minute / 60 + dt.second / 3600


# ===== 构造 timeline =====
def build_timeline(df_person):
    records = []

    if df_person.empty:
        return pd.DataFrame(records)

    task_groups = []

    for task_id, group in df_person.groupby(TASK_COL):

        group = group.sort_values(by='拣货完成时间')
        task_receive = group['任务领取时间'].iloc[0]

        if len(group) == 1:
            task_end_pick = group['拣货完成时间'].iloc[0]
            task_start_pick = max(
                task_receive,
                task_end_pick - pd.Timedelta(minutes=2)
            )
            is_burst = False

        else:
            time_counts = group['拣货完成时间'].value_counts()
            max_count = time_counts.iloc[0]
            ratio = max_count / len(group)

            if ratio >= 0.8:
                task_end_pick = time_counts.index[0]
                minutes = len(group)
                task_start_pick = task_end_pick - pd.Timedelta(minutes=minutes)
                is_burst = True
            else:
                task_start_pick = group['拣货完成时间'].min()
                task_end_pick = group['拣货完成时间'].max()
                is_burst = False

        task_groups.append({
            'receive': task_receive,
            'start_pick': task_start_pick,
            'end_pick': task_end_pick,
            'is_burst': is_burst
        })

    task_groups = sorted(task_groups, key=lambda x: x['start_pick'])

    lunch_used = False

    for i in range(len(task_groups)):

        task = task_groups[i]

        if i == 0:
            start = min(task['receive'], task['start_pick'])
        else:
            start = task['start_pick']

        end = task['end_pick']

        records.append({
            'start': start,
            'end': end,
            'type': 'burst' if task['is_burst'] else 'work'
        })

        if i < len(task_groups) - 1:
            next_task = task_groups[i + 1]

            idle_start = end
            idle_end = next_task['start_pick']

            if idle_end > idle_start:

                lunch_start = idle_start.replace(hour=12, minute=0, second=0)
                lunch_end = idle_start.replace(hour=13, minute=30, second=0)

                valid_start = max(idle_start, lunch_start)
                valid_end = min(idle_end, lunch_end)

                valid_seconds = (valid_end - valid_start).total_seconds()

                if valid_seconds >= 1800 and not lunch_used:

                    lunch_real_start = valid_start
                    lunch_real_end = lunch_real_start + pd.Timedelta(minutes=30)

                    if lunch_real_start > idle_start:
                        records.append({'start': idle_start, 'end': lunch_real_start, 'type': 'idle'})

                    records.append({'start': lunch_real_start, 'end': lunch_real_end, 'type': 'lunch'})

                    lunch_used = True

                    if idle_end > lunch_real_end:
                        records.append({'start': lunch_real_end, 'end': idle_end, 'type': 'idle'})
                else:
                    records.append({'start': idle_start, 'end': idle_end, 'type': 'idle'})

    return pd.DataFrame(records)


def calculate_idle_hours(timeline_df):
    total = 0
    for _, row in timeline_df.iterrows():
        if row['type'] == 'idle':
            total += (row['end'] - row['start']).total_seconds()
    return total / 3600


# ===== UI =====
root = tk.Tk()
root.title("Efficiency Dashboard")
root.geometry("1300x800")

selected_date = tk.StringVar()

dropdown = ttk.Combobox(root, textvariable=selected_date, state="readonly", width=30)
dropdown.pack(pady=10)

# ===== 上传按钮 =====
def load_file():
    global df, all_dates

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    df = pd.read_excel(file_path)

    df['任务领取时间'] = pd.to_datetime(df['任务领取时间'], errors='coerce')
    df['拣货完成时间'] = pd.to_datetime(df['拣货完成时间'], errors='coerce')

    df.dropna(subset=['姓名', '任务领取时间', '拣货完成时间'], inplace=True)

    df['date'] = df['任务领取时间'].dt.date
    all_dates = sorted(df['date'].astype(str).unique())

    dropdown['values'] = all_dates
    dropdown.current(0)

    update_plot()


upload_btn = tk.Button(root, text="上传Excel", command=load_file)
upload_btn.pack(pady=10)


# ===== 图 =====
fig, ax = plt.subplots(figsize=(13, 8))
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)


# ===== 画图 =====
def update_plot(event=None):
    if df.empty:
        return

    date = selected_date.get()
    df_day = df[df['date'].astype(str) == date]

    ax.clear()

    people_raw = sorted(df_day['姓名'].astype(str).unique())

    # ===== 先计算每个人 idle =====
    people_idle_map = {}

    for person in people_raw:
        df_person = df_day[df_day['姓名'] == person]
        timeline_df = build_timeline(df_person)
        idle_hours = calculate_idle_hours(timeline_df)
        people_idle_map[person] = idle_hours

    # ===== 按 idle 排序（从少到多）=====
    people = sorted(people_raw, key=lambda x: people_idle_map[x])

    xmin, xmax = 24, 0
    idle_results = []

    for i, person in enumerate(people):
        df_person = df_day[df_day['姓名'] == person]
        timeline_df = build_timeline(df_person)

        idle_h = calculate_idle_hours(timeline_df)
        idle_results.append(idle_h)

        for _, row in timeline_df.iterrows():
            start = time_to_hour(row['start'])
            end = time_to_hour(row['end'])

            width = end - start
            if width <= 0:
                continue

            xmin = min(xmin, start)
            xmax = max(xmax, end)

            color = {
                'work': '#2ca02c',
                'burst': '#1f77b4',
                'idle': '#ff7f0e',
                'lunch': '#ffffff'
            }[row['type']]

            ax.barh(i, width, left=start, height=0.5, color=color)

    for i, idle_h in enumerate(idle_results):
        ax.text(xmax + 0.3, i, f"{idle_h:.2f} h", va='center', color='red')

    ax.set_yticks(range(len(people)))
    ax.set_yticklabels(people)

    ax.set_xlim(max(0, xmin - 0.2), min(24, xmax + 0.5))
    ax.set_xlabel("Time")

    ax.legend(
        handles=[
            Patch(color='#2ca02c', label='Work'),
            Patch(color='#1f77b4', label='Burst'),
            Patch(color='#ff7f0e', label='Idle'),
            Patch(facecolor='#ffffff', edgecolor='black', label='Lunch')
        ],
        loc='upper left',
        bbox_to_anchor=(1.01, 1.08)
    )

    canvas.draw()


dropdown.bind("<<ComboboxSelected>>", update_plot)

root.mainloop()
import re
import warnings
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================
# Theme
# =========================
THEME = {
    "bg": "#F6F8FC",
    "surface": "#FFFFFF",
    "surface_alt": "#F8FAFC",
    "border": "#E5E7EB",
    "primary": "#2563EB",
    "text": "#0F172A",
    "muted": "#64748B",
    "success": "#16A34A",
    "danger": "#DC2626",
    "work": "#7EC398",
    "overnight": "#8B7ABD",
    "idle": "#C6D7EB",
    "break": "#D9E6AF",
}

TYPE_COLOR = {
    "work": THEME["work"],
    "overnight": THEME["overnight"],
    "idle": THEME["idle"],
    "break": THEME["break"],
}


# =========================
# Tools
# =========================
def classify_file(file):
    try:
        df = pd.read_excel(file, nrows=5)
        cols = [str(c).strip() for c in df.columns]

        # ===== df1：任务 =====
        if "任务单开始时间" in cols and "任务单结束时间" in cols:
            return "task"

        # ===== df2：班次 =====
        if "实际上班时间" in cols and "实际下班时间" in cols:
            return "attendance"

        # ===== df3：件量 =====
        if "打包完成时间" in cols and "件数" in cols:
            return "volume"

        # ===== df4：打卡 =====
        if "打卡时间" in cols and "进出仓标识" in "".join(cols):
            return "punch"

        return "unknown"

    except:
        return "unknown"
def adjust_work_date(dt):
    if pd.isna(dt):
        return None
    if dt.hour < 6:
        return (dt - pd.Timedelta(days=1)).date()
    return dt.date()


def extract_warehouse(name):
    if pd.isna(name):
        return "Unknown"
    match = re.search(r"\d+号仓", str(name))
    return match.group() if match else str(name)


def safe_time(v):
    if pd.isna(v):
        return "缺失"
    return pd.to_datetime(v).strftime("%H:%M:%S")


# =========================
# Data Logic
# =========================
def load_data(file1, file2):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    df2["employee_no"] = df2.iloc[:, 1].astype(str).str.strip()
    df2["班次名称"] = df2.iloc[:, 7].astype(str).str.strip()
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
            "任务单开始时间": "start",
            "任务单结束时间": "end",
        }
    ).copy()

    if "warehouse_name" in df1.columns:
        df["warehouse_name"] = df1["warehouse_name"]
    else:
        df["warehouse_name"] = "Unknown"

    if "employee_no" not in df.columns:
        df["employee_no"] = pd.NA

    if "operator" not in df.columns:
        df["operator"] = "Unknown"

    df["employee_no"] = (
        df["employee_no"].astype(str).str.strip().replace(["nan", "None", ""], pd.NA)
    )
    df["start"] = pd.to_datetime(df["start"], errors="coerce")
    df["end"] = pd.to_datetime(df["end"], errors="coerce")

    def fix_isc_datetime(row):
        start = row["start"]
        end = row["end"]

        if pd.isna(start) or pd.isna(end):
            return start, end

        if start.hour < 6:
            start = start + pd.Timedelta(days=1)
            end = end + pd.Timedelta(days=1)

        return start, end

    df[["start", "end"]] = df.apply(lambda r: pd.Series(fix_isc_datetime(r)), axis=1)

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


def load_volume_data(file3):
    if file3 is None:
        return {}

    try:
        df = pd.read_excel(file3)
        df.columns = df.columns.str.strip()

        df["打包完成时间"] = pd.to_datetime(df["打包完成时间"], errors="coerce")

        if "是否取消" in df.columns:
            df = df[~df["是否取消"].astype(str).str.contains("是", na=False)]

        df["件数"] = pd.to_numeric(df.iloc[:, 6], errors="coerce").fillna(0)
        df["日期"] = df["打包完成时间"].dt.date

        return df.groupby("日期")["件数"].sum().to_dict()

    except Exception as e:
        st.warning(f"Volume读取失败: {e}")
        return {}


def load_break_data(file4):
    if file4 is None:
        return pd.DataFrame()

    try:
        df4 = pd.read_excel(file4)

        df = pd.DataFrame(
            {
                "emp_no": df4.iloc[:, 3].astype(str).str.strip(),
                "time": pd.to_datetime(df4.iloc[:, 12], errors="coerce"),
                "action": df4.iloc[:, 18].astype(str),
                "group": df4.iloc[:, 2].astype(str).str.strip(),
            }
        ).dropna(subset=["time"])

        df = df.sort_values(["emp_no", "time"])

        breaks = []

        for emp, group in df.groupby("emp_no"):
            group = group.sort_values("time")

            out_time = None

            for _, row in group.iterrows():
                action = str(row["action"])
                t = row["time"]

                # 出仓 = break开始
                if "出仓" in action:
                    out_time = t

                # 进仓 = break结束
                elif "进仓" in action and out_time is not None:

                    if t > out_time:

                        # ✅ 判断整个区间是否属于午休
                        lunch_start = out_time.replace(hour=11, minute=0, second=0)
                        lunch_end   = out_time.replace(hour=14, minute=0, second=0)

                        # 👉 只要有交集就算午休
                        if not (t < lunch_start or out_time > lunch_end):

                            breaks.append({
                                "emp_no": emp,
                                "start": out_time,
                                "end": t,
                                "work_date": adjust_work_date(out_time),
                                "type": "break",
                                "group": row["group"],
                            })

                    out_time = None

        return pd.DataFrame(breaks)

    except Exception as e:
        st.warning(f"Punch读取失败: {e}")
        return pd.DataFrame()


def load_indirect_data(file1):
    if file1 is None:
        return pd.DataFrame()

    try:
        df = pd.read_excel(file1, sheet_name="间接工时明细-间接工时")
        df.columns = df.columns.str.strip()

        df["employee_no"] = df["employee_no"].astype(str).str.strip()
        df["开始时间"] = pd.to_datetime(df["开始时间"], errors="coerce")
        df["结束时间"] = pd.to_datetime(df["结束时间"], errors="coerce")
        df["操作分类"] = df["操作分类"].astype(str).str.strip().replace("nan", "未分类")
        df["间接工时(分)"] = pd.to_numeric(df["间接工时(分)"], errors="coerce").fillna(0)

        df = df.dropna(subset=["开始时间", "结束时间"])
        df = df[df["结束时间"] > df["开始时间"]]

        df["work_date"] = df["开始时间"].apply(adjust_work_date)
        df["display_name"] = df["employee_no"]
        df["duration"] = (df["结束时间"] - df["开始时间"]).dt.total_seconds() / 3600

        return df

    except Exception:
        return pd.DataFrame()


def find_no_operation_people(df_isc, df_iams):
    iams_valid = df_iams[
        df_iams["上班时间"].notna() & df_iams["下班时间"].notna()
    ].copy()

    iams_valid["employee_no"] = iams_valid["employee_no"].astype(str).str.strip()

    isc_valid = df_isc[df_isc["employee_no"].notna()].copy()
    isc_valid["employee_no"] = isc_valid["employee_no"].astype(str).str.strip()

    isc_emp_set = set(isc_valid["employee_no"])

    no_op = iams_valid[~iams_valid["employee_no"].isin(isc_emp_set)].copy()

    return no_op[["employee_no", "real_name", "上班时间", "下班时间"]]


def build_timeline(df, df2, group_mapping, df_breaks=None):
    records = []
    processed = set()

    def get_isc_time(emp_no, day):
        emp_df = df[
            (df["employee_no"].astype(str) == emp_no) &
            (df["work_date"] == day)
        ].copy()

        if emp_df.empty:
            return None, None

        start = pd.to_datetime(emp_df["上班时间"].min(), errors="coerce") if "上班时间" in emp_df.columns else None
        end = pd.to_datetime(emp_df["下班时间"].max(), errors="coerce") if "下班时间" in emp_df.columns else None
        return start, end

    for _, row in df2.iterrows():
        emp_no = str(row["employee_no"])
        name = str(row["real_name"])
        display = f"{name} ({emp_no})"
        group_name = str(row["attendance_group"])

        work_start = row["上班时间"]
        work_end = row["下班时间"]

        day = pd.to_datetime(work_start).date() if not pd.isna(work_start) else None

        isc_start, isc_end = get_isc_time(emp_no, day)

        if pd.isna(work_start) and isc_start is not None:
            work_start = isc_start

        if pd.isna(work_end) and isc_end is not None:
            work_end = isc_end

        if pd.isna(work_start) or pd.isna(work_end):
            continue

        day = work_start.date()
        processed.add((emp_no, day))

        # ✅ 优先用班次名称判断
        shift_name = str(row.get("班次名称", ""))

        if any(x in shift_name for x in ["早", "白"]):
            shift = "morning"
        elif "晚" in shift_name:
            shift = "evening"
        else:
            shift = "morning" if work_start.hour < 15 else "evening"

        tasks = df[
            (df["employee_no"].astype(str) == emp_no) &
            (df["work_date"] == day)
        ].copy()

        tasks["start"] = pd.to_datetime(tasks["start"], errors="coerce")
        tasks["end"] = pd.to_datetime(tasks["end"], errors="coerce")
        tasks["type"] = "work"

        tasks = tasks[tasks["start"] < work_end].copy()
        tasks["start"] = tasks["start"].clip(lower=work_start)

        if df_breaks is not None and not df_breaks.empty:
            p_breaks = df_breaks[
                (df_breaks["emp_no"].astype(str).str.strip() == emp_no) &
                (df_breaks["work_date"] == day) 
            ].copy()

            p_breaks["start"] = pd.to_datetime(p_breaks["start"], errors="coerce")
            p_breaks["end"] = pd.to_datetime(p_breaks["end"], errors="coerce")
            p_breaks["type"] = "break"

            p_breaks = p_breaks[
                (p_breaks["end"] > work_start) &
                (p_breaks["start"] < work_end)
            ].copy()

            p_breaks["start"] = p_breaks["start"].clip(lower=work_start, upper=work_end)
            p_breaks["end"] = p_breaks["end"].clip(lower=work_start, upper=work_end)
            p_breaks["employee_no"] = emp_no

            tasks = pd.concat(
                [tasks, p_breaks[["start", "end", "type", "employee_no"]]],
                ignore_index=True
            )

        tasks = tasks.sort_values("start")
        cursor = work_start

        for _, t in tasks.iterrows():
            if "employee_no" in t and str(t["employee_no"]) != emp_no:
                continue

            s = max(t["start"], work_start)
            e = min(t["end"], work_end)

            if s >= e:
                continue

            if s > cursor:
                records.append([display, day, cursor, s, "idle", shift, False, group_name])

            records.append([display, day, s, e, t["type"], shift, False, group_name])
            cursor = e

        if pd.notna(cursor) and cursor < work_end:
            records.append([display, day, cursor, work_end, "idle", shift, False, group_name])

    for day in sorted(df["work_date"].dropna().unique()):
        day_df = df[df["work_date"] == day]

        for emp_no, group in day_df.groupby("employee_no"):
            if (emp_no, day) in processed:
                continue

            emp_no = str(emp_no)
            display = group["display_name"].iloc[0]
            group_name = group_mapping.get(emp_no, "Unknown")

            work_start = pd.to_datetime(group["上班时间"].min(), errors="coerce") if "上班时间" in group.columns else pd.NaT
            work_end = pd.to_datetime(group["下班时间"].max(), errors="coerce") if "下班时间" in group.columns else pd.NaT

            task_start = pd.to_datetime(group["start"].min(), errors="coerce")
            task_end = pd.to_datetime(group["end"].max(), errors="coerce")

            if pd.isna(work_start):
                work_start = task_start

            if pd.isna(work_end):
                work_end = task_end

            if pd.isna(work_start) or pd.isna(work_end):
                continue

            shift_name = str(group.iloc[0].get("班次名称", ""))

            if any(x in shift_name for x in ["早", "白"]):
                shift = "morning"
            elif "晚" in shift_name:
                shift = "evening"
            else:
                shift = "morning" if work_start.hour < 15 else "evening"

            group = group.copy()
            group["employee_no"] = emp_no
            group["start"] = pd.to_datetime(group["start"], errors="coerce")
            group["end"] = pd.to_datetime(group["end"], errors="coerce")
            group["type"] = "work"

            group = group[group["start"] < work_end].copy()
            group["start"] = group["start"].clip(lower=work_start)

            if df_breaks is not None and not df_breaks.empty:
                p_breaks = df_breaks[
                    (df_breaks["emp_no"].astype(str).str.strip() == emp_no) &
                    (df_breaks["work_date"] == day) &
                    (df_breaks["group"].astype(str).str.strip() == group_name)
                ].copy()

                p_breaks["start"] = pd.to_datetime(p_breaks["start"], errors="coerce")
                p_breaks["end"] = pd.to_datetime(p_breaks["end"], errors="coerce")
                p_breaks["type"] = "break"

                p_breaks = p_breaks[
                    (p_breaks["end"] > work_start) &
                    (p_breaks["start"] < work_end)
                ].copy()

                p_breaks["start"] = p_breaks["start"].clip(lower=work_start, upper=work_end)
                p_breaks["end"] = p_breaks["end"].clip(lower=work_start, upper=work_end)
                p_breaks["employee_no"] = emp_no

                group = pd.concat(
                    [
                        group[["start", "end", "type"]],
                        p_breaks[["start", "end", "type", "employee_no"]],
                    ],
                    ignore_index=True
                )
            else:
                group = group[["start", "end", "type"]]

            group = group.sort_values("start")
            cursor = work_start

            for _, t in group.iterrows():
                s = max(t["start"], work_start)
                e = min(t["end"], work_end)

                if s >= e:
                    continue

                if s > cursor:
                    records.append([display, day, cursor, s, "idle", shift, True, group_name])

                if pd.notna(t["end"]) and t["end"].date() > work_end.date():
                    t_type = "overnight"
                else:
                    t_type = t["type"]

                records.append([display, day, s, e, t_type, shift, True, group_name])
                cursor = e

            if cursor < work_end:
                records.append([display, day, cursor, work_end, "idle", shift, True, group_name])

    timeline = pd.DataFrame(
        records,
        columns=["name", "date", "start", "end", "type", "shift", "is_absent", "group"],
    )

    if timeline.empty:
        return timeline

    timeline["duration"] = (
        timeline["end"] - timeline["start"]
    ).dt.total_seconds() / 3600

    return timeline

def remove_break_overlap(timeline):
    if timeline.empty:
        return timeline

    result = []

    for (name, date), group in timeline.groupby(["name", "date"]):
        breaks = group[group["type"] == "break"].copy()
        others = group[group["type"] != "break"].copy()

        # 先处理非 break
        for _, row in others.iterrows():
            segments = [(row["start"], row["end"])]

            for _, br in breaks.iterrows():
                new_segments = []

                for s, e in segments:
                    br_s = br["start"]
                    br_e = br["end"]

                    # 没有重叠
                    if e <= br_s or s >= br_e:
                        new_segments.append((s, e))
                    else:
                        # break 前面的部分保留
                        if s < br_s:
                            new_segments.append((s, br_s))

                        # break 后面的部分保留
                        if e > br_e:
                            new_segments.append((br_e, e))

                segments = new_segments

            for s, e in segments:
                if s < e:
                    new_row = row.copy()
                    new_row["start"] = s
                    new_row["end"] = e
                    new_row["duration"] = (e - s).total_seconds() / 3600
                    result.append(new_row)

        # break 本身保留
        for _, br in breaks.iterrows():
            result.append(br)

    out = pd.DataFrame(result)

    if not out.empty:
        out = out.sort_values(["name", "start", "type"])
        out["duration"] = (
            out["end"] - out["start"]
        ).dt.total_seconds() / 3600

    return out


# =========================
# Summary Logic
# =========================
def build_history(timeline):
    history_ratios = {}
    history_wait = {}

    if timeline.empty:
        return history_ratios, history_wait

    for d, day_data in timeline.groupby("date"):
        d_str = pd.to_datetime(d).strftime("%Y-%m-%d")
        history_ratios[d_str] = {}
        history_wait[d_str] = {}

        for name, group in day_data.groupby("name"):
            work_duration = group[group["type"].isin(["work", "overnight"])]["duration"].sum()
            total_duration = group["duration"].sum()

            history_ratios[d_str][name] = (
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

                history_wait[d_str][emp_no] = wait_seconds

    return history_ratios, history_wait


def build_summary(df_day, df, df2, current_day):
    summary = {}

    for name, group in df_day.groupby("name"):
        work_duration = group[group["type"].isin(["work", "overnight"])]["duration"].sum()
        break_duration = group[group["type"] == "break"]["duration"].sum()
        total_duration = group["duration"].sum()
        effective_duration = total_duration - break_duration

        ratio = (work_duration / effective_duration * 100) if effective_duration > 0 else 0

        if "(" in name:
            emp_no = name.split("(")[-1].replace(")", "").strip()
        else:
            emp_no = name

        in_time = None
        out_time = None

        row_iams = df2[
            (df2["employee_no"].astype(str) == emp_no) &
            (df2["上班时间"].apply(adjust_work_date) == current_day)
        ]

        if not row_iams.empty:
            in_time = row_iams["上班时间"].iloc[0]
            out_time = row_iams["下班时间"].iloc[0]

        if pd.isna(in_time) or pd.isna(out_time):
            row_isc = df[
                (df["employee_no"].astype(str) == emp_no) &
                (df["work_date"] == current_day)
            ]

            if not row_isc.empty:
                if "上班时间" in row_isc.columns:
                    isc_in = pd.to_datetime(row_isc["上班时间"].min(), errors="coerce")
                else:
                    isc_in = pd.NaT

                if "下班时间" in row_isc.columns:
                    isc_out = pd.to_datetime(row_isc["下班时间"].max(), errors="coerce")
                else:
                    isc_out = pd.NaT

                if pd.isna(in_time):
                    in_time = isc_in
                if pd.isna(out_time):
                    out_time = isc_out

        summary[name] = {
            "ratio": ratio,
            "in": safe_time(in_time),
            "out": safe_time(out_time),
        }

    return summary


def calc_total_manhours(df_day, df, df2, current_day):
    total = 0

    for name, group in df_day.groupby("name"):
        if "(" in name:
            emp_no = name.split("(")[-1].replace(")", "").strip()
        else:
            emp_no = name.strip()

        punch_in = None
        punch_out = None

        row_iams = df2[
            (df2["employee_no"].astype(str) == str(emp_no)) &
            (df2["上班时间"].apply(adjust_work_date) == current_day)
        ]

        if not row_iams.empty:
            punch_in = pd.to_datetime(row_iams["上班时间"].iloc[0], errors="coerce")
            punch_out = pd.to_datetime(row_iams["下班时间"].iloc[0], errors="coerce")

        if punch_in is None or pd.isna(punch_in) or pd.isna(punch_out):
            row_isc = df[
                (df["employee_no"].astype(str) == str(emp_no)) &
                (df["work_date"] == current_day)
            ]

            if not row_isc.empty:
                if punch_in is None or pd.isna(punch_in):
                    punch_in = pd.to_datetime(row_isc["上班时间"].min(), errors="coerce") if "上班时间" in row_isc.columns else pd.NaT

                if punch_out is None or pd.isna(punch_out):
                    punch_out = pd.to_datetime(row_isc["下班时间"].max(), errors="coerce") if "下班时间" in row_isc.columns else pd.NaT

        if pd.isna(punch_in):
            punch_in = group["start"].min()

        if pd.isna(punch_out):
            punch_out = group["end"].max()

        if pd.notna(punch_in) and pd.notna(punch_out) and punch_out > punch_in:
            total += (punch_out - punch_in).total_seconds() / 3600

    return total


# =========================
# Charts
# =========================
def draw_gantt(df_shift, selected_date, shift, summary, history_ratios, df_day, df2):
    if df_shift.empty:
        return None

    current_date = pd.to_datetime(selected_date).date()
    prev_date = current_date - pd.Timedelta(days=1)
    prev_date_str = prev_date.strftime("%Y-%m-%d")

    names = sorted(
        df_shift["name"].unique(),
        key=lambda n: summary.get(n, {}).get("ratio", 0),
    )

    df_plot = df_shift.copy()
    df_plot["start"] = pd.to_datetime(df_plot["start"])
    df_plot["end"] = pd.to_datetime(df_plot["end"])

    df_plot["efficiency"] = df_plot["name"].map(lambda n: f"{summary.get(n, {}).get('ratio', 0):.1f}%")
    df_plot["in_time"] = df_plot["name"].map(lambda n: summary.get(n, {}).get("in", "--"))
    df_plot["out_time"] = df_plot["name"].map(lambda n: summary.get(n, {}).get("out", "--"))
    # 👉 每个人的 total work
    def calc_unique_work_hours(group):
        work = group[group["type"].isin(["work", "overnight"])].copy()

        if work.empty:
            return 0

        # 按时间排序
        work = work.sort_values("start")

        merged = []
        current_start = work.iloc[0]["start"]
        current_end = work.iloc[0]["end"]

        for _, row in work.iloc[1:].iterrows():
            s = row["start"]
            e = row["end"]

            if s <= current_end:
                current_end = max(current_end, e)
            else:
                merged.append((current_start, current_end))
                current_start = s
                current_end = e

        merged.append((current_start, current_end))

        # 计算总时长（去重）
        total = sum((e - s).total_seconds() for s, e in merged) / 3600

        return total


    work_map = df_day.groupby("name").apply(calc_unique_work_hours).to_dict()

    # 👉 每个人的 attendance
    attendance_map = {}

    for name, group in df_day.groupby("name"):

        if "(" in name:
            emp_no = name.split("(")[-1].replace(")", "").strip()
        else:
            emp_no = name

        # ✅ 优先用 iAMS
        row_iams = df2[
            (df2["employee_no"].astype(str) == emp_no) &
            (df2["上班时间"].apply(adjust_work_date) == pd.to_datetime(selected_date).date())
        ]

        if not row_iams.empty:
            start = row_iams["上班时间"].iloc[0]
            end = row_iams["下班时间"].iloc[0]
        else:
            # fallback
            start = group["start"].min()
            end = group["end"].max()

        if pd.notna(start) and pd.notna(end) and end > start:
            attendance_map[name] = (end - start).total_seconds() / 3600
        else:
            attendance_map[name] = 0
    df_plot["work_hours"] = df_plot["name"].map(work_map)
    df_plot["attendance_hours"] = df_plot["name"].map(attendance_map)
    df_plot["task_start_str"] = df_plot["start"].dt.strftime("%H:%M:%S")
    df_plot["task_end_str"] = df_plot["end"].dt.strftime("%H:%M:%S")
    df_plot["type_order"] = df_plot["type"].map({
        "work": 0,
        "idle": 1,
        "overnight": 2,
        "break": 3
    })
    df_plot = df_plot.sort_values("type_order")
    # ✅ 分层画：先画非 break，再画 break
    df_base = df_plot[df_plot["type"] != "break"].copy()
    df_break = df_plot[df_plot["type"] == "break"].copy()

    fig = px.timeline(
        df_base,
        x_start="start",
        x_end="end",
        y="name",
        color="type",
        color_discrete_map=TYPE_COLOR,
        category_orders={"name": names},
        custom_data=[
            "type",
            "name",
            "task_start_str",
            "task_end_str",
            "duration",
            "work_hours",
            "attendance_hours",
            "efficiency"
        ]
    )
    fig.update_traces(
        hovertemplate=
        "type: %{customdata[0]}<br>" +
        "name: %{customdata[1]}<br>" +
        "task start: %{customdata[2]}<br>" +
        "task end: %{customdata[3]}<br>" +
        "task duration: %{customdata[4]:.2f}h<br>" +
        "total work: %{customdata[5]:.2f}h<br>" +
        "attendance: %{customdata[6]:.2f}h<br>" +
        "efficiency: %{customdata[7]}<br>" +
        "<extra></extra>"
    )

    if not df_break.empty:
        fig_break = px.timeline(
            df_break,
            x_start="start",
            x_end="end",
            y="name",
            color="type",
            color_discrete_map=TYPE_COLOR,
            category_orders={"name": names},
            custom_data=[
                "type",
                "name",
                "task_start_str",
                "task_end_str",
                "duration",
                "work_hours",
                "attendance_hours",
                "efficiency"
            ]
        )
        fig_break.update_traces(
            hovertemplate=
            "type: %{customdata[0]}<br>" +
            "name: %{customdata[1]}<br>" +
            "task start: %{customdata[2]}<br>" +
            "task end: %{customdata[3]}<br>" +
            "task duration: %{customdata[4]:.2f}h<br>" +
            "total work: %{customdata[5]:.2f}h<br>" +
            "attendance: %{customdata[6]:.2f}h<br>" +
            "efficiency: %{customdata[7]}<br>" +
            "<extra></extra>"
        )

        # ✅ break trace 最后加入，hover 优先
        for trace in fig_break.data:
            trace.name = "break"
            trace.legendgroup = "break"
            trace.showlegend = True
            fig.add_trace(trace)

    fig.update_yaxes(autorange="reversed")

    if shift == "morning":
        fig.update_xaxes(
            range=[
                pd.to_datetime(f"{selected_date} 06:00"),
                pd.to_datetime(f"{selected_date} 18:00"),
            ]
        )

    fig.update_layout(
        height=max(450, len(names) * 32),
        plot_bgcolor="white",
        paper_bgcolor="white",

        # ✅ 右边留更多空间
        margin=dict(l=25, r=160, t=30, b=30),

        xaxis_title="Time",
        yaxis_title="Employee",

        # ✅ 图例位置
        legend=dict(
            x=1.09,          # 👉 往右移
            y=1,
            xanchor="left",
            yanchor="top"
        )
    )

    annotations = []
    for i, name in enumerate(names):
        ratio = summary.get(name, {}).get("ratio", 0)
        prev_ratio = history_ratios.get(prev_date_str, {}).get(name)

        arrow = ""
        color = THEME["text"]

        if prev_ratio is not None:
            if ratio > prev_ratio:
                arrow = " ↑"
                color = THEME["success"]
            elif ratio < prev_ratio:
                arrow = " ↓"
                color = THEME["danger"]

        annotations.append(
            dict(
                x=1.01,
                y=name,
                xref="paper",
                yref="y",
                text=f"{ratio:.1f}%{arrow}",
                showarrow=False,
                font=dict(size=11, color=color),
                xanchor="left",
            )
        )

    fig.update_layout(annotations=annotations)

    return fig


def draw_efficiency_donut(ratio, total_attendance, total_work):
    idle = max(0, total_attendance - total_work)

    fig = go.Figure(
        data=[
            go.Pie(
                labels=["Work", "Idle"],
                values=[total_work, idle],  # 用时间，不用百分比
                hole=0.65,
                marker=dict(colors=[THEME["primary"], "#D7E3FF"]),
                textinfo="none",

                # ✅ 核心：hover内容
                hovertemplate=
                "%{label}: %{percent:.1%} (%{value:.1f}h)<extra></extra>"
            )
        ]
    )

    fig.update_layout(
        height=180,
        margin=dict(l=5, r=5, t=5, b=5),
        annotations=[
            dict(
                text=f"{ratio:.1f}%",
                x=0.5,
                y=0.5,
                showarrow=False,
                font=dict(size=22, color=THEME["primary"]),
            )
        ],
        showlegend=False,
    )

    return fig


def draw_indirect_chart(df_day):
    if df_day.empty:
        return None

    df_plot = df_day.copy()
    df_plot["开始时间"] = pd.to_datetime(df_plot["开始时间"])
    df_plot["结束时间"] = pd.to_datetime(df_plot["结束时间"])

    names = sorted(df_plot["employee_no"].astype(str).unique())

    fig = px.timeline(
        df_plot,
        x_start="开始时间",
        x_end="结束时间",
        y="employee_no",
        color="操作分类",
        category_orders={"employee_no": names},
        hover_data={
            "employee_no": True,
            "操作分类": True,
            "duration": ":.2f",
            "开始时间": True,
            "结束时间": True,
        },
    )

    fig.update_yaxes(autorange="reversed")

    fig.update_layout(
        height=max(450, len(names) * 32),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=40, t=30, b=30),
        xaxis_title="Time",
        yaxis_title="Employee",
        legend_title_text="Category",
    )

    return fig


# =========================
# Tables
# =========================
def build_first_action_table(df_day, df, df2, selected_date, history_wait):
    if df_day.empty:
        return pd.DataFrame()

    current_day = pd.to_datetime(selected_date).date()
    prev_date = (pd.to_datetime(selected_date) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")

    records = []

    for name, group in df_day.groupby("name"):
        emp_no = name.split("(")[-1].replace(")", "").strip() if "(" in name else name

        work = group[group["type"].isin(["work", "overnight"])].sort_values("start")

        punch_in = None

        row_iams = df2[
            (df2["employee_no"].astype(str) == emp_no) &
            (df2["上班时间"].apply(adjust_work_date) == current_day)
        ]

        if not row_iams.empty:
            punch_in = row_iams["上班时间"].iloc[0]

        if pd.isna(punch_in):
            row_isc = df[
                (df["employee_no"].astype(str) == emp_no) &
                (df["work_date"] == current_day)
            ]

            if not row_isc.empty and "上班时间" in row_isc.columns:
                punch_in = pd.to_datetime(row_isc["上班时间"].min(), errors="coerce")

        if pd.isna(punch_in):
            punch_in = group["start"].min()

        if not work.empty:
            first_task = work.iloc[0]["start"]
            wait_seconds = (first_task - punch_in).total_seconds()

            prev_day = (pd.to_datetime(selected_date) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
            prev_wait = history_wait.get(prev_day, {}).get(emp_no)

            trend = "-"

            if prev_wait is not None:
                diff = wait_seconds - prev_wait

                diff_abs = abs(diff)
                minutes = int(diff_abs // 60)
                seconds = int(diff_abs % 60)

                if diff > 0:
                    trend = f"↑ {minutes}m {seconds}s"
                elif diff < 0:
                    trend = f"↓ {minutes}m {seconds}s"
                else:
                    trend = "="

            records.append(
                {
                    "Name": name,
                    "Group": group["group"].iloc[0],
                    "Shift": group["shift"].iloc[0],
                    "Punch-in": safe_time(punch_in),
                    "First Task": safe_time(first_task),
                    "Wait": f"{int(wait_seconds // 60)}m {int(wait_seconds % 60)}s",
                    "Trend": trend,
                    "_wait_seconds": wait_seconds,

                    # ✅ 新增
                    "is_late": wait_seconds > 300
                }
            )

    out = pd.DataFrame(records)

    if out.empty:
        return out

    out = out.sort_values("_wait_seconds", ascending=False)
    out = out.drop(columns=["_wait_seconds", "is_late"])

    return out


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="WorkSight Pro", layout="wide")

st.markdown(
    """
    <style>
    .main {
        background-color: #F6F8FC;
    }
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }
    div[data-testid="stMetric"] {
        background-color: #EEF4FF;
        padding: 18px;
        border-radius: 16px;
        border: 1px solid #DBEAFE;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("WorkSight Pro")
st.caption("Advanced productivity analytics dashboard for ISC, iAMS, volume, and punch data")

with st.expander("Data Sources", expanded=True):
    st.markdown("**Upload all files**")
    st.caption("📄 ISC Task Data：来自 ISC员工操作时长（任务单明细）")
    st.caption("👥 iAMS Attendance：来自 iAMS班次明细（实际上下班时间）")
    st.caption("📦 Volume Data：来自 iWMS销售单综合查询（打包完成时间 & 件数）")
    st.caption("🧾 Punch Data：来自 iAMS打卡流水（进出仓记录）")

    files = st.file_uploader(
        "Upload all files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
generate_btn = st.button("Generate Dashboard", type="primary")

if generate_btn:
    if not files:
        st.warning("请先上传文件")
        st.stop()
    file1 = file2 = file3 = file4 = None

    for f in files:
        f_type = classify_file(f)

        if f_type == "task":
            file1 = f
        elif f_type == "attendance":
            file2 = f
        elif f_type == "volume":
            file3 = f
        elif f_type == "punch":
            file4 = f

    # ✅ 必须校验
    if file1 is None:
        st.error("缺少任务单明细（df1）")
        st.stop()

    if file2 is None:
        st.error("缺少班次数据（df2）")
        st.stop()

    if file3 is None:
        st.error("缺少件量数据（df3）")
        st.stop()

    if file4 is None:
        st.error("缺少打卡数据（df4）")
        st.stop()
    
    # ===== loading UI 初始化（必须在 try 前）=====
    status_box = st.empty()

    def update_status(msg, type="info"):
        if type == "success":
            status_box.success(msg)
        elif type == "error":
            status_box.error(msg)
        else:
            status_box.info(msg)

    progress = st.progress(0)

    try:
        update_status(f"📄 任务数据：{file1.name}")
        update_status(f"👥 班次数据：{file2.name}")
        update_status(f"📦 件量数据：{file3.name}")
        update_status(f"🧾 打卡数据：{file4.name}")

        df, df2, group_map = load_data(file1, file2)
        progress.progress(20)

        update_status("📦 读取件量数据...")
        volume_data = load_volume_data(file3)
        progress.progress(40)

        update_status("🧾 读取打卡数据...")
        break_df = load_break_data(file4)
        progress.progress(60)

        update_status("📊 读取间接工时...")
        indirect_df = load_indirect_data(file1)
        progress.progress(70)

        update_status("🧠 构建 timeline...")
        timeline = build_timeline(df, df2, group_map, break_df)
        timeline = remove_break_overlap(timeline)
        progress.progress(85)

        update_status("📈 构建历史数据...")
        history_ratios, history_wait = build_history(timeline)
        progress.progress(95)

        df["warehouse_short"] = df["warehouse_name"].apply(extract_warehouse)
        current_warehouse = df["warehouse_short"].mode()[0] if not df.empty else "Unknown"

        st.session_state["df"] = df
        st.session_state["df2"] = df2
        st.session_state["timeline"] = timeline
        st.session_state["volume_data"] = volume_data
        st.session_state["indirect_df"] = indirect_df
        st.session_state["current_warehouse"] = current_warehouse
        st.session_state["history_ratios"] = history_ratios
        st.session_state["history_wait"] = history_wait

        progress.progress(100)
        update_status("✅ Dashboard 生成完成", "success")

    except Exception as e:
        update_status(f"❌ 出错：{str(e)}", "error")
        st.stop()

if "timeline" not in st.session_state:
    st.info("上传文件后点击 Generate Dashboard。")
    st.stop()

df = st.session_state["df"]
df2 = st.session_state["df2"]
timeline = st.session_state["timeline"]
volume_data = st.session_state["volume_data"]
indirect_df = st.session_state["indirect_df"]
current_warehouse = st.session_state["current_warehouse"]
history_ratios = st.session_state["history_ratios"]
history_wait = st.session_state["history_wait"]

if timeline.empty:
    st.warning("没有生成 timeline 数据，请检查 ISC 和 iAMS 文件时间字段。")
    st.stop()

dates = sorted(timeline["date"].astype(str).unique())

selected_date = st.selectbox("Date", dates, index=0)

current_day = pd.to_datetime(selected_date).date()

df_day = timeline[timeline["date"].astype(str) == selected_date].copy()

summary = build_summary(df_day, df, df2, current_day)

# =========================
# KPI
# =========================
st.subheader("KPI Dashboard")

def calc_attendance_and_work(df_day, df, df2, current_day):

    # ✅ 1. 计算总 work（绿色）
    total_work = df_day[df_day["type"].isin(["work", "overnight"])]["duration"].sum()

    # ✅ 2. 计算总 attendance（所有人上班时间）
    total_attendance = df_day["duration"].sum()

    # ✅ 3. 比例
    ratio = (total_work / total_attendance * 100) if total_attendance > 0 else 0

    return total_work, total_attendance, ratio
volume = int(volume_data.get(current_day, 0)) if volume_data else 0

k1, k2, k3, k4, k5 = st.columns([1, 1, 1, 1.2, 1])
total_work, total_attendance, eff_ratio = calc_attendance_and_work(
    df_day, df, df2, current_day
)
total_workers = df_day["name"].nunique()
k1.metric("WAREHOUSE", current_warehouse)
k2.metric("TOTAL WORKERS", total_workers)
k3.metric("TOTAL ATTENDANCE", f"{total_attendance:.1f}h")
k4.plotly_chart(
    draw_efficiency_donut(eff_ratio, total_attendance, total_work),
    width="stretch"
)
k5.metric("TOTAL VOLUME", f"{volume:,} pcs")

# =========================
# Tabs
# =========================
tab_m, tab_e, tab_fa, tab_no_op, tab_indirect = st.tabs(
    [
        "Morning Shift",
        "Evening Shift",
        "First Action Analysis",
        "No Operation Analysis",
        "Indirect Time",
    ]
)

with tab_m:
    st.subheader("Morning Shift Timeline")

    df_m = df_day[df_day["shift"] == "morning"].copy()
    groups_m = ["All Groups"] + sorted(df_m["group"].dropna().astype(str).unique())

    selected_group_m = st.selectbox("Filter Group - Morning", groups_m)

    if selected_group_m != "All Groups":
        df_m = df_m[df_m["group"] == selected_group_m]

    fig_m = draw_gantt(df_m, selected_date, "morning", summary, history_ratios, df_day, df2)
    if fig_m is None:
        st.info("No morning shift data.")
    else:
        st.plotly_chart(fig_m, width="stretch")

with tab_e:
    st.subheader("Evening Shift Timeline")

    df_e = df_day[df_day["shift"] == "evening"].copy()
    groups_e = ["All Groups"] + sorted(df_e["group"].dropna().astype(str).unique())

    selected_group_e = st.selectbox("Filter Group - Evening", groups_e)

    if selected_group_e != "All Groups":
        df_e = df_e[df_e["group"] == selected_group_e]

    fig_e = draw_gantt(df_e, selected_date, "evening", summary, history_ratios, df_day, df2)
    if fig_e is None:
        st.info("No evening shift data.")
    else:
        st.plotly_chart(fig_e, width="stretch")

with tab_fa:
    st.subheader("First Action Analysis")
    st.caption("Wait time from punch-in to first task, with day-over-day trend.")

    fa_df = build_first_action_table(df_day, df, df2, selected_date, history_wait)

    if fa_df.empty:
        st.info("No first action data.")
    else:
        def highlight_wait(row):
            styles = [""] * len(row)

            # 👉 从 Wait 字符串解析分钟
            wait_str = row["Wait"]

            try:
                minutes = int(wait_str.split("m")[0])
            except:
                minutes = 0

            if minutes > 5:
                styles[row.index.get_loc("Wait")] = "color: red"

            # 👉 Trend 颜色（可选）
            trend_idx = row.index.get_loc("Trend")

            if "↑" in row["Trend"]:
                styles[trend_idx] = "color: red"
            elif "↓" in row["Trend"]:
                styles[trend_idx] = "color: green"

            return styles

        st.dataframe(
            fa_df.style.apply(highlight_wait, axis=1),
            width="stretch",
            height=600
        )

with tab_no_op:
    st.subheader("No System Operation Analysis")
    st.caption("Employees who clocked in but had no system operations.")

    df2_day = df2[
        df2["上班时间"].apply(adjust_work_date) == current_day
    ].copy()

    df_day_isc = df[
        df["work_date"] == current_day
    ].copy()

    no_op_df = find_no_operation_people(df_day_isc, df2_day)

    if no_op_df.empty:
        st.success("No employees found with punch records but no system operations.")
    else:
        st.dataframe(no_op_df, width="stretch", height=600)

with tab_indirect:
    st.subheader("Indirect Time Timeline")
    st.caption("Indirect work time by employee, filtered by operation category.")

    if indirect_df.empty:
        st.info("没有找到间接工时表，或者 sheet 名不是：间接工时明细-间接工时")
    else:
        indirect_day = indirect_df[
            indirect_df["work_date"] == current_day
        ].copy()

        categories = ["All Categories"] + sorted(
            indirect_day["操作分类"].dropna().astype(str).unique()
        )

        selected_category = st.selectbox("Filter Category", categories)

        if selected_category != "All Categories":
            indirect_day = indirect_day[
                indirect_day["操作分类"] == selected_category
            ]

        fig_indirect = draw_indirect_chart(indirect_day)

        if fig_indirect is None:
            st.info("No indirect time data for this date.")
        else:
            st.plotly_chart(fig_indirect, width="stretch")
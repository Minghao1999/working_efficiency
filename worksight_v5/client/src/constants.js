export const API = import.meta.env.VITE_API_URL || "http://127.0.0.1:3001";

export const TYPE_COLORS = {
  work: "#86AEE8",
  overnight: "#111827",
  idle: "#D8E3F0",
  break: "#F3E6B8"
};

export const COLUMN_LABELS = {
  业务日期: "Business Date",
  单量: "Orders",
  件量: "Units",
  总工时: "Total Hours",
  出勤人数: "Headcount",
  仓库: "Warehouse",
  目标UPPH: "Target UPPH",
  日期: "Date",
  工号: "Employee ID",
  姓名: "Name",
  总件数: "Total Units",
  件数: "Units",
  有效工时: "Effective Hours",
  总时长: "Total Duration",
  考勤时长: "Attendance Hours",
  有效工时占比: "Effective Hours Ratio",
  拣非爆品效率: "Non-Burst Picking Efficiency",
  总效率: "Total Efficiency",
  低于目标: "Below Target",
  删除: "Delete",
  上班时间: "Clock In",
  下班时间: "Clock Out",
  employee_no: "Employee ID",
  real_name: "Name",
  操作分类: "Category",
  开始时间: "Start",
  结束时间: "End"
};

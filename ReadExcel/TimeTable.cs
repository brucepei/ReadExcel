using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ReadExcel
{
    class TimeTable : ExcelTable, IExcelTable
    {
        private static Dictionary<String, String> timeAttrTitles = new Dictionary<String, String> {
            {"EmpId", "工号"}, {"Name", "姓名"}, {"SickLeave", "病假（天）"}, {"CasualLeave", "事假（天）"}, {"OverTime", "法定节假日加班（小时）"},
            {"NightWorkDays", "夜班天数（天）"}, {"MidWorkDays", "中班天数（天）"},
        };
        private static Int32[] timeRows = new Int32[] { 1 };
        public TimeTable(String fileName)
            : base(fileName, "员工考勤", timeRows, timeAttrTitles)
        {
        }
        private Dictionary<string, int> updateStatus = new Dictionary<string, int>();
        public Dictionary<string, int> UpdateStatus
        {
            get { return updateStatus; }
        }

        private Regex isNumeric = new Regex(@"^\d+$");

        public void updateEmployment(Employment em)
        {
            if (em == null || em.AllEmployee.Count == 0)
            {
                Logging.logMessage("没有员工信息，请先读取员工表!", LogType.ERROR);
                return;
            }
            var nameCols = getNameCols();
            int empIdIndex = nameCols["EmpId"];
            nameCols.Remove("EmpId");
            for (int i = ContentRow; i < Table.Rows.Count; i++)
            {
                string empId = Table.Rows[i][empIdIndex].ToString();
                if (string.IsNullOrWhiteSpace(empId))
                {
                    Logging.logMessage(String.Format("{0}(行{1}列{2})为空, 忽略该行!", timeAttrTitles["EmpId"], i + 1, empIdIndex + 1), LogType.DEBUG);
                    continue;
                }
                else if (!isNumeric.IsMatch(empId))
                {
                    Logging.logMessage(String.Format("{0}(行{1}列{2})非数字：{3}, 忽略该行!", timeAttrTitles["EmpId"], i + 1, empIdIndex + 1, empId), LogType.DEBUG);
                    continue;
                }
                else
                {
                    Employee e = em.getEmployee(empId);
                    if (e == null)
                    {
                        Logging.logMessage(String.Format("在考勤表中发现员工号{0}({1} row={2})，但员工表中不存在该员工，忽略该员工！", empId, Table.Rows[i][nameCols["Name"]], i + 1), LogType.WARNING);
                        continue;
                    }
                    else
                    {
                        foreach (KeyValuePair<String, int> kv in nameCols)
                        {
                            String timeProp = kv.Key;
                            int colIndex = kv.Value;
                            try
                            {
                                e.setPropertyByString(timeProp, Table.Rows[i][colIndex].ToString());
                            }
                            catch (Exception ex)
                            {
                                Logging.logMessage(String.Format("更新员工 {0}({1}) 考勤属性 {2}(Row {3} Column {4}) 失败: {5}!", empId, Table.Rows[i][nameCols["Name"]], timeAttrTitles[timeProp], i + 1, colIndex + 1, ex.GetOriginalException().Message), LogType.ERROR);
                            }
                        }
                        if (updateStatus.ContainsKey(empId))
                            updateStatus[empId]++;
                        else
                            updateStatus[empId] = 1;
                    }
                }
            }
            foreach (string knownEmpId in em.AllEmployee.Keys)
            {
                if (updateStatus.ContainsKey(knownEmpId))
                {
                    if (updateStatus[knownEmpId] != 1)
                    {
                        Logging.logMessage(String.Format("考勤表表中员工号{0}({1})重复：{2}次！", knownEmpId, em.AllEmployee[knownEmpId].Name, updateStatus[knownEmpId]), LogType.ERROR);
                    }
                }
                else
                {
                    Logging.logMessage(String.Format("考勤表中找不到员工号{0}({1})！", knownEmpId, em.AllEmployee[knownEmpId].Name), LogType.ERROR);
                }
            }
            Logging.logMessage(String.Format("考勤表 {0} 更新完成！", SheetName));
        }
    }
}

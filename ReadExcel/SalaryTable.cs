using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ReadExcel
{
    class SalaryTable : ExcelTable, IExcelTable
    {
        private static Dictionary<String, String> salaryAttrTitles = new Dictionary<String, String> {
            {"EmpId", "ID"}, {"Name", "姓名"}, 
            {"BaseSalary", "月基本工资"}, {"MonthBonus", "月岗位奖金"}, {"YearBase", "年基本绩效"}, {"YearBonus", "年岗位绩效"},
            {"TrafficAllowance", "交通补贴基数"}, {"CommAllowance", "通讯补贴基数"}, {"OutsideAllowance", "异地津贴"}, {"CompanyAllowance", "企业补贴基数"},
            {"LaborAllowance", "劳防津贴"}, {"HotAllowance", "高温津贴"}, {"BirthAllowance", "生日津贴"}, {"NeedTaxSalaryAdjust", "税前工资调整"}, {"NoTaxSalaryAdjust", "税后工资调整"},
            {"EnterpriceYearFond", "企业年金基数"}, {"SingleChildFee", "独生子女费"}, {"HealthInsurance", "税优健康险扣款（月度）"},
        };
        private static Int32[] salaryRows = new Int32[] {0, 1, 2};
        public SalaryTable(String fileName)
            : base(fileName, "明细", salaryRows, salaryAttrTitles)
        {
        }
        private Dictionary<string, int> updateStatus = new Dictionary<string, int>();
        public Dictionary<string, int> UpdateStatus
        {
            get { return updateStatus; }
        }

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
                    Logging.logMessage(String.Format("没有{0}(行{1}列{2}), 忽略该行!", salaryAttrTitles["EmpId"], i + 1, empIdIndex + 1), LogType.DEBUG);
                    continue;
                }
                else
                {
                    Employee e = em.getEmployee(empId);
                    if (e == null)
                    {
                        Logging.logMessage(String.Format("在薪资表中发现员工号{0}({1} row={2})，但员工表中不存在该员工，忽略该员工！", empId, Table.Rows[i][nameCols["Name"]], i + 1), LogType.ERROR);
                        continue;
                    }
                    else
                    {
                        foreach (KeyValuePair<String, int> kv in nameCols)
                        {
                            String salaryProp = kv.Key;
                            int colIndex = kv.Value;
                            try
                            {
                                e.setPropertyByString(salaryProp, Table.Rows[i][colIndex].ToString());
                            }
                            catch (Exception ex)
                            {
                                Logging.logMessage(String.Format("更新员工 {0}({1}) 薪资属性 {2}(Row {3} Column {4}) 失败: {5}!", empId, Table.Rows[i][nameCols["Name"]], salaryAttrTitles[salaryProp], i + 1, colIndex + 1, ex.GetOriginalException().Message), LogType.ERROR);
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
                        Logging.logMessage(String.Format("薪资表中员工号{0}({1})重复：{2}次！", knownEmpId, em.AllEmployee[knownEmpId].Name, updateStatus[knownEmpId]), LogType.ERROR);
                    }
                }
                else
                {
                    Logging.logMessage(String.Format("薪资表中找不到员工号{0}({1})！", knownEmpId, em.AllEmployee[knownEmpId].Name), LogType.ERROR);
                }
            }
            Logging.logMessage(String.Format("薪资表 {0} 更新完成！", SheetName));
        }
    }
}

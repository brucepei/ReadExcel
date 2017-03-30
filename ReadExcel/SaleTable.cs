using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcel
{
    class SaleTable : ExcelTable, IExcelTable
    {
        private static Dictionary<String, String> saleAttrTitles = new Dictionary<String, String> {
            {"Name", "姓名"}, {"BaseSalary", "职级底薪"}, {"TrialCommission", "筹备津贴"}, {"MonthCommission", "月度绩效"},
        };
        private static Int32[] saleRows = new Int32[] { 2 };
        public SaleTable(String fileName)
            : base(fileName, "Sheet1", saleRows, saleAttrTitles)
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
            int empNameIndex = nameCols["Name"];
            nameCols.Remove("Name");
            for (int i = ContentRow; i < Table.Rows.Count; i++)
            {
                string empName = Table.Rows[i][empNameIndex].ToString();
                if (string.IsNullOrWhiteSpace(empName))
                {
                    Logging.logMessage(String.Format("没有{0}(行{1}列{2}), 忽略该行!", saleAttrTitles["Name"], i + 1, empNameIndex + 1), LogType.DEBUG);
                    continue;
                }
                else
                {
                    Employee e = null;
                    try
                    {
                        e = em.getEmployeeByName(empName);
                    }
                    catch (ArgumentException ex)
                    {
                        Logging.logMessage(String.Format("在销售表中发现员工名{0}(row={1})，但员工表中无法确认，忽略该员工: {2}！", empName, i + 1, ex.GetOriginalException().Message), LogType.ERROR);
                        continue;
                    }
                    if (e == null)
                    {
                        Logging.logMessage(String.Format("在销售表中发现员工名{0}(row={1})，但员工表中不无法找到，忽略该员工！", empName, i + 1), LogType.ERROR);
                        continue;
                    }
                    else
                    {
                        e.IsSale = true;
                        foreach (KeyValuePair<String, int> kv in nameCols)
                        {
                            String saleProp = kv.Key;
                            int colIndex = kv.Value;
                            try
                            {
                                Object new_val = Table.Rows[i][colIndex];
                                if ("BaseSalary" == saleProp)
                                {
                                    Decimal orig_val = (Decimal)e.getPropertyByString(saleProp);
                                    Decimal actual_new_val = 0m;
                                    Decimal.TryParse(new_val.ToString(), out actual_new_val);
                                    if (orig_val != actual_new_val)
                                    {
                                        Logging.logMessage(String.Format("员工 {0} 销售属性 {1}(Row {2} Column {3}): {4} 与员工表的基本月薪 {5} 不一致，强制更新为销售表数据!", empName, saleAttrTitles[saleProp], i + 1, colIndex + 1, actual_new_val, orig_val), LogType.WARNING);
                                    }
                                }
                                e.setPropertyByString(saleProp, new_val.ToString());
                            }
                            catch (Exception ex)
                            {
                                Logging.logMessage(String.Format("更新员工 {0} 销售属性 {1}(Row {2} Column {3}) 失败: {4}!", empName, saleAttrTitles[saleProp], i + 1, colIndex + 1, ex.GetOriginalException().Message), LogType.ERROR);
                            }
                        }
                        if (updateStatus.ContainsKey(empName))
                            updateStatus[empName]++;
                        else
                            updateStatus[empName] = 1;
                    }
                }
            }
            foreach (string empName in updateStatus.Keys)
            {
                if (updateStatus[empName] != 1)
                {
                    Logging.logMessage(String.Format("销售表中员工号{0}重复：{1}次！", empName, updateStatus[empName]), LogType.ERROR);
                }
            }
            Logging.logMessage(String.Format("销售表 {0} 更新完成！", SheetName));
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcel
{
    class InsuranceTable : IExcelTable
    {
        private List<IExcelTable> tableList = new List<IExcelTable>();

        public InsuranceTable(String fileName)
        {
            tableList.Add(new HeadInsuranceTable(fileName));
            tableList.Add(new BjInsuranceTable(fileName));
            tableList.Add(new ShInsuranceTable(fileName));
            tableList.Add(new GzInsuranceTable(fileName));
            tableList.Add(new CdInsuranceTable(fileName));
        }
        private Dictionary<string, int> updateStatus = new Dictionary<string, int>();
        public Dictionary<string, int> UpdateStatus
        {
            get { return updateStatus; }
        }

        public void updateEmployment(Employment em)
        {
            foreach (var table in tableList)
            {
                table.updateEmployment(em);
                foreach (var kv in table.UpdateStatus)
                {
                    updateStatus[kv.Key] = kv.Value;
                }
            }
            foreach (var employee in em.AllEmployee.Values)
            {
                String id = employee.Id;
                if (updateStatus.ContainsKey(id))
                {
                    if (updateStatus[id] > 1)
                    {
                        Logging.logMessage(String.Format("所有社保表中 ID 号{0}({1})重复：{2} 次！", id, employee.Name, updateStatus[id]), LogType.ERROR);
                    }
                }
                else
                {
                    Logging.logMessage(String.Format("所有社保表中找不到员工: {0}！", employee), LogType.ERROR);
                }
            }
        }
    }

    class InsuranceBaseTable : ExcelTable, IExcelTable
    {
        private static Int32[] insuranceRows = new Int32[] { 0 };
        public InsuranceBaseTable(String fileName, String sheetName, Dictionary<string, string> insuranceAttrTitles)
            : base(fileName, sheetName, insuranceRows, insuranceAttrTitles)
        {
        }

        private Dictionary<string, int> updateStatus = new Dictionary<string, int> ();
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
            int idIndex = nameCols["Id"];
            nameCols.Remove("Id");
            for (int i = ContentRow; i < Table.Rows.Count; i++)
            {
                string id = Table.Rows[i][idIndex].ToString().ToUpper();
                if (string.IsNullOrWhiteSpace(id))
                {
                    Logging.logMessage(String.Format("没有{0}(行{1}列{2}), 忽略该行!", NameTitles["Id"], i + 1, idIndex + 1), LogType.DEBUG);
                    continue;
                }
                else
                {
                    Employee e = em.getEmployeeById(id);
                    if (e == null)
                    {
                        Logging.logMessage(String.Format("在社保表 {0} 中发现员工号{1}({2} row={3})，但员工表中不存在该员工，忽略该员工！", this.SheetName, id, Table.Rows[i][nameCols["Name"]], i + 1), LogType.ERROR);
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
                                Logging.logMessage(String.Format("更新员工 {0}({1}) 社保属性 {2}(Row {3} Column {4}) 失败: {5}!", id, Table.Rows[i][nameCols["Name"]], NameTitles[timeProp], i + 1, colIndex + 1, ex.GetOriginalException().Message), LogType.ERROR);
                            }
                        }
                        if (updateStatus.ContainsKey(id))
                            updateStatus[id]++;
                        else
                            updateStatus[id] = 1;
                    }
                }
            }
            Logging.logMessage(String.Format("社保表 {0} 更新完成！", SheetName));
        }
    }

    class HeadInsuranceTable : InsuranceBaseTable
    {
        private static Dictionary<String, String> insuranceAttrTitles = new Dictionary<String, String> {
            {"Id", "身份证"}, {"Name", "姓名"}, {"Endowment", "个人养老"}, {"Medical", "个人医疗"}, {"Unemployment", "个人失业"},
            {"Housing", "个人公积金"}, {"SuppleHousing", "个人补充公积金"}, {"BenefitCity", "福利城市"},
        };
        public HeadInsuranceTable(String fileName)
            : base(fileName, "总部", insuranceAttrTitles)
        {
        }
    }

    class BjInsuranceTable : InsuranceBaseTable
    {
        private static Dictionary<String, String> insuranceAttrTitles = new Dictionary<String, String> {
            {"Id", "身份证"}, {"Name", "姓名"}, {"Endowment", "个人养老"}, {"Medical", "个人医疗"}, {"Unemployment", "个人失业"},
            {"Housing", "个人公积金"},  {"BenefitCity", "福利城市"},
        };
        public BjInsuranceTable(String fileName)
            : base(fileName, "北分", insuranceAttrTitles)
        {
        }
    }

    class ShInsuranceTable : InsuranceBaseTable
    {
        private static Dictionary<String, String> insuranceAttrTitles = new Dictionary<String, String> {
            {"Id", "身份证号码"}, {"Name", "姓名"}, {"Endowment", "个人养老"}, {"Medical", "个人医疗"}, {"Unemployment", "个人失业"},
            {"Housing", "个人公积金"}, {"SuppleHousing", "个人补充公积金"}, {"BenefitCity", "工作城市"},
        };
        public ShInsuranceTable(String fileName)
            : base(fileName, "上营", insuranceAttrTitles)
        {
        }
    }

    class GzInsuranceTable : InsuranceBaseTable
    {
        private static Dictionary<String, String> insuranceAttrTitles = new Dictionary<String, String> {
            {"Id", "身份证"}, {"Name", "姓名"}, {"Endowment", "个人养老"}, {"Medical", "个人医疗"}, {"Unemployment", "个人失业"},
            {"Housing", "个人公积金"}, {"BenefitCity", "福利城市"},
        };
        public GzInsuranceTable(String fileName)
            : base(fileName, "广分", insuranceAttrTitles)
        {
        }
    }

    class CdInsuranceTable : InsuranceBaseTable
    {
        private static Dictionary<String, String> insuranceAttrTitles = new Dictionary<String, String> {
            {"Id", "身份证"}, {"Name", "姓名"}, {"Endowment", "个人养老"}, {"Medical", "个人医疗"}, {"Unemployment", "个人失业"},
            {"Housing", "个人公积金"}, {"BenefitCity", "福利城市"}, 
        };
        public CdInsuranceTable(String fileName)
            : base(fileName, "成都分", insuranceAttrTitles)
        {
        }
    }
}

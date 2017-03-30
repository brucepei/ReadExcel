using System;
using System.Collections.Generic;
using System.Data;

namespace ReadExcel
{
    class EmployeeTable : ExcelTable
    {
        private static Dictionary<String, String> employeeAttrTitles = new Dictionary<String, String> {
            {"EmpId", "工号"}, {"Name", "列2"}, {"Company", "公司"}, {"Department", "新部门"}, {"Title", "岗位"}, {"Level", "职级"}, 
            {"Birthday", "出生日期"}, {"WorkDate", "社会工作日"}, {"EmpDate", "入职时间"}, {"EmpOrigDate", "加入太保日期"}, {"Id", "证件号"},
        };
        private static Int32[] employeeRows = new Int32[] {0};
        public EmployeeTable(String fileName)
            : base(fileName, "员工信息总表", employeeRows, employeeAttrTitles)
        {
        }

        public Dictionary<string, Employee> getAllEmployee()
        {
            var allEmployee = new Dictionary<string, Employee>();
            var nameCols = getNameCols();
            int empIdIndex = nameCols["EmpId"];
            nameCols.Remove("EmpId");
            for (int i = ContentRow; i < Table.Rows.Count; i++)
            {
                string empId = Table.Rows[i][empIdIndex].ToString();
                if (string.IsNullOrWhiteSpace(empId))
                {
                    Logging.logMessage(String.Format("没有{0}(行{1}列{2}), 忽略该行!", employeeAttrTitles["EmpId"], i + 1, empIdIndex + 1), LogType.DEBUG);
                    continue;
                }
                Employee existedEmployee;
                allEmployee.TryGetValue(empId, out existedEmployee);
                Employee e = null;
                if (existedEmployee == null)
                {
                    e = new Employee(empId);
                }
                else
                {
                    Logging.logMessage(String.Format("发现重复的员工号: {0}({1}) 在(行{2}列{3}), 忽略该行！", empId, Table.Rows[i][nameCols["Name"]], i + 1, empIdIndex + 1), LogType.ERROR);
                    continue;
                }
                Boolean needAddEmployee = true;
                foreach (KeyValuePair<String, int> kv in nameCols)
                {
                    String empProp = kv.Key;
                    int colIndex = kv.Value;
                    try
                    {
                        e.setPropertyByString(empProp, Table.Rows[i][colIndex].ToString());
                    }
                    catch (Exception ex)
                    {
                        Logging.logMessage(String.Format("设置员工{0}({1})属性 {2} 失败(行{3}, 列{4}): 不再添加该员工: {5}!", empId, e.Name, employeeAttrTitles[empProp], i + 1, colIndex + 1, ex.GetOriginalException().Message), LogType.ERROR);
                        needAddEmployee = false;
                        break;
                    }
                }
                if (needAddEmployee) allEmployee.Add(e.EmpId, e);
            }
            Logging.logMessage(String.Format("员工表 {0} 更新完成！", SheetName));
            return allEmployee;
        }
    }
}

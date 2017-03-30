using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.ComponentModel;

namespace ReadExcel
{
    class Employment : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged(String info)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(info));
            }
        }

        public readonly Decimal AverageMonthDays = 21.75m;

        private Boolean useAverageMonthDays = true;
        public Boolean UseAverageMonthDays
        {
            get { return useAverageMonthDays; }
            set
            {
                if (value != useAverageMonthDays)
                {
                    useAverageMonthDays = value;
                    NotifyPropertyChanged("UseAverageMonthDays");
                }
            }
        }

        private DateTime salaryThisDate;
        public DateTime SalaryThisDate
        {
            get { return salaryThisDate; }
            set
            {
                if (value != salaryThisDate)
                {
                    salaryThisDate = value;
                    salaryLastDate = value.AddMonths(-1);
                    salaryThisDateLastDay = value.AddMonths(1).AddDays(-1);
                    salaryLastDateLastDay = value.AddDays(-1);
                    NotifyPropertyChanged("SalaryThisDate");
                }
            }
        }

        private DateTime salaryThisDateLastDay;
        public DateTime SalaryThisDateLastDay
        {
            get { return salaryThisDateLastDay; }
        }

        private DateTime salaryLastDate;
        public DateTime SalaryLastDate
        {
            get { return salaryLastDate; }
        }

        private DateTime salaryLastDateLastDay;
        public DateTime SalaryLastDateLastDay
        {
            get { return salaryLastDateLastDay; }
        }

        private Decimal salaryThisMonthDays = 0m;
        public Decimal SalaryThisMonthDays
        {
            get { return salaryThisMonthDays; }
            set
            {
                if (value != salaryThisMonthDays)
                {
                    salaryThisMonthDays = value;
                }
            }
        }

        private Decimal salaryLastMonthDays = 0m;
        public Decimal SalaryLastMonthDays
        {
            get { return salaryLastMonthDays; }
            set {
                if (value != salaryLastMonthDays)
                {
                    salaryLastMonthDays = value;
                }
            }
        }

        private String employeeXLSX;
        public String EmployeeXLSX
        {
            get { return employeeXLSX; }
            set { employeeXLSX = value; }
        }

        private EmployeeTable employeeTb;
        public EmployeeTable EmployeeTb
        {
            get { return employeeTb; }
        }

        private SalaryTable salaryTb;
        public SalaryTable SalaryTb
        {
            get { return salaryTb; }
        }

        private TimeTable timeTb;
        public TimeTable TimeTb
        {
            get { return timeTb; }
        }

        private SaleTable saleTb;
        public SaleTable SaleTb
        {
            get { return saleTb; }
        }

        private InsuranceTable insuranceTb;
        public InsuranceTable InsuranceTb
        {
            get { return insuranceTb; }
        }

        private String salaryXLSX;

        public String SalaryXLSX
        {
            get { return salaryXLSX; }
            set { salaryXLSX = value; }
        }

        private String saleXLSX;

        public String SaleXLSX
        {
            get { return saleXLSX; }
            set { saleXLSX = value; }
        }

        private String timeXLSX;

        public String TimeXLSX
        {
            get { return timeXLSX; }
            set { timeXLSX = value; }
        }

        private String insuranceXLSX;

        public String InsuranceXLSX
        {
            get { return insuranceXLSX; }
            set { insuranceXLSX = value; }
        }

        private Dictionary<string, Employee> allEmployee;

        internal Dictionary<string, Employee> AllEmployee
        {
            get { return allEmployee; }
        }

        private Boolean initialized = false;
        public Boolean Initialized
        {
            get { return initialized; }
        }

        public void clear()
        {
            allEmployee = null;
            initialized = false;
        }

        public void init(String employeeXLSX)
        {
            if (employeeXLSX.Length > 0)
            {
                this.employeeXLSX = employeeXLSX;
                try
                {
                    var et = new EmployeeTable(EmployeeXLSX);
                    employeeTb = et;
                    allEmployee = et.getAllEmployee();
                    initialized = true;
                }
                catch (ArgumentException ex)
                {
                    throw new ArgumentException(String.Format("读取员工表失败：\n{0}", ex.GetOriginalException().Message));
                }
            }
            else
            {
                throw new ArgumentException("员工表路径不能为空!");
            }
            Logging.logMessage(String.Format("总员工数目: {0}", allEmployee.Count), LogType.INFO);
        }

        public void updateOthers(String otherName, String otherXLSX)
        {
            if (otherXLSX.Length > 0)
            {
                try
                {
                    IExcelTable table = null;
                    if (otherName == "SalaryTable")
                    {
                        salaryXLSX = otherXLSX;
                        table = new SalaryTable(otherXLSX);
                        salaryTb = table as SalaryTable;
                    }
                    else if (otherName == "SaleTable")
                    {
                        saleXLSX = otherXLSX;
                        table = new SaleTable(otherXLSX);
                        saleTb = table as SaleTable;
                    }
                    else if (otherName == "TimeTable")
                    {
                        timeXLSX = otherXLSX;
                        table = new TimeTable(otherXLSX);
                        timeTb = table as TimeTable;
                    }
                    else if (otherName == "InsuranceTable")
                    {
                        insuranceXLSX = otherXLSX;
                        table = new InsuranceTable(otherXLSX);
                        insuranceTb = table as InsuranceTable;
                    }
                    else
                    {
                        throw new ArgumentException(String.Format("不支持的表 {0}:{1}", otherName, otherXLSX));
                    }
                    table.updateEmployment(this);
                }
                catch (ArgumentException ex)
                {
                    throw new ArgumentException(String.Format("加载表 {0}:{1} 失败: {2}", otherName, otherXLSX, ex.GetOriginalException().Message));
                }
            }
        }

        public void saveSalaryDetailCSV()
        {
            DataTable dt = new DataTable();
            dt.TableName = "全体员工薪酬当月数据";
            foreach (var kv in Employee.getAttrNames())
            {
                dt.Columns.Add(kv.Value);
            }
            foreach (var empId in allEmployee.Keys)
            {
                var e = allEmployee[empId];
                dt.Rows.Add(e.getAttrs(this));
            }
            ExcelHelper.writeToCSV(dt, String.Format("工资明细 {0}{1:D2}.csv", salaryThisDate.Year, salaryThisDate.Month));
        }

        public void summary()
        {
            Logging.logMessage(String.Format("{0:D2}月份({1}天) 员工汇总数据：", salaryThisDate.Month, salaryThisMonthDays), LogType.NOTE);
            Logging.logMessage(String.Format("员工表: {0} ({1}人)", employeeXLSX, allEmployee.Count), LogType.NOTE, 1);
            Logging.logMessage(String.Format("薪资表: {0} ({1}人)", salaryXLSX, salaryTb == null ? 0 : salaryTb.UpdateStatus.Count), LogType.NOTE, 1);
            Logging.logMessage(String.Format("销售表: {0} ({1}人)", saleXLSX, saleTb == null ? 0 : saleTb.UpdateStatus.Count), LogType.NOTE, 1);
            Logging.logMessage(String.Format("考勤表: {0} ({1}人)", timeXLSX, timeTb == null ? 0 : timeTb.UpdateStatus.Count), LogType.NOTE, 1);
            Logging.logMessage(String.Format("社保表: {0} ({1}人)", insuranceXLSX, insuranceTb == null ? 0 : insuranceTb.UpdateStatus.Count), LogType.NOTE, 1);
            Logging.logMessage("员工信息统计如下:", LogType.NOTE, 1);
            Int32 saleCount = 0;
            var companyDepartmentEmployee = new Dictionary<String, Dictionary<String, Int32>>();
            foreach (var emp in allEmployee.Values)
            {
                if (!companyDepartmentEmployee.ContainsKey(emp.Company))
                {
                    companyDepartmentEmployee[emp.Company] = new Dictionary<String, Int32>();
                }
                if (!companyDepartmentEmployee[emp.Company].ContainsKey(emp.Department))
                {
                    companyDepartmentEmployee[emp.Company][emp.Department] = 0;
                }
                companyDepartmentEmployee[emp.Company][emp.Department]++;
                if (!emp.onBoard(this))
                {
                    Logging.logMessage(String.Format("员工{0}({1})入职日期 {2} 还未到！", emp.EmpId, emp.Name, emp.EmpDate.ToShortDateString()), LogType.NOTE, 2);
                }
                if (emp.IsSale)
                {
                    saleCount++;
                    Logging.logMessage(String.Format("销售员工{0:D2}: {1}({2})", saleCount, emp.Name, emp.EmpId), LogType.DEBUG, 2);
                }
            }
            Logging.logMessage(String.Format("全体员工 {0}, 销售员工: {1}", allEmployee.Count, saleCount), LogType.NOTE, 2);
            Logging.logMessage(String.Format("总计共 {0} 子公司", companyDepartmentEmployee.Count), LogType.NOTE, 2);
            foreach (var kv1 in companyDepartmentEmployee)
            {
                var company = kv1.Key;
                var departmentEmployee = kv1.Value;
                Logging.logMessage(String.Format("{0}, {1} 部门:", company, departmentEmployee.Count), LogType.NOTE, 2);
                foreach (var kv2 in departmentEmployee)
                {
                    var department = kv2.Key;
                    var employeeCount = kv2.Value;
                    Logging.logMessage(String.Format("部门: {1} 员工数: {2}", company, department, employeeCount), LogType.NOTE, 3);
                }
            }
        }

        public Boolean addEmployee(Employee e)
        {
            Boolean result = true;
            if (string.IsNullOrWhiteSpace(e.EmpId))
            {
                Logging.logMessage("发现空的员工号: " + e.EmpId, LogType.DEBUG);
                result = false;
            }
            else if(allEmployee.ContainsKey(e.EmpId))
            {
                Logging.logMessage("发现重复的员工号: " + e.EmpId, LogType.WARNING);
                result = false;
            }
            else
            {
                allEmployee.Add(e.EmpId, e);
            }
            return result;
        }

        public Employee getEmployee(string empId)
        {
            Employee e = null;
            if (allEmployee.ContainsKey(empId))
            {
                e = allEmployee[empId];
            }
            return e;
        }

        public Employee getEmployeeByName(string empName)
        {
            Employee e = null;
            Int32 found = 0;
            foreach (var emp in allEmployee.Values)
            {
                if (emp.Name == empName)
                {
                    found++;
                    e = emp;
                }
            }
            if (found > 1)
                throw new ArgumentException(String.Format("存在 {0} 名重复的员工名 {0}，无法确认员工！", found, empName));
            return e;
        }

        public Employee getEmployeeById(string id)
        {
            Employee e = null;
            Int32 found = 0;
            foreach (var emp in allEmployee.Values)
            {
                if (emp.Id.ToUpper() == id.ToUpper())
                {
                    found++;
                    e = emp;
                }
            }
            if (found > 1)
                throw new ArgumentException(String.Format("存在 {0} 名重复的员工名 {0}，无法确认员工！", found, id));
            return e;
        }

        public static Dictionary<string, string> getCompanySummaryNames()
        {
            Dictionary<string, string> result = new Dictionary<string, string> {
                {"Department", "部门"}, {"employeeCount", "人数"}, 
                {"allBaseSalary", "基本工资"}, {"allSickCausualFee", "病事假扣款"},  {"allBaseTotal", "小计"},
                {"allMonthBonus", "岗位月奖"}, {"allMonthAdjust", "税前调整"}, {"curMonthSalary", "月度绩效"}, 
                {"curMonthSalaryAdjust", "季度绩效"},{"curMonthSalaryAdjust", "年度绩效"},{"curMonthSalaryAdjust", "首年津贴"},
                {"curMonthSalaryAdjust", "小计"},
                {"curMonthBonusSum", "加班费"}, {"MonthBonus", "交通补贴"}, {"curMonthBonus", "通讯补贴"}, {"curMonthBonusAdjust", "企业补贴"},
                {"curMonthBonusSum", "生日津贴"}, {"MonthBonus", "劳防补贴"}, {"curMonthBonus", "轮班津贴"}, {"curMonthBonusAdjust", "推荐奖金"},
                {"overTimeFee", "小计"},{"overTimeFee", "高温津贴"},{"overTimeFee", "独子贴"},{"overTimeFee", "税优健康险补贴"},{"overTimeFee", "扣前合计"},
                {"overTimeFee", "住房公积金"},{"overTimeFee", "补充公积金"},{"overTimeFee", "养老保险金"},{"overTimeFee", "医疗保险金"},{"overTimeFee", "失业保险金"},
                {"overTimeFee", "企业年金"},{"overTimeFee", "工会费"},{"overTimeFee", "个人所得税"},{"overTimeFee", "税优健康险扣款"},{"overTimeFee", "实发金额"},
            };
            Logging.logMessage(String.Format("共有{0}列!", result.Count));
            return result;
        }

        public void saveCompanySummaryCSV()
        {
            //Double result = 0L;
            throw new NotImplementedException("还未实现!");
            //var names = getCompanySummaryNames();
            //foreach (var emp in allEmployee.Values)
            //{
            //}
            //result = Math.Round(result, 2, MidpointRounding.AwayFromZero);
        } 
    }
}

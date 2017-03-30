using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace ReadExcel
{
    class Employee
    {
        public Employee(string empId)
        {
            this.empId = empId;
        }

        public static Int32 autoId = 1;
        public static Dictionary<String, Decimal> SocialAveMaxSalary = new Dictionary<String, Decimal>{
            {"上海", 19512m}, {"北京", 21258m}, {"广州", 20292m}, {"成都", 14370m}
        };
        public static Decimal FreeChineseTaxSalary = 3500m;
        public static Decimal FreeForeignTaxSalary = 4800m;

        #region EmployeeBasicAttrs
        private DateTime birthday;
        public DateTime Birthday
        {
            get { return birthday; }
            set { birthday = value; }
        }

        private string id;
        public string Id
        {
            get { return id; }
            set { 
                id = value.ToUpper();
                if (id.Length != 18 && id.Length != 15)
                {
                    isChinese = false;
                    Logging.logMessage(String.Format("发现外国人 {0}({1})，税率将与中国人不同!", name, id), LogType.NOTE);
                }
            }
        }

        private Boolean isChinese = true;
        public Boolean IsChinese
        {
            get { return isChinese; }
        }

        private string empId;
        public string EmpId
        {
            get { return empId; }
            set { empId = value; }
        }

        private string name;
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        private Decimal baseSalary = 0m;
        public Decimal BaseSalary
        {
            get { return baseSalary; }
            set {
                if (value == 0m)
                {
                    Logging.logMessage(String.Format("注意员工号{0} 月基本工资为0!", this.empId), LogType.WARNING);
                }
                baseSalary = value;
            }
        }

        private Decimal monthBonus = 0m;
        public Decimal MonthBonus
        {
            get { return monthBonus; }
            set { monthBonus = value; }
        }

        private Decimal yearBase = 0m;
        public Decimal YearBase
        {
            get { return yearBase; }
            set { yearBase = value; }
        }

        private Decimal yearBonus = 0m;
        public Decimal YearBonus
        {
            get { return yearBonus; }
            set { yearBonus = value; }
        }

        private Decimal trafficAllowance = 0m;
        public Decimal TrafficAllowance
        {
            get { return trafficAllowance; }
            set { trafficAllowance = value; }
        }

        private Decimal commAllowance = 0m;
        public Decimal CommAllowance
        {
            get { return commAllowance; }
            set { commAllowance = value; }
        }

        private Decimal outsideAllowance = 0m;
        public Decimal OutsideAllowance
        {
            get { return outsideAllowance; }
            set { outsideAllowance = value; }
        }

        private Decimal companyAllowance = 0m;
        public Decimal CompanyAllowance
        {
            get { return companyAllowance; }
            set { companyAllowance = value; }
        }

        private Decimal laborAllowance = 0m;
        public Decimal LaborAllowance
        {
            get { return laborAllowance; }
            set { laborAllowance = value; }
        }

        private Decimal hotAllowance = 0m;
        public Decimal HotAllowance
        {
            get { return hotAllowance; }
            set { hotAllowance = value; }
        }

        private Decimal birthAllowance = 0m;
        public Decimal BirthAllowance
        {
            get { return birthAllowance; }
            set { birthAllowance = value; }
        }

        private Decimal saleAllowance = 0m;
        public Decimal SaleAllowance
        {
            get { return saleAllowance; }
            set { saleAllowance = value; }
        }

        private Decimal saleBonus = 0m;
        public Decimal SaleBonus
        {
            get { return saleBonus; }
            set { saleBonus = value; }
        }

        private string company;
        public string Company
        {
            get { return company; }
            set { company = value; }
        }

        private string department;
        public string Department
        {
            get { return department; }
            set { department = value; }
        }

        private string oldDepartment;
        public string OldDepartment
        {
            get { return oldDepartment; }
            set { oldDepartment = value; }
        }

        private string title;
        public string Title
        {
            get { return title; }
            set { title = value; }
        }

        private int level;
        public int Level
        {
            get { return level; }
            set { level = value; }
        }

        private string benefitCity;
        public string BenefitCity
        {
            get { return benefitCity; }
            set { benefitCity = value; }
        }

        private DateTime workDate;
        public DateTime WorkDate
        {
            get { return workDate; }
            set { workDate = value; }
        }

        private DateTime empDate;
        public DateTime EmpDate
        {
            get { return empDate; }
            set { empDate = value; }
        }

        private DateTime empOrigDate;
        public DateTime EmpOrigDate
        {
            get { return empOrigDate; }
            set { empOrigDate = value; }
        }

        private Decimal casualLeave = 0m;
        public Decimal CasualLeave
        {
            get { return casualLeave; }
            set {
                if (value % 0.5m == 0)
                {
                    casualLeave = value;
                }
                else
                {
                    throw new ArgumentException(String.Format("事假 {0} 不是0.5的倍数!", value));
                }
            }
        }

        private Decimal sickLeave = 0m;
        public Decimal SickLeave
        {
            get { return sickLeave; }
            set {
                if (value % 0.5m == 0)
                {
                    sickLeave = value;
                }
                else
                {
                    throw new ArgumentException(String.Format("病假 {0} 不是0.5的倍数!", value));
                }
            }
        }

        private Decimal overTime = 0m;
        public Decimal OverTime
        {
            get { return overTime; }
            set { overTime = value; }
        }

        private Int32 nightWorkDays = 0;
        public Int32 NightWorkDays
        {
            get { return nightWorkDays; }
            set { nightWorkDays = value; }
        }

        private Int32 midWorkDays = 0;
        public Int32 MidWorkDays
        {
            get { return midWorkDays; }
            set { midWorkDays = value; }
        }

        private Decimal trialCommission = 0m;
        public Decimal TrialCommission
        {
            get { return trialCommission; }
            set { trialCommission = value; }
        }

        private Decimal monthCommission = 0m;
        public Decimal MonthCommission
        {
            get { return monthCommission; }
            set { monthCommission = value; }
        }

        private Boolean isSale;
        public Boolean IsSale
        {
            get { return isSale; }
            set { isSale = value; }
        }

        private Decimal endowment;
        public Decimal Endowment
        {
            get { return endowment; }
            set { endowment = value; }
        }

        private Decimal medical;
        public Decimal Medical
        {
            get { return medical; }
            set { medical = value; }
        }

        private Decimal unemployment;
        public Decimal Unemployment
        {
            get { return unemployment; }
            set { unemployment = value; }
        }

        private Decimal housing;
        public Decimal Housing
        {
            get { return housing; }
            set { housing = value; }
        }

        private Decimal suppleHousing;
        public Decimal SuppleHousing
        {
            get { return suppleHousing; }
            set { suppleHousing = value; }
        }

        private Decimal healthInsurance;
        public Decimal HealthInsurance
        {
            get { return healthInsurance; }
            set { healthInsurance = value; }
        }

        private Decimal enterpriceYearFond;
        public Decimal EnterpriceYearFond
        {
            get { return enterpriceYearFond; }
            set { enterpriceYearFond = value; }
        }

        private Decimal singleChildFee;
        public Decimal SingleChildFee
        {
            get { return singleChildFee; }
            set { singleChildFee = value; }
        }

        private Decimal needTaxSalaryAdjust;
        public Decimal NeedTaxSalaryAdjust
        {
            get { return needTaxSalaryAdjust; }
            set { needTaxSalaryAdjust = value; }
        }

        private Decimal noTaxSalaryAdjust;
        public Decimal NoTaxSalaryAdjust
        {
            get { return noTaxSalaryAdjust; }
            set { noTaxSalaryAdjust = value; }
        }
        
        #endregion

        #region EmployeeComplexAttrs
        public Decimal workDay(Employment em)
        {
            Decimal days = 0;
            if (empDate < em.SalaryLastDate)
            {
                days = em.SalaryThisMonthDays;
            }
            else
            {
                days = getWorkDays(empDate, em.SalaryThisDateLastDay);
            }
            return days;
        }

        public Decimal curMonthSalary(Employment em)
        {
            return curMonthIncome(em, baseSalary);
        }

        public Decimal curMonthSalaryAdjust(Employment em)
        {
            return curMonthIncomeAdjust(em, baseSalary);
        }

        public Decimal curMonthSalarySum(Employment em)
        {
            Decimal result = curMonthSalary(em) + curMonthSalaryAdjust(em);
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal curMonthBonus(Employment em)
        {
            return curMonthIncome(em, monthBonus);
        }

        public Decimal curMonthBonusAdjust(Employment em)
        {
            return curMonthIncomeAdjust(em, monthBonus);
        }

        public Decimal curMonthBonusSum(Employment em)
        {
            Decimal result = curMonthBonus(em) + curMonthBonusAdjust(em);
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal curMonthTrafficAllowance(Employment em)
        {
            return curMonthIncome(em, trafficAllowance);
        }

        public Decimal curMonthTrafficAllowanceAdjust(Employment em)
        {
            return curMonthIncomeAdjust(em, trafficAllowance);
        }

        public Decimal curMonthTrafficAllowanceSum(Employment em)
        {
            Decimal result = curMonthTrafficAllowance(em) + curMonthTrafficAllowanceAdjust(em);
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal curMonthCommAllowance(Employment em)
        {
            return curMonthIncome(em, commAllowance);
        }

        public Decimal curMonthCommAllowanceAdjust(Employment em)
        {
            return curMonthIncomeAdjust(em, commAllowance);
        }

        public Decimal curMonthCommAllowanceSum(Employment em)
        {
            Decimal result = curMonthCommAllowance(em) + curMonthCommAllowanceAdjust(em);
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal curMonthCompanyAllowance(Employment em)
        {
            return curMonthIncome(em, companyAllowance);
        }

        public Decimal curMonthCompanyAllowanceAdjust(Employment em)
        {
            return curMonthIncomeAdjust(em, companyAllowance);
        }

        public Decimal curMonthCompanyAllowanceSum(Employment em)
        {
            Decimal result = curMonthCompanyAllowance(em) + curMonthCompanyAllowanceAdjust(em);
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal curMonthHotAllowance(Employment em)
        {
            Decimal income = 0.0m;
            if (em.SalaryThisDate.Month >= 6 && em.SalaryThisDate.Month <= 9)
                income = curMonthIncome(em, hotAllowance);
            return Math.Round(income, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal curMonthHotAllowanceAdjust(Employment em)
        {
            Decimal income = 0.0m;
            if (em.SalaryLastDate.Month >= 6 && em.SalaryLastDate.Month <= 9)
                income = curMonthIncomeAdjust(em, hotAllowance);
            return Math.Round(income, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal curMonthHotAllowanceSum(Employment em)
        {
            Decimal result = curMonthHotAllowance(em) + curMonthHotAllowanceAdjust(em);
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal overTimeFee(Employment em)
        {
            Decimal result = 0.0m;
            if (overTime > 0m)
            {
                Int32 times = 3;
                result = overTime * times * (baseSalary + monthBonus) / (em.SalaryLastMonthDays * 8.0m);
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal birthdayFee(Employment em)
        {
            Decimal result = 0.0m;
            if (birthday.Month == em.SalaryThisDate.Month)
            {
                if (empOrigDate.AddMonths(6).Year >= em.SalaryThisDate.Year)
                {
                    result = birthAllowance / 2;
                }
                else
                {
                    result = birthAllowance;
                }
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal healthInsurancePersonFee(Employment em)
        {
            Decimal result = 2 * healthInsurance;
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal laborAllowanceFee(Employment em)
        {
            Decimal result = 0.0m;
            if (em.SalaryThisDate.Month == 4 && workDay(em) >= em.SalaryThisMonthDays)
            {
                result = laborAllowance;
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal turnWorkFee(Employment em)
        {
            Decimal result = 0.0m;
            if (nightWorkDays > 0 || midWorkDays > 0)
            {
                result = nightWorkDays * 100 + midWorkDays * 50;
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal trialCommissionFee(Employment em)
        {
            Decimal result = 0.0m;
            if (trialCommission > 0)
            {
                Decimal days = workDay(em);
                result = trialCommission / em.SalaryThisMonthDays * days;
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal monthCommissionFee(Employment em)
        {
            Decimal result = 0.0m;
            if (monthCommission > 0)
            {
                Decimal days = workDay(em);
                result = monthCommission / em.SalaryThisMonthDays * days;
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal sickLeaveBaseFee(Employment em)
        {
            Decimal result = 0.0m;
            if (sickLeave > 0)
            {
                var discount_1_date = em.SalaryThisDateLastDay.AddYears(-8);
                var discount_2_date = em.SalaryThisDateLastDay.AddYears(-6);
                var discount_3_date = em.SalaryThisDateLastDay.AddYears(-4);
                var discount_4_date = em.SalaryThisDateLastDay.AddYears(-2);
                if (empOrigDate < discount_1_date)
                {
                    result = 0m;
                }
                else if (empOrigDate < discount_2_date)
                {
                    result = baseSalary * 0.1m / em.SalaryThisMonthDays * sickLeave;
                }
                else if (empOrigDate < discount_3_date)
                {
                    result = baseSalary * 0.2m / em.SalaryThisMonthDays * sickLeave;
                }
                else if (empOrigDate < discount_4_date)
                {
                    result = baseSalary * 0.3m / em.SalaryThisMonthDays * sickLeave;
                }
                else
                {
                    result = baseSalary * 0.4m / em.SalaryThisMonthDays * sickLeave;
                }
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal sickLeaveBonusFee(Employment em)
        {
            Decimal result = 0.0m;
            if (sickLeave > 0)
            {
                result = (monthBonus + commAllowance + companyAllowance + trafficAllowance) / em.SalaryThisMonthDays * sickLeave;
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal casualLeaveBaseFee(Employment em)
        {
            Decimal result = 0.0m;
            if (casualLeave > 0)
            {
                result = baseSalary / em.SalaryThisMonthDays * casualLeave;
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal casualLeaveBonusFee(Employment em)
        {
            Decimal result = 0.0m;
            if (casualLeave > 0)
            {
                result = (monthBonus + commAllowance + companyAllowance + trafficAllowance) / em.SalaryThisMonthDays * casualLeave;
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal workUnionFee(Employment em)
        {
            Decimal result = curMonthSalarySum(em) * 0.005m;
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal monthNeedPaySalary(Employment em)
        {
            Decimal result = curMonthSalarySum(em) + curMonthBonusSum(em) + overTimeFee(em) + curMonthTrafficAllowanceSum(em)
                + curMonthCommAllowanceSum(em) + curMonthCompanyAllowanceSum(em) + curMonthHotAllowanceSum(em)
                + outsideAllowance + birthdayFee(em) + laborAllowanceFee(em) + turnWorkFee(em) + trialCommissionFee(em)
                + monthCommissionFee(em) + HealthInsurance + needTaxSalaryAdjust - sickLeaveBaseFee(em) - sickLeaveBonusFee(em)
                - casualLeaveBaseFee(em) - casualLeaveBonusFee(em);
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal socialInsuranceFee(Employment em)
        {
            Decimal result = endowment + medical + unemployment + housing + suppleHousing;
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal enterpriceYearFondPersonFee(Employment em)
        {
            Decimal result = 0m;
            if (enterpriceYearFond > 0)
            {
                result = enterpriceYearFond * 0.05m * 0.25m;
                Decimal max = 0m;
                if (SocialAveMaxSalary.ContainsKey(benefitCity))
                {
                    max = SocialAveMaxSalary[benefitCity] * 0.04m;
                }
                else
                {
                    throw new ArgumentException(String.Format("无法为员工 {0} 获取城市 {1} 的社保基数3倍值！", benefitCity, name));
                }
                if (result > max)
                {
                    result = max;
                }
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal needTaxPaySalary(Employment em)
        {
            Decimal result = monthNeedPaySalary(em) - socialInsuranceFee(em) - enterpriceYearFondPersonFee(em) - healthInsurancePersonFee(em);
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal personTax(Employment em)
        {
            Decimal result = 0m;
            Decimal freeTaxSalary = FreeChineseTaxSalary;
            if (!isChinese)
            {
                freeTaxSalary = FreeForeignTaxSalary;
            }
            Decimal needTaxIncome = needTaxPaySalary(em) - freeTaxSalary;
            if (needTaxIncome > 80000)
            {
                result = 0.45m * needTaxIncome - 13505m;
            }
            else if (needTaxIncome > 55000)
            {
                result = 0.35m * needTaxIncome - 5505m;
            }
            else if (needTaxIncome > 35000)
            {
                result = 0.30m * needTaxIncome - 2755m;
            }
            else if (needTaxIncome > 9000)
            {
                result = 0.25m * needTaxIncome - 1005m;
            }
            else if (needTaxIncome > 4500)
            {
                result = 0.20m * needTaxIncome - 555m;
            }
            else if (needTaxIncome > 1500)
            {
                result = 0.10m * needTaxIncome - 105m;
            }
            else if (needTaxIncome > 0)
            {
                result = 0.03m * needTaxIncome;
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal enterpriceYearFondPersonAfterTaxFee(Employment em)
        {
            Decimal result = 0m;
            if (enterpriceYearFond > 0)
            {
                result = enterpriceYearFond * 0.05m * 0.25m - enterpriceYearFondPersonFee(em);
                if (result < 0.005m)
                {
                    result = 0m;
                }
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal actualPaySalary(Employment em)
        {
            Decimal result = needTaxPaySalary(em) - personTax(em) + singleChildFee - enterpriceYearFondPersonAfterTaxFee(em) - workUnionFee(em) + noTaxSalaryAdjust;
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        #endregion

        #region EmployeeDynamicInvokes
        public static Dictionary<string, string> getAttrNames()
        {
            Dictionary<string, string> result = new Dictionary<string,string> {
                {"__id", "序号"}, {"EmpId", "工号"}, {"Name", "姓名"}, {"Company", "成本归属"}, {"Department", "部门"}, 
                {"EmpDate", "入职日期"}, {"QuitDate", "离职日期"}/*not implemented*/,  {"workDay", "工作日"},
                {"curMonthSalarySum", "基本工资合计"}, {"BaseSalary", "基本工资基数"}, {"curMonthSalary", "当月基本工资"}, {"curMonthSalaryAdjust", "基本工资调整"},
                {"curMonthBonusSum", "岗位奖金合计"}, {"MonthBonus", "岗位奖金基数"}, {"curMonthBonus", "当月岗位奖金"}, {"curMonthBonusAdjust", "岗位奖金调整"},
                {"overTimeFee", "加班费"},
                {"curMonthTrafficAllowanceSum", "交通补贴合计"}, {"TrafficAllowance", "交通补贴基数"}, {"curMonthTrafficAllowance", "当月交通补贴"}, {"curMonthTrafficAllowanceAdjust", "交通补贴调整"},
                {"curMonthCommAllowanceSum", "通讯补贴合计"}, {"CommAllowance", "通讯补贴基数"}, {"curMonthCommAllowance", "当月通讯补贴"}, {"curMonthCommAllowanceAdjust", "通讯补贴调整"},
                {"curMonthCompanyAllowanceSum", "企业补贴合计"}, {"CompanyAllowance", "企业补贴基数"}, {"curMonthCompanyAllowance", "当月企业补贴"}, {"curMonthCompanyAllowanceAdjust", "企业补贴调整"},
                {"curMonthHotAllowanceSum", "高温津贴合计"}, {"HotAllowance", "高温补贴基数"}, {"curMonthHotAllowance", "当月高温补贴"}, {"curMonthHotAllowanceAdjust", "高温补贴调整"},
                {"OutsideAllowance", "异地津贴"}, {"birthdayFee", "生日津贴"}, {"LaborAllowance", "劳防津贴基数"}, {"laborAllowanceFee", " 劳防补贴"},
                {"turnWorkFee", "轮班津贴"}, 
                {"trialCommissionFee", "首年津贴"}, {"recommendFee", "推荐奖金"}/*not implemented*/, 
                {"HealthInsurance", "税优健康险企业补贴"}, {"monthCommissionFee", "月度绩效奖金"}, 
                {"seasonCommissionFee", "季度绩效奖金"}/*not implemented*/, {"yearCommissionFee", "年度绩效奖金"}/*not implemented*/, 
                {"sickLeaveBaseFee", "病假基本工资扣款"}, {"sickLeaveBonusFee", "病假岗位月奖扣款"}, 
                {"casualLeaveBaseFee", "事假基本工资扣款"}, {"casualLeaveBonusFee", "事假岗位月奖扣款"}, 
                {"monthNeedPaySalary", "本月应发工资"},
                {"Endowment", "养老(个人)"}, {"Medical", "医疗(个人)"}, {"Unemployment", "失业(个人)"}, {"Housing", "住房公积金(个人)"}, {"SuppleHousing", "补充公积金（个人）"}, 
                {"socialInsuranceFee", "社保代扣小计"}, {"enterpriceYearFondPersonFee", "企业年金个人税前扣款"},
                {"healthInsurancePersonFee", "税优健康险税前扣款"}, {"NeedTaxSalaryAdjust", "税前工资调整"},
                {"needTaxPaySalary", "税前应发工资"}, {"personTax", "个人所得税"},
                {"DecimalChildFee", "独生子女费"}, {"enterpriceYearFondPersonAfterTaxFee", "企业年金个人税后扣款"},
                {"workUnionFee", "个人缴纳工会费"}, {"NoTaxSalaryAdjust", "税后工资调整"}, {"actualPaySalary", "实发工资"}
            };
            return result;
        }

        public object[] getAttrs(Employment em)
        {
            var names = getAttrNames();
            object[] result = new object[names.Count];
            int i = 0;
            foreach (var kv in names)
            {
                var name = kv.Key;
                var displayName = kv.Value;
                if (name == "__id")
                {
                    result[i] = autoId++;
                }
                else if (hasProperty(name))
                {
                    result[i] = getPropertyByString(name);
                }
                else
                {
                    var method = getMethod(name);
                    if (method != null)
                    {
                        try
                        {
                            result[i] = method.Invoke(this, new object[] { em }); //suppose each method need 'Employment instance' argument!
                        }
                        catch (Exception ex)
                        {
                            throw new ArgumentException(String.Format("自动调用员工方法 {0} 失败: {1}", name, ex.GetOriginalException().Message));
                        }
                    }
                    else
                    {
                        result[i] = String.Empty;
                    }
                }
                i++;
            }
            return result;
        }

        public MethodInfo getMethod(String methodName)
        {
            MethodInfo result = null;
            var t = GetType();
            MethodInfo[] methods = t.GetMethods(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public);
            foreach (var method in methods)
            {
                if (method.Name == methodName)
                {
                    result = method;
                }
            }
            return result;
        }

        public Boolean hasProperty(String propName)
        {
            Boolean result = false;
            try
            {
                var property = this.GetType().GetProperty(propName);
                if (property != null)
                    result = true;
            }
            catch
            {
                result = true; //more than one property
            }
            return result;
        }

        public object getPropertyByString(String propName)
        {
            PropertyInfo property;
            try
            {
                property = this.GetType().GetProperty(propName);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(String.Format("无法确认属性 {0}: {2}", propName, ex.GetOriginalException().Message));
            }
            if (property == null)
                throw new ArgumentException(String.Format("找不到属性 {0}", propName));
            var propValue = property.GetValue(this, null);
            if (propValue is DateTime)
                propValue = ((DateTime)propValue).ToShortDateString();
            return propValue;
        }

        public Boolean setPropertyByString(String propName, string propValue)
        {
            Boolean result = true;
            PropertyInfo property;
            try
            {
                property = this.GetType().GetProperty(propName);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(String.Format("无法确认属性 {0}={1}: {2}", propName, propValue, ex.GetOriginalException().Message));
            }
            if (property == null)
                throw new ArgumentException(String.Format("找不到属性 {0}={1}", propName, propValue));
            string propType = property.PropertyType.FullName;
            if (propType == "System.String")
            {
                property.SetValue(this, propValue, null); 
            }
            else if (propType == "System.Int32")
            {
                int val = 0;
                if (string.IsNullOrWhiteSpace(propValue) || int.TryParse(propValue, out val)) 
                {
                    property.SetValue(this, val, null);
                }
                else
                {
                    throw new ArgumentException(String.Format("员工 {0}：无法转换 {1} 为整数", propName, propValue));
                }
            }
            else if (propType == "System.DateTime")
            {
                DateTime val;
                if (DateTime.TryParse(propValue, out val))
                {
                    property.SetValue(this, val, null);
                }
                else
                {
                    int digitVal = 0;
                    if (int.TryParse(propValue, out digitVal))
                    {
                        property.SetValue(this, DateTime.FromOADate(digitVal), null);
                    }
                    else
                    {
                        throw new ArgumentException(String.Format("员工 {0}：无法转换 {1} 为日期", propName, propValue));
                    }
                }
            }
            else if (propType == "System.Decimal")
            {
                propValue = FilterChars(propValue);
                Decimal val = 0m;
                if (string.IsNullOrWhiteSpace(propValue) || Decimal.TryParse(propValue, out val))
                {
                    property.SetValue(this, val, null);
                }
                else
                {
                    throw new ArgumentException(String.Format("员工 {0}：无法转换 {1} 为小数", propName, propValue));
                }
            }
            else
            {
                throw new ArgumentException(String.Format("员工 {0}：不支持将 {1} 转换为 {2}", propName, propValue, propType));
            }
            return result;
        }

        #endregion

        #region StaticMethod
        public static String FilterChars(String str, String[] replaceStrs = null)
        {
            String result = str;
            result = result.Trim();
            String[] specialStrs = { new String('-', result.Length), new String('—', result.Length) };
            foreach (var specialStr in specialStrs)
            {
                if (result == specialStr)
                {
                    result = String.Empty;
                    break;
                }
            }
            if (replaceStrs != null && result.Length > 0)
            {
                foreach (var replaceStr in replaceStrs)
                {
                    result = result.Replace(replaceStr, "");
                }
            }
            return result;
        }

        public static Int32 getWorkDays(DateTime start, DateTime end)
        {
            if (start > end)
            {
                return 0;
            }
            Int32 diffDays = (end - start).Days + 1;
            Int32 weekdays = 0;
            for (DateTime temp = start; temp <= end; temp = temp.AddDays(1))
            {
                if (temp.DayOfWeek == DayOfWeek.Saturday || temp.DayOfWeek == DayOfWeek.Sunday)
                {
                    weekdays++;
                }
            }
            return diffDays - weekdays;
        }

        #endregion

        #region InstanceMethod
        public override string ToString()
        {
            return String.Format("{0}(员工号={1}, ID={2}, 入职日期={3})", name, empId, id, empDate.ToShortDateString());
        }

        public Boolean onBoard(Employment em)
        {
            var onBoard = true;
            if (empDate != null && empDate >= em.SalaryThisDateLastDay)
            {
                onBoard = false;
            }
            return onBoard;
        }

        public Decimal curMonthIncome(Employment em, Decimal baseIncome)
        {
            Decimal result = 0.0m;
            if (baseIncome > 0)
            {
                if (empDate < em.SalaryThisDate)
                {
                    result = baseIncome;
                }
                else if (empDate <= em.SalaryThisDateLastDay)
                {
                    result = (baseIncome / em.SalaryThisMonthDays) * getWorkDays(empDate, em.SalaryThisDateLastDay);
                }
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        public Decimal curMonthIncomeAdjust(Employment em, Decimal baseIncome)
        {
            Decimal result = 0.0m;
            if (baseIncome > 0 && empDate >= em.SalaryLastDate && empDate < em.SalaryThisDate)
            {
                result = (baseIncome / em.SalaryLastMonthDays) * getWorkDays(empDate, em.SalaryLastDateLastDay);
            }
            return Math.Round(result, 2, MidpointRounding.AwayFromZero);
        }

        #endregion
    }
}

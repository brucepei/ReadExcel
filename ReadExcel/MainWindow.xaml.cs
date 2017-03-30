using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Ookii.Dialogs.Wpf;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.ComponentModel;

namespace ReadExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private Employment employment = new Employment();
        private Int32 initYear = 2010;
        private Int32 initMonth = 1;
        public UpdateLogDelegate updateLogDelegate;
        private const int ATTACH_PARENT_PROCESS = -1;
        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int dwProcessId);
        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();
        [DllImport("kernel32.dll")]
        private static extern bool FreeConsole();

        public MainWindow()
        {
            //AllocConsole();
            this.Loaded += new RoutedEventHandler(OnLoaded);
            updateLogDelegate = new UpdateLogDelegate(AppendRichTextToLogBox);
            InitializeComponent();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Logging.ui = this;
            this.DataContext = employment;
            logBox.Document.Blocks.Clear();

            List<String> allYears = new List<String>();
            for (Int32 i = initYear; i < 2050; i++)
            {
                allYears.Add(String.Format("{0}年", i));
            }
            yearComboBox.ItemsSource = allYears;

            List<String> allMonths = new List<String>();
            for (Int32 i = 1; i < 13; i++)
            {
                allMonths.Add(String.Format("{0}月", i));
            }
            monthComboBox.ItemsSource = allMonths;

            var now = DateTime.Now;
            employment.PropertyChanged += employmentDateChanged;
            employment.UseAverageMonthDays = true;
            employment.SalaryThisDate = Convert.ToDateTime(String.Format("{0}-{1}-{2}", now.Year, now.Month, 1));
        }

        public delegate void UpdateLogDelegate(string log, Color color);

        private void AppendRichTextToLogBox(string text, Color color)
        {
            bool focused = this.logBox.IsFocused;
            if (!focused)
            {
                this.logBox.Focus();
            }
            var para = new Paragraph { Margin = new Thickness(0) };
            logBox.Document.Blocks.Add(para);
            Run run = new Run() { Text = text, Foreground = new SolidColorBrush(color) };
            para.Inlines.Add(run);
            logBox.ScrollToEnd();
        }

        private string getFileByDialog(string title, TextBox inputBox, string filter="")
        {
            String initialDir = String.Empty;
            if (inputBox.Text != String.Empty)
            {
                if (System.IO.File.Exists(inputBox.Text))
                {
                    initialDir = System.IO.Path.GetDirectoryName(inputBox.Text);
                }
                else
                {
                    inputBox.Text = String.Empty;
                }
            }
            var dialog = new VistaOpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = title;
            dialog.Filter = filter;
            Logging.logMessage("Show dialog with initial dir: " + initialDir, LogType.DEBUG);
            dialog.InitialDirectory = initialDir;
            dialog.RestoreDirectory = true;
            if ((bool)dialog.ShowDialog())
            {
                inputBox.Text = dialog.FileName;
                return dialog.FileName;
            }
            else
            {
                return string.Empty;
            }
        }

        private string getFileByDialog(string title, string filter = "")
        {
            String initialDir = String.Empty;
            var dialog = new VistaOpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = title;
            dialog.Filter = filter;
            Logging.logMessage("Show dialog with initial dir: " + initialDir, LogType.DEBUG);
            dialog.InitialDirectory = initialDir;
            dialog.RestoreDirectory = true;
            if ((bool)dialog.ShowDialog())
            {
                return dialog.FileName;
            }
            else
            {
                return string.Empty;
            }
        }

        private void ClearLogMenuItem_Click(object sender, RoutedEventArgs e)
        {
            logBox.Document.Blocks.Clear();
        }

        private void GenerateSalaryDetailMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (this.employment.Initialized)
            {
                try
                {
                    this.employment.saveSalaryDetailCSV();
                }
                catch (Exception ex)
                {
                    Logging.logMessage(String.Format("错误：\n{0}", ex.GetOriginalException().Message), LogType.ERROR);
                }
            }
            else
            {
                Logging.logMessage("还未打开员工表，无法生成员工数据", LogType.INFO);
            }
        }

        private void GenerateCompanySummaryMenuItem_Click(Object sender, RoutedEventArgs e)
        {
            if (this.employment.Initialized)
            {
                try
                {
                    this.employment.saveCompanySummaryCSV();
                }
                catch (Exception ex)
                {
                    Logging.logMessage(String.Format("错误：\n{0}", ex.GetOriginalException().Message), LogType.ERROR);
                }
            }
            else
            {
                Logging.logMessage("还未打开员工表，无法生成员工数据", LogType.INFO);
            }
        }

        private void SummaryDataMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (this.employment.Initialized)
            {
                this.employment.summary();
            }
            else
            {
                Logging.logMessage("还未打开员工表，没有员工数据", LogType.INFO);
            }
        }

        private void employmentDateChanged(object sender, PropertyChangedEventArgs e)
        {
            Boolean needUpdate = false;
            switch (e.PropertyName)
            {
                case "SalaryThisDate":
                    yearComboBox.SelectedIndex = employment.SalaryThisDate.Year - initYear;
                    monthComboBox.SelectedIndex = employment.SalaryThisDate.Month - initMonth;
                    if (!employment.UseAverageMonthDays)
                    {
                        needUpdate = true;
                    }
                    break;
                default:
                    break;
            }
            if (needUpdate)
            {
                updateEmploymentLastMonthWorkDays(employment.SalaryLastDate, employment.SalaryLastDateLastDay);
                updateEmploymentThisMonthWorkDays(employment.SalaryThisDate, employment.SalaryThisDateLastDay);
            }
        }

        private void updateEmploymentLastMonthWorkDays(DateTime start, DateTime end)
        {
            lastMonthTextBox.Text = String.Format("上月({0}/{1})工作日:", start.Year, start.Month);
            Int32 weekdays = 0;
            List<Int32> days = new List<Int32>();
            Int32 i = 1;
            for (var day = start; day <= end; day = day.AddDays(1))
            {
                if (day.DayOfWeek == DayOfWeek.Saturday || day.DayOfWeek == DayOfWeek.Sunday)
                {
                    weekdays++;
                }
                days.Add(i++);
            }
            lastWorkDayComboBox.ItemsSource = days;
            lastWorkDayComboBox.SelectedIndex = days.Count - weekdays - 1;
            employment.SalaryLastMonthDays = days.Count - weekdays;
        }

        private void updateEmploymentThisMonthWorkDays(DateTime start, DateTime end)
        {
            thisMonthTextBox.Text = String.Format("当月({0}/{1})工作日:", start.Year, start.Month);
            Int32 weekdays = 0;
            List<Int32> days = new List<Int32>();
            Int32 i = 1;
            for (var day = start; day <= end; day = day.AddDays(1))
            {
                if (day.DayOfWeek == DayOfWeek.Saturday || day.DayOfWeek == DayOfWeek.Sunday)
                {
                    weekdays++;
                }
                days.Add(i++);
            }
            thisWorkDayComboBox.ItemsSource = days;
            thisWorkDayComboBox.SelectedIndex = days.Count - weekdays - 1;
            employment.SalaryThisMonthDays = days.Count - weekdays;
        }

        private void yearComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Int32 thisYear = yearComboBox.SelectedIndex + initYear;
            if (employment.SalaryThisDate.Year != thisYear)
            {
                employment.SalaryThisDate = Convert.ToDateTime(String.Format("{0}-{1}-{2}", thisYear, employment.SalaryThisDate.Month, 1));
            }
            Logging.logMessage(String.Format("生成结果年份修改为 {0} 年!", employment.SalaryThisDate.Year));
        }

        private void monthComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Int32 thisMonth = monthComboBox.SelectedIndex + initMonth;
            if (employment.SalaryThisDate.Month != thisMonth)
            {
                employment.SalaryThisDate = Convert.ToDateTime(String.Format("{0}-{1}-{2}", employment.SalaryThisDate.Year, thisMonth, 1));
            }
            Logging.logMessage(String.Format("生成结果月份修改为 {0} 月 ({1}天/月)!", employment.SalaryThisDate.Month, employment.SalaryThisMonthDays));
        }

        private void EmployeeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            string basicXLSX = getFileByDialog(Properties.Resources.XLSXFileDialogTitle, Properties.Resources.XLSXFileDialogFilter);
            Logging.logMessage(String.Format("选择文件 {0}", basicXLSX), LogType.INFO);
            if (basicXLSX.Length > 0)
            {
                try
                {
                    this.employment.init(basicXLSX);
                    ((MenuItem)sender).IsChecked = true;
                }
                catch (Exception ex)
                {
                    Logging.logMessage(String.Format("错误：\n{0}", ex.GetOriginalException().Message), LogType.ERROR);
                }
            }
        }

        private void OtherTableMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (!this.employment.Initialized)
            {
                Logging.logMessage(String.Format("请先加载基本信息：员工表！"), LogType.WARNING);
            }
            else
            {
                string otherXLSX = getFileByDialog(Properties.Resources.XLSXFileDialogTitle, Properties.Resources.XLSXFileDialogFilter);
                if (otherXLSX.Length > 0)
                {
                    try
                    {
                        var item = (MenuItem)sender;
                        Logging.logMessage(String.Format("选择文件 {0}: {1}", item.Name, otherXLSX), LogType.INFO);
                        this.employment.updateOthers(item.Name, otherXLSX);
                        ((MenuItem)sender).IsChecked = true;
                    }
                    catch (Exception ex)
                    {
                        Logging.logMessage(String.Format("错误：\n{0}", ex.GetOriginalException().Message), LogType.ERROR);
                    }
                }
            }
        }

        private void ClearDataMenuItem_Click(object sender, RoutedEventArgs e)
        {
            this.employment.clear();
            foreach (var item in menu1.Items)
            {
                unCheckMenuItem(item);
            }
        }

        private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var aboutWindow = new AboutWindow();
            aboutWindow.verTextBlock.Text = App.ResourceAssembly.GetName(false).Version.ToString();
            aboutWindow.ShowDialog();
        }

        private void unCheckMenuItem(object menuItem)
        {
            var mi = menuItem as MenuItem;
            if (mi == null) return;
            mi.IsChecked = false;
            foreach (var item in mi.Items)
            {
                unCheckMenuItem(item);
            }
        }

        private void UseAverageRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            lastMonthTextBox.Visibility = Visibility.Hidden;
            lastWorkDayComboBox.Visibility = Visibility.Hidden;
            thisWorkDayComboBox.Visibility = Visibility.Hidden;
            thisMonthTextBox.Visibility = Visibility.Hidden;
            employment.SalaryThisMonthDays = employment.AverageMonthDays;
            employment.SalaryLastMonthDays = employment.AverageMonthDays;
            Logging.logMessage(String.Format("修改为 {0}天/月!", employment.SalaryThisMonthDays));
        }

        private void UseAverageRadioButton_Unchecked(object sender, RoutedEventArgs e)
        {
            lastMonthTextBox.Visibility = Visibility.Visible;
            lastWorkDayComboBox.Visibility = Visibility.Visible;
            thisWorkDayComboBox.Visibility = Visibility.Visible;
            thisMonthTextBox.Visibility = Visibility.Visible;
            updateEmploymentLastMonthWorkDays(employment.SalaryLastDate, employment.SalaryLastDateLastDay);
            updateEmploymentThisMonthWorkDays(employment.SalaryThisDate, employment.SalaryThisDateLastDay);
            Logging.logMessage(String.Format("修改为 {0}天/月!", employment.SalaryThisMonthDays));
        }
    }
}

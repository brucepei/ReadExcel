using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.IO;

namespace ReadExcel
{
    class ExcelHelper
    {
        public DataTable readToDataTable(string fileName, string sheetName)
        {
            string fileType = System.IO.Path.GetExtension(fileName);
            //System.Windows.MessageBox.Show("Test read: " + fileName);
            string connStr = String.Empty;
            if (fileType == ".xls")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\"";
            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            DataTable dataTable = new DataTable();
            try
            {              
                conn = new OleDbConnection(connStr);
                conn.Open();
                da = new OleDbDataAdapter();
                da.SelectCommand = new OleDbCommand(String.Format("Select * FROM [{0}$]", sheetName), conn);
                da.Fill(dataTable);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(String.Format("读取Excel‘{0}’的sheet‘{1}’错误: {2}", fileName, sheetName, ex.GetOriginalException().Message));
            }
            finally
            {                  
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }
            return dataTable;
        }

        public static void writeToCSV(DataTable dt, string fullPath)
        {
            FileInfo fi = new FileInfo(fullPath);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            FileStream fs = new FileStream(fullPath, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            //StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);
            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            string data = "";
            //写出列名称
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                data += dt.Columns[i].ColumnName.ToString();
                if (i < dt.Columns.Count - 1)
                {
                    data += ",";
                }
            }
            sw.WriteLine(data);
            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string str = dt.Rows[i][j].ToString();
                    str = str.Replace("\"", "\"\"");//替换英文冒号 英文冒号需要换成两个冒号
                    if (str.Contains(',') || str.Contains('"')
                        || str.Contains('\r') || str.Contains('\n')) //含逗号 冒号 换行符的需要放到引号中
                    {
                        str = string.Format("\"{0}\"", str);
                    }

                    data += str;
                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();
            Logging.logMessage(String.Format("CSV文件 {0} 保存成功！", fullPath), LogType.INFO);
        } 
    }
}

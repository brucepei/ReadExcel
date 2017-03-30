using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ReadExcel
{
    class ExcelTable
    {
        private String fileName;
        private String sheetName;

        public String SheetName
        {
            get { return sheetName; }
        }
        private Int32[] titleRows;
        private DataTable dataTable;
        private Dictionary<String, String> nameTitles;

        public Dictionary<String, String> NameTitles
        {
            get { return nameTitles; }
        }

        public Int32 ContentRow
        {
            get { return titleRows.Max() + 1; }
        }
        public DataTable Table
        {
            get { return dataTable; }
        }

        public ExcelTable(String fileName, String sheetName, Int32[] titleRows, Dictionary<String, String> nameTitles)
        {
            this.fileName = fileName;
            this.sheetName = sheetName;
            this.titleRows = titleRows;
            this.nameTitles = nameTitles;
            ExcelHelper reader = new ExcelHelper();
            var dt = reader.readToDataTable(fileName, sheetName);
            Logging.logMessage(String.Format("成功打开Excel文件:{0}, 表 {1}!", fileName, sheetName), LogType.INFO);
            this.dataTable = dt;
        }

        public Dictionary<String, Int32> getNameCols()
        {
            Dictionary<String, Int32> nameCols = null;
            if (dataTable == null)
            {
                throw new ArgumentException(Properties.Resources.ErrorNullDataTable);
            }
            else
            {
                Int32 titleRowsMax = titleRows.Max();
                if (dataTable.Rows.Count > titleRowsMax && dataTable.Columns.Count >= nameTitles.Count) //it is a valid table even it only contains the title rows
                {
                    nameCols = new Dictionary<String, Int32>();
                    foreach (var kv in nameTitles)
                    {
                        String name = kv.Key;
                        String title = kv.Value;
                        Boolean found = false;
                        for (int i = 0; i < dataTable.Columns.Count; i++)
                        {
                            foreach (int titleRow in titleRows)
                            {
                                if (dataTable.Rows[titleRow][i].ToString() == title)
                                {
                                    nameCols[name] = i;
                                    found = true;
                                    goto FoundColIndex;
                                }
                            }
                        }
                    FoundColIndex:
                        if (!found)
                        {
                            throw new ArgumentException(String.Format("Cannot find {0} above Row {1} in Sheet {2}!", title, titleRowsMax + 1, sheetName));
                        }
                    }
                }
                else
                {
                    throw new ArgumentException(String.Format("DataTable is too small(row={0}, column={1}) to get the tilte!", dataTable.Rows.Count, dataTable.Columns.Count));
                }
            }
            return nameCols;
        }
    }
}

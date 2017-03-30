using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcel
{
    interface IExcelTable
    {
        Dictionary<string, int> UpdateStatus { get; }
        void updateEmployment(Employment em);
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excelt = AdvancedExcelFunctions.ExcelFunctions;

namespace Testing
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename = "C:\\Users\\sudhe\\OneDrive\\Documents\\ExcelMetabotTesting\\Exceltesting.xlsx";
            string filename1 = "C:\\Users\\sudhe\\OneDrive\\Documents\\ExcelMetabotTesting\\Exceltesting1.xlsx";
            string sheetname = "";
            string cellValue = "Hello Testing!";
            string cellName = "C2";
            string result = "";
            excelt excel = new excelt();
            result = excel.OpenExcel(filename,sheetname);
            result = excel.ChangeSourceDataPivotTable(filename, "pivot2", "Sudheer", "A1:G87", "PivotTable2");
            Console.ReadLine();
        }
    }
}

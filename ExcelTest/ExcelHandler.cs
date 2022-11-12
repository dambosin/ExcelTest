using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelTest
{
    public static class ExcelHandler
    {
        public static Dictionary<String, ExcelApplication> Applications { get; set; } = new();

        public static void Open(string name, string source)
        {
            if (Applications.ContainsKey(name))
            {
                Console.WriteLine($"File with name:{name} already exists");
                return;
            }
            Applications.Add(name, new ExcelApplication(source));
        }

        public static void Show(string name)
        {
            if (Applications.ContainsKey(name))
            {
                Applications[name].SetVisibility(true);
            }
            else
            {
                Console.WriteLine($"Application with name:{name} doesn't exist");
            }
            
        }
        /*static void Displayl(IEnumerable<Account> accounts)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;

            excelApp.Workbooks.Add();
            Excel._Worksheet worksheet = (Excel.Worksheet)excelApp.ActiveSheet;


            worksheet.Cells[1, "A"] = "ID Number";
            worksheet.Cells[1, "B"] = "Current Balance";

            var row = 2;
            foreach (var account in accounts)
            {
                worksheet.Cells[row, "A"] = account.ID;
                worksheet.Cells[row, "B"] = account.Balance;
                row++;
            }

            ((Excel.Range)worksheet.Columns[1]).AutoFit();
            ((Excel.Range)worksheet.Columns[2]).AutoFit();
        }*/

    }
}

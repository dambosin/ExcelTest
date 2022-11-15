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
        public static Dictionary<String, PostExcelApplication> Applications { get; set; } = new();
        delegate void Execute();
        public static void Open<T>(string name, string source)where T : IExcelApplication, new()
        {
            if (!CheckKey(name, () => AppAlreadyExistAlert(name)))
            {
             //   Applications.Add(name, new T(source));
            }
        }

        public static void Show(string name)
        {
            if (!CheckKey(name, () => Applications[name].SetVisibility(true)))
            {
                AppNotExistAlert(name);
            }   
        }

        public static void Hide(string name)
        {
            if (!CheckKey(name, () => Applications[name].SetVisibility(false)))
            {
                AppNotExistAlert(name);
            }
        }
        public static void Save(string name, string path)
        {
            if (!CheckKey(name, () => Applications[name].Save(path)))
            {
                AppNotExistAlert(name);
            }
        }

        public static void Configure(string name, PostModel model)
        {
            if (!CheckKey(name, () => Applications[name].Configure(model)))
            {
                AppNotExistAlert(name);
            }
        }

        private static bool CheckKey(string name, Execute execute)
        {
            if (Applications.ContainsKey(name))
            {
                execute();
                return true;
            }
            return false;
        }
        private static void AppAlreadyExistAlert(string name)
        {
            Console.WriteLine($"File with name:{name} already exists");
        }

        private static void AppNotExistAlert(string name)
        {
            Console.WriteLine($"Application with name:{name} doesn't exist");
        }

        public static void Close(string name)
        {
            if(!CheckKey(name, () => Applications[name].Close()))
            {
                AppNotExistAlert(name);
            }
            else
            {
                Applications.Remove(name);
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

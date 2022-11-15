using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelTest
{
    public class BaseExcelApplication : IExcelApplication
    {
        public Application App { get; set; } = new();
        public Workbook Workbook { get; set; }
        public Worksheet Worksheet { get; set; }

        public BaseExcelApplication(string source)
        {
            Workbook = App.Workbooks.Open(source);
            Worksheet = App.ActiveSheet;
        }

        public void SetVisibility(bool isVisible) 
        {
            App.Visible = isVisible;
        }

        public Excel.Range GetData()
        {
            return Worksheet.UsedRange;
        }

        public void Save(string path)
        {
            Worksheet.SaveAs(path);
        }
        public void Close()
        {
            Workbook.Close(0);
            App.Quit();
        }
    }
}

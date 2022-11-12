using Microsoft.Office.Interop.Excel;
namespace ExcelTest
{
    public class BaseExcelApplication : IExcelApplication
    {
        public Application App { get; set; } = new();
        public Workbook Workbook { get; set; }

        public BaseExcelApplication(string source)
        {
            Workbook = App.Workbooks.Open(source);
        }

        public void SetVisibility(bool isVisible)
        {
            App.Visible = isVisible;
        }
    }
}

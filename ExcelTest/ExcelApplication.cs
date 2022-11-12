using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    public class ExcelApplication
    {
        private Excel.Application app = new();
        private Excel.Workbook workbook;
        public ExcelApplication(string source)
        {
             workbook = app.Workbooks.Open(source);
        }

        public void SetVisibility(bool visible)
        {
            app.Visible = visible;
        }


    }
}
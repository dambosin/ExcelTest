using ExcelTest.Interfaces;
using ExcelTest.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    public class ExcelApplication : IExcelApplication
    {
        public Excel.Application App { get; set; } = new();
        public Excel.Workbook? Workbook { get; set; }
        public Excel.Worksheet? Worksheet { get; set; }

        public ExcelApplication() {}
        public ExcelApplication(string path)
        {
            Open(path);
        }
        public virtual void Configure(BaseModel model)
        {
            throw new NotImplementedException();
        }

        public void SetVisibility(bool isVisibe)
        {
            App.Visible = isVisibe;
        }

        public void Close()
        {
            if (Workbook == null) throw new NullReferenceException("Workbook is null");
            Workbook.Close();
            App.Quit();
        }

        public Excel.Range GetData()
        {
            if (Worksheet == null) throw new NullReferenceException("Worksheet is null");
            return Worksheet.UsedRange;
        }

        public void Save(string path)
        {
            if (Worksheet == null) throw new NullReferenceException("Worksheet is null");
            Worksheet.SaveAs(path);
        }

        public void Open(string path)
        {
            Workbook = App.Workbooks.Open(path);
            Worksheet = App.ActiveSheet;
        }

    }
}

using Excel = Microsoft.Office.Interop.Excel;
using ExcelTest.Models;

namespace ExcelTest.Interfaces
{
    public interface IExcelApplication
    {
        Excel.Application App { get; set; }
        Excel.Workbook? Workbook { get; set; }
        Excel.Worksheet? Worksheet { get; set; }

        public void Configure(BaseModel model);

        public void SetVisibility(bool isVisibe);

        public void Close();

        public Excel.Range GetData();

        public void Save(string path);

        public void Open(string path);
    }
}
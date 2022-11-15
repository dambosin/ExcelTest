using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    public interface IExcelApplication
    {
        Excel.Application App { get; set; }
        Excel.Workbook Workbook { get; set; }

        public void SetVisibility(bool isVisibe);

        public void Close();

        public Excel.Range GetData();

        public void Save(string path);

    }
}
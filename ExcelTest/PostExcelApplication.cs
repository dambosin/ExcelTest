using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    public class PostExcelApplication : BaseExcelApplication, IExcelApplication
    {
        public PostExcelApplication(string source) : base(source)
        {
        }

        public void Configure(PostModel model)
        {
            _Worksheet worksheet = (Worksheet)App.ActiveSheet;
            worksheet.Cells[20, "C"] = $"{model.Id} {model.Size.Width}x{model.Size.Height}";
            worksheet.Cells[3, "F"] = model.Price;
            worksheet.Cells[4, "G"] = model.Price * 100 % 100;
            worksheet.Cells[4, "D"] = model.PriceInText;
            worksheet.Cells[12, "E"] = model.Name;
            worksheet.Cells[15, "E"] = model.Adress;
            worksheet.Cells[20, "F"] = model.Phone;
            worksheet.SaveAs($"D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\PostDocs\\{model.Id.ToString()}");
            Workbook.Close(0);
            App.Quit();

        }
    }
}
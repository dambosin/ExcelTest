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
            Worksheet = App.ActiveSheet;
            Worksheet.Cells[20, "C"] = $"{model.Id} {model.Size.Width}x{model.Size.Height}";
            Worksheet.Cells[3, "F"] = model.Price;
            Worksheet.Cells[4, "G"] = model.Price * 100 % 100;
            Worksheet.Cells[4, "D"] = model.PriceInText;
            Worksheet.Cells[12, "E"] = model.Name;
            Worksheet.Cells[15, "E"] = model.Adress;
            Worksheet.Cells[20, "F"] = model.Phone;
        }
    }
}
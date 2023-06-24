using ExcelTest.Interfaces;
using ExcelTest.Models;

namespace ExcelTest
{
    public class PostExcelApplication : ExcelApplication
    {
        public PostExcelApplication() : base() { }
        public PostExcelApplication(string source) : base(source) { }

        public override void Configure(BaseModel model)
        {
            if (model is not PostModel postModel)
            {
                throw new MemberAccessException("Wrong model used");
            }
            Worksheet = App.ActiveSheet;
            Worksheet.Cells[20, "C"] = $"{postModel.Id} {postModel.Size}";
            Worksheet.Cells[3, "F"] = postModel.Price;
            Worksheet.Cells[4, "G"] = postModel.Price * 100 % 100;
            Worksheet.Cells[4, "D"] = postModel.PriceInText;
            Worksheet.Cells[12, "E"] = postModel.Name;
            Worksheet.Cells[15, "E"] = postModel.Adress;
            Worksheet.Cells[20, "F"] = postModel.Phone;
            if (postModel.IsCarefully) {
                Worksheet.Cells[20, "A"].Font.Bold = true;
                Worksheet.Cells[20, "A"] = "X ОСТОРОЖНО";
            }
            else
            {
                Worksheet.Cells[20, "A"].Font.Bold = false;
                Worksheet.Cells[20, "A"] = "□ ОСТОРОЖНО";
            }
        }
    }
}
using ExcelTest;
using ExcelTest.Models;

try
{
    ExcelHandler.Open<ExcelApplication>("Data", "D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\Data.xlsx");
    var range = ExcelHandler.GetData("Data");

    List<PostModel> models = Parser.Parse(range);
    ExcelHandler.Close("Data");

    ExcelHandler.Open<PostExcelApplication>("Post", "D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\PostExample.xlsx");
    foreach (var model in models)
    {
        ExcelHandler.Configure("Post", model);
        ExcelHandler.Save("Post", $"D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\PostDocs\\{model.Id}");
    }
    ExcelHandler.Close("Post");
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
/*ExcelApplication dataExcel = new("D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\Data.xlsx");
var range = dataExcel.GetData();

List<PostModel> models = Parser.Parse(range);
dataExcel.Close();

PostExcelApplication post = new("D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\PostExample.xlsx");
foreach (var model in models)
{
    post.Configure(model);
    post.Save($"D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\PostDocs\\{model.Id}");
}
post.Close();*/

Console.WriteLine("WTF is hapenning?");
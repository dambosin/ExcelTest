using ExcelTest;

BaseExcelApplication dataExcel = new("D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\Data.xlsx");
var range = dataExcel.GetData();

List<PostModel> models = Parser.Parse(range);
dataExcel.Close();

PostExcelApplication post = new("D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\PostExample.xlsx");
foreach (var model in models)
{
    post.Configure(model);
    post.Save($"D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\PostDocs\\{model.Id}");
}
post.Close();

Console.WriteLine("WTF is hapenning?");
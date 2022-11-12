using ExcelTest;
using System.Drawing;

ExcelHandler.Open("Post", "D:\\Repos\\ExcelTest\\ExcelTest\\bin\\Debug\\net6.0\\PostExample.xlsx");

var model = new PostModel
{
    Id = 1234,
    Size = new Rectangle(0, 0, 20, 20),
    Phone = "447868522",
    Name = "Голубцова Ксения Валерьевна",
    Adress = "г.Могилев, ул. Залуцкого д3,кв 412., 212040",
    Price = 135.15
};
model.PriceInText = PriceConverter.Convert((int)model.Price);

ExcelHandler.Configure("Post", model);

ExcelHandler.Show("Post");



Console.WriteLine("WTF is hapenning?");
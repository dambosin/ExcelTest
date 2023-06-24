using System.Drawing;

namespace ExcelTest.Models
{
    public class PostModel : BaseModel
    {
        public string Size { get; set; }
        public string Name { get; set; }
        public string Adress { get; set; }
        public string Phone { get; set; }
        public double Price { get; set; }
        public string PriceInText { get; set; }
        public bool IsCarefully { get; set; } = false;

    }
}

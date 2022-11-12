using System.Drawing;

namespace ExcelTest
{
    public class PostModel
    {
        public int Id { get; set; }
        public Rectangle Size { get; set; }
        public string Name { get; set; }
        public string Adress { get; set; } 
        public string Phone { get; set; }
        public double Price { get; set; }
        public string PriceInText { get; set; }

    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest
{
    public static class Parser
    {
        public static List<PostModel> Parse(Microsoft.Office.Interop.Excel.Range range)
        {
            List<PostModel> list = new List<PostModel>();
            for(int i = 1; i <= range.Columns.Count; i++)
            {
                if (range.Cells[i, 1].Value2 != null)
                {
                    PostModel post = new();
                    post.Id = (int)(range.Cells[i, 1].Value2);
                    int width;
                    int height;
                    int price;
                    string t = range.Cells[i, 3].Value2.ToString();
                    int.TryParse(t.AsSpan(0, t.IndexOf('х')), out width);
                    int.TryParse(t.AsSpan(t.IndexOf('х') + 1, t.IndexOf('(') - t.IndexOf('х') - 2), out height);
                    int.TryParse(t.AsSpan(t.IndexOf('(') + 1, t.IndexOf(')') - t.IndexOf('(') - 1), out price);
                    post.Size = new Rectangle(0, 0, width, height);
                    post.Price = price;
                    post.PriceInText = PriceConverter.Convert(price);
                    string phone = range.Cells[i, 7].Value2.ToString();
                    string phone2 = "";
                    foreach (var letter in phone)
                    {
                        if (IsDigit(letter))
                        {
                            phone2 += letter;
                        }
                    }
                    post.Phone = phone2;
                    t = range.Cells[i, 8].Value2;
                    while (t.Contains('\n'))
                    {
                        t = t.Remove(t.IndexOf('\n'), 1);
                    }
                    post.Name = t;
                    t = range.Cells[i, 9].Value2;
                    while (t.Contains('\n'))
                    {
                        t = t.Remove(t.IndexOf('\n'), 1);
                    }
                    post.Adress = t;
                    list.Add(post);
                }
            }

            return list;
        }


        public static bool IsDigit(char x)
        {
            bool result = false;
            for (int i = 0; i < 10; i++)
            {
                if (x == i.ToString()[0])
                {
                    result = true;
                }
            }
            return result;

        }
    }
}

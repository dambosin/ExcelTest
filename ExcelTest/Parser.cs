using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTest.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    public static class Parser
    {
        public static List<PostModel> Parse(Microsoft.Office.Interop.Excel.Range range)
        {
            List<PostModel> list = new();
            for(int i = 1; i <= range.Columns.Count; i++)
            {
                if (range.Cells[i, 1].Value2 != null)
                {
                    list.Add(ParseRow(range, i));
                }
            }
            return list;
        }

        private static PostModel ParseRow(Microsoft.Office.Interop.Excel.Range range, int i)
        {
            PostModel post = new()
            {
                Id = (int)(range.Cells[i, 1].Value2)
            };
            string temp = range.Cells[i, 3].Value2;
            _ = double.TryParse(temp.AsSpan(temp.IndexOf('(') + 1, temp.IndexOf(')') - temp.IndexOf('(') - 1), out double price);
            post.Price = price;
            post.PriceInText = PriceConverter.Convert(price);
            post.Size = temp.AsSpan(0, temp.IndexOf('(')).ToString().Trim();
            post.Phone = range.Cells[i, 7].Value2.ToString();
            post.Name = range.Cells[i, 8].Value2; ;
            post.Adress = range.Cells[i, 9].Value2;
            post.Size = post.Size.Replace('х', 'x');
            _ = int.TryParse(post.Size.AsSpan(0, post.Size.IndexOf('x')), out int size1);
            int size2;
            if(post.Size.IndexOf('x') == post.Size.LastIndexOf('x'))
            _ = int.TryParse(post.Size.AsSpan(post.Size.IndexOf('x') + 1,post.Size.Length - post.Size.IndexOf('x') - 1), out size2);
            else _ = int.TryParse(post.Size.AsSpan(post.Size.IndexOf('x') + 1, post.Size.LastIndexOf('x') - post.Size.IndexOf('x') - 1), out size2);
            if (size1 >= 70 || size2 >= 70)
            {
                post.IsCarefully = true;
            }
            return post;
        }


        public static bool IsDigit(char x)
        {
            string digits = "0123456789";
            if (digits.Contains(x)) return true;
            return false;

        }
    }
}

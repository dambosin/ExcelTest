using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest
{
    public static class PriceConverter
    {

        public static string Convert(int price)
        {
            string result = "";
            switch (price / 100)
            {
                case 9: 
                    result += "Девятьсот ";
                    break;
                case 8: 
                    result += "Восемьсот ";
                    break;
                case 7:
                    result += "Семьсот ";
                    break;
                case 6:
                    result += "Шестьсот ";
                    break;
                case 5:
                    result += "Пятьсот ";
                    break;
                case 4:
                    result += "Четыреста ";
                    break;
                case 3:
                    result += "Триста ";
                    break;
                case 2:
                    result += "Двести ";
                    break;
                case 1:
                    result += "Сто ";
                    break;
                default:
                    break;
            }
            price %= 100;
            switch (price / 10)
            {
                case 9:
                    result += "Девяносто ";
                    break;
                case 8:
                    result += "Восемьдесят ";
                    break;
                case 7:
                    result += "Семьдесят ";
                    break;
                case 6:
                    result += "Шестьдесят ";
                    break;
                case 5:
                    result += "Пятьдесят ";
                    break;
                case 4:
                    result += "Сорок ";
                    break;
                case 3:
                    result += "Тридцать ";
                    break;
                case 2:
                    result += "Двадцать ";
                    break;
                case 1:
                    switch (price%10)
                    {
                        case 9:
                            result += "Девятнадцать ";
                            break;
                        case 8:
                            result += "Восемнадцать ";
                            break;
                        case 7:
                            result += "Семнадцать ";
                            break;
                        case 6:
                            result += "Шестнадцать ";
                            break;
                        case 5:
                            result += "Пятнадцать ";
                            break;
                        case 4:
                            result += "Четырнадцать ";
                            break;
                        case 3:
                            result += "Тринадцать ";
                            break;
                        case 2:
                            result += "Двенадцать ";
                            break;
                        case 1:
                            result += "Одиннадцать ";
                            break;
                        default:
                            result += "Десять ";
                            break;
                    }
                    break;
                default:
                    break;
            }
            if(price / 10 == 1)
            {
                return result;
            }
            switch (price % 10)
            {
                case 9:
                    result += "Девять ";
                    break;
                case 8:
                    result += "Восемь ";
                    break;
                case 7:
                    result += "Семь ";
                    break;
                case 6:
                    result += "Шесть ";
                    break;
                case 5:
                    result += "Пять ";
                    break;
                case 4:
                    result += "Четыре ";
                    break;
                case 3:
                    result += "Три ";
                    break;
                case 2:
                    result += "Два ";
                    break;
                case 1:
                    result += "Один ";
                    break;
                default:
                    if(result == "") 
                    { 
                        result += "Ноль "; 
                    }
                    break;
            }
            return result;
        }

    }
}

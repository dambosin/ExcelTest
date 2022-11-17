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
        public static string Convert(double price)
        {
            string result = "";
            result += (int)(price / 100) switch
            {
                9 => "Девятьсот ",
                8 => "Восемьсот ",
                7 => "Семьсот ",
                6 => "Шестьсот ",
                5 => "Пятьсот ",
                4 => "Четыреста ",
                3 => "Триста ",
                2 => "Двести ",
                1 => "Сто ",
                _ => ""
            };
            price %= 100;
            result += (int)(price / 10) switch
            {
                9 => "Девяносто ",
                8 => "Восемьдесят ",
                7 => "Семьдесят ",
                6 => "Шестьдесят ",
                5 => "Пятьдесят ",
                4 => "Сорок ",
                3 => "Тридцать ",
                2 => "Двадцать ",
                1 => (int)(price % 10) switch
                {
                    9 => "Девятнадцать ",
                    8 => "Восемнадцать ",
                    7 => "Семнадцать ",
                    6 => "Шестнадцать ",
                    5 => "Пятнадцать ",
                    4 => "Четырнадцать ",
                    3 => "Тринадцать ",
                    2 => "Двенадцать ",
                    1 => "Одиннадцать ",
                    _ => "Десять ",
                },
                _ => ""
            };
            if(price / 10 == 1) return result;
            result += (int)(price % 10) switch
            {
                9 => "Девять ",
                8 => "Восемь ",
                7 => "Семь ",
                6 => "Шесть ",
                5 => "Пять ",
                4 => "Четыре ",
                3 => "Три ",
                2 => "Два ",
                1 => "Один ",
                _ => result == "" ? result += "Ноль " : ""
            };
            return result;
        }

    }
}

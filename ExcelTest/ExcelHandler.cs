using ExcelTest.Exceptions;
using ExcelTest.Interfaces;
using ExcelTest.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelTest
{
    public static class ExcelHandler
    {
        public static Dictionary<String, ExcelApplication> Applications { get; set; } = new();
        public static void Open<T>(string name, string path)where T : ExcelApplication, new()
        {
            if (CheckKey(name)) throw new AlredyExistException($"Excel application {name} already exist");
            Applications.Add(name, new T());
            Applications[name].Open(path);
        }

        public static void Show(string name)
        {
            if (!CheckKey(name)) throw new NotExistException($"Excel application {name} does not exist");
            Applications[name].SetVisibility(true);
        }

        public static void Hide(string name)
        {
            if (!CheckKey(name)) throw new NotExistException($"Excel application {name} does not exist");
            Applications[name].SetVisibility(false);
        }
        public static void Save(string name, string path)
        {
            if (!CheckKey(name)) throw new NotExistException($"Excel application {name} does not exist");
            Applications[name].Save(path);
        }

        public static void Configure(string name, BaseModel model)
        {
            if (!CheckKey(name)) throw new NotExistException($"Excel application {name} does not exist");
            Applications[name].Configure(model);
        }
        public static void Close(string name)
        {
            if (!CheckKey(name)) throw new NotExistException($"Excel application {name} does not exist");
            Applications[name].Close();
            Applications.Remove(name);
        }

        public static Excel.Range GetData(string name)
        {
            if (!CheckKey(name)) throw new NotExistException($"Excel application {name} does not exist");
            return Applications[name].GetData();
        }

        private static bool CheckKey(string name)
        {
            if (Applications.ContainsKey(name)) return true;
            return false;
        }

    }
}

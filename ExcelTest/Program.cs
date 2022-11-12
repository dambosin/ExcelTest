
using ExcelTest;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


var bankAccounts = new List<Account> {
    new Account
    {
        ID =129421,
        Balance = 124.315
    },
    new Account
    {
        ID = 138508,
        Balance = 87135.185
    }
};
static void DisplayInExcel(IEnumerable<Account> accounts) 
{
    var excelApp = new Excel.Application();
    excelApp.Visible = true;

    excelApp.Workbooks.Add();
    Excel._Worksheet worksheet = (Excel.Worksheet)excelApp.ActiveSheet;


    worksheet.Cells[1, "A"] = "ID Number";
    worksheet.Cells[1, "B"] = "Current Balance";

    var row = 2;
    foreach (var account in accounts)
    {
        worksheet.Cells[row, "A"] = account.ID;
        worksheet.Cells[row, "B"] = account.Balance;
        row++;
    }

    ((Excel.Range)worksheet.Columns[1]).AutoFit();
    ((Excel.Range)worksheet.Columns[2]).AutoFit();
}



DisplayInExcel(bankAccounts);

Console.WriteLine("WTF is hapenning?");
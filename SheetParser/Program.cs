// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
try
{
    var package = new ExcelPackage(new FileInfo("C:\\Users\\gauravgoel\\source\\prsnl\\SheetParser\\SheetParser\\stmt.xlsx"));
    var sheet = package.Workbook.Worksheets[0];
    int rowStart = sheet.Dimension.Start.Row;
    int rowEnd = sheet.Dimension.End.Row;
    string cellRange = rowStart.ToString() + ":" + rowEnd.ToString();
    var searchCell = from cell in sheet.Cells[cellRange] 
                     where cell.Value.ToString() == "Txn Date"
                     select cell.Start.Row;

    int rowNum = searchCell.First();
    for(int start = rowNum+1; start < rowEnd - 1; start++)
    {
        for(int col = 1; col <=7; col++)
        {
            Console.Write(sheet.Cells[start, col].Text);
            Console.Write("  ");
        }
        Console.WriteLine();
    }
} catch (Exception ex)
{
    Console.WriteLine(ex.Message);
    throw;
}
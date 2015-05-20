using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace helloExcel
{
    class Program
    {
        private static void Main(string[] args)
        {

            using (SpreadsheetDocument xl = SpreadsheetDocument.Open("test.xlsx", true))
            {
                foreach (WorksheetPart wsp in xl.WorkbookPart.WorksheetParts)
                {
                    foreach (TableDefinitionPart tdp in wsp.TableDefinitionParts)
                    {
                        Console.WriteLine("test");
                        // for example
                        // tdp.Table.AutoFilter = new AutoFilter() { Reference = "B2:D3" };
                    }
                }
            }








            List<Import> list = new List<Import>();
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open("test.xlsx", false))
            {
                var workbookPart = document.WorkbookPart;
                var tmp = document.WorkbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();

                var iter = GetStrings(workbookPart, tmp);
                bool startParse = false;

                foreach (var item in iter)
                {
                    if (item == "Комнатность")
                    {
                        startParse = true;
                        continue;
                    }
                    if (startParse)
                    {
                       list.Add(new Import()); 
                    }
                }



            }
        }
        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }

        static IEnumerable<string> GetStrings(WorkbookPart wk,SheetData sheetData)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    string cellValue = string.Empty;

                    if (cell.DataType != null)
                    {
                        if (cell.DataType == CellValues.SharedString)
                        {
                            int id = -1;

                            if (Int32.TryParse(cell.InnerText, out id))
                            {
                                SharedStringItem item = GetSharedStringItemById(wk, id);

                                if (item.Text != null)
                                {
                                    cellValue = item.Text.Text;
                                }
                                else if (item.InnerText != null)
                                {
                                    cellValue = item.InnerText;
                                }
                                else if (item.InnerXml != null)
                                {
                                    cellValue = item.InnerXml;
                                }
                            }
                        }
                        yield return cellValue;
                       // Console.WriteLine(cellValue);
                    }
                    else
                    {
                        yield return cell.InnerText;
                       // Console.WriteLine(cell.InnerText);
                    }
                }
            }
        } 
    }

}

class Import
{
    public string ObjectBilder { get; set; }
    public string K { get; set; }
    public Status Status { get; set; }
    public double Area { get; set; }
    public decimal PriceMeter { get; set; }
    public decimal PriceApartment { get; set; }
    public int CountDayArmor { get; set; }
    public DateTime DayArmor { get; set; }
    public int Access { get; set; }
    public int Floor { get; set; }
    public int LevelRoom { get; set; }
    public string Room { get; set; }
}

public enum Status : byte
{
    Free, Reservations
}

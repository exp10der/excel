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
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open("test.xlsx", false))
            {
                var workbookPart = document.WorkbookPart;
                var tmp = document.WorkbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();

                foreach (var row in tmp.Elements<Row>())
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
                                    SharedStringItem item = GetSharedStringItemById(workbookPart, id);

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
                            Console.WriteLine(cellValue);
                        } else 
                        {
                            Console.WriteLine(cell.InnerText);
                        }
                    } 
                }
            }
    }
        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
    }
}

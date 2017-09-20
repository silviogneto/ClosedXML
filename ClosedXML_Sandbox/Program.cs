using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ClosedXML_Sandbox
{
    public class Pastry
    {
        public Pastry(string name, int amount, string month)
        {
            Month = month;
            Name = name;
            NumberOfOrders = amount;
        }

        public string Name { get; set; }
        public int NumberOfOrders { get; set; }
        public string Month { get; set; }
    }


    internal class Program
    {
        public static void GeneratePivotTable()
        {
            var pastries = new List<Pastry>
            {
              new Pastry("Croissant", 150, "Apr"),
              new Pastry("Croissant", 250, "May"),
              new Pastry("Croissant", 134, "June"),
              new Pastry("Doughnut", 250, "Apr"),
              new Pastry("Doughnut", 225, "May"),
              new Pastry("Doughnut", 210, "June"),
              new Pastry("Bearclaw", 134, "Apr"),
              new Pastry("Bearclaw", 184, "May"),
              new Pastry("Bearclaw", 124, "June"),
              new Pastry("Danish", 394, "Apr"),
              new Pastry("Danish", 190, "May"),
              new Pastry("Danish", 221, "June"),
              new Pastry("Scone", 135, "Apr"),
              new Pastry("Scone", 122, "May"),
              new Pastry("Scone", 243, "June")
            };

            var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("PastrySalesData");
            IXLRange rRange = sheet.Cell(1, 1).InsertData(pastries);
            IXLTable tTable = rRange.CreateTable();
            tTable.Theme = XLTableTheme.TableStyleDark1;

            var range = tTable.DataRange;
            var header = sheet.Range(1, 1, 1, 3);
            var dataRange = sheet.Range(header.FirstCell(), range.LastCell());

            var ptSheet = workbook.Worksheets.Add("PivotTable");
            var pt = ptSheet.PivotTables.AddNew("PivotTable", ptSheet.Cell(1, 1), dataRange);
            pt.RowLabels.Add("Name");
            pt.ColumnLabels.Add("Month");
            pt.Values.Add("NumberOfOrders");

            workbook.SaveAs(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.xlsx"));
        }

        private static void Main(string[] args)
        {
            GeneratePivotTable();

            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.ConditionalFormatting;
using GemBox.Spreadsheet.PivotTables;
using GemBox.Spreadsheet.Tables;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using Spire.Xls;
using Spire.Xls.Collections;

namespace AutoFilterExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            ActivateAutoFilter(@"C:\temp\Email_Gonderim.xlsx");

            //@"C:\temp\Email_Gonderim.xlsx"
        }
        static void ActivateAutoFilter(string path)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(path);
            Worksheet sheet = workbook.Worksheets[0];
            AutoFiltersCollection filters = sheet.AutoFilters;
            filters.Range = sheet.Range[1, 6, sheet.LastRow, 6];
            filters.AddFilter(0, "FİLO YÖNETİCİSİ");
            filters.AddFilter(0, "TPC SORUMLUSU");
            filters.Filter();
            workbook.SaveToFile(@"C:\temp\output.xlsx", ExcelVersion.Version2010);

        }

    }
}

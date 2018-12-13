using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using Spire.Xls.Collections;

namespace AutoFilterYourExcel
{
    public class Class1
    {
        public static void ActivateAutoFilter(string path, string dest)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(path);
            Worksheet sheet = workbook.Worksheets[0];


            workbook.DataSorter.SortColumns.Add(5, OrderBy.Descending);
            workbook.DataSorter.Sort(sheet["A1:Z10000"]);

            AutoFiltersCollection filters = sheet.AutoFilters;
            filters.Range = sheet.Range[1, 6, sheet.LastRow, 6];
            filters.AddFilter(0, "FİLO YÖNETİCİSİ");
            filters.AddFilter(0, "TPC SORUMLUSU");
            filters.Filter();

           

            workbook.SaveToFile(dest, ExcelVersion.Version2010);





        }
    }
}

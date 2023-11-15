using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Excel.Range;
namespace ExcelConsole
{
    public class Program
    {
        static void Main(string[] args)
        {
             Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = true;           

            string fileName = "G:\\Untitled spreadsheet.xlsx";
            Workbook workbook = _excelApp.Workbooks.Open(fileName);

            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            Range excelRange = worksheet.UsedRange;
            object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

            for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
            {
                for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                {
                    Console.Write(valueArray[row, col]?.ToString() + "\t");
                }
                Console.WriteLine();
            }
            Console.ReadLine();
        }
    }
}



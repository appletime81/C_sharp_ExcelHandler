using System;
using OfficeOpenXml;
using System.IO;

namespace ExcelHandler
{
    public static class ReadExcel
    {
        public static void ReadExcelFile(string path)
        {
            Console.WriteLine("-------------------------------------------------------------");
            Console.WriteLine(path);
            Console.WriteLine("-------------------------------------------------------------");
            var file = new FileInfo(path);

            if (!file.Exists)
            {
                Console.WriteLine($"Error: File not found at {path}");
                return;
            }

            using var package = new ExcelPackage(file);

            // 取得第一個工作表
            var worksheet = package.Workbook.Worksheets[0];

            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // 逐列讀取
            for (var row = 1; row <= rowCount; row++)
            {
                // 逐欄讀取
                for (var col = 1; col <= colCount; col++)
                {
                    // 讀取儲存格的值
                    var cellValue = worksheet.Cells[row, col].Value;
                    Console.Write(cellValue + "\t");
                }

                Console.WriteLine();
            }
        }
    }
}
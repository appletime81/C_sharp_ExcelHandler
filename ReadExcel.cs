using System;
using OfficeOpenXml;
using System.IO;

public static class ReadExcel
{
    public static void ReadExcelFile(string path)
    {
        Console.WriteLine("-------------------------------------------------------------");
        Console.WriteLine(path);
        Console.WriteLine("-------------------------------------------------------------");
        FileInfo file = new FileInfo(path);

        if (!file.Exists)
        {
            Console.WriteLine($"Error: File not found at {path}");
            return;
        }

        using (ExcelPackage package = new ExcelPackage(file))
        {
            // 取得第一個工作表
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // 逐列讀取
            for (int row = 1; row <= rowCount; row++)
            {
                // 逐欄讀取
                for (int col = 1; col <= colCount; col++)
                {
                    // 讀取儲存格的值
                    object cellValue = worksheet.Cells[row, col].Value;
                    Console.Write(cellValue + "\t");
                }
                Console.WriteLine();
            }
        }
    }
}

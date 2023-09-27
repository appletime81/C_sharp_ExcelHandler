namespace ExcelHandler;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

public class CombineAllSubExcels
{
    public void MergeExcelFiles(List<string> filePaths, string outputPath)
    {
        var combinedData = new List<RowData>();

        foreach (var path in filePaths)
        {
            Console.WriteLine($"Loading {path}");
            LoadAndAppendData(path, combinedData);
        }

        // Remove duplicates, prioritizing yellow highlighted rows
        combinedData = combinedData
            .GroupBy(d => new { d.Zh, d.En, d.Th, d.Vi })
            .Select(g => g.OrderByDescending(d => d.IsHighlighted).First())
            .ToList();

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Merged Data");

            // Add headers
            worksheet.Cells[1, 1].Value = "Zh";
            worksheet.Cells[1, 2].Value = "En";
            worksheet.Cells[1, 3].Value = "Th";
            worksheet.Cells[1, 4].Value = "Vi";

            // Modify loop to start from the second row as the first row now contains headers
            for (var i = 0; i < combinedData.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = combinedData[i].Zh;
                worksheet.Cells[i + 2, 2].Value = combinedData[i].En;
                worksheet.Cells[i + 2, 3].Value = combinedData[i].Th;
                worksheet.Cells[i + 2, 4].Value = combinedData[i].Vi;

                if (combinedData[i].IsHighlighted)
                {
                    worksheet.Cells[i + 2, 1, i + 2, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[i + 2, 1, i + 2, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }
            }

            package.SaveAs(new FileInfo(outputPath));
        }

        Console.WriteLine($"Saved to {outputPath}");
    }

    private void LoadAndAppendData(string path, List<RowData> combinedData)
    {
        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            var worksheet = package.Workbook.Worksheets[0];

            // 打印header
            Console.WriteLine(
                $"  {worksheet.Cells[1, 1].Text} | {worksheet.Cells[1, 2].Text} | {worksheet.Cells[1, 3].Text} | {worksheet.Cells[1, 4].Text}");
            Console.WriteLine(
                "---------------------------------------------------------------------------------------------------------------------------------------------------------");

            // check header, its need including "zh", "en", "th", "vi", maybe lowerCase or upperCase, always convert into lowerCase to compare
            // if no have "zh", "en", "th", "vi" one of them, give a column with empty values
            var zhIndex = -1;
            var enIndex = -1;
            var thIndex = -1;
            var viIndex = -1;
            for (var i = 1; i <= worksheet.Dimension.Columns; i++)
            {
                var header = worksheet.Cells[1, i].Text.ToLower();
                if (header == "zh")
                {
                    zhIndex = i;
                }
                else if (header == "en")
                {
                    enIndex = i;
                }
                else if (header == "th")
                {
                    thIndex = i;
                }
                else if (header == "vi")
                {
                    viIndex = i;
                }
            }

            // combine data and column name is "zh", "en", "th", "vi"
            for (var i = 2; i <= worksheet.Dimension.Rows; i++)
            {
                var zh = zhIndex == -1 ? "" : worksheet.Cells[i, zhIndex].Text;
                var en = enIndex == -1 ? "" : worksheet.Cells[i, enIndex].Text;
                var th = thIndex == -1 ? "" : worksheet.Cells[i, thIndex].Text;
                var vi = viIndex == -1 ? "" : worksheet.Cells[i, viIndex].Text;
                var isHighlighted = worksheet.Cells[i, 1].Style.Fill.BackgroundColor.Rgb == "FFFFFF00";

                combinedData.Add(new RowData
                {
                    Zh = zh,
                    En = en,
                    Th = th,
                    Vi = vi,
                    IsHighlighted = isHighlighted
                });

                // 打印每一行的資料
                Console.WriteLine($"  {zh} | {en} | {th} | {vi}");
                Console.WriteLine(
                    "*********************************************************************************************************************************************************");
            }
        }
    }

    class RowData
    {
        public string? Zh { get; set; }
        public string? En { get; set; }
        public string? Th { get; set; }
        public string? Vi { get; set; }
        public bool IsHighlighted { get; set; }
    }
}
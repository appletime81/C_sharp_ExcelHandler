namespace ExcelHandler;

using OfficeOpenXml;
using System.IO;

public class ExcelProcessor
{
    public static void AppendExcelWithoutDuplicates(string sourceFilePath, string targetFilePath)
    {
        var sourceFile = new FileInfo(sourceFilePath);
        var targetFile = new FileInfo(targetFilePath);
        using var sourcePackage = new ExcelPackage(sourceFile);
        using var targetPackage = new ExcelPackage(targetFile);
        var sourceWorksheet = sourcePackage.Workbook.Worksheets[0];
        // 如果sourceWorksheet有column name為'line'或'Key'的話，就把它drop掉
        var sourceColumnCount = sourceWorksheet.Dimension.End.Column;
        for (var col = 1; col <= sourceColumnCount; col++)
        {
            var columnName = sourceWorksheet.Cells[1, col].Text;
            if (columnName == "line" || columnName == "Key")
            {
                sourceWorksheet.DeleteColumn(col);
                col--;
                sourceColumnCount--;
            }
        }

        var targetWorksheet = targetPackage.Workbook.Worksheets[0];
        var lastRowTarget = targetWorksheet.Dimension?.End.Row ?? 0;

        for (var sourceRow = 2; sourceRow <= sourceWorksheet.Dimension.End.Row; sourceRow++)
        {
            var isDuplicate = false;

            for (var targetRow = 2; targetRow <= lastRowTarget; targetRow++)
            {
                if (!IsRowDuplicate(sourceWorksheet, sourceRow, targetWorksheet, targetRow))
                    continue;

                isDuplicate = true;
                break;
            }

            if (isDuplicate)
                continue;

            lastRowTarget++;
            for (var col = 1; col <= sourceWorksheet.Dimension.End.Column; col++)
            {
                targetWorksheet.Cells[lastRowTarget, col].Value = sourceWorksheet.Cells[sourceRow, col].Value;
            }
        }

        targetPackage.Save();
    }

    private static bool IsRowDuplicate(ExcelWorksheet sourceWorksheet, int sourceRow, ExcelWorksheet targetWorksheet,
        int targetRow)
    {
        for (var col = 1; col <= sourceWorksheet.Dimension.End.Column; col++)
        {
            if (sourceWorksheet.Cells[sourceRow, col].Text != targetWorksheet.Cells[targetRow, col].Text)
            {
                return false; // 有不匹配的儲存格，所以行不重複
            }
        }

        return true; // 所有儲存格都匹配，行重複
    }
}
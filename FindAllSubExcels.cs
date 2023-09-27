namespace ExcelHandler;

using System;
using System.Collections.Generic;
using System.IO;

public class FindAllSubExcels
{
    public List<string> GetExcelFiles(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
        {
            throw new DirectoryNotFoundException($"Directory '{directoryPath}' not found.");
        }

        // Search for both .xlsx and .xls files
        var excelFiles = new List<string>();
        excelFiles.AddRange(Directory.GetFiles(directoryPath, "*.xlsx", SearchOption.AllDirectories));
        excelFiles.AddRange(Directory.GetFiles(directoryPath, "*.xls", SearchOption.AllDirectories));

        return excelFiles;
    }
}
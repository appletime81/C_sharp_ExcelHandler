public static class FindAllExcel
{
    // 遍歷尋找某個目錄底下所有的excel檔案
    public static List<string> FindAllExcelFiles(string path)
    {
        // 取得目錄底下所有的檔案
        var files = Directory.GetFiles(path);
        var allExcelFiles = new List<string>();

        foreach (var file in files)
        {
            // 判斷副檔名是否為excel
            if (Path.GetExtension(file) == ".xlsx" && Path.GetExtension(file) != ".xls")
            {
                allExcelFiles.Add(file);
                // Console.WriteLine(file);
            }
        }

        // 取得目錄底下所有的資料夾
        var directories = Directory.GetDirectories(path);
        foreach (var directory in directories)
        {
            // 遞迴呼叫
            allExcelFiles.AddRange(FindAllExcelFiles(directory));
        }
        return allExcelFiles;
    }
}
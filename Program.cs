using OfficeOpenXml;


namespace ExcelHandler
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            // ------------------------------ 宣告路變數 ------------------------------
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var assemblyDirectory = System.IO.Path.GetDirectoryName(assemblyLocation);
            const string relativePath = "mams";

            string directoryPath;
            directoryPath = System.IO.Path.Combine(assemblyDirectory, relativePath);

            string savePath1;
            string savePath2;
            savePath1 = System.IO.Path.Combine(assemblyDirectory, "output1.xlsx");
            savePath2 = System.IO.Path.Combine(assemblyDirectory, "final_glossary.xlsx");

            var tempExcelFiles = new List<string>();
            tempExcelFiles.Add(System.IO.Path.Combine(assemblyDirectory, "glossary.xlsx"));
            tempExcelFiles.Add(savePath1);
            // -----------------------------------------------------------------------

            // ------------------------------ 遍歷搜尋mams folder 底下的 Excel ------------------------------
            var finder = new FindAllSubExcels();
            var excelFiles = finder.GetExcelFiles(directoryPath);
            // -------------------------------------------------------------------------------------------

            // ------------------------------ 開始合併 ------------------------------
            var combiner = new CombineAllSubExcels();
            combiner.MergeExcelFiles(excelFiles, savePath1);
            combiner.MergeExcelFiles(tempExcelFiles, savePath2);
            // --------------------------------------------------------------------

            // ------------------------------ 設定 「"final_glossary.xlsx"」 的格式 ------------------------------
            using (var package = new ExcelPackage(new FileInfo(savePath2)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                worksheet.Column(1).Width = 180;
                worksheet.Column(2).Width = 180;
                worksheet.Column(3).Width = 180;
                worksheet.Column(4).Width = 180;

                // 設定預設可以篩選
                worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].AutoFilter = true;

                // 儲存
                package.Save();
            }
            // -----------------------------------------------------------------------------------------------
        }
    }
}
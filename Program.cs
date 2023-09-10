using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using static ReadExcel;


// 遍歷某個路徑底下所有資料夾及子資料夾，並找出所有excel檔案
namespace ExcelConsoleApp // <-- 添加此命名空間宣告
{
    // 遍歷某個路徑底下所有資料夾及子資料夾，並找出所有excel檔案
    public static class Program
    {
        private static void Main(string[] args)
        {
            string? assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string? assemblyDirectory = System.IO.Path.GetDirectoryName(assemblyLocation);
            string relativePath = "mams";
            // if not null
            if (assemblyDirectory != null)
            {
                string fullPath = System.IO.Path.Combine(assemblyDirectory, relativePath);
                // Console.WriteLine("-----------------------------");
                // Console.WriteLine(fullPath);
                // Console.WriteLine("-----------------------------");
                // 呼叫FindAllExcelFiles方法
                var allExcelFiles = FindAllExcel.FindAllExcelFiles(fullPath);
                foreach (var file in allExcelFiles)
                {
                    Console.WriteLine(file);
                    // ReadExcel.ReadExcelFile(file);
                }
                Console.WriteLine("-------------------------------------------------------------");
                foreach (var file in allExcelFiles)
                {
                    // Console.WriteLine(file);
                    ReadExcel.ReadExcelFile(file);
                }
            }
            else
            {
                Console.WriteLine("assemblyDirectory is null");
            }
        }
    }
}
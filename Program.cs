﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


// 遍歷某個路徑底下所有資料夾及子資料夾，並找出所有excel檔案
namespace ExcelHandler // <-- 添加此命名空間宣告
{
    // 遍歷某個路徑底下所有資料夾及子資料夾，並找出所有excel檔案
    public static class Program
    {
        private static void Main(string[] args)
        {
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var assemblyDirectory = System.IO.Path.GetDirectoryName(assemblyLocation);
            const string relativePath = "mams";
            
            if (assemblyDirectory != null)
            {
                var fullPath = System.IO.Path.Combine(assemblyDirectory, relativePath);
                var allExcelFiles = FindAllExcel.FindAllExcelFiles(fullPath);
                foreach (var file in allExcelFiles)
                {
                    // Console.WriteLine(file);
                }
                
                foreach (var file in allExcelFiles)
                {
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
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Match_Multiple_Strings_inside_Multiple_Files
{
    class Program
    {
        private static Dictionary<string, string> readSourceKeys(string filepath)
        {
            Dictionary<string, string> sourceKeys = new Dictionary<string, string>();
            ExcelPackage.LicenseContext = new LicenseContext?(LicenseContext.NonCommercial);
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filepath)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.First<ExcelWorksheet>();
                int row = excelWorksheet.Dimension.End.Row;
                int column = excelWorksheet.Dimension.End.Column;
                for (int index = 1; index <= row; ++index)
                {
                    IEnumerable<string> source = excelWorksheet.Cells[index, 1, index, column].Select<ExcelRangeBase, string>((Func<ExcelRangeBase, string>)(c => c.Value != null ? c.Value.ToString() : string.Empty));
                    sourceKeys.Add(source.FirstOrDefault<string>(), source.LastOrDefault<string>());
                }
            }
            return sourceKeys;
        }

        private static List<string> readSourceFile(string filepath)
        {
            List<string> allFolders = new List<string>();
            ExcelPackage.LicenseContext = new LicenseContext?(LicenseContext.NonCommercial);
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filepath)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.First<ExcelWorksheet>();
                int row = excelWorksheet.Dimension.End.Row;
                int column = excelWorksheet.Dimension.End.Column;
                for (int index = 1; index <= row; ++index)
                {
                    IEnumerable<string> source = excelWorksheet.Cells[index, 1, index, column].Select<ExcelRangeBase, string>((Func<ExcelRangeBase, string>)(c => c.Value != null ? c.Value.ToString() : string.Empty));
                    allFolders.Add(source.FirstOrDefault<string>());
                }
            }
            return allFolders;
        }

        private static void Main(string[] args)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            
            Console.WriteLine("Enter the excel sheet path for source key-value file - ");
            string keysFilePath = Console.ReadLine();
            
            Console.WriteLine("Enter the excel sheet path which has all folders files from which strings to be matches with the keys - ");
            string allFolderPaths = Console.ReadLine();
            
            Console.WriteLine("Enter the source path where you want result file - ");
            string resultPath = Console.ReadLine();

            Dictionary<string, string> sourceKeyValuePairs = Program.readSourceKeys(keysFilePath);
            List<string> allFolders = Program.readSourceFile(allFolderPaths);

            Dictionary<string, string> finalKeyValuePairs = new Dictionary<string, string>();
            
            foreach (string folderPath in allFolders)
            {
                string fileName = folderPath.Substring(folderPath.LastIndexOf("\\") + 1);
                int duplicatesCounter = 0;
                
                IEnumerable<string> files = Directory.EnumerateFiles(folderPath, "*.*").Where<string>((Func<string, bool>)(s => s.EndsWith(".cshtml") || s.EndsWith(".cs") || s.EndsWith(".ascx")));

                finalKeyValuePairs.Clear();

                foreach (string file in files)
                {
                    using (StreamReader streamReader = File.OpenText(file))
                    {
                        string end = streamReader.ReadToEnd();
                        foreach (KeyValuePair<string, string> keyValuePair in sourceKeyValuePairs)
                        {
                            if (end.IndexOf(keyValuePair.Key) > 0 && !finalKeyValuePairs.ContainsKey(keyValuePair.Key))
                                finalKeyValuePairs.Add(keyValuePair.Key, keyValuePair.Value);
                        }
                    }
                }

                if (finalKeyValuePairs != null && finalKeyValuePairs.Count > 0)
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(fileName)))
                    {
                        foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                        {
                            if (worksheet.Name.ToLower().Contains(fileName.ToLower()))
                                duplicatesCounter++;
                        }
                        fileName = duplicatesCounter > 0 ? fileName + duplicatesCounter.ToString() : fileName;

                        ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add(fileName);
                        excelWorksheet.Cells[1, 1].Value = "final keys of " + fileName;
                        int row = 2;

                        foreach (KeyValuePair<string, string> keyValuePair in finalKeyValuePairs)
                        {
                            excelWorksheet.Cells[row, 1].Value = keyValuePair.Key;
                            excelWorksheet.Cells[row, 2].Value = keyValuePair.Value;
                            row++;
                        }
                        excelPackage.Save();
                    }
                }
            }
            stopwatch.Stop();
            Console.WriteLine("Task completed to check all keys in " + stopwatch.Elapsed.TotalSeconds.ToString() + " seconds");
            Console.ReadKey();
        }
    }
}
using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                int lastRow;
                Console.WriteLine("Podaj ścieżke do folderu:");
                string path = Console.ReadLine();
                Console.WriteLine("Jak ma sie nazywac plik?");
                string targetFile = path + @"/" + Console.ReadLine() + @".xlsx";
                using (ExcelPackage targetWorkbook = new ExcelPackage())
                {
                    foreach (string file in Directory.GetFiles(path, "*.xlsx"))
                    {
                        using (ExcelPackage sourceWorkbook = new ExcelPackage(new FileInfo(file)))
                        {
                            for (int i = 1; i < sourceWorkbook.Workbook.Worksheets.Count; i++)
                            {
                                if (i <= 5) continue;
                                ExcelWorksheet sourceSheet = sourceWorkbook.Workbook.Worksheets[i];
                                string sheetName = sourceSheet.Name;
                                ExcelWorksheet targetSheet = targetWorkbook.Workbook.Worksheets[sheetName];
                                if (targetSheet == null)
                                {
                                    targetSheet = targetWorkbook.Workbook.Worksheets.Add(sheetName);
                                    lastRow = 0;
                                    targetSheet.InsertRow(lastRow + 1, 7, lastRow + 1);
                                    targetSheet.Cells[lastRow + 1, 1, 7, sourceSheet.Dimension.Columns].Value = sourceSheet.Cells[1, 1, 7, sourceSheet.Dimension.Columns].Value;

                                }
                                lastRow = targetSheet.Dimension.End.Row;

                                if (sourceSheet.Dimension.Rows - 7 < 0) continue;

                                targetSheet.InsertRow(lastRow + 1, sourceSheet.Dimension.Rows - 7, lastRow + 1);
                                targetSheet.Cells[lastRow + 1, 1, lastRow + sourceSheet.Dimension.Rows - 7, sourceSheet.Dimension.Columns].Value = sourceSheet.Cells[7, 1, sourceSheet.Dimension.Rows, sourceSheet.Dimension.Columns].Value;

                                for (int row = targetSheet.Dimension.End.Row; row >= 8; row--)
                                {
                                    if (targetSheet.Cells[row, 1].Value == null)
                                    {
                                        targetSheet.DeleteRow(row);
                                    }
                                }
                                targetSheet.Column(1).Style.Numberformat.Format = "dd-mm-yy";
                            }
                        }
                    }
                    targetWorkbook.SaveAs(new FileInfo(targetFile));
                }
                Console.WriteLine("Plik utworzony");
                Console.WriteLine();
            }
        }
    }
}
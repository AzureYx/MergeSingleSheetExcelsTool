using System;
using System.IO;
using NPOI.HSSF.UserModel; // for .xls
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // for .xlsx

namespace ExcelMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter the source directory path:");
            Console.WriteLine("Example: C:\\Users\\chaiy\\Desktop\\sourceExcel");
            string sourceDirectory = Console.ReadLine();

            Console.WriteLine("\nPlease enter the target file path:");
            Console.WriteLine("Example: C:\\Users\\chaiy\\Desktop\\destExcel");
            string targetFilePath = Console.ReadLine();

            Console.WriteLine("\nPlease enter target file name:");
            Console.WriteLine("Example: MergedExcel.xlsx");
            string targetFileName = Console.ReadLine();

            targetFilePath = Path.Combine(targetFilePath, targetFileName);

            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine("The source directory does not exist. Please check the path and try again.");
                return;
            }

            string targetDirectory = Path.GetDirectoryName(targetFilePath);
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
                Console.WriteLine("Target directory created.");
            }

            MergeExcelFiles(sourceDirectory, targetFilePath);
        }

        static void MergeExcelFiles(string sourceDirectory, string targetFilePath)
        {
            try
            {
                IWorkbook targetWorkbook = new XSSFWorkbook();
                DirectoryInfo dir = new DirectoryInfo(sourceDirectory);
                FileInfo[] sourceFiles = dir.GetFiles("*.xls*");

                foreach (FileInfo file in sourceFiles)
                {
                    using (FileStream sourceStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook sourceWorkbook;
                        if (file.Extension == ".xls")
                        {
                            sourceWorkbook = new HSSFWorkbook(sourceStream); // for .xls
                        }
                        else
                        {
                            sourceWorkbook = new XSSFWorkbook(sourceStream); // for .xlsx
                        }

                        ISheet sourceSheet = sourceWorkbook.GetSheetAt(0); // Assuming each source file has only one sheet
                        string sheetName = Path.GetFileNameWithoutExtension(file.Name);

                        ISheet targetSheet = targetWorkbook.CreateSheet(sheetName);
                        CopySheet(sourceSheet, targetSheet, targetWorkbook);
                    }
                }

                using (FileStream targetStream = new FileStream(targetFilePath, FileMode.Create, FileAccess.Write))
                {
                    targetWorkbook.Write(targetStream);
                }

                Console.WriteLine("Excel files merged successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        static void CopySheet(ISheet sourceSheet, ISheet targetSheet, IWorkbook targetWorkbook)
        {
            for (int rowIndex = sourceSheet.FirstRowNum; rowIndex <= sourceSheet.LastRowNum; rowIndex++)
            {
                IRow sourceRow = sourceSheet.GetRow(rowIndex);
                IRow targetRow = targetSheet.CreateRow(rowIndex);

                if (sourceRow != null)
                {
                    CopyRow(sourceRow, targetRow, targetWorkbook);
                }
            }

            // Copy column widths
            for (int colIndex = 0; colIndex <= sourceSheet.LastRowNum; colIndex++)
            {
                targetSheet.SetColumnWidth(colIndex, sourceSheet.GetColumnWidth(colIndex));
            }
        }

        static void CopyRow(IRow sourceRow, IRow targetRow, IWorkbook targetWorkbook)
        {
            for (int colIndex = sourceRow.FirstCellNum; colIndex < sourceRow.LastCellNum; colIndex++)
            {
                ICell sourceCell = sourceRow.GetCell(colIndex);
                ICell targetCell = targetRow.CreateCell(colIndex);

                if (sourceCell != null)
                {
                    CopyCell(sourceCell, targetCell, targetWorkbook);
                }
            }

            // Copy row height
            targetRow.Height = sourceRow.Height;
        }

        static void CopyCell(ICell sourceCell, ICell targetCell, IWorkbook targetWorkbook)
        {
            switch (sourceCell.CellType)
            {
                case CellType.Boolean:
                    targetCell.SetCellValue(sourceCell.BooleanCellValue);
                    break;
                case CellType.Numeric:
                    targetCell.SetCellValue(sourceCell.NumericCellValue);
                    break;
                case CellType.String:
                    targetCell.SetCellValue(sourceCell.StringCellValue);
                    break;
                case CellType.Error:
                    targetCell.SetCellValue(sourceCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    targetCell.CellFormula = sourceCell.CellFormula;
                    break;
                case CellType.Blank:
                    targetCell.SetCellValue(sourceCell.StringCellValue);
                    break;
                default:
                    targetCell.SetCellValue(sourceCell.StringCellValue);
                    break;
            }

            // Copy cell style
            //ICellStyle newCellStyle = targetWorkbook.CreateCellStyle();
            //if (sourceCell.CellStyle != null)
            //{
            //    newCellStyle.CloneStyleFrom(sourceCell.CellStyle);

            //    // Apply font style
            //    IFont sourceFont = sourceCell.CellStyle.GetFont(sourceCell.Sheet.Workbook);
            //    IFont targetFont = targetWorkbook.CreateFont();
            //    targetFont.Boldweight = sourceFont.Boldweight;
            //    targetFont.Color = sourceFont.Color;
            //    targetFont.FontHeight = sourceFont.FontHeight;
            //    targetFont.FontName = sourceFont.FontName;
            //    targetFont.IsItalic = sourceFont.IsItalic;
            //    targetFont.IsStrikeout = sourceFont.IsStrikeout;
            //    targetFont.Underline = sourceFont.Underline;
            //    newCellStyle.SetFont(targetFont);
            //}

            //targetCell.CellStyle = newCellStyle;
        }
    }
}

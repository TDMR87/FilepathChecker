using ClosedXML.Excel;
using FilepathCheckerWPF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

namespace FilepathCheckerWPF
{
    public static class Methods
    {
        public static async Task<List<string>> GetFilepathsFromFileParallelAsync(
            string filepath, 
            int columnNumber,
            IProgress<ProgressReportModelV2> progress, 
            ParallelOptions parallelOptions)
        {
            List<string> filepaths = new List<string>();
            ProgressReportModelV2 report = new ProgressReportModelV2();
            XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet;

            // Try to open excel-file
            try
            {
                workbook = new XLWorkbook(filepath);
            }
            catch (Exception)
            {
                return filepaths;
            }
            finally
            {
                workbook.Dispose();
            }

            // Try to get the first sheet in the workbook
            try
            {
                worksheet = workbook.Worksheets.First();
            }
            catch (Exception)
            {
                return filepaths;
            }

            await Task.Run(() =>
            {
                try
                {
                    // Iterate each row
                    Parallel.ForEach<IXLRow>(worksheet.Rows(), parallelOptions, async (row) =>
                    {
                        // Cancel Parallel.ForEach()
                        parallelOptions.CancellationToken.ThrowIfCancellationRequested();

                        // Skip the title row
                        if (row.RowNumber() == 1)
                            return;

                        // Get the cell in specified column
                        IXLCell cell = row.Cell(columnNumber);

                        // Get filepaths in the cell
                        foreach (string path in cell.Value.ToString().Split('|').ToList())
                        {
                            filepaths.Add(path);
                        }

                        // Report progress
                        report.Filepaths = filepaths;
                        report.PercentageCompleted = (filepaths.Count * 100) / worksheet.Rows().Count();
                        progress.Report(report);
                    });
                }
                catch (OperationCanceledException ex)
                {
                }

            }).ConfigureAwait(true);

            // Return result
            return filepaths;
        }

        public static async Task<List<string>> GetFilepathsFromFileAsync(
            string filepath,
            int columnNumber,
            IProgress<ProgressReportModelV2> progress, 
            ParallelOptions parallelOptions)
        {
            List<string> filepaths = new List<string>();
            ProgressReportModelV2 report = new ProgressReportModelV2();
            XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet;

            // Try to open excel-file
            try
            {
                workbook = new XLWorkbook(filepath);
            }
            catch (Exception)
            {
                return filepaths;
            }
            finally
            {
                workbook.Dispose();
            }

            // Try to get the first sheet in the workbook
            try
            {
                worksheet = workbook.Worksheets.First();
            }
            catch (Exception)
            {
                return filepaths;
            }

            await Task.Run(() =>
            {
                try
                {
                    foreach (var row in worksheet.Rows())
                    {
                        // Cancel Parallel.ForEach()
                        parallelOptions.CancellationToken.ThrowIfCancellationRequested();

                        // Skip the title row
                        if (row.RowNumber() == 1)
                            continue;

                        // Get the cell in specified column
                        IXLCell cell = row.Cell(columnNumber);

                        // Get filepaths in the cell
                        foreach (string path in cell.Value.ToString().Split('|').ToList())
                        {
                            filepaths.Add(path);
                        }

                        // Report progress
                        report.Filepaths = filepaths;
                        report.PercentageCompleted = (filepaths.Count * 100) / worksheet.Rows().Count();
                        progress.Report(report);
                    }
                }
                catch (OperationCanceledException ex)
                {
                }
            }).ConfigureAwait(true);

            // Return result
            return filepaths;
        }

        public static async Task<List<string>> ReadFileUsingOpenXMLAsync(
            string filepath,
            string columnCharacter,
            IProgress<ProgressReportModelV2> progress,
            ParallelOptions parallelOptions)
        {
            List<string> filepaths = new List<string>();
            ProgressReportModelV2 report = new ProgressReportModelV2();

            await Task.Run(() =>
            {
                // Open a SpreadsheetDocument for read-only access based on a filepath.
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
                {
                    int sheetNo = 0; // Index 0 => sheet 1
                    WorkbookPart workbookPart;
                    Sheet sheet;
                    WorksheetPart worksheetPart;
                    SheetData sheetData;

                    try
                    {
                        workbookPart = spreadsheetDocument.WorkbookPart;
                        sheet = workbookPart.Workbook
                         .Descendants<Sheet>()
                         .ElementAt(sheetNo);
                        worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                        sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    List<Row> rows = sheetData.Elements<Row>().ToList();
                    int rowAmount = rows.Count;
                    int rowCounter = 1;

                    try
                    {
                        foreach (Row row in rows)
                        {
                            parallelOptions.CancellationToken.ThrowIfCancellationRequested();

                            // Gets all cell elements from the row
                            List<Cell> cells = row.Elements<Cell>().ToList();

                            foreach (Cell cell in cells)
                            {
                                // If cell's datatype is text
                                // If column letter (cell reference) equals the user specified column 
                                // If cell is not on the first row (usually a title row)
                                if (cell.DataType != null
                                    && cell.DataType == CellValues.SharedString
                                    && cell.CellReference.InnerText.Equals(String.Join("",columnCharacter,rowCounter))
                                    && cell.CellReference.InnerText != columnCharacter + "1") // Exclude headings
                                {
                                    //it's a shared string so use the cell inner text as the index into the 
                                    //shared strings table
                                    var stringId = Convert.ToInt32(cell.InnerText);
                                    string cellValue = workbookPart.SharedStringTablePart.SharedStringTable
                                        .Elements<SharedStringItem>()
                                        .ElementAt(stringId).InnerText;

                                    // If cell is empty, move on to the next row.
                                    if (string.IsNullOrWhiteSpace(cellValue))
                                        break;

                                    // Get filepaths in the cell
                                    // filepaths may be sepratated by a pipe character
                                    foreach (string path in cellValue.Split('|').ToList())
                                    {
                                        filepaths.Add(path);
                                    }

                                    break; // No need to check other cells, move on to the next row.
                                }
                            }

                            // Report progress after each row
                            report.Filepaths = filepaths;
                            report.PercentageCompleted = (rowCounter * 100) / rowAmount;
                            progress.Report(report);

                            rowCounter++;
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        return;
                    }
                }
            }).ConfigureAwait(true);

            return filepaths;
        }

        public static async Task<FileModel> CheckFileExistsAsync(
            string path)
        {
            FileModel file = new FileModel();

            await Task.Run(() =>
            {
                string name = "";
                try
                {
                    name = System.IO.Path.GetFileName(path);
                }
                catch (Exception)
                {
                    // 
                }

                if (File.Exists(path))
                {
                    file.FileExists = true;
                    file.Filepath = $"{name}";
                }
                else
                {
                    file.FileExists = false;
                    file.Filepath = $"{name}";
                }
            }).ConfigureAwait(true);

            return file;
        }

        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException(nameof(columnName));

            columnName = columnName.ToUpperInvariant();
            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
    }
}

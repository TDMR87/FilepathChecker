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
    public static class FileReader
    {
        /// <summary>
        /// Reads an excel file using Open XML library and extracts the values in a specific column.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="columnCharacter"></param>
        /// <param name="progress"></param>
        /// <param name="parallelOptions"></param>
        /// <returns></returns>
        public static async Task<List<string>> ReadFileUsingOpenXMLAsync(
            string filepath,
            string columnCharacter,
            IProgress<ProgressReportModelV2> progress,
            ParallelOptions parallelOptions)
        {
            List<string> filepaths = new List<string>();
            ProgressReportModelV2 report = new ProgressReportModelV2();

            // Start a task in the background that we can cancel using the ParallelOptions cancellation token.
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
                        // Add a workbook part
                        workbookPart = spreadsheetDocument.WorkbookPart;

                        // Get the first sheet
                        sheet = workbookPart.Workbook
                         .Descendants<Sheet>()
                         .ElementAt(sheetNo);

                        // Add a worksheet part
                        worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

                        // Data inside the first sheet
                        sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    // Get all the rows in the first sheet
                    List<Row> rows = sheetData.Elements<Row>().ToList();
                    int rowAmount = rows.Count;
                    int rowCounter = 1;

                    try
                    {
                        // Iterate over all the rows
                        foreach (Row row in rows)
                        {
                            // Throw if cancelled by the user
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

        /// <summary>
        /// Checks the given UNC filepath if it exists.
        /// Returns a FileModel object and sets its FileExists property
        /// to True or False.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static async Task<FileModel> CheckFileExistsAsync(
            string path)
        {
            FileModel file = new FileModel();

            await Task.Run(() =>
            {
                string filename = "";
                try
                {
                    filename = System.IO.Path.GetFileName(path);
                }
                catch (Exception)
                {
                    throw;
                }

                if (File.Exists(path))
                {
                    file.FileExists = true;
                    file.Filepath = $"{filename}";
                }
                else
                {
                    file.FileExists = false;
                    file.Filepath = $"{filename}";
                }
            }).ConfigureAwait(true);

            return file;
        }

        /// <summary>
        /// Helper method for resolving column names (e.g. A,B) to their corresponding number
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
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

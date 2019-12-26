using FilepathCheckerWPF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;

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
                SpreadsheetDocument spreadsheetDocument;
                try
                {
                    spreadsheetDocument = SpreadsheetDocument.Open(filepath, false);
                }
                catch (OpenXmlPackageException ex)
                {
                    filepaths.Add(ex.Message);
                    return;
                }
                catch (ArgumentException ex)
                {
                    filepaths.Add(ex.Message);
                    return;
                }
                catch (IOException ex)
                {
                    filepaths.Add(ex.Message);
                    return;
                }
                catch (FileFormatException ex)
                {
                    filepaths.Add(ex.Message);
                    return;
                }

                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook
                    .Descendants<Sheet>()
                    .ElementAt(0); // Get the first sheet, index 0

                WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Get all the rows in the first sheet
                List<Row> rows = sheetData.Elements<Row>().ToList();
                int rowAmount = rows.Count;
                int rowCounter = 1;

                // Iterate through all the rows
                foreach (Row row in rows)
                {
                    try
                    {
                        parallelOptions.CancellationToken.ThrowIfCancellationRequested();
                    }
                    catch (OperationCanceledException)
                    {
                        return;
                    }

                    string column = string.Join("", columnCharacter, rowCounter);

                    // Get all cell elements from the row that belong to the specified column
                    // and are of correct datatype
                    List<Cell> cells = row.Elements<Cell>().Where(cell => 
                        cell.DataType != null 
                        && cell.DataType == CellValues.SharedString
                        && cell.CellReference.InnerText.Equals(column, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    // Iterate through the cells in the current row
                    foreach (Cell cell in cells)
                    {
                        // The cell value is a shared string so use the cell's inner text as the index into the 
                        // shared strings table
                        int stringId = Convert.ToInt32(cell.InnerText);
                        string cellValue = workbookPart.SharedStringTablePart.SharedStringTable
                            .Elements<SharedStringItem>()
                            .ElementAt(stringId).InnerText;

                        // Get filepath values in the cell value
                        // filepaths may be separated with a pipe character
                        foreach (string path in cellValue.Split('|').ToList())
                        {
                            filepaths.Add(path);
                        }
                    }

                    // Report progress after each row
                    report.Filepaths = filepaths;
                    report.PercentageCompleted = (rowCounter * 100) / rowAmount;
                    progress.Report(report);

                    rowCounter++;
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

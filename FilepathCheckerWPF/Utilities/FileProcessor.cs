using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;

namespace FilepathCheckerWPF
{
    public static class FileProcessor
    {
        /// <summary>
        /// Opens an excel file from the provided filepath using Open XML library. Extracts all the values from the specified column.
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
                catch(Exception ex) when (ex is OpenXmlPackageException
                                        || ex is ArgumentException
                                        || ex is IOException
                                        || ex is FileFormatException)
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

                    // Concatenates the specified column character with the current row.
                    // e.g A1, A2, A3 and so on...
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

                        if (string.IsNullOrWhiteSpace(cellValue)) { continue; }

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

        public static async Task<List<IFileModel>> ProcessFilepaths(List<string> filepaths, IProgress<ProgressReportModel> progress, ParallelOptions parallelOptions)
        {
            List<IFileModel> output = new List<IFileModel>();
            ProgressReportModel report = new ProgressReportModel();
            CsvLogger logger = new CsvLogger();

            await Task.Run(async () =>
            {
                foreach (string path in filepaths)
                {
                    // Cancel if user presses Stop
                    try
                    {
                        parallelOptions.CancellationToken.ThrowIfCancellationRequested();
                    }
                    catch (OperationCanceledException)
                    {
                        return;
                    }

                    // Wrap each filepath into a IFileModel object
                    IFileModel file = FileProcessor.CreateFileModel(path);

                    // Add the processed file to the output collection
                    output.Add((FileModel)file);

                    // If file does not exist, log the filepath.
                    if (file.FileExists == false)
                        logger.WriteLine(file.Filepath);

                    // Report progress after each processed filepath
                    report.FilesProcessed = output;
                    report.PercentageCompleted = (output.Count * 100) / filepaths.Count;
                    progress.Report(report);
                }
            }).ConfigureAwait(true);

            // Close the logger
            logger.Close();
            logger.Dispose();

            return output;
        }

        /// <summary>
        /// Checks the given UNC filepath and returns a FileModel object. The FileExists property
        /// will be set to True or False depending if the provided filepath exists or not.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static async Task<IFileModel> CreateFileModelAsync(
            string path)
        {
            FileModel file = new FileModel();

            // Create an instance of FileModel and check if file exists
            await Task.Run(() =>
            {
                string filepath = "";
                try
                {
                    filepath = System.IO.Path.GetFileName(path);
                }
                catch (Exception)
                {
                    throw;
                }

                if (File.Exists(path))
                {
                    file.FileExists = true;
                    file.Filepath = $"{filepath}";
                }
                else
                {
                    file.FileExists = false;
                    file.Filepath = $"{filepath}";
                }
            }).ConfigureAwait(true);

            return file;
        }

        /// <summary>
        /// Checks the given UNC filepath and returns a FileModel object. The FileExists property
        /// will be set to True or False depending if the provided filepath exists or not.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static IFileModel CreateFileModel(
            string path)
        {
            FileModel file = new FileModel();

            string filepath = "";
            try
            {
                filepath = System.IO.Path.GetFileName(path);
            }
            catch (Exception)
            {
                throw;
            }

            if (File.Exists(path))
            {
                file.FileExists = true;
                file.Filepath = $"{filepath}";
            }
            else
            {
                file.FileExists = false;
                file.Filepath = $"{filepath}";
            }

            return file;
        }

        /// <summary>
        /// Resolves column names (e.g. A) to their corresponding number (A = 1)
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

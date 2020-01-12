using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;
using System.Globalization;
using DocumentFormat.OpenXml;

namespace FilepathCheckerWPF
{
    public static class FileProcessor
    {
        /// <summary>
        /// Opens an excel file from the provided filepath using Open XML library. 
        /// Uses the DOM approach that requires loading entire Open XML parts into memory, 
        /// which can cause an Out of Memory exception when working with really large files.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="columnCharacter"></param>
        /// <param name="progress"></param>
        /// <param name="parallelOptions"></param>
        /// <returns></returns>
        public static async Task<List<string>> ReadExcelFileDOMAsync(
            string filepath,
            string columnCharacter,
            IProgress<ProgressReportModelV2> progress,
            ParallelOptions parallelOptions)
        {
            List<string> output = new List<string>();
            ProgressReportModelV2 report = new ProgressReportModelV2();

            // Start a task in the background that we can cancel using the ParallelOptions cancellation token.
            await Task.Run(() =>
            {
                SpreadsheetDocument spreadsheetDocument;
                try
                {
                    // Open a SpreadsheetDocument in read-only mode.
                    spreadsheetDocument = SpreadsheetDocument.Open(filepath, false);
                }
                catch(Exception ex) when (ex is OpenXmlPackageException
                                        || ex is ArgumentException
                                        || ex is IOException
                                        || ex is FileFormatException)
                {
                    output.Add(ex.Message);
                    return;
                }

                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                // Gets the first sheet, index 0
                Sheet sheet = workbookPart.Workbook
                    .Descendants<Sheet>()
                    .ElementAt(0); 

                WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Get all the rows in the first sheet
                List<Row> rows = sheetData.Elements<Row>().ToList();
                int rowAmount = rows.Count;
                int currentRow = 1;

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

                    // Set the cell name we are looking for in each row.
                    // Concatenates the specified column character with the current row number.
                    // e.g A1, A2, A3 and so on...
                    string columnName = string.Join("", columnCharacter, currentRow);

                    // Get all cell elements from the row that belong to the specified column
                    // and are of correct datatype
                    List<Cell> cells = row.Elements<Cell>().Where(cell =>
                                cell.DataType != null
                                && cell.DataType == CellValues.SharedString
                                && cell.CellReference.InnerText.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                                .ToList();

                    // Iterate through the cells
                    foreach (Cell cell in cells)
                    {
                        // The cell value is a shared string so use the cell's inner text as the index into the 
                        // shared strings table
                        int stringId = Convert.ToInt32(cell.InnerText, CultureInfo.InvariantCulture);
                        string cellValue = workbookPart.SharedStringTablePart.SharedStringTable
                            .Elements<SharedStringItem>()
                            .ElementAt(stringId).InnerText;

                        // If cell is empty, continue to the next row.
                        if (string.IsNullOrWhiteSpace(cellValue)) { continue; }

                        // Get filepath values in the cell value
                        // filepaths may be separated with a pipe character
                        foreach (string path in cellValue.Split('|').ToList())
                        {
                            output.Add(path);
                        }
                    }

                    // Report progress after each row
                    report.Filepaths = output;
                    report.PercentageCompleted = (currentRow * 100) / rowAmount;
                    progress.Report(report);

                    currentRow++;
                }
            }).ConfigureAwait(false);

            return output;
        }

        /// <summary>
        /// Opens an excel file from the provided filepath using Open XML library. 
        /// Using the SAX approach, you can employ an OpenXMLReader to read the XML in the file one element at a time, 
        /// without having to load the entire file into memory. Consider using SAX when handling very large files.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="columnCharacter"></param>
        /// <param name="progress"></param>
        /// <param name="parallelOptions"></param>
        /// <returns></returns>
        public static async Task<List<string>> ReadExcelFileSAXAsync(
            string filepath,
            string columnCharacter,
            IProgress<ProgressReportModelV2> progress,
            ParallelOptions parallelOptions)
        {
            List<string> output = new List<string>();
            ProgressReportModelV2 report = new ProgressReportModelV2();

            await Task.Run(() =>
            {
                SpreadsheetDocument spreadsheetDocument;
                try
                {
                    // Open a SpreadsheetDocument in read-only mode.
                    spreadsheetDocument = SpreadsheetDocument.Open(filepath, false);
                }
                catch (Exception ex) when (ex is OpenXmlPackageException
                                        || ex is ArgumentException
                                        || ex is IOException
                                        || ex is FileFormatException)
                {
                    output.Add("File open error. " + ex.Message);
                    return;
                }

                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                int totalRows = 0;
                int currentRow = 1;
                int sharedStringIndex;
                string sharedStringValue;

                // Read the file to the end in order to calculate the amount of rows
                while (reader.Read() && !reader.EOF)
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        totalRows++;
                        reader.Skip(); // Skip rest of the elements and read the next row
                    }

                    try
                    {
                        // Cancel task if user pressed stop
                        parallelOptions.CancellationToken.ThrowIfCancellationRequested();
                    }
                    catch (OperationCanceledException)
                    {
                        return;
                    }
                }

                // Re-create the reader to start reading from the top again.
                reader = OpenXmlReader.Create(worksheetPart);

                while (reader.Read())
                {
                    // Concatenate a column name
                    string columnName = string.Join("", columnCharacter, currentRow);

                    // Iterate through the XML elements and find the ones that are of the Cell type
                    if (reader.ElementType == typeof(Cell))
                    {
                        Cell cell = (Cell)reader.LoadCurrentElement();

                        // Last condition checks that the cell is in the user specified column
                        if (cell.DataType != null
                            && cell.DataType == CellValues.SharedString
                            && string.IsNullOrWhiteSpace(cell.InnerText) == false
                            && cell.CellReference.InnerText.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                        {
                            sharedStringIndex = Convert.ToInt32(cell.InnerText, CultureInfo.InvariantCulture);
                            sharedStringValue = workbookPart.SharedStringTablePart.SharedStringTable
                                    .Elements<SharedStringItem>()
                                    .ElementAt(sharedStringIndex).InnerText;

                            output.Add(sharedStringValue);
                            currentRow++;

                            // Report progress after each row
                            report.Filepaths = output;
                            report.PercentageCompleted = (currentRow * 100) / totalRows;
                            progress.Report(report);
                        }

                        try
                        {
                            // Cancel task if user pressed stop
                            parallelOptions.CancellationToken.ThrowIfCancellationRequested();
                        }
                        catch (OperationCanceledException)
                        {
                            return;
                        }
                    }
                }

                reader.Dispose();

            }).ConfigureAwait(false);

            return output;
        }

        /// <summary>
        /// Takes a list of filepaths and wraps each filepath into an IFileModel object. 
        /// Returns the objects as a List.
        /// </summary>
        /// <param name="filepaths"></param>
        /// <param name="progress"></param>
        /// <param name="parallelOptions"></param>
        /// <returns></returns>
        public static async Task<List<IFileModel>> ProcessFilepaths(List<string> filepaths, IProgress<ProgressReportModel> progress, ParallelOptions parallelOptions)
        {
            List<IFileModel> output = new List<IFileModel>();
            ProgressReportModel report = new ProgressReportModel();
            CsvLogger logger = new CsvLogger();

            await Task.Run(() =>
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
            }).ConfigureAwait(false);

            // Close the logger
            logger.Close();
            logger.Dispose();

            return output;
        }

        /// <summary>
        /// Checks the given UNC filepath and returns a FileModel object. The FileExists property
        /// will be set to True or False depending if the provided filepath exists or not.
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        private static async Task<IFileModel> CreateFileModelAsync(
            string filepath)
        {
            FileModel file = new FileModel();

            // Create an instance of FileModel and check if file exists
            await Task.Run(() =>
            {
                file.Filepath = filepath;
                file.FileExists = File.Exists(filepath) ? true : false;

            }).ConfigureAwait(false);

            return file;
        }

        /// <summary>
        /// Checks the given UNC filepath and returns a FileModel object. The FileExists property
        /// will be set to True or False depending if the provided filepath exists or not.
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        private static IFileModel CreateFileModel(
            string filepath)
        {
            FileModel file = new FileModel();

            file.Filepath = filepath;
            file.FileExists = File.Exists(filepath) ? true : false;

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

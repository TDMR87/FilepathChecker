using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;
using System.Globalization;
using DocumentFormat.OpenXml;
using System.Diagnostics;

namespace FilepathCheckerWPF
{
    public class FileProcessor
    {
        private readonly ILogger _logger;
        private static Func<IFileModel> _fileModelFactory;
        private readonly SpreadsheetDocument _spreadsheetDocument;

        public FileProcessor(SpreadsheetDocument spreadsheetDocument, Func<IFileModel> fileModelFactory, ILogger logger)
        {
            _logger = logger;
            _spreadsheetDocument = spreadsheetDocument;
            _fileModelFactory = fileModelFactory;
        }

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
        public async Task<List<string>> ReadColumnDomAsync(
            string columnCharacter,
            IProgress<ProgressReportModelV2> progress,
            ParallelOptions parallelOptions)
        {
            List<string> output = new List<string>();
            ProgressReportModelV2 report = new ProgressReportModelV2();

            // Start a task in the background that we can cancel using the ParallelOptions cancellation token.
            await Task.Run(() =>
            {
                WorkbookPart workbookPart = _spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().ElementAt(0); // Gets the first sheet, index 0
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
                    List<Cell> cells = row.Elements<Cell>().Where(cell => cell.DataType != null && 
                                                                  cell.DataType == CellValues.SharedString && 
                                                                  cell.CellReference.InnerText.Equals(
                                                                  columnName, 
                                                                  StringComparison.OrdinalIgnoreCase))
                                                                  .ToList();

                    // Iterate through the cells
                    foreach (Cell cell in cells)
                    {
                        // The cell value is actually an index to shared string table
                        int stringId = Convert.ToInt32(cell.InnerText, CultureInfo.InvariantCulture);

                        // The cell value is a shared string so use the cell's inner text as the index into the 
                        // shared strings table
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
        public async Task<List<string>> ReadColumnSaxAsync(
            string columnCharacter,
            IProgress<ProgressReportModelV2> progress,
            ParallelOptions parallelOptions)
        {
            // The output of this method
            List<string> output = new List<string>();

            // Report object
            ProgressReportModelV2 report = new ProgressReportModelV2();

            // Local variables
            Cell cell;
            int totalRows = 0;
            int currentRow = 1;
            int sharedStringIndex;
            string sharedStringValue;

            // Start a task
            await Task.Run(() =>
            {     
                // Open the spreadsheet document
                WorkbookPart workbookPart = _spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().ElementAt(0); // Gets the first sheet, index 0
                WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                // First, we want to calculate the total amount of rows in a SAX way.
                // We don't want to load all rows into memory and possibly have a OutOfMemoryException.
                while (reader.Read() && !reader.EOF)
                {
                    // If element is type Row
                    if (reader.ElementType == typeof(Row))
                    {
                        // While the row has sibling rows
                        while (reader.ReadNextSibling())
                        {
                            // Increment total rows
                            totalRows++;

                            // Cancel task if user pressed stop
                            parallelOptions.CancellationToken.ThrowIfCancellationRequested();
                        }

                        // Break from the loop after counting the rows
                        break;
                    }
                }

                // Re-open the worksheet to start reading from the beginning
                reader = OpenXmlReader.Create(worksheetPart);

                // Read all XML elements
                while (reader.Read())
                {
                    // Throw if cancelled
                    parallelOptions.CancellationToken.ThrowIfCancellationRequested();

                    // If the XML element is of type Cell
                    if (reader.ElementType == typeof(Cell))
                    {
                        // Load the cell
                        cell = reader.LoadCurrentElement() as Cell;

                        // If cell matches the following conditions
                        if (cell.DataType == CellValues.SharedString &&
                            CellReferenceToColumnName(cell.CellReference.Value).Equals(columnCharacter, StringComparison.OrdinalIgnoreCase))
                        {
                            // The cell value is an index to shared string table
                            sharedStringIndex = Convert.ToInt32(cell.InnerText, CultureInfo.InvariantCulture);

                            // Get the value in the specified index from the shared string table
                            sharedStringValue = workbookPart.SharedStringTablePart.SharedStringTable
                                                .Elements<SharedStringItem>()
                                                .ElementAt(sharedStringIndex).InnerText;

                            // One cell might contain several filepaths separated by a pipe character
                            foreach (string filepath in sharedStringValue.Split('|'))
                            {
                                // Add the filepath to the output list
                                output.Add(filepath);
                            }

                            // Report progress after each row
                            report.Filepaths = output;
                            report.PercentageCompleted = (currentRow * 100) / totalRows;
                            progress.Report(report);
                        }
                    }
                }

                // Cleanup
                reader.Close();
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
        public async Task<List<IFileModel>> ProcessFilepaths(
            List<string> filepaths,
            IProgress<ProgressReportModel> progress, 
            ParallelOptions parallelOptions)
        {
            List<IFileModel> output = new List<IFileModel>();
            ProgressReportModel report = new ProgressReportModel();

            await Task.Run(() =>
            {
                foreach (string path in filepaths)
                {
                    try
                    {
                        parallelOptions.CancellationToken.ThrowIfCancellationRequested();
                    }
                    catch (OperationCanceledException)
                    {
                        return;
                    }

                    // Create a IFileModel object for each filepath
                    var file = _fileModelFactory();
                    file.Filepath = path;
                    file.FileExists = File.Exists(path) ? true : false;

                    // If file does not exist, log the filepath.
                    if (file.FileExists == false)
                        _logger.Write($"File not found;{file.Filepath}");

                    // Add the file to the output collection
                    output.Add(file);

                    // Report progress after each processed filepath
                    report.FilesProcessed = output;
                    report.PercentageCompleted = (output.Count * 100) / filepaths.Count;
                    progress.Report(report);
                }
            }).ConfigureAwait(false);

            // Close and dispose the logger after wiriting
            _logger.Close();
            _logger.Dispose();

            return output;
        }

        /// <summary>
        /// Resolves column names to their corresponding number (e.g A = 1)
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

        /// <summary>
        /// Resolves a cell reference character to it's index number variant.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static int CellReferenceToIndex(Cell cell)
        {
            int index = 0;
            string cellReference = cell.CellReference.ToString().ToUpper();

            foreach (char character in cellReference)
            {
                if (Char.IsLetter(character))
                {
                    int value = (int)character - (int)'A';
                    index = (index == 0) ? value : ((index + 1) * 26) + value;
                }
                else
                    return index;
            }
            return index;
        }

        /// <summary>
        /// Returns the alphabetical letter porsion of a cell reference value.
        /// For example, an input of "A1" returns "A".
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string CellReferenceToColumnName(string text)
        {
            string output = "";

            char[] chars = text?.ToCharArray();

            foreach (char c in chars)
            {
                if (Char.IsNumber(c))
                {
                    break;
                }

                output += c;
            }

            return output;
        }
    }
}

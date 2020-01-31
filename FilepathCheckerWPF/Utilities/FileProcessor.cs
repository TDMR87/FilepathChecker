using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;
using System.Globalization;
using DocumentFormat.OpenXml;
using ClosedXML.Excel;

namespace FilepathCheckerWPF
{
    public class FileProcessor
    {
        private ILogger _logger;
        private IFileModel _fileModel;
        private SpreadsheetDocument _spreadsheetDocument;

        public FileProcessor(SpreadsheetDocument spreadsheetDocument, IFileModel fileModel, ILogger logger)
        {
            _logger = logger;
            _spreadsheetDocument = spreadsheetDocument;
            _fileModel = fileModel;
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
        public async Task<List<string>> ReadColumnSaxAsync(
            string columnCharacter,
            IProgress<ProgressReportModelV2> progress,
            ParallelOptions parallelOptions)
        {
            List<string> output = new List<string>();
            ProgressReportModelV2 report = new ProgressReportModelV2();

            await Task.Run(() =>
            {           
                WorkbookPart workbookPart = _spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().ElementAt(0); // Gets the first sheet, index 0
                WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                int totalRows = 0;

                // Read the file to the end in order to calculate the amount of rows
                while (reader.Read() && !reader.EOF)
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        while (reader.ReadNextSibling())
                        {
                            totalRows++;

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

                        // Break from the loop after counting the rows
                        break;
                    }
                }

                // Re-create the reader to start reading from the top again.
                reader.Close();
                reader.Dispose();
                reader = OpenXmlReader.Create(worksheetPart);

                int currentRow = 1;
                int sharedStringIndex;
                string sharedStringValue;

                while (reader.Read())
                {
                    // Construct the name of the column we are looking for
                    string cellName = string.Join("", columnCharacter, currentRow);

                    // Iterate through the XML elements and find the ones that are of the Cell type
                    if (reader.ElementType == typeof(Cell))
                    {
                        Cell cell = (Cell)reader.LoadCurrentElement();

                        // If cell matches the conditions
                        if (cell.DataType != null
                            && cell.DataType == CellValues.SharedString
                            && cell.CellReference.InnerText.Equals(cellName, StringComparison.OrdinalIgnoreCase)
                            && !string.IsNullOrWhiteSpace(cell.InnerText))
                        {
                            // The cell value is actually an index to shared string table
                            sharedStringIndex = Convert.ToInt32(cell.InnerText, CultureInfo.InvariantCulture);

                            // Get the value in the specified index from the shared string table
                            sharedStringValue = workbookPart.SharedStringTablePart.SharedStringTable
                                    .Elements<SharedStringItem>()
                                    .ElementAt(sharedStringIndex).InnerText;

                            // One cell might contain several filepaths separated by a pipe character
                            foreach (string filepath in sharedStringValue.Split('|').ToList())
                            {
                                // Add the filepath to the output list
                                output.Add(filepath);
                            }

                            // Report progress after each row
                            report.Filepaths = output;
                            report.PercentageCompleted = (currentRow * 100) / totalRows;
                            progress.Report(report);

                            // We found what we wanted from the current row. Move to next row
                            currentRow++;
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
                    output.Add(file);

                    // If file does not exist, log the filepath.
                    if (file.FileExists == false)
                        _logger.LogFileNotFound(file.Filepath);

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
        /// Checks the given UNC filepath and returns a FileModel object. The FileExists property
        /// will be set to True or False depending if the provided filepath exists or not.
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        private static async Task<IFileModel> CreateFileModelAsync(
            string filepath)
        {
            return await Task<IFileModel>.Run(() =>
            {
                return new FileModelV1()
                {
                    Filepath = filepath,
                    FileExists = File.Exists(filepath) ? true : false
                };

            }).ConfigureAwait(false);
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
            return new FileModelV1()
            {
                Filepath = filepath,
                FileExists = File.Exists(filepath) ? true : false
            };
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

        private static int CellReferenceToIndex(Cell cell)
        {
            int index = 0;
            string reference = cell.CellReference.ToString().ToUpper();
            foreach (char ch in reference)
            {
                if (Char.IsLetter(ch))
                {
                    int value = (int)ch - (int)'A';
                    index = (index == 0) ? value : ((index + 1) * 26) + value;
                }
                else
                    return index;
            }
            return index;
        }
    }
}

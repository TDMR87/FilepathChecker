using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;
using System.Globalization;
using DocumentFormat.OpenXml;
using System.Threading;
using System.Diagnostics;

namespace FilepathCheckerWPF
{
    public class ExcelFileProcessor
    {
        private readonly ILogger _logger;
        private static Func<IFileWrapper> _fileWrapperFactory;
        private readonly SpreadsheetDocument _spreadsheetDocument;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="fileWrapperFactory"></param>
        /// <param name="logger"></param>
        public ExcelFileProcessor(
            SpreadsheetDocument spreadsheetDocument, 
            Func<IFileWrapper> fileWrapperFactory, 
            ILogger logger)
        {
            _logger = logger;
            _spreadsheetDocument = spreadsheetDocument;
            _fileWrapperFactory = fileWrapperFactory;
        }

        /// <summary>
        /// Reads values from the specified column using Open XML library with SAX approach.
        /// Using the SAX approach, the OpenXMLReader reads the XML in the file one element at a time, 
        /// without having to load the entire file into memory. Consider using SAX when handling very large files.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="columnCharacter"></param>
        /// <param name="progress"></param>
        /// <param name="parallelOptions"></param>
        /// <returns></returns>
        public async Task<List<string>> ReadColumnSax(
            string columnCharacter,
            IProgress<ProgressStatus> progress,
            CancellationToken ct)
        {
            var output = new List<string>();
            var status = new ProgressStatus();
            Cell cell;
            int totalRows = 0;
            int currentRow = 1;
            int sharedStringIndex;
            string sharedStringValue;

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
                        ct.ThrowIfCancellationRequested();
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
                ct.ThrowIfCancellationRequested();

                // If the XML element is of type Cell
                if (reader.ElementType == typeof(Cell))
                {
                    // Load the cell
                    cell = reader.LoadCurrentElement() as Cell;

                    // If cell matches the following conditions
                    if (cell.DataType != null && cell.DataType == CellValues.SharedString &&
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

                        // Assume that each time we hit this part of code,
                        // we are on a new row.
                        currentRow++;

                        // Report progress
                        status.PercentageCompleted = (currentRow * 100) / totalRows;
                        progress.Report(status);
                    }
                }
            }

            // Cleanup
            reader.Close();
            reader.Dispose();
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
        public async Task<List<IFileWrapper>> ProcessFilepaths(
            List<string> filepaths,
            IProgress<ProgressStatus> progress, 
            CancellationToken ct)
        {
            var status = new ProgressStatus();
            var output = new List<IFileWrapper>();

            foreach (string path in filepaths)
            {
                // Throw if cancelled
                ct.ThrowIfCancellationRequested();

                // Create a FileWrapper object
                IFileWrapper file = _fileWrapperFactory();
                file.Filepath = path;
                file.FileExists = File.Exists(path);

                // If file does not exist, log the filepath.
                if (file.FileExists == false)
                    _logger.Write($"File not found;{file.Filepath}");

                // Add the file to the output collection
                output.Add(file);

                // Report progress after each processed filepath
                status.PercentageCompleted = (output.Count * 100) / filepaths.Count;
                progress.Report(status);
            }

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

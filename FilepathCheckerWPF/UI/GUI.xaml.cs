using DocumentFormat.OpenXml.Packaging;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace FilepathCheckerWPF
{
    /// <summary>
    /// Interaction logic for GUI.xaml
    /// </summary>
    public partial class GUI : Window
    {
        private static CancellationTokenSource cancellationSource = new CancellationTokenSource();
        private static List<IFileWrapper> processedFilepaths = new List<IFileWrapper>();
        private static string excelFilepath = "";
        private static string excelFilename = "";

        public GUI()
        {
            InitializeComponent();
            Title = "Filepath Validator";
            PrintMessage("Open a .xlsx file and specify a column that contains local filepaths.\n" +
                         "The program performs a check for each filepath to see if the file exists.");
        }

        /// <summary>
        /// Executes the application.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Start_Clicked(object sender, RoutedEventArgs e)
        {
            // Reset the application state
            ResetAppState();

            // If user did not provide input
            if (string.IsNullOrWhiteSpace(textboxSelectedColumn.Text))
                PrintMessage("You must specify a column.");

            // The user specified column letter
            string column =
                textboxSelectedColumn.Text.ToUpper(CultureInfo.InvariantCulture);

            // Update UI buttons
            buttonStart.IsEnabled = false;
            buttonStop.IsEnabled = true;
            buttonStart.Visibility = Visibility.Hidden;
            buttonStop.Visibility = Visibility.Visible;

            // Create a timer and start timing
            var timer = Stopwatch.StartNew();

            try
            {
                // Open the Excel spreadsheet
                using SpreadsheetDocument excelFile =
                    SpreadsheetDocument.Open(excelFilepath, false);

                // Create a file processor
                var fileProcessor =
                    new ExcelFileProcessor(excelFile,
                        () => new FileWrapper(), new CsvLogger());

                // Update status label
                labelProgressBar1.Content = "Reading the file ...";

                // Create an object for reporting progress from a background task
                var progressReport = new Progress<ProgressStatus>();

                // Add a callback for the progress report
                progressReport.ProgressChanged += UpdateProgressBar1;

                // Create a task that reads values in the specified column
                var task1 = Task.Run(() => 
                    fileProcessor.ReadColumnSax(
                        column, progressReport, cancellationSource.Token), 
                            cancellationSource.Token);
                
                // Run the task
                List<string> filepaths = await task1;

                // If task was completed
                if (task1.IsCompleted)
                    imageFileReadStatus.Source =
                        new BitmapImage(new Uri(new Checkmark().Path(), UriKind.Relative));

                // Update status label
                labelProgressBar2.Content = "Validating filepaths ...";

                // Remove previous callback from the progress report
                progressReport.ProgressChanged -= UpdateProgressBar1;

                // Add a new callback to the progress report
                progressReport.ProgressChanged += UpdateProgressBar2;

                // Create a task that processes the filepaths
                var task2 = Task.Run(() => 
                    fileProcessor.ProcessFilepaths(
                        filepaths, progressReport, cancellationSource.Token), 
                            cancellationSource.Token);

                // Run the task
                processedFilepaths = await task2;

                // If task was completed
                if (task2.IsCompleted)
                    imageFileExistsStatus.Source =
                        new BitmapImage(new Uri(new Checkmark().Path(), UriKind.Relative));
            }
            catch (Exception ex) when (ex is OpenXmlPackageException ||
                                       ex is ArgumentException ||
                                       ex is IOException ||
                                       ex is FileFormatException ||
                                       ex is OperationCanceledException)
            {
                PrintMessage(ex.Message);
            }
            finally
            {
                // Go to the last item in the listbox
                listboxResultsWindow.SelectedIndex = listboxResultsWindow.Items.Count - 1;
                listboxResultsWindow.ScrollIntoView(listboxResultsWindow.SelectedItem);

                // Get the amount of missing files
                int missingAmount = (from file in processedFilepaths
                                     where file.FileExists == false
                                     select file).Count();

                // Print results
                PrintMessage($"DONE! \n" +
                             $"Time elapsed: {timer.Elapsed.ToString("hh\\:mm\\:ss", CultureInfo.InvariantCulture)}\n" +
                             $"Filepaths validated: {processedFilepaths.Count}\n" +
                             $"Missing files: {missingAmount} \n" +
                             $"Log file has been created in the application folder.");

                // Enable/disable UI buttons
                buttonStart.IsEnabled = true;
                buttonStop.IsEnabled = false;
                buttonStart.Visibility = Visibility.Visible;
                buttonStop.Visibility = Visibility.Hidden;

                // Stop the timer
                timer.Stop();
            }
        }

        /// <summary>
        /// Starts the open-file-dialog and let's the user choose a file to be opened.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFile_Clicked(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (.xlsx)|*.xlsx";

            if (openFileDialog.ShowDialog() == true)
            {
                excelFilepath = openFileDialog.FileName;

                if (!string.IsNullOrWhiteSpace(excelFilepath))
                {
                    excelFilename = Path.GetFileNameWithoutExtension(excelFilepath);
                }

                // Show filename on the UI
                labelSelectedFile.Content = excelFilename;

                // Enable Start button
                buttonStart.IsEnabled = true;
            }
        }

        /// <summary>
        /// Tries to cancel any ongoing tasks
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Stop_Clicked(object sender, RoutedEventArgs e)
        {
            try
            {
                // Call cancellation token to cancel any ongoing tasks.
                cancellationSource.Cancel();
            }
            catch (ObjectDisposedException ex)
            {
                listboxResultsWindow.Items.Add(new ResultMessage
                {
                    Content = ex.Message
                });
            }
            catch (AggregateException ex)
            {
                listboxResultsWindow.Items.Add(new ResultMessage
                {
                    Content = ex.Message
                });
            }
            finally
            {
                // Update UI buttons
                buttonStart.IsEnabled = true;
                buttonStop.IsEnabled = false;
                buttonStart.Visibility = Visibility.Visible;
                buttonStop.Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        /// Updates the progress bar when reading the Excel-file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateProgressBar2(object sender, ProgressStatus e)
        {
            progressBar2.Value = e.PercentageCompleted;
        }

        /// <summary>
        /// Updates the progress bar when checking if the files exist
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateProgressBar1(object sender, ProgressStatus e)
        {
            progressBar1.Value = e.PercentageCompleted;
        }

        /// <summary>
        /// Prints the specified text to the UI
        /// </summary>
        /// <param name="message"></param>
        private void PrintMessage(string message)
        {
            listboxResultsWindow.Items.Add(new ResultMessage
            {
                Content = message
            });
        }

        /// <summary>
        /// Resets the state of the application
        /// </summary>
        private void ResetAppState()
        {
            cancellationSource = new CancellationTokenSource();
            processedFilepaths.Clear();
            listboxResultsWindow.Items.Clear();
            labelProgressBar1.Content = "";
            labelProgressBar2.Content = "";
            progressBar1.Value = 0;
            progressBar2.Value = 0;
            imageFileExistsStatus.Source = new BitmapImage(new Uri(new RedCross().Path(), UriKind.Relative));
            imageFileReadStatus.Source = new BitmapImage(new Uri(new RedCross().Path(), UriKind.Relative));
        }
    }
}

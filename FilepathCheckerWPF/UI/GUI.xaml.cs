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
        Progress<ProgressReportModelV2> progressModelV2 = new Progress<ProgressReportModelV2>();
        Progress<ProgressReportModel> progressModel = new Progress<ProgressReportModel>();
        private static CancellationTokenSource cancellationSource = new CancellationTokenSource();
        private static List<IFileModel> processedFilepaths = new List<IFileModel>();
        private static string excelFilepath = "";
        private static string excelFilename = "";

        public GUI()
        {
            InitializeComponent();
            this.Title = "Filepath Checker";
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
        /// Executes the application.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Start_Clicked(object sender, RoutedEventArgs e)
        {
            // Reset the application state and progress bars
            ResetAppState();

            // If user did not provide input into the column selection box
            if (string.IsNullOrWhiteSpace(textboxSelectedColumn.Text))
            {
                // Add a result message to the results wwindow
                listboxResultsWindow.Items.Add(new ResultMessage
                {
                    Content = "Please specify a column."
                });

                // Stop execution
                return;
            }

            // Get the column letter
            string columnLetter = textboxSelectedColumn.Text.ToUpper(CultureInfo.InvariantCulture);

            // Update UI buttons
            buttonStart.IsEnabled = false;
            buttonStop.IsEnabled = true;
            buttonStart.Visibility = Visibility.Hidden;
            buttonStop.Visibility = Visibility.Visible;

            // Objects for transferring information about the progress of the ongoing tasks
            progressModelV2.ProgressChanged += UpdateProgressBar1;
            progressModel.ProgressChanged += UpdateProgressBar2;

            // Create a timer and start timing
            Stopwatch timer = Stopwatch.StartNew();

            try
            {
                // Open the Excel spreadsheet
                using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilepath, false);

                // Create a file processor
                FileProcessor fileProcessor = new FileProcessor(spreadsheetDocument, () => new FileModel(), new CsvLogger());

                // Options object for cancelling parallel foreach loops
                ParallelOptions cancellationOptions = new ParallelOptions
                {
                    CancellationToken = cancellationSource.Token,
                    MaxDegreeOfParallelism = Environment.ProcessorCount
                };

                // Start reading a specific column from the spreadsheet
                labelProgressBar1.Content = "Reading the file ...";
                List<string> filepaths = await fileProcessor.ReadColumnSaxAsync(
                                               columnLetter,
                                               progressModelV2,
                                               cancellationOptions)
                                               .ConfigureAwait(true);

                // File was read successfully
                imageFileReadStatus.Source = new BitmapImage(new Uri(new Checkmark().Path(), UriKind.Relative));

                // Process the filepaths
                labelProgressBar2.Content = "Checking filepaths ...";
                processedFilepaths = await fileProcessor.ProcessFilepaths(
                                           filepaths,
                                           progressModel,
                                           cancellationOptions)
                                           .ConfigureAwait(true);

                // Update image
                imageFileExistsStatus.Source = new BitmapImage(new Uri(new Checkmark().Path(), UriKind.Relative));
            }
            catch (Exception ex) when (ex is OpenXmlPackageException || 
                                       ex is ArgumentException || 
                                       ex is IOException || 
                                       ex is FileFormatException ||
                                       ex is OperationCanceledException)
            {
                // Print the exception message to the UI
                listboxResultsWindow.Items.Add(new ErrorMessage
                {
                    Content = ex.Message
                });
            }
            finally
            {
                // Go to the last item in the listbox
                listboxResultsWindow.SelectedIndex = listboxResultsWindow.Items.Count - 1;
                listboxResultsWindow.ScrollIntoView(listboxResultsWindow.SelectedItem);

                // Stop timing
                timer.Stop();

                // Get the amount of missing files
                int missingAmount = (from file in processedFilepaths
                                     where file.FileExists == false
                                     select file).Count();

                // Print results to the UI
                listboxResultsWindow.Items.Add(new ResultMessage
                {
                    Content = $"DONE! \n" +
                    $"Time elapsed: {timer.Elapsed.ToString("hh\\:mm\\:ss", CultureInfo.InvariantCulture)}\n" +
                    $"Filepaths checked: {processedFilepaths.Count}\n" +
                    $"Missing files: {missingAmount} \n" +
                    $"Log file has been created in the application folder."
                });

                // Enable/disable UI buttons
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
        private void UpdateProgressBar2(object sender, ProgressReportModel e)
        {
            progressBar2.Value = e.PercentageCompleted;
        }

        /// <summary>
        /// Updates the progress bar when checking if the files exist
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateProgressBar1(object sender, ProgressReportModelV2 e)
        {
            progressBar1.Value = e.PercentageCompleted;
        }

        /// <summary>
        /// Resets the state of the application
        /// </summary>
        private void ResetAppState()
        {
            cancellationSource = new CancellationTokenSource();
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

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
        private static string excelFilepath = "";
        private static string excelFilename = "";

        public GUI()
        {
            InitializeComponent();
            this.Title = "Filepath Checker";
        }

        /// <summary>
        /// Starts the open-file-dialog and let's the user choose an Excel-file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFile_Clicked(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                excelFilepath = openFileDialog.FileName;

                if (!String.IsNullOrWhiteSpace(excelFilepath))
                {
                    excelFilename = System.IO.Path.GetFileNameWithoutExtension(excelFilepath);
                }

                labelSelectedFile.Content = excelFilename;
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
        /// Starts the process of reading the opened Excel-file
        /// Iterates through all the rows in selected column and checks if files exist
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Start_Clicked(object sender, RoutedEventArgs e)
        {
            // Reset the application state and progress bars first, 
            // so that we are not using any previous data
            AppReset();
            progressBar1.Value = 0;
            progressBar2.Value = 0;

            // Check that the user has provided input
            if (String.IsNullOrWhiteSpace(textboxSelectedColumn.Text))
            {
                listboxResultsWindow.Items.Add(new ResultMessage
                {
                    Content = "Please specify a column."
                });

                return;
            }

            // Get the column letter that the user specified.
            string column = textboxSelectedColumn.Text.ToUpper(System.Globalization.CultureInfo.InvariantCulture);

            // Update UI buttons
            buttonStart.IsEnabled = false;
            buttonStop.IsEnabled = true;
            buttonStart.Visibility = Visibility.Hidden;
            buttonStop.Visibility = Visibility.Visible;

            // Objects for transferring information about the progress of the ongoing tasks
            Progress<ProgressReportModelV2> progressModelV2 = new Progress<ProgressReportModelV2>();
            Progress<ProgressReportModel> progressModel = new Progress<ProgressReportModel>();
            progressModelV2.ProgressChanged += UpdateProgressBar1;
            progressModel.ProgressChanged += UpdateProgressBar2;

            // Object for cancelling parallel foreach loops
            ParallelOptions cancellationOptions = new ParallelOptions();
            cancellationOptions.CancellationToken = cancellationSource.Token;
            cancellationOptions.MaxDegreeOfParallelism = System.Environment.ProcessorCount;

            try
            {
                // Create a logger and open the Excel spreadsheet.
                using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilepath, false);

                // Create a file processor
                FileProcessor fileProcessor = new FileProcessor(spreadsheetDocument, new FileModelV1(), new CsvLogger());

                // Create a timer and start timing
                Stopwatch timer = Stopwatch.StartNew();

                // Start reading a specific column from the spreadsheet
                labelProgressBar1.Content = "Reading the file ...";
                List<string> filepaths = await fileProcessor.ReadColumnSaxAsync(
                    column,
                    progressModelV2,
                    cancellationOptions)
                    .ConfigureAwait(true);

                // File was read successfully.
                imageFileReadStatus.Source = new BitmapImage(new Uri(new Checkmark().Path(), UriKind.Relative));

                labelProgressBar2.Content = "Checking filepaths ..."; // Check if filepaths exist
                List<IFileModel> processedFilepaths = await fileProcessor.ProcessFilepaths(
                    filepaths,
                    progressModel,
                    cancellationOptions)
                    .ConfigureAwait(true);

                // Processing done.
                imageFileExistsStatus.Source = new BitmapImage(new Uri(new Checkmark().Path(), UriKind.Relative));

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
            }
            catch (Exception ex) when
            (ex is OpenXmlPackageException
                || ex is ArgumentException
                || ex is IOException
                || ex is FileFormatException)
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

                // Enable/disable UI buttons
                buttonStart.IsEnabled = true;
                buttonStop.IsEnabled = false;
                buttonStart.Visibility = Visibility.Visible;
                buttonStop.Visibility = Visibility.Hidden;

                // Release resources (free-up ram)
                //GC.Collect();
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
        private void AppReset()
        {
            cancellationSource = new CancellationTokenSource();

            listboxResultsWindow.Items.Clear();
            labelProgressBar1.Content = "";
            labelProgressBar2.Content = "";

            imageFileExistsStatus.Source = new BitmapImage(new Uri(new RedCross().Path(), UriKind.Relative));
            imageFileReadStatus.Source = new BitmapImage(new Uri(new RedCross().Path(), UriKind.Relative));
        }
    }
}

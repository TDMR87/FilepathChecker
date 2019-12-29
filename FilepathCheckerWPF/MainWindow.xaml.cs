using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace FilepathCheckerWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static CancellationTokenSource cancellationSource = new CancellationTokenSource();
        private static List<string> allFilepaths = new List<string>();
        private static List<IFileModel> listOfFilesNotExist = new List<IFileModel>();

        private static string openedFile_Path = "";
        private static string openedFile_Name = "";
        private static string logFileUNCPath = "";

        private static int processedFilesCount = 0;

        public MainWindow()
        {
            InitializeComponent();
            this.Title = "Filepath Checker";
        }

        /// <summary>
        /// Starts the open file dialog for opening an Excel-file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFile_Clicked(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                openedFile_Path = openFileDialog.FileName;

                if (!String.IsNullOrWhiteSpace(openedFile_Path))
                {
                    openedFile_Name = System.IO.Path.GetFileNameWithoutExtension(openedFile_Path);
                }

                labelSelectedFile.Content = openedFile_Name;
                buttonStart.IsEnabled = true;
            }
        }

        /// <summary>
        /// Tries to set the cancellation token and resets the application state
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Stop_Clicked(object sender, RoutedEventArgs e)
        {
            try
            {
                // Set cancellation token to cancel any ongoing tasks.
                cancellationSource.Cancel();
            }
            catch (ObjectDisposedException ex)
            {
                listboxFilepaths.Items.Add(new FileModel
                {
                    Filepath = ex.Message
                });
            }
            catch (AggregateException ex)
            {
                listboxFilepaths.Items.Add(new FileModel
                {
                    Filepath = ex.Message
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

            // Input sanitation check
            if (String.IsNullOrWhiteSpace(textboxSelectedColumn.Text))
            {
                MessageBox.Show("Please specify a column.");
                return;
            }

            // Get the user-specified column letter
            string column = textboxSelectedColumn.Text.ToUpper(System.Globalization.CultureInfo.InvariantCulture);

            // Update UI buttons
            buttonStart.IsEnabled = false;
            buttonStop.IsEnabled = true;
            buttonStart.Visibility = Visibility.Hidden;
            buttonStop.Visibility = Visibility.Visible;

            // Objects for transferring information about the progress of the ongoing tasks
            Progress<ProgressReportModelV2> progress = new Progress<ProgressReportModelV2>();
            Progress<ProgressReportModel> checkExistsProgress = new Progress<ProgressReportModel>();
            progress.ProgressChanged += UpdateProgressBar1;
            checkExistsProgress.ProgressChanged += UpdateProgressBar2;

            // Object for cancelling parallel foreach loops
            ParallelOptions parallelOptions = new ParallelOptions();
            parallelOptions.CancellationToken = cancellationSource.Token;
            parallelOptions.MaxDegreeOfParallelism = System.Environment.ProcessorCount;

            // Start timing
            Stopwatch timer = Stopwatch.StartNew();

            // Update progress bar label
            labelReadFileProgressStatus.Content = "Reading the file";

            // Start a task of opening the excel-file and start reading the filepaths from the user-specified column
            // Extracts the filepaths into a collection
            allFilepaths = await FileReader.ReadFileUsingOpenXMLAsync(
                openedFile_Path, 
                column, 
                progress, 
                parallelOptions).ConfigureAwait(true);

            // File read done. Update progress bar image and label
            imageFileReadStatus.Source = new BitmapImage(new Uri(ImageModel.ImageFound, UriKind.Relative));
            labelFileExistsProgressStatus.Content = "Checking filepaths";

            // If any filepaths are found, check each filepath if it exists
            if (allFilepaths.Count > 0)
            {
                // Progress bar and logger objects
                ProgressReportModel report = new ProgressReportModel();
                CsvLogger logger = new CsvLogger();

                // Start a new task for iterating through all the filepaths. 
                // Task is needed here so we can cancel the process by calling
                // the cancellation token
                await Task.Run(async() =>
                {
                    try
                    {
                        foreach (string path in allFilepaths)
                        {
                            // Ignore possible empty paths
                            if (String.IsNullOrWhiteSpace(path)) { continue; }

                            // Cancel if user presses Stop
                            cancellationSource.Token.ThrowIfCancellationRequested();

                            // Wrap each filepath into a IFileModel object
                            // Sets the object's FileExists-property to true/false
                            IFileModel file = await FileReader.CheckFileModelExistsAsync(path).ConfigureAwait(true);

                            // If file does not exist, add it to a list
                            if (!file.FileExists)
                            {
                                listOfFilesNotExist.Add(file);

                                // Log it
                                await logger.WriteLineAsync(path).ConfigureAwait(true);
                            }

                            // Increment the count of processed filepaths for updating the progress bar
                            processedFilesCount++;

                            // Report progress after each processed filepath
                            report.PercentageCompleted = (processedFilesCount * 100) / allFilepaths.Count;
                            SendProgressReport(checkExistsProgress, report);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        return;
                    }
                    
                }).ConfigureAwait(true);

                // Close the logger
                logFileUNCPath = CsvLogger.GetPath();
                logger.Close(); 
                logger.Dispose();
            }

            // Stop timing the process
            timer.Stop();

            // Filepaths checking is finished or stopped. Set progressbar image.
            imageFileExistsStatus.Source = new BitmapImage(new Uri(ImageModel.ImageFound, UriKind.Relative));

            // Print results to the UI
            listboxFilepaths.Items.Add(new FileModel
            {
                Filepath = $"DONE! \n" +
                $"Time elapsed: {timer.Elapsed} ms.\n" +
                $"Filepaths checked: {allFilepaths.Count}\n" +
                $"Missing files: {listOfFilesNotExist.Count}"
            });

            // Print the location of the log file
            listboxFilepaths.Items.Add(new FileModel
            {
                Filepath = $"Log file has been created in:\n {logFileUNCPath}"
            });

            // Go to the last item in the listbox
            listboxFilepaths.SelectedIndex = listboxFilepaths.Items.Count - 1;
            listboxFilepaths.ScrollIntoView(listboxFilepaths.SelectedItem);

            // Enable/disable UI buttons
            buttonStart.IsEnabled = true;
            buttonStop.IsEnabled = false;
            buttonStart.Visibility = Visibility.Visible;
            buttonStop.Visibility = Visibility.Hidden;
        }

        /// <summary>
        /// Method for passing through an instance of IProgress interface object
        /// </summary>
        /// <param name="progress"></param>
        /// <param name="report"></param>
        void SendProgressReport(IProgress<ProgressReportModel> progress, ProgressReportModel report)
        {
            progress.Report(report);
        }

        /// <summary>
        /// Progress bar for updating the status of reading the Excel-file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateProgressBar2(object sender, ProgressReportModel e)
        {
            progressBar2.Value = e.PercentageCompleted;
        }

        /// <summary>
        /// Progress bar for updating the status of iterating the rows int he Excel-file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateProgressBar1(object sender, ProgressReportModelV2 e)
        {
            progressBar1.Value = e.PercentageCompleted;
        }

        /// <summary>
        /// Reset the state of the application
        /// </summary>
        private void AppReset()
        {
            cancellationSource = new CancellationTokenSource();
            allFilepaths.Clear();
            listOfFilesNotExist.Clear();
            listboxFilepaths.Items.Clear();
            processedFilesCount = 0;
            imageFileExistsStatus.Source = new BitmapImage(new Uri(ImageModel.ImageNotFound, UriKind.Relative));
            imageFileReadStatus.Source = new BitmapImage(new Uri(ImageModel.ImageNotFound, UriKind.Relative));
        }
    }
}

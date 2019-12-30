using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
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
        private static string excelFile_Path = "";
        private static string excelFile_Name = "";

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
                excelFile_Path = openFileDialog.FileName;

                if (!String.IsNullOrWhiteSpace(excelFile_Path))
                {
                    excelFile_Name = System.IO.Path.GetFileNameWithoutExtension(excelFile_Path);
                }

                labelSelectedFile.Content = excelFile_Name;
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

            // Input check
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
            Progress<ProgressReportModelV2> openFileProcess = new Progress<ProgressReportModelV2>();
            Progress<ProgressReportModel> processFilepathsProgress = new Progress<ProgressReportModel>();
            openFileProcess.ProgressChanged += UpdateProgressBar1;
            processFilepathsProgress.ProgressChanged += UpdateProgressBar2;

            // Object for cancelling parallel foreach loops
            ParallelOptions parallelOptions = new ParallelOptions();
            parallelOptions.CancellationToken = cancellationSource.Token;
            parallelOptions.MaxDegreeOfParallelism = System.Environment.ProcessorCount;

            // Start timing
            Stopwatch timer = Stopwatch.StartNew();

            // Start opening the excel-file and start reading the filepaths from the user-specified column
            labelProgressBar1.Content = "Reading the file ...";
            List<string> filepaths = await FileProcessor.ReadFileUsingOpenXMLAsync(
                excelFile_Path, 
                column, 
                openFileProcess, 
                parallelOptions)
                .ConfigureAwait(true);

            // File read done.
            imageFileReadStatus.Source = new BitmapImage(new Uri(ImageModel.ImageFound, UriKind.Relative));

            // Process the filepaths and resolve them into IFileModel objects
            labelProgressBar2.Content = "Checking filepaths ...";
            List<IFileModel> processedFiles = await FileProcessor.ProcessFilepaths(
                filepaths, 
                processFilepathsProgress, 
                parallelOptions)
                .ConfigureAwait(true);

            // Processing done.
            imageFileExistsStatus.Source = new BitmapImage(new Uri(ImageModel.ImageFound, UriKind.Relative));

            // Stop timing
            timer.Stop();

            // Get the amount of missing files
            int missing = processedFiles.Where(file => !file.FileExists).Count();

            // Print results to the UI
            listboxFilepaths.Items.Add(new FileModel
            {
                Filepath = $"DONE! \n" +
                $"Time elapsed: {timer.Elapsed} ms.\n" +
                $"Filepaths checked: {filepaths.Count}\n" +
                $"Missing files: {missing}" + 
                $"Check the log file in the application folder."
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
        /// Updates the progress bar status of reading the Excel-file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateProgressBar2(object sender, ProgressReportModel e)
        {
            progressBar2.Value = e.PercentageCompleted;
        }

        /// <summary>
        /// Updates the progress bar status of iterating through the extracted filepaths
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

            listboxFilepaths.Items.Clear();
            labelProgressBar1.Content = "";
            labelProgressBar2.Content = "";

            imageFileExistsStatus.Source = new BitmapImage(new Uri(ImageModel.ImageNotFound, UriKind.Relative));
            imageFileReadStatus.Source = new BitmapImage(new Uri(ImageModel.ImageNotFound, UriKind.Relative));
        }
    }
}

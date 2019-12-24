using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace FilepathChecker
{
    class Program
    {
        static List<string> pathsNotFound = new List<string>();
        static List<string> filepathsList = new List<string>();
        static string appPath = AppDomain.CurrentDomain.BaseDirectory;
        static string logFilename = "";
        static int found = 0;
        static int missing = 0;
        static int counter = 0;

        static async Task Main(string[] args)
        {
            Console.WriteLine("\n** Filepath checker **");
            Console.WriteLine("\nHint: Remember to encode your csv-file as UTF-8.");

            Start:
            Console.Write("\nEnter csv filepath: ");
            string filepath = Console.ReadLine();

            while (true)
            {
                if (!File.Exists(filepath))
                {
                    Console.WriteLine("\n** File not found **\n");
                    goto Start;
                }

                Console.WriteLine($"\nChecking filepaths...\n");
                Thread.Sleep(1500);

                // Read the file line by line and check each filepath.
                using (StreamReader reader = new StreamReader(filepath, Encoding.UTF8))
                {
                    string line;
                    int i = 1;

                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        // Do not handle empty filepaths
                        if (line != "")
                        {
                            if (line.Contains("|"))
                            {
                                SplitFilepathsAndCheckIfExists(line);
                            }
                            else
                            {
                                CheckIfExists(line);
                            }
                        }
                    }
                }

                // Create log file of missing files, if any.
                if (missing > 0)
                {
                    logFilename = $"{DateTime.Now.ToLongDateString()}_{DateTime.Now.ToLongTimeString()}";

                    using (StreamWriter writer = new StreamWriter($"{appPath}\\{logFilename}.csv"))
                    {
                        writer.WriteLine("Error;Filepath");

                        foreach (string path in pathsNotFound)
                        {
                            writer.WriteLine($"Source file not found;{path}");
                        }
                    }
                }

                // Print results.
                Console.WriteLine("\n*******************************************\n");
                Console.WriteLine($"{found} file(s) found.\n");
                Console.WriteLine($"{missing} file(s) missing.\n");
                Console.WriteLine($"Log file of the missing files has been created in: {appPath}.");
                Console.WriteLine("\n*******************************************\n");
                Console.ReadKey();
            }
        }

        private static void CheckIfExists(string path)
        {
            if (File.Exists(path))
            {
                Console.WriteLine($"{counter}: FOUND {path}");
                found++;
            }
            else
            {
                Console.WriteLine($"{counter}: NOT FOUND {path}");
                pathsNotFound.Add(path);
                missing++;
            }
            counter++;
        }

        private static void SplitFilepathsAndCheckIfExists(string line)
        {
            List<string> paths = line.Split('|').ToList();

            foreach (string path in paths)
            {
                if (File.Exists(path))
                {
                    Console.WriteLine($"{counter}: FOUND {path}");
                    found++;
                }
                else
                {
                    Console.WriteLine($"{counter}: NOT FOUND {path}");
                    pathsNotFound.Add(path);
                    missing++;
                }
                counter++;
            }
        }
    }
}

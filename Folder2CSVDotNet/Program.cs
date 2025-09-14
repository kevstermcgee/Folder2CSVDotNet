using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.IO.Hashing;
using FxResources.System.IO.Hashing;
using System.Security.Cryptography;
using CsvHelper;
using CsvHelper.Configuration;

namespace Folder2CSVDotNet
{
    public class Record
    {
        public required string Filelink { get; set; }
        public required string Filepath { get; set; }
        public required string Folder { get; set; }
        public required string Filename { get; set; }
        public DateTime ModifiedDate { get; set; }
        public long Size { get; set; }
        public required string PrettySize { get; set; }
        public required string Extension { get; set; }
        public required string FileHash { get; set; }
    }

    class Program
    {
        private static void Main()
        {
            Console.Write("Enter root folder path: ");
            var rootPath = Console.ReadLine()?.Trim('"', ' ');
            
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            if (string.IsNullOrWhiteSpace(rootPath))
            {
                Console.WriteLine("No folder path entered!");
                return;
            }

            // Normalize the path
            rootPath = Path.GetFullPath(rootPath);

            if (!Directory.Exists(rootPath))
            {
                Console.WriteLine($"Error: Root folder not found: {rootPath}");
                Console.WriteLine("Please check the path and ensure the folder is available offline.");
                return; // Exit the Main method cleanly
            }

            // Get the desktop path
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var csvFilePath = GetUniqueFilePath(desktopPath, "data", ".csv");
            Console.WriteLine($"CSV will be saved as: {csvFilePath}");
            
            // If CSV exists, delete it to start fresh
            if (File.Exists(csvFilePath))
            {
                File.Delete(csvFilePath);
            }

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
            };

            using (var writer = new StreamWriter(csvFilePath))
            using (var csv = new CsvWriter(writer, config))
            {
                // Write header once
                csv.WriteHeader<Record>();
                csv.NextRecord();

                // Start recursive scraping
                ScrapeFolder(rootPath, csv);
            }

            Console.WriteLine("CSV saved successfully!");
            stopwatch.Stop();
            var elapsed = stopwatch.Elapsed;
            Console.WriteLine($"Time taken: {elapsed.ToString(@"hh\:mm\:ss")}");
            Console.ReadLine();
        }

        private static void ScrapeFolder(string rootPath, CsvWriter csv)
        {
            var fileQueue = new ConcurrentQueue<string>();
            var directories = new Stack<string>();
            directories.Push(rootPath);

            // Traverse directories iteratively to avoid stack overflow
            while (directories.Count > 0)
            {
                var currentDir = directories.Pop();
                try
                {
                    foreach (var file in Directory.EnumerateFiles(currentDir))
                    {
                        fileQueue.Enqueue(file);
                    }
                    foreach (var dir in Directory.EnumerateDirectories(currentDir))
                    {
                        directories.Push(dir);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    Console.WriteLine("Access denied: " + ex.Message);
                }
            }

            // Limit the degree of parallelism to avoid overwhelming the system
            int maxDegreeOfParallelism = Math.Max(Environment.ProcessorCount - 1, 1);

            Parallel.ForEach(
                fileQueue,
                new ParallelOptions { MaxDegreeOfParallelism = maxDegreeOfParallelism },
                file =>
                {
                    var record = GetMetaData(file);
                    lock (csv)
                    {
                        csv.WriteRecord(record);
                        csv.NextRecord();
                    }
                });
        }
        
        private static string GetUniqueFilePath(string folderPath, string baseFileName, string extension)
        {
            string filePath = Path.Combine(folderPath, baseFileName + extension);
            int counter = 1;

            while (File.Exists(filePath))
            {
                filePath = Path.Combine(folderPath, $"{baseFileName}({counter++}){extension}");
            }

            return filePath;
        }

        private static string ComputeHash(string path)
        {
            try
            {
                using (var stream = File.OpenRead(path))
                {
                    // Creates an instance of the hash algorithm
                    var hasher = new XxHash64();

                    // Append the stream to the hasher
                    hasher.Append(stream);

                    // Convert to hex string
                    return Convert.ToHexString(hasher.GetHashAndReset()).ToLowerInvariant();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string FormatSize(long bytes)
        {
            const long KB = 1024;
            const long MB = KB * 1024;
            const long GB = MB * 1024;
            const long TB = GB * 1024;

            if (bytes >= TB)
                return $"{bytes / (double)TB:0.##} TB";
            if (bytes >= GB)
                return $"{bytes / (double)GB:0.##} GB";
            if (bytes >= MB)
                return $"{bytes / (double)MB:0.##} MB";
            if (bytes >= KB)
                return $"{bytes / (double)KB:0.##} KB";
            return $"{bytes} B";
        }


        private static Record GetMetaData(string path)
        {
            var fileInfo = new FileInfo(path);
            return new Record
            {
                Filelink = $"=HYPERLINK(\"{fileInfo.FullName}\", \"Open\")", // Excel formula
                Filepath = fileInfo.FullName,
                Folder = fileInfo.DirectoryName ?? string.Empty,
                Filename = fileInfo.Name,
                ModifiedDate = fileInfo.LastWriteTime,
                Size = fileInfo.Length,
                PrettySize = FormatSize(fileInfo.Length),
                Extension = fileInfo.Extension,
                // Below is great for file hashing but takes a long time
                FileHash = ComputeHash(fileInfo.FullName)
            };
        }
    }
}

using System.Diagnostics;
using System.Globalization;
using CsvHelper;

namespace Folder2CSVDotNet
{
    public class Record
    {
        public string Filepath { get; set; }
        public string Folder { get; set; }
        public string Filename { get; set; }
        public DateTime ModifiedDate { get; set; }
        public long Size { get; set; }
        public string Extension { get; set; }
    }

    class Program
    {
        static void Main()
        {
            Console.WriteLine("Enter root folder path: ");
            string rootPath = Console.ReadLine()?.Trim('"', ' ');
            
            Stopwatch stopwatch = new Stopwatch();
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

            string csvFilePath = GetUniqueFilePath(rootPath, "data", ".csv");
            Console.WriteLine($"CSV will be saved as: {csvFilePath}");
            
            // If CSV exists, delete it to start fresh
            if (File.Exists(csvFilePath))
                File.Delete(csvFilePath);

            using (var writer = new StreamWriter(csvFilePath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                // Write header once
                csv.WriteHeader<Record>();
                csv.NextRecord();

                // Start recursive scraping
                ScrapeFolder(rootPath, csv);
            }

            Console.WriteLine("CSV saved successfully!");
            stopwatch.Stop();
            Console.WriteLine($"Time taken: {stopwatch.ElapsedMilliseconds}ms");
            Console.ReadLine();
        }

        static void ScrapeFolder(string path, CsvWriter csv)
        {
            try
            {
                foreach (string file in Directory.GetFiles(path))
                {
                    var record = GetMetaData(file);
                    csv.WriteRecord(record);
                    csv.NextRecord(); // Move to next line
                }

                foreach (string dir in Directory.GetDirectories(path))
                {
                    ScrapeFolder(dir, csv); // Recursive call
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine("Access denied: " + ex.Message);
            }
        }
        
        static string GetUniqueFilePath(string folderPath, string baseFileName, string extension)
        {
            string filePath = Path.Combine(folderPath, baseFileName + extension);
            int counter = 1;

            while (File.Exists(filePath))
            {
                filePath = Path.Combine(folderPath, $"{baseFileName}({counter++}){extension}");
            }

            return filePath;
        }


        static Record GetMetaData(string path)
        {
            var fileInfo = new FileInfo(path);
            return new Record
            {
                Filepath = fileInfo.FullName,
                Folder = fileInfo.DirectoryName,
                Filename = fileInfo.Name,
                ModifiedDate = fileInfo.LastWriteTime,
                Size = fileInfo.Length,
                Extension = fileInfo.Extension
            };
        }
    }
}

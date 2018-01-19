using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SubmissionIndesignExtractor
{
    class Program
    {
        static int Main(string[] args)
        {
            var year = GetYear(args);
            var searchPath = $@"P:\{year}\{year.ToString().Substring(2, 2)}900_Submissions";

            var documents = GetDocumentsList(searchPath, year);

            var success = CopyDocuments(documents, year);

            if (success)
            {
                Console.WriteLine("Copy completed, press any key to exit");
                Console.ReadKey();
                return 0;
            }

            Console.WriteLine("Copy unsuccessful, press any key to exit");
            Console.ReadKey();
            return 1;
        }

        static int GetYear(string[] args)
        {
            var input = String.Empty;

            if (!args.Any())
            {
                Console.Write("Please enter a valid year and press enter: ");
                input = Console.ReadLine();
            }
            else
            {
                input = args[0];
            }

            var isInputNumeric = int.TryParse(input, out int year);

            while (!isInputNumeric || year < 2010)
            {
                Console.Write("Please enter a valid year and press enter: ");
                input = Console.ReadLine();
                isInputNumeric = int.TryParse(input, out year);
            }

            return year;
        }

        static IEnumerable<Document> GetDocumentsList(string searchPath, int year)
        {
            var documents = Directory.EnumerateFiles(searchPath, "*.indd", SearchOption.AllDirectories)
                                .Where(d => d.Contains("Indesign"))
                                .Select(d => new Document
                                {
                                    FullPath = d,
                                    FileName = d.Substring(d.LastIndexOf('\\') + 1),
                                    FolderName = d.Split('\\')[3]
                                });

            return documents;
        }

        static bool CopyDocuments(IEnumerable<Document> documents, int year)
        {
            var documentsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var pathRoot = $@"{documentsFolderPath}\InDesign Files\{year}\";

            var success = CreateFolderIfNotExists(pathRoot);

            if (!success)
            {
                return false;
            }

            foreach (var d in documents)
            {
                try
                {
                    var destinationPath = Path.Combine(pathRoot, $"{d.FolderName} ({d.FileName.Replace(".indd", ").indd")}");
                    for (var i = 1;  File.Exists(destinationPath); i++)
                    {
                        destinationPath = destinationPath.Replace((i > 1 ? (i - 1).ToString() : "") + ".indd", i + ".indd");
                    }
                    File.Copy(d.FullPath, destinationPath);
                }
                catch (IOException copyError)
                {
                    Console.WriteLine(copyError.Message);
                }
            }

            return true;
        }

        static bool CreateFolderIfNotExists(string path)
        {
            if (!Directory.Exists(path))
            {
                try
                {
                    Directory.CreateDirectory(path);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return false;
                }
            }

            return true;
        }
    }

    class Document
    {
        public string FullPath { get; set; }
        public string FileName { get; set; }
        public string FolderName { get; set; }
    }
}

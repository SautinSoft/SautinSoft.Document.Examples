using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text.RegularExpressions;



namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            string searchDir = Path.GetFullPath(@"..\..\..\searching\");
            string searchText = "with";
            FullTextSearching(searchDir, searchText);
        }

        /// <summary>
        /// This sample shows how to launch full text search in the specific directory.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/full-text-searching-in-documents-net-csharp-vb.php
        /// </remarks>
        public static void FullTextSearching(string searchPath, string searchText)
        {
            DirectoryInfo searchDir = new DirectoryInfo(searchPath);
            List<string> supportedFiles = new List<string>();

            // 1. Find theS files to make search.
            // Specify to make the search only in *.docx, *.rtf, *.pdf and *.html files,
            // including subdirectories.
            foreach (string file in Directory.GetFiles(searchDir.FullName, "*.*", SearchOption.AllDirectories))
            {
                string ext = Path.GetExtension(file).ToLower();

                if (ext == ".docx" || ext == ".pdf" || ext == ".html" || ext == ".rtf")
                    supportedFiles.Add(file);
            }

            // 2. Perform the text search in the each file using a loop.
            // We'll search the word "video" in the each and count how many times the file contains it.
            Console.WriteLine($"The results for \"{searchText}\":");

            int totalFiles = 0, totalMatches = 0;
            foreach (string file in supportedFiles)
            {
                DocumentCore dc = DocumentCore.Load(file);
                totalFiles++;
                Regex regex = new Regex($"\\b({searchText})\\b", RegexOptions.IgnoreCase);

                // Show also subfolder if we aren't in the root folder.
                DirectoryInfo dirInfo = new DirectoryInfo(Path.GetDirectoryName(file));
                string fileName = String.Empty;

                if (dirInfo.FullName.TrimEnd(new char[] { '\\' }) != searchDir.FullName.TrimEnd(new char[] { '\\' }))
                    fileName = file.Substring(searchPath.Length, file.Length - searchPath.Length);
                else
                    // We are in the root folder.
                    fileName = Path.GetFileName(file);

                int matches = dc.Content.Find(regex).Count();
                totalMatches += matches;

                Console.WriteLine($"{totalFiles:D3} from {supportedFiles.Count} {fileName} - {matches} matches.");
            }
            Console.WriteLine($"\nSearching finished. {supportedFiles.Count} file(s) has been processed. Total matches: {totalMatches}.");
            Console.WriteLine("Press any key ...");
            Console.ReadKey();
        }
    }
}
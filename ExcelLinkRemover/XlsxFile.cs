using System;
using System.IO;
using System.IO.Compression;

namespace ExcelLinkRemover
{
    /// <summary>
    /// Class for working with Xlsx file.
    /// </summary>
    public class XlsxFile
    {
        /// <summary>
        /// Removes all external links in the .xlsx file.
        /// </summary>
        /// <param name="fileName">Full path to the .xlsx file.</param>
        /// <returns>Returns True if the operation completed successfully or False otherwise.</returns>
        public static bool Process(string fileName)
        {
            var tempPath = Path.GetTempPath() + "ExcelLinkRemover\\" + new Random().Next();
            var newPath = new DirectoryInfo(fileName).Parent.FullName + "\\ExcelLinkRemover\\";

            // Unzip it into a temporary folder.
            try
            {
                ZipFile.ExtractToDirectory(fileName, tempPath);
            }
            catch (IOException ex)
            {
                Console.WriteLine(Environment.NewLine + ex.Message.Replace($" \"{fileName}\"", ""));
                return false;
            }

            // Document content processing.
            int result = Workbook.Process(tempPath);

            // Save in new file if changes have been made.
            if (result > 0)
            {
                var newFile = new FileInfo(fileName).Name;
                try
                {
                    if (!Directory.Exists(newPath))
                    {
                        Directory.CreateDirectory(newPath);
                    }

                    if (File.Exists(newPath + newFile))
                    {
                        File.Delete(newPath + newFile);
                        return false;
                    }
                    ZipFile.CreateFromDirectory(tempPath, newPath + newFile);
                }
                catch (IOException ex)
                {
                    Console.WriteLine(Environment.NewLine + ex.Message.Replace($" \"{fileName}\"", ""));
                    return false;
                }
            }
            else
            {
                Console.WriteLine(Environment.NewLine + "No external links found.");
            }

            // Delete temporary files.
            try
            {
                Directory.Delete(tempPath, true);
            }
            catch
            { }

            return true;
        }
    }
}
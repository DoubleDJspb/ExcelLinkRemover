using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelLinkRemover
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                //args = new string[] { @"Debug.xlsx" };
            }

            if (args != null && args.Length > 0)
            {
                var files = new List<string>();
                for (int i = 0; i < args.Length; i++)
                {
                    if (File.Exists(args[i]) && args[i].EndsWith(".xlsx"))
                    {
                        files.Add(args[i]);
                    }
                }

                if (files.Count == 0)
                {
                    Console.WriteLine("No files found.");
                }
                else
                {
                    Console.WriteLine("Processing in progress:");

                    foreach (var file in files)
                    {
                        Console.Write(file + "... ");

                        if (XlsxFile.Process(file))
                        {
                            Console.WriteLine("OK");
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }

            Console.WriteLine(Environment.NewLine + "Press any key to exit...");
            Console.ReadKey();
        }
    }
}

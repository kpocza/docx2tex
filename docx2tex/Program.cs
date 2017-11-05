using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Configuration;
using System.Diagnostics;
using docx2tex.Library;

namespace docx2tex
{
    class Program
    {
        static void Main(string[] args)
        {
            ConsoleOutput consoleOutput = new ConsoleOutput();

            consoleOutput.WriteLine("docx2tex was created by Krisztian Pocza in 2007-2008 under the terms of the BSD licence");
            consoleOutput.WriteLine("info: kpocza@kpocza.net");
            consoleOutput.WriteLine();
//            Console.ReadLine();
            if (args.Length < 2)
            {
                consoleOutput.WriteLine("Usage:");
                consoleOutput.WriteLine("docx2tex.exe source.docx dest.tex");
                return;
            }

            string inputDocxPath = ResolveFullPath(args[0]);
            string outputTexPath = ResolveFullPath(args[1]);

            StaticConfigHelper.DocxPath = inputDocxPath;

            if (new Docx2TexWorker().Process(inputDocxPath, outputTexPath, consoleOutput))
            {
                using (StreamReader sr = new StreamReader(outputTexPath))
                {
                    Console.WriteLine(sr.ReadToEnd());
                }
            }
        }

        private static string ResolveFullPath(string path)
        {
            if (Path.IsPathRooted(path))
            {
                return path;
            }
            return Path.Combine(Environment.CurrentDirectory, path);
        }
    }
}

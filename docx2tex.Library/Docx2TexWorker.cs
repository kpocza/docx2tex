using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.IO.Packaging;

namespace docx2tex.Library
{
    public class Docx2TexWorker
    {
        public bool Process(string inputDocxPath, string outputTexPath, IStatusInformation statusInfo)
        {
            string documentPath = Path.GetDirectoryName(outputTexPath);
            if (documentPath == null)
            {
                documentPath = Path.GetPathRoot(outputTexPath);
            }

            EnsureMediaPath(documentPath);

            statusInfo.WriteCR("Opening document...");
            Package pkg = null;
            try
            {
                pkg = Package.Open(inputDocxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            }
            catch (Exception ex)
            {
                // this happens mostly when the user leaves the Word file open
                statusInfo.WriteLine(ex.Message);
                return false;
            }

            ZipPackagePart documentPart = (ZipPackagePart)pkg.GetPart(new Uri("/word/document.xml", UriKind.Relative));

            //numbering part may not exist for simple documents
            ZipPackagePart numberingPart = null;
            if (pkg.PartExists(new Uri("/word/numbering.xml", UriKind.Relative)))
            {
                numberingPart = (ZipPackagePart)pkg.GetPart(new Uri("/word/numbering.xml", UriKind.Relative));
            }

            Numbering numbering = new Numbering(numberingPart);
            Imaging imaging = new Imaging(documentPart, inputDocxPath, outputTexPath);

            using (Stream documentXmlStream = documentPart.GetStream())
            {
                Engine engine = new Engine(documentXmlStream, numbering, imaging, statusInfo);
                statusInfo.WriteLine("Document opened.        ");

                string outputString = engine.Process();
                string latexSource = ReplaceSomeCharacters(outputString);

                Encoding encoding = Encoding.Default;

                var enc = docx2tex.Library.Data.InputEnc.Instance.CurrentEncoding;
                if (enc != null)
                {
                    encoding = Encoding.GetEncoding(enc.DotNetEncoding);
                }

                byte[] data = encoding.GetBytes(latexSource);
                using (FileStream fs = new FileStream(outputTexPath, FileMode.Create, FileAccess.Write))
                {
                    using (BinaryWriter outputTexStream = new BinaryWriter(fs))
                    {
                        outputTexStream.Write(data);
                    }
                }
            }
            pkg.Close();
            return true;
        }

        private static string ReplaceSomeCharacters(string latexSource)
        {
            latexSource = latexSource.Replace("!!!DOLLARSIGN!!!", "\\$");
            return latexSource;
        }

        private static void EnsureMediaPath(string documentPath)
        {
            string mediaPath = Path.Combine(documentPath, "media");
            if (!Directory.Exists(mediaPath))
            {
                Directory.CreateDirectory(mediaPath);
            }
        }

    }
}

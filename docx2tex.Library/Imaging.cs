using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Packaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Diagnostics;

namespace docx2tex.Library
{
    class Imaging
    {
        private ZipPackagePart _docRelPart;
        private string _documentDirectory;
        private string _latexDirectory;
        private string _inputDocxPath;

        public Imaging(ZipPackagePart docRelPart, string inputDocxPath, string outputLatexPath)
        {
            _docRelPart = docRelPart;
            _inputDocxPath = inputDocxPath;
            _documentDirectory = Path.GetDirectoryName(inputDocxPath);
            _latexDirectory = Path.GetDirectoryName(outputLatexPath);
        }

        public string ResolveImage(string imageId, IStatusInformation statusInfo)
        {
            PackageRelationship rs = _docRelPart.GetRelationship(imageId);

            string imageUrl = rs.TargetUri.OriginalString;
            string orginalImagePath = Path.Combine(_latexDirectory, imageUrl);
            string newImagePath = Path.Combine(_latexDirectory, imageUrl);

            ZipPackagePart imagePackagePart = (ZipPackagePart)rs.Package.GetPart(new Uri("/word/" + imageUrl, UriKind.Relative));

            using (Stream contentStream = imagePackagePart.GetStream())
            {
                byte[] content = new byte[contentStream.Length];
                contentStream.Read(content, 0, (int)contentStream.Length);

                using (FileStream fs = new FileStream(orginalImagePath, FileMode.Create))
                {
                    using (BinaryWriter bwImage = new BinaryWriter(fs))
                    {
                        bwImage.Write(content, 0, (int)contentStream.Length);
                    }
                }
            }

            ConvertImageToEPS(orginalImagePath, newImagePath, statusInfo);

            return Path.ChangeExtension(imageUrl, "eps"); ;
        }

        public string GetWidthAndHeightFromStyle(string style)
        {
            try
            {
                Regex styleRegEx = new Regex("width:(?<Width>.+?);height:(?<Height>.+?)(;|$)", RegexOptions.Compiled);
                Match match = styleRegEx.Match(style);
                if (match.Success)
                {
                    return string.Format("width={0},height={1}", match.Groups["Width"].Value, match.Groups["Height"].Value);
                }
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public string GetWidthAndHeightFromStyle(int? cx, int? cy)
        {
            try
            {
                if (cx.HasValue && cy.HasValue)
                {
                    string width = string.Format("{0:f2}", (float)cx / 360100).Replace(',', '.');
                    string height = string.Format("{0:f2}", (float)cy / 360100).Replace(',', '.');
                    return string.Format("width={0}cm,height={1}cm", width, height);
                }
                else
                {
                    return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        #region Helper methods

        private static void ConvertImageToEPS(string orginalImagePath, string newImagePath, IStatusInformation statusInfo)
        {
            string epsImagePath = Path.ChangeExtension(newImagePath, "eps");
            string imageMagickPath = Config.Instance.Infra.ImageMagickPath;
            
            if(string.IsNullOrEmpty(imageMagickPath))
            {
                statusInfo.WriteLine("ERROR: Unable to read configuration setting of ImageMagick's path");
                return;
            }
            try
            {
                Process proc = Process.Start(imageMagickPath, string.Format("\"{0}\" \"{1}\"", orginalImagePath, epsImagePath));
                proc.WaitForExit(60 * 1000); // wait one minute
            }
            catch
            {
                statusInfo.WriteLine("ERROR: Unable to start ImageMagicK");
            }
        }

        #endregion
    }
}

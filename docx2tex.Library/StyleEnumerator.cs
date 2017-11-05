using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;
using System.IO;

namespace docx2tex.Library
{
    /// <summary>
    /// Style enumerator helper class
    /// </summary>
    public static class StyleEnumerator
    {
        /// <summary>
        /// Enumerate all styles in a document
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<string> Enumerate(string path)
        {
            var styles = new List<string>();

            try
            {
                var pkg = Package.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);

                var stylesPart = (ZipPackagePart)pkg.GetPart(new Uri("/word/styles.xml", UriKind.Relative));
                var stylesStream = stylesPart.GetStream();

                var nt = new NameTable();
                var nsmgr = new XmlNamespaceManager(nt);
                nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                var stylesDoc = new XmlDocument(nt);
                stylesDoc.Load(stylesStream);

                foreach (XmlNode xmlNode in stylesDoc.DocumentElement.SelectNodes("/w:styles/w:style", nsmgr))
                {
                    styles.Add(xmlNode.Attributes["w:styleId"].Value.ToLower());
                }
                pkg.Close();
            }
            catch
            {
            }

            styles.Sort();

            return styles;
        }
    }
}

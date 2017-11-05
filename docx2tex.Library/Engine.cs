using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;

namespace docx2tex.Library
{
    partial class Engine
    {
        private XmlDocument _doc;
        private Store _tex;
        private XmlNamespaceManager _nsmgr;
        private Numbering _numberingFn;
        private Styling _stylingFn;
        private Taging _tagingFn;
        private Imaging _imagingFn;
        private TeXing _texingFn;
        private IStatusInformation _statusInfo;

        /// <summary>
        /// Setup helpers and namespaces
        /// </summary>
        /// <param name="documentXmlStream"></param>
        /// <param name="dotnetFn"></param>
        public Engine(Stream documentXmlStream, Numbering numberingFn, Imaging imagingFn, IStatusInformation statusInfo)
        {
            _statusInfo = statusInfo;
            _doc = new XmlDocument();
            _doc.Load(documentXmlStream);
            _texingFn = new TeXing();
            _stylingFn = new Styling();
            var tagingFn = new Taging();
            _tex = new Store(_stylingFn, tagingFn, statusInfo);

            _nsmgr = new XmlNamespaceManager(_doc.NameTable);
            _nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _nsmgr.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            _nsmgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            _nsmgr.AddNamespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            _nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            _nsmgr.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            _nsmgr.AddNamespace("v", "urn:schemas-microsoft-com:vml");
        
            _numberingFn = numberingFn;
            _imagingFn = imagingFn;

            InitMathTables();
            CacheResolvedStyles();
            CacheBookmarks();
        }

        /// <summary>
        /// Entry method
        /// </summary>
        /// <returns></returns>
        public string Process()
        {
            Header();

            int cntNodes = CountNodes(_doc, "/w:document/w:body/*");

            int cnt = 0;
            foreach (XmlNode paragraphNode in GetNodes(_doc, "/w:document/w:body/*"))
            {
                MainProcessor(paragraphNode, false, true);
                cnt++;
                _statusInfo.WriteCR(string.Format("Processed: {0} percent", cnt * 100 / cntNodes));
            }

            Footer();
            _statusInfo.WriteLine("Temporary data structure generated.");

            string tex = _tex.ConvertToString();
            _statusInfo.WriteLine("done.");
            return tex;
        }

        /// <summary>
        /// Processes a paragraph or a table
        /// </summary>
        /// <param name="paragraphNode"></param>
        private void MainProcessor(XmlNode paragraphNode, bool inTable, bool drawNewLine)
        {
            if (paragraphNode.Name == "w:p")
            {
                ProcessParagraph(paragraphNode, paragraphNode.PreviousSibling, paragraphNode.NextSibling, inTable, drawNewLine);
            }
            else if (paragraphNode.Name == "w:tbl")
            {
                ProcessTable(paragraphNode);
            }
        }

        /// <summary>
        /// Header
        /// </summary>
        private void Header()
        {
            // the specified document class or article if not specified
            string docClass = Config.Instance.Infra.DocumentClass ?? "article";

            // properties
            var propsList = new List<string>();
            if (!string.IsNullOrEmpty(Config.Instance.Infra.FontSize))
            {
                propsList.Add(Config.Instance.Infra.FontSize);
            }
            if (!string.IsNullOrEmpty(Config.Instance.Infra.PaperSize))
            {
                propsList.Add(Config.Instance.Infra.PaperSize);
            }
            if (Config.Instance.Infra.Landscape == true)
            {
                propsList.Add("landscape");
            }
            propsList.RemoveAll(p => string.IsNullOrEmpty(p));

            var props = string.Join(",", propsList.ToArray());

            _tex.AddTextNL(@"\documentclass[" + props + "]{" + docClass + "}");

            var enc = docx2tex.Library.Data.InputEnc.Instance.CurrentEncoding;
            if (enc != null)
            {
                _tex.AddTextNL(@"\usepackage[" + enc.InputEncoding + "]{inputenc}");
            }
            _tex.AddTextNL(@"\usepackage{graphicx}");
            _tex.AddTextNL(@"\usepackage{ulem}");
            _tex.AddTextNL(@"\usepackage{amsmath}");
            
            _tex.AddTextNL(@"\begin{document}");
        }

        /// <summary>
        /// Footer
        /// </summary>
        private void Footer()
        {
            _tex.AddTextNL(@"\end{document}");
        }
    }
}

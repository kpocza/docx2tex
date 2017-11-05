using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Xml;

namespace docx2tex.Library.Data
{
    /// <summary>
    /// Codetable for math and other LaTeX symbols
    /// </summary>
    public class CodeTable
    {
        /// <summary>
        /// This is for math symbols
        /// </summary>
        public Dictionary<string, CodeTableInfo> MathOnlyTable { get; private set; }
        /// <summary>
        /// This is for non-math symbols and single symbols that switch to math
        /// </summary>
        public Dictionary<string, CodeTableInfo> NonMathTable { get; private set; }

        private CodeTable()
        {
            MathOnlyTable = new Dictionary<string, CodeTableInfo>();
            NonMathTable = new Dictionary<string, CodeTableInfo>();

            Assembly ass = Assembly.GetExecutingAssembly();

            // load embedded resource
            using (Stream stream = ass.GetManifestResourceStream("docx2tex.Library.Data.CodeTable.xml"))
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(stream);

                // fill dictionary
                foreach (XmlNode xmlNode in doc.SelectNodes("/Codes/Code"))
                {
                    var word = xmlNode.Attributes["Word"].Value;
                    var tex = xmlNode.Attributes["TeX"].Value;
                    var mathModeStr = xmlNode.Attributes["MathMode"].Value;
                    var mathMode = (MathMode)Enum.Parse(typeof(MathMode), mathModeStr);

                    if ((mathMode & MathMode.Yes) == MathMode.Yes || (mathMode & MathMode.Switch) == MathMode.Switch)
                    {
                        MathOnlyTable.Add(word, new CodeTableInfo(word, tex, mathMode));
                    }

                    if ((mathMode & MathMode.No) == MathMode.No || (mathMode & MathMode.Switch) == MathMode.Switch)
                    {
                        NonMathTable.Add(word, new CodeTableInfo(word, tex, mathMode));
                    }
                }
            }
        }

        /// <summary>
        /// Singleton instance
        /// </summary>
        public static readonly CodeTable Instance = new CodeTable();
    }

    [Flags]
    public enum MathMode
    {
        No = 1,
        Switch = 2,
        Yes = 4
    }

    public class CodeTableInfo
    {
        internal CodeTableInfo(string word, string tex, MathMode mm)
        {
            Word = word;
            TeX = tex;
            MathMode = mm;
        }

        public string Word { get; private set; }
        public string TeX { get; private set; }
        public MathMode MathMode { get; private set; }
    }
}

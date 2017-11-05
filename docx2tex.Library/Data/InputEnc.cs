using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Xml;

namespace docx2tex.Library.Data
{
    /// <summary>
    /// Input encodings
    /// </summary>
    public class InputEnc
    {
        /// <summary>
        /// This is for all encodings
        /// </summary>
        public List<InputEncInfo> InputEncs { get; private set; }
        /// <summary>
        /// This is the current encoding
        /// </summary>
        public InputEncInfo CurrentEncoding 
        {
            get
            {
                string enc = Config.Instance.Infra.InputEnc;

                //if not found or enc is null then return null else the correct object
                return InputEncs.Find(e => e.InputEncoding == enc);
            }
        }

        private InputEnc()
        {
            InputEncs = new  List<InputEncInfo>();

            Assembly ass = Assembly.GetExecutingAssembly();

            // load embedded resource
            using (Stream stream = ass.GetManifestResourceStream("docx2tex.Library.Data.InputEncs.xml"))
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(stream);

                // fill dictionary
                foreach (XmlNode xmlNode in doc.SelectNodes("/InputEncs/InputEnc"))
                {
                    var inputenc = xmlNode.Attributes["inputenc"].Value;
                    var enc = xmlNode.Attributes["encoding"].Value;
                    var desc = xmlNode.Attributes["description"].Value;

                    InputEncs.Add(new InputEncInfo(inputenc, enc, desc));
                }
            }
        }

        /// <summary>
        /// Singleton instance
        /// </summary>
        public static readonly InputEnc Instance = new InputEnc();
    }

    public class InputEncInfo
    {
        public InputEncInfo(string inputenc, string enc, string desc)
        {
            InputEncoding = inputenc;
            DotNetEncoding = enc;
            Description = desc;
        }

        public string InputEncoding { get; private set; }
        public string DotNetEncoding { get; private set; }
        public string Description { get; private set; }
    }
}

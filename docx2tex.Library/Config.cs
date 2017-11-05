using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;
using docx2tex.Library.Data;
using System.Xml.Serialization;

namespace docx2tex.Library
{
    /// <summary>
    /// Store data that has to be available when the singleton instance of Config is constructed
    /// </summary>
    sealed public class StaticConfigHelper
    {
        private static string _docxPath;

        public static string DocxPath
        {
            get { return _docxPath; }
            set
            {
                _docxPath = value;
                Config.Instance = new Config();
            }
        }
    }

    /// <summary>
    /// Global system configuration
    /// </summary>
    sealed public class Config
    {
        internal Config()
        {
            // read system config
            Docx2TexConfig systemConfig = LoadSystemConfig();

            // read user config
            Docx2TexConfig userConfig = LoadUserConfig();

            // read document config
            Docx2TexConfig documentConfig = LoadDocumentConfig();

            // fill system data structures
            LaTeXTags = new LaTeXTags(systemConfig.LaTeXTags, userConfig != null ? userConfig.LaTeXTags : null, documentConfig != null ? documentConfig.LaTeXTags : null);
            Infra = new Infra(systemConfig.Infra, userConfig != null ? userConfig.Infra : null, documentConfig != null ? documentConfig.Infra : null);
            StyleMap = new StyleMap(systemConfig.StyleMap, userConfig != null ? userConfig.StyleMap : null, documentConfig != null ? documentConfig.StyleMap : null);
        }

        public static Docx2TexConfig LoadSystemConfig()
        {
            string mainModuleDirPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            return GetConfig(mainModuleDirPath, "docx2tex.SystemConfig");
        }

        public static Docx2TexConfig LoadUserConfig()
        {
            string addDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string userConfigPath = Path.Combine(addDataPath, "docx2tex");
            return GetConfig(userConfigPath, "docx2tex.UserConfig");
        }

        public static Docx2TexConfig LoadDocumentConfig()
        {
            if (StaticConfigHelper.DocxPath != null)
            {
                string docxDirPath = Path.GetDirectoryName(StaticConfigHelper.DocxPath);
                string configFileName = Path.ChangeExtension(Path.GetFileName(StaticConfigHelper.DocxPath), ".docx2texConfig");
                return GetConfig(docxDirPath, configFileName);
            }
            return null;
        }

        public static void SaveConfig(Docx2TexConfig docx2TexConfig)
        {
            string configFilePath = docx2TexConfig.ConfigurationFilePath;

            Directory.CreateDirectory(Path.GetDirectoryName(configFilePath));

            // config serializer
            XmlSerializer docx2texConfigSerialier = new XmlSerializer(typeof(Docx2TexConfig));

            using (StreamWriter systemConfigWriter = new StreamWriter(configFilePath))
            {
                docx2texConfigSerialier.Serialize(systemConfigWriter, docx2TexConfig);
            }
        }

        private static Docx2TexConfig GetConfig(string moduleDir, string fileName)
        {
            string configPath = Path.Combine(moduleDir, fileName);

            Docx2TexConfig config = null;
            if (File.Exists(configPath))
            {
                // config serializer
                XmlSerializer docx2texConfigSerialier = new XmlSerializer(typeof(Docx2TexConfig));

                using (StreamReader systemConfigReader = new StreamReader(configPath))
                {
                    config = (Docx2TexConfig)docx2texConfigSerialier.Deserialize(systemConfigReader);
                    config.ConfigurationFilePath = configPath;
                }
            }
            if (config == null)
            {
                config = new Docx2TexConfig();
                config.ConfigurationFilePath = configPath;
            }
            return config;
        }

        /// <summary>
        /// Singleton instance (not readonly because it has to be overwritten some times)
        /// </summary>
        internal static Config Instance = new Config();

        /// <summary>
        /// running infrastructure
        /// </summary>
        public Infra Infra { get; set; }
        /// <summary>
        /// Tags
        /// </summary>
        public LaTeXTags LaTeXTags { get; set; }
        /// <summary>
        /// Style mappings
        /// </summary>
        public StyleMap StyleMap { get; set; }
    }
}

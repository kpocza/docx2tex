using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.Reflection;

namespace docx2tex.Library.Data
{

    /// <summary>
    /// Root of config data to be serialized
    /// </summary>
    sealed public class Docx2TexConfig
    {
        public Docx2TexConfig()
        {
            CleanProperties();
        }

        public Infra Infra { get; set; }
        public LaTeXTags LaTeXTags { get; set; }
        public StyleMap StyleMap { get; set; }

        [XmlIgnore]
        public string ConfigurationFilePath {get;set;}

        public void CleanProperties()
        {
            Infra = new Infra();
            LaTeXTags = new LaTeXTags();
            StyleMap = new StyleMap();
        }
    }

    /// <summary>
    /// docx2tex running infrastructure
    /// </summary>
    sealed public class Infra : Docx2TexAutoConfig
    {
        public Infra()
        {
        }

        public Infra(Infra system, Infra user, Infra document)
            : base(system, user, document)
        {
        }

        /// <summary>
        /// Path to the imagemagick program
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string ImageMagickPath { get; set; }

        /// <summary>
        /// Input encoding
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string InputEnc { get; set; }

        /// <summary>
        /// Line length
        /// </summary>
        [XmlElement(IsNullable = true)]
        [Docx2TexAutoConfig]
        public int? LineLength { get; set; }

        /// <summary>
        /// Document class type
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string DocumentClass { get; set; }

        /// <summary>
        /// default font size
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string FontSize { get; set; }

        /// <summary>
        /// paper size
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string PaperSize { get; set; }

        /// <summary>
        /// is landscape
        /// </summary>
        [XmlElement(IsNullable = true)]
        [Docx2TexAutoConfig]
        public bool? Landscape { get; set; }
    }

    /// <summary>
    /// Main latex tags
    /// </summary>
    sealed public class LaTeXTags : Docx2TexAutoConfig
    {
        public LaTeXTags()
        {
            Breaks = new Breaks();
            StylePair = new StylePair();
            TagPair = new TagPair();
        }

        public LaTeXTags(LaTeXTags system, LaTeXTags user, LaTeXTags document)
            : base(system, user, document)
        {
            Breaks = new Breaks(system.Breaks, user != null ? user.Breaks : null, document != null ? document.Breaks : null);
            StylePair = new StylePair(system.StylePair, user != null ? user.StylePair : null, document != null ? document.StylePair : null);
            TagPair = new TagPair(system.TagPair, user != null ? user.TagPair : null, document != null ? document.TagPair : null);
        }

        /// <summary>
        /// Heading 1
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Section { get; set; }

        /// <summary>
        /// Heading 2
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string SubSection { get; set; }

        /// <summary>
        /// Heading 3
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string SubSubSection { get; set; }

        /// <summary>
        /// Breaks
        /// </summary>
        public Breaks Breaks { get; set; }

        /// <summary>
        /// Styles
        /// </summary>
        public StylePair StylePair { get; set; }

        /// <summary>
        /// Tags
        /// </summary>
        public TagPair TagPair { get; set; }

        /// <summary>
        /// Process figures
        /// </summary>
        [Docx2TexAutoConfig]
        public bool? ProcessFigures { get; set; }

        /// <summary>
        /// Center figures
        /// </summary>
        [Docx2TexAutoConfig]
        public bool? CenterFigures { get; set; }

        /// <summary>
        /// Placement of figure
        /// </summary>
        [Docx2TexAutoConfig]
        [XmlAttribute]
        public string FigurePlacement { get; set; }

        /// <summary>
        /// Center tables
        /// </summary>
        [Docx2TexAutoConfig]
        public bool? CenterTables { get; set; }

        /// <summary>
        /// Placement of table
        /// </summary>
        [Docx2TexAutoConfig]
        [XmlAttribute]
        public string TablePlacement { get; set; }

        /// <summary>
        /// Allow continuous lists
        /// </summary>
        [Docx2TexAutoConfig]
        public bool? AllowContinuousLists { get; set; }

        /// <summary>
        /// Put figure cross-references
        /// </summary>
        [Docx2TexAutoConfig]
        public bool? PutFigureReferences { get; set; }

        /// <summary>
        /// Put table cross-references
        /// </summary>
        [Docx2TexAutoConfig]
        public bool? PutTableReferences { get; set; }

        /// <summary>
        /// Put section cross-references
        /// </summary>
        [Docx2TexAutoConfig]
        public bool? PutSectionReferences { get; set; }

        /// <summary>
        /// Put listing cross-references
        /// </summary>
        [Docx2TexAutoConfig]
        public bool? PutListingReferences { get; set; }
    }

    /// <summary>
    /// Break representations
    /// </summary>
    sealed public class Breaks : Docx2TexAutoConfig
    {
        public Breaks()
        {
        }

        public Breaks(Breaks system, Breaks user, Breaks document)
            : base(system, user, document)
        {
        }

        /// <summary>
        /// Page break
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Page { get; set; }

        /// <summary>
        /// Line break
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Line { get; set; }
    }

    /// <summary>
    /// Pair of begin and end styles
    /// </summary>
    sealed public class StylePair : Docx2TexAutoConfig
    {
        public StylePair()
        {
            Begin = new Style();
            End = new Style();
        }

        public StylePair(StylePair system, StylePair user, StylePair document)
            : base(system, user, document)
        {
            Begin = new Style(system.Begin, user != null ? user.Begin : null, document != null ? document.Begin : null);
            End = new Style(system.End, user != null ? user.End : null, document != null ? document.End : null);
        }

        /// <summary>
        /// Begin styles
        /// </summary>
        public Style Begin { get; set; }
        /// <summary>
        /// End styles
        /// </summary>
        public Style End { get; set; }
    }

    /// <summary>
    /// Style info
    /// </summary>
    sealed public class Style : Docx2TexAutoConfig
    {
        public Style()
        {
        }

        public Style(Style system, Style user, Style document)
            : base(system, user, document)
        {

        }

        /// <summary>
        /// Italic
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string TextIt { get; set; }

        /// <summary>
        /// Bold
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string TextBf { get; set; }

        /// <summary>
        /// Underline
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Underline { get; set; }

        /// <summary>
        /// Strike out
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Sout { get; set; }

        /// <summary>
        /// Double strike out
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Xout { get; set; }

        /// <summary>
        /// Small capitals
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string TextSc { get; set; }

        /// <summary>
        /// All capitals
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string TextC { get; set; }

        /// <summary>
        /// Superscript
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string SuperScript { get; set; }

        /// <summary>
        /// Subscript
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string SubScript { get; set; }

        /// <summary>
        /// Flush right paragraph
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string ParaFlushRight { get; set; }

        /// <summary>
        /// Flush left paragraph
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string ParaCenter { get; set; }
    }

    /// <summary>
    /// Style mapping (map style in word doc to latex style)
    /// </summary>
    sealed public class StyleMap : Docx2TexAutoConfig
    {
        public StyleMap()
        {
        }

        public StyleMap(StyleMap system, StyleMap user, StyleMap document)
            : base(system, user, document)
        {

        }

        /// <summary>
        /// Name of the section style
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Section { get; set; }

        /// <summary>
        /// Name of the subsection style
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string SubSection { get; set; }

        /// <summary>
        /// Name of the subsubsection style
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string SubSubSection { get; set; }

        /// <summary>
        /// Name of the verbatim style
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Verbatim { get; set; }
    }


    /// <summary>
    /// Pair of begin and end tags
    /// </summary>
    sealed public class TagPair : Docx2TexAutoConfig
    {
        public TagPair ()
        {
            Begin = new Tag();
            End = new Tag();
        }

        public TagPair(TagPair system, TagPair user, TagPair document)
            : base(system, user, document)
        {
            Begin = new Tag(system.Begin, user != null ? user.Begin : null, document != null ? document.Begin : null);
            End = new Tag(system.End, user != null ? user.End : null, document != null ? document.End : null);
        }

        /// <summary>
        /// Begin tags
        /// </summary>
        public Tag Begin { get; set; }
        /// <summary>
        /// End tags
        /// </summary>
        public Tag End { get; set; }
    }


    /// <summary>
    /// Tag info
    /// </summary>
    sealed public class Tag : Docx2TexAutoConfig
    {
        public Tag()
        {
        }

        public Tag(Tag system, Tag user, Tag document)
            : base(system, user, document)
        {

        }

        /// <summary>
        /// Verbatim
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Verbatim { get; set; }

        /// <summary>
        /// Math
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Math { get; set; }

        /// <summary>
        /// Figure
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Figure { get; set; }

        /// <summary>
        /// Enumerate
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Enumerate { get; set; }

        /// <summary>
        /// Itemize
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Itemize { get; set; }

        /// <summary>
        /// Table
        /// </summary>
        [XmlAttribute]
        [Docx2TexAutoConfig]
        public string Table { get; set; }
    }

    /// <summary>
    /// These are mapped
    /// </summary>
    sealed public class Docx2TexAutoConfigAttribute : Attribute
    {
    }

    /// <summary>
    /// Automatic configuration mapper class
    /// </summary>
    abstract public class Docx2TexAutoConfig
    {
        protected Docx2TexAutoConfig()
        {
        }

        /// <summary>
        /// Constructore that does the job
        /// </summary>
        /// <param name="system"></param>
        /// <param name="user"></param>
        /// <param name="document"></param>
        protected Docx2TexAutoConfig(Docx2TexAutoConfig system, Docx2TexAutoConfig user, Docx2TexAutoConfig document)
        {
            SetAutoConfigProperties(system, user, document);
        }

        /// <summary>
        /// Property mapper
        /// The same properties are read from the system level, the user level, and the document level configuration.
        /// User level properties can override system level properties while document level properties are specific
        /// only for a single document
        /// </summary>
        /// <param name="system"></param>
        /// <param name="user"></param>
        /// <param name="document"></param>
        private void SetAutoConfigProperties(Docx2TexAutoConfig system, Docx2TexAutoConfig user, Docx2TexAutoConfig document)
        {
            // all public instance properties
            var properties = this.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);

            foreach (var prop in properties)
            {
                // properties that has the Docx2TexAutoConfigAttribute attribute
                if (prop.GetCustomAttributes(typeof(Docx2TexAutoConfigAttribute), false).Length > 0)
                {
                    var systemPropVal = prop.GetValue(system, null);
                    var userPropVal = user != null ? prop.GetValue(user, null) : null;
                    var documentPropVal = document != null ? prop.GetValue(document, null) : null;

                    // if document level property is set, then it wins. If not then try user level. 
                    // If no user level property found then user the system level property.
                    var resultVal = documentPropVal != null ? documentPropVal :
                        (userPropVal != null ? userPropVal : systemPropVal);
                    prop.SetValue(this, resultVal, null);
                }
            }
        }
    }
}

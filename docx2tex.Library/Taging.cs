using System;
using System.Collections.Generic;
using System.Text;

namespace docx2tex.Library
{
    class Taging
    {
 
        #region Lifecycle methods

        public Taging()
        {
        }

        #endregion

        public string Enum2TextStart(TagEnum tagEnum)
        {
            switch (tagEnum)
            {
                case TagEnum.Verbatim:
                    return Config.Instance.LaTeXTags.TagPair.Begin.Verbatim;
                case TagEnum.Math:
                    return Config.Instance.LaTeXTags.TagPair.Begin.Math;
                case TagEnum.Figure:
                    return Config.Instance.LaTeXTags.TagPair.Begin.Figure;
                case TagEnum.Enumerate:
                    return Config.Instance.LaTeXTags.TagPair.Begin.Enumerate;
                case TagEnum.Itemize:
                    return Config.Instance.LaTeXTags.TagPair.Begin.Itemize;
                case TagEnum.Table:
                    return Config.Instance.LaTeXTags.TagPair.Begin.Table;
            }
            return string.Empty;
        }


        public string Enum2TextEnd(TagEnum tagEnum)
        {
            switch (tagEnum)
            {
                case TagEnum.Verbatim:
                    return Config.Instance.LaTeXTags.TagPair.End.Verbatim;
                case TagEnum.Math:
                    return Config.Instance.LaTeXTags.TagPair.End.Math;
                case TagEnum.Figure:
                    return Config.Instance.LaTeXTags.TagPair.End.Figure;
                case TagEnum.Enumerate:
                    return Config.Instance.LaTeXTags.TagPair.End.Enumerate;
                case TagEnum.Itemize:
                    return Config.Instance.LaTeXTags.TagPair.End.Itemize;
                case TagEnum.Table:
                    return Config.Instance.LaTeXTags.TagPair.End.Table;
            }
            return string.Empty;
        }
    }

    enum TagEnum
    {
        // tags
        Verbatim,
        Math,
        Figure,
        Enumerate,
        Itemize,
        Table,
    }
}

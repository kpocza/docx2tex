using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace docx2tex.Library
{
    class Styling
    {
        #region Fields
        
        private static Dictionary<string, string> _paraStylePairs;
        private static Dictionary<string, string> _runStylePairs;

        #endregion    
        
        #region Lifecycle methods

        public Styling()
        {
            InitStylePairs();
        }

        #endregion

        #region Public methods

        public string ResolveParaStyle(string styleName)
        {
            if (_paraStylePairs.ContainsKey(styleName.ToLower()))
            {
                return _paraStylePairs[styleName];
            }
            return styleName;
        }

        public string ResolveRunStyle(string styleName)
        {
            if (_runStylePairs.ContainsKey(styleName.ToLower()))
            {
                return _runStylePairs[styleName];
            }
            return styleName;
        }

        #endregion

        #region Initialization

        private void InitStylePairs()
        {
            _paraStylePairs = new Dictionary<string, string>();
            _paraStylePairs.Add("section", Config.Instance.StyleMap.Section);
            _paraStylePairs.Add("subsection", Config.Instance.StyleMap.SubSection);
            _paraStylePairs.Add("subsubsection", Config.Instance.StyleMap.SubSubSection);
            _paraStylePairs.Add("verbatim", Config.Instance.StyleMap.Verbatim);

            _runStylePairs = new Dictionary<string, string>();
            _runStylePairs.Add("section", Config.Instance.StyleMap.Section);
            _runStylePairs.Add("subsection", Config.Instance.StyleMap.SubSection);
            _runStylePairs.Add("subsubsection", Config.Instance.StyleMap.SubSubSection);
            _runStylePairs.Add("verbatim", Config.Instance.StyleMap.Verbatim);
        }

        #endregion

        public string Enum2TextStart(StyleEnum styleEnum)
        {
            switch (styleEnum)
            {
                case StyleEnum.TextIt:
                    return Config.Instance.LaTeXTags.StylePair.Begin.TextIt;
                case StyleEnum.TextBf:
                    return Config.Instance.LaTeXTags.StylePair.Begin.TextBf;
                case StyleEnum.Underline:
                    return Config.Instance.LaTeXTags.StylePair.Begin.Underline;
                case StyleEnum.Sout:
                    return Config.Instance.LaTeXTags.StylePair.Begin.Sout;
                case StyleEnum.Xout:
                    return Config.Instance.LaTeXTags.StylePair.Begin.Xout;
                case StyleEnum.TextSc:
                    return Config.Instance.LaTeXTags.StylePair.Begin.TextSc;
                case StyleEnum.TextC:
                    return Config.Instance.LaTeXTags.StylePair.Begin.TextC;
                case StyleEnum.SuperScript:
                    return Config.Instance.LaTeXTags.StylePair.Begin.SuperScript;
                case StyleEnum.SubScript:
                    return Config.Instance.LaTeXTags.StylePair.Begin.SubScript;
                case StyleEnum.ParaFlushRight:
                    return Config.Instance.LaTeXTags.StylePair.Begin.ParaFlushRight;
                case StyleEnum.ParaCenter:
                    return Config.Instance.LaTeXTags.StylePair.Begin.ParaCenter;
            }
            return string.Empty;
        }


        public string Enum2TextEnd(StyleEnum styleEnum)
        {
            switch (styleEnum)
            {
                case StyleEnum.TextIt:
                    return Config.Instance.LaTeXTags.StylePair.End.TextIt;
                case StyleEnum.TextBf:
                    return Config.Instance.LaTeXTags.StylePair.End.TextBf;
                case StyleEnum.Underline:
                    return Config.Instance.LaTeXTags.StylePair.End.Underline;
                case StyleEnum.Sout:
                    return Config.Instance.LaTeXTags.StylePair.End.Sout;
                case StyleEnum.Xout:
                    return Config.Instance.LaTeXTags.StylePair.End.Xout;
                case StyleEnum.TextSc:
                    return Config.Instance.LaTeXTags.StylePair.End.TextSc;
                case StyleEnum.TextC:
                    return Config.Instance.LaTeXTags.StylePair.End.TextC;
                case StyleEnum.SuperScript:
                    return Config.Instance.LaTeXTags.StylePair.End.SuperScript;
                case StyleEnum.SubScript:
                    return Config.Instance.LaTeXTags.StylePair.End.SubScript;
                case StyleEnum.ParaFlushRight:
                    return Config.Instance.LaTeXTags.StylePair.End.ParaFlushRight;
                case StyleEnum.ParaCenter:
                    return Config.Instance.LaTeXTags.StylePair.End.ParaCenter;
            }
            return string.Empty;
        }
    }

    enum StyleEnum
    {
        // run styles
        TextIt,
        TextBf,
        Underline,
        Sout,
        Xout,
        TextSc,
        TextC,
        SuperScript,
        SubScript,

        // paragraph styles
        ParaFlushRight,
        ParaCenter,
    }
}

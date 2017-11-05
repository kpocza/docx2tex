using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using docx2tex.Library.Data;

namespace docx2tex.Library
{
    class TeXing
    {
        #region Fields

        private Dictionary<string, string> _replacerPairs;
        private Regex _refRegEx;

        #endregion

        #region Lifecycle methods

        public TeXing()
        {
            InitReplacesPairs();
        }

        #endregion

        #region Initialization methods

        private void InitReplacesPairs()
        {
            _refRegEx = new Regex(" REF (?<Ref>.+?) ", RegexOptions.Compiled);
            _replacerPairs = new Dictionary<string, string>();

            foreach (var ent in CodeTable.Instance.NonMathTable)
            {
                if ((ent.Value.MathMode & MathMode.No) == MathMode.No)
                {
                    _replacerPairs.Add(ent.Key, ent.Value.TeX);
                }
                if ((ent.Value.MathMode & MathMode.Switch) == MathMode.Switch)
                {
                    _replacerPairs.Add(ent.Key, string.Format("${0}$", ent.Value.TeX));
                }
            }
        }                      
        #endregion

        #region Referencing

        public string ResolveBookmarkRef(string reference)
        {
            try
            {
                Match match = _refRegEx.Match(reference);
                return match.Groups["Ref"].Value;
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion

        #region TeXize

        public string TeXizeText(string original)
        {
            string replaced = original;
            foreach (KeyValuePair<string, string> replacerPair in _replacerPairs)
            {
                replaced = replaced.Replace(replacerPair.Key, replacerPair.Value);
            }
            return replaced;
        }

        /// <summary>
        /// TODO: put to codetable
        /// </summary>
        /// <param name="original"></param>
        /// <returns></returns>
        public string VerbatimizeText(string original)
        {
            string newText = original.Replace("’", "'").Replace("“", "\"").Replace("”", "\"");
            newText = newText.Replace("…", "...");

            return newText;
        }                    

        #endregion
    }
}

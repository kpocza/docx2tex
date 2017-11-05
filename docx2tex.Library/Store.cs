using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace docx2tex.Library
{
    class Store
    {
        #region Constants
        
        public static readonly int LINELENGTH = 72;

        #endregion

        #region Fields

        private Styling _stylingFn;
        private Taging _tagingFn;
        private List<Run> _runs;
        private IStatusInformation _statusInfo;

        #endregion

        #region Lifecycle methods

        static Store()
        {
            LINELENGTH = Config.Instance.Infra.LineLength.Value;
        }

        public Store(Styling stylingFn, Taging tagingFn, IStatusInformation statusInfo)
        {
            _stylingFn = stylingFn;
            _tagingFn = tagingFn;
            _runs = new List<Run>();
            _statusInfo = statusInfo;
        }

        #endregion

        #region Add Runs

        public void AddText(string text)
        {
            _runs.Add(new TextRun(text));
        }

        public void AddVerbatim(string text)
        {
            _runs.Add(new VerbatimRun(text));
        }

        public void AddNL()
        {
            _runs.Add(new NewLineRun());
        }

        public void AddTextNL(string text)
        {
            AddText(text);
            AddNL();
        }

        public void AddStartStyle(StyleEnum styleEnum)
        {
            _runs.Add(new StyleStartRun(styleEnum, _stylingFn));
        }

        public void AddEndStyle(StyleEnum styleEnum)
        {
            _runs.Add(new StyleEndRun(styleEnum, _stylingFn));
        }

        public void AddStyleKiller()
        {
            _runs.Add(new StyleKillerRun());
        }

        public void AddStartTag(TagEnum tagEnum)
        {
            _runs.Add(new TagStartRun(tagEnum, _tagingFn));
        }

        public void AddEndTag(TagEnum tagEnum)
        {
            _runs.Add(new TagEndRun(tagEnum, _tagingFn));
        }

        #endregion

        #region Convert to "TeXString"

		public string ConvertToString()
        {
            var originalRuns = _runs;
            List<Run> simplifiedRuns;

            _statusInfo.WriteLine("Removing unusable styles...");
            // run style killer
            // delete styles that suround some float
            originalRuns = RunStyleKillers(originalRuns);

            _statusInfo.WriteLine("Removing empty styles...");
            // kill style pairs that do not have content while there are any
            while (KillEmptyStyles(out simplifiedRuns, originalRuns))
            {
                originalRuns = simplifiedRuns;
            }

            _statusInfo.WriteLine("Compacting runs...");
            // simplify the runs as long as they can be simplified
            while (Simplify(out simplifiedRuns, originalRuns))
            {
                originalRuns = simplifiedRuns;
            }

            _statusInfo.WriteLine("Merging text runs...");
            // merge textrun siblings
            MergeTextRuns(simplifiedRuns);

            _statusInfo.WriteLine("Correcting line lengths...");
            // split the line lengths
            return CompileOutputText(simplifiedRuns);
        }

        #endregion

        #region Helper : Kill extra Styles

        private List<Run> RunStyleKillers(List<Run> originalRuns)
        {
            List<Run> simplifiedRuns = new List<Run>(originalRuns);
            int cntActiveStyles = 0;

            for (int i = 0; i < simplifiedRuns.Count; i++)
            {
                Run run = simplifiedRuns[i];

                if (run is StyleStartRun)
                {
                    cntActiveStyles++;
                }
                else if (run is StyleEndRun)
                {
                    cntActiveStyles--;
                }

                if (run is StyleKillerRun && cntActiveStyles > 0)
                {
                    int cnt = cntActiveStyles;

                    int effectiveCnt = 0;
                    int j = i - 1;
                    while (cnt > 0 && j > 0)
                    {
                        if (simplifiedRuns[j] is StyleStartRun)
                        {
                            effectiveCnt++;
                            if (effectiveCnt > 0)
                            {
                                simplifiedRuns[j] = new NullRun();
                                //simplifiedRuns.RemoveAt(j);
                                effectiveCnt--;
                                cnt--;
                                //i--;
                            }
                        }
                        else if (simplifiedRuns[j] is StyleEndRun)
                        {
                            effectiveCnt--;
                        }
                        j--;
                    }

                    cnt = cntActiveStyles;

                    effectiveCnt = 0;
                    j = i + 1;

                    while (cnt > 0 && j < simplifiedRuns.Count)
                    {
                        if (simplifiedRuns[j] is StyleEndRun)
                        {
                            effectiveCnt++;
                            if (effectiveCnt > 0)
                            {
                                simplifiedRuns[j] = new NullRun();
                                //simplifiedRuns.RemoveAt(j);
                                effectiveCnt--;
                                cnt--;
                                //j--;
                            }
                        }
                        else if (simplifiedRuns[j] is StyleStartRun)
                        {
                            effectiveCnt--;
                        }
                        j++;
                    }
                }
            }
            return simplifiedRuns;
        }

        #endregion

        #region Helper : Kill empty Styles

        private bool KillEmptyStyles(out List<Run> simplifiedRuns, List<Run> originalRuns)
        {
            bool didKill = false;
            simplifiedRuns = new List<Run>(originalRuns);

            simplifiedRuns.RemoveAll(r => (r is TextRun) && string.IsNullOrEmpty((r as TextRun).Text));

            for (int i = 0; i < simplifiedRuns.Count - 1; i++)
            {
                Run run1 = simplifiedRuns[i];
                Run run2 = simplifiedRuns[i + 1];

                if (run1 is StyleStartRun && run2 is StyleEndRun && 
                    (run1 as StyleRun).Style == (run2 as StyleRun).Style)
                {
                    simplifiedRuns[i] = new NullRun();
                    simplifiedRuns[i+1] = new NullRun();
                    //simplifiedRuns.RemoveAt(i); //ith
                    //simplifiedRuns.RemoveAt(i); // i+1th
                    i++;
                    didKill = true;
                }
            }
            return didKill;
        }

        #endregion

        #region Helper : Simplify

        private bool Simplify(out List<Run> simplifiedRuns, List<Run> originalRuns)
        {
            simplifiedRuns = new List<Run>();

            bool didSimplify = false;
            StyleEndRun lastStyleEndRun = null;

            var runEnum = originalRuns.GetEnumerator();
            while (runEnum.MoveNext())
            {
                Run run = runEnum.Current;
                // if newline or text
                if (run is NewLineRun || run is TextRun || run is VerbatimRun)
                {
                    // if a style ending run found then flush it
                    if (lastStyleEndRun != null)
                    {
                        simplifiedRuns.Add(lastStyleEndRun);
                        lastStyleEndRun = null;
                    }
                    // add run
                    simplifiedRuns.Add(run);
                }
                // if tag
                if (run is TagRun)
                {
                    // if a style ending run found then flush it
                    if (lastStyleEndRun != null)
                    {
                        simplifiedRuns.Add(lastStyleEndRun);
                        lastStyleEndRun = null;
                    }
                    // add run
                    simplifiedRuns.Add(run);
                }
                else if (run is StyleStartRun) // style start run
                {
                    // if a style ending run found then process it
                    if (lastStyleEndRun != null)
                    {
                        // if the style of the end is not the same and the start
                        if (((StyleStartRun)run).Style != lastStyleEndRun.Style)
                        {
                            // flush style end
                            simplifiedRuns.Add(lastStyleEndRun);
                            // add start run
                            simplifiedRuns.Add(run);
                        }
                        else
                        {
                            didSimplify = true;
                        }
                        lastStyleEndRun = null;
                    }
                    else // no style ending run found
                    {
                        // add run
                        simplifiedRuns.Add(run);
                    }
                }
                else if (run is StyleEndRun)
                {
                    // if an other style end run found
                    if (lastStyleEndRun != null)
                    {
                        // flush it
                        simplifiedRuns.Add(lastStyleEndRun);
                        lastStyleEndRun = null;
                    }
                    // save the style end run
                    lastStyleEndRun = (StyleEndRun)run;
                }
            }
            // if an style end run found
            if (lastStyleEndRun != null)
            {
                // flush it
                simplifiedRuns.Add(lastStyleEndRun);
            }
            return didSimplify;
        }

	    #endregion


        #region Helper : MergeTextRuns

        private void MergeTextRuns(List<Run> simplifiedRuns)
        {
            for (int i = 0; i < simplifiedRuns.Count; i++)
            {
                if (simplifiedRuns[i] is TextRun)
                {
                    string finalText = (simplifiedRuns[i] as TextRun).Text;
                    bool found = false;
                    int j = i + 1;
                    while (j < simplifiedRuns.Count && simplifiedRuns[j] is TextRun)
                    {
                        finalText += (simplifiedRuns[j] as TextRun).Text;
                        simplifiedRuns[j] = new NullRun();
                        found = true;
                        j++;
                    }
                    if (found)
                    {
                        simplifiedRuns[i] = new TextRun(finalText);
                    }
                }
            }
        }

        #endregion

        #region Helper : CompileOutputText

        private string CompileOutputText(List<Run> simplifiedRuns)
        {
            StringBuilder sb = new StringBuilder();
            int lastLineLength = 0;

            foreach (Run r in simplifiedRuns)
            {
                if (r is VerbatimRun)
                {
                    sb.Append(r.TeXText);
                }
                else if (r is TextRun || r is StyleStartRun || r is StyleEndRun || r is TagRun)
                {
                    string parts = r.TeXText;
                    lastLineLength = AddRun(sb, r.TeXText, lastLineLength);
                }
                else if (r is NewLineRun)
                {
                    sb.Append(r.TeXText);
                    lastLineLength = 0;
                }
            }
            return sb.ToString();
        }

        private int AddRun(StringBuilder sb, string parts, int lastLineLength)
        {
            StringBuilder broken = new StringBuilder();

            foreach (string part in parts.Split(' '))
            {
                if (lastLineLength + part.Length > LINELENGTH)
                {
                    broken.Append(Environment.NewLine);
                    lastLineLength = 0;
                }
                if (!string.IsNullOrEmpty(part))
                {
                    broken.Append(string.Format("{0} ", part));

                    lastLineLength += part.Length + 1;
                }
            }

            string res = broken.ToString();

            if (!parts.EndsWith(" "))
            {
                res = res.TrimEnd();
            }

            if (parts.StartsWith(" ") && !res.StartsWith(" "))
            {
                res = " " + res;
            }
            sb.Append(res);

            return lastLineLength;
        }

        #endregion
    }

    #region Runs

	abstract class Run
    {
        public abstract string TeXText { get; }
    }

    class NullRun : Run
    {
        public override string TeXText
        {
            get { return string.Empty; }
        }
    }

    class TextRun : Run
    {
        public string Text { get; private set ; }

        public TextRun(string text)
        {
            this.Text = text;
        }

        public override string TeXText
        {
            get { return Text; }
        }
    }

    class VerbatimRun : Run
    {
        public string Text { get; private set; }

        public VerbatimRun(string text)
        {
            this.Text = text;
        }

        public override string TeXText
        {
            get { return Text; }
        }
    }

    class NewLineRun : Run
    {
        public override string TeXText
        {
            get { return Environment.NewLine; }
        }
    }

    abstract class StyleRun : Run
    {
        public StyleEnum Style { get; private set; }
        protected Styling _stylingFn;

        public StyleRun(StyleEnum style, Styling stylingFn)
        {
            this.Style = style;
            this._stylingFn = stylingFn;
        }
    }

    class StyleStartRun : StyleRun
    {
        public StyleStartRun(StyleEnum style, Styling stylingFn)
            : base(style, stylingFn)
        {
        }

        public override string TeXText
        {
            get { return _stylingFn.Enum2TextStart(this.Style); }
        }
    }

    class StyleEndRun : StyleRun
    {
        public StyleEndRun(StyleEnum style, Styling stylingFn)
            : base(style, stylingFn)
        {
        }

        public override string TeXText
        {
            get { return _stylingFn.Enum2TextEnd(this.Style); }
        }
    }

    class StyleKillerRun : Run
    {
        public override string TeXText
        {
            get { return string.Empty; }
        }
    }


    abstract class TagRun : Run
    {
        public TagEnum Tag { get; private set; }
        protected Taging _tagingFn;

        public TagRun(TagEnum tag, Taging tagingFn)
        {
            this.Tag = tag;
            this._tagingFn = tagingFn;
        }
    }

    class TagStartRun : TagRun
    {
        public TagStartRun(TagEnum tag, Taging tagingFn)
            : base(tag, tagingFn)
        {
        }

        public override string TeXText
        {
            get { return _tagingFn.Enum2TextStart(this.Tag); }
        }
    }

    class TagEndRun : TagRun
    {
        public TagEndRun(TagEnum tag, Taging tagingFn)
            : base(tag, tagingFn)
        {
        }

        public override string TeXText
        {
            get { return _tagingFn.Enum2TextEnd(this.Tag); }
        }
    }

	#endregion
}

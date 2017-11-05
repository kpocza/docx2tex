using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace docx2tex.Library
{
    partial class Engine
    {
        /// <summary>
        /// Process normal paragraph runs
        /// </summary>
        /// <param name="paraNode"></param>
        /// <param name="inTable"></param>
        /// <param name="inVerbatim"></param>
        private void ParagraphRuns(XmlNode paraNode, bool inTable, bool inVerbatim)
        {
            // apply paragraph level styling for standard paragraphs
            if (!inTable && !inVerbatim)
            {
                TextParaStyleStart(GetNode(paraNode, "./w:pPr"));
            }

            var lastRunInfo = new RunInfo();
            //string lastCFC = "!!!NONE!!!";

            // process all runs
            foreach (XmlNode run in GetNodes(paraNode, "./w:r|./m:oMathPara|./m:oMath|./w:smartTag|./w:hyperlink"))
            {
                // normal runs
                if (run.Name == "w:r")
                {
                    lastRunInfo = ProcessSingleRun(inVerbatim, lastRunInfo, run, inTable);
                }
                else if(run.Name == "w:smartTag" || run.Name == "w:hyperlink")
                {   // when runs are under smartTags or hyperlinks
                    foreach (XmlNode stRun in GetNodes(run, ".//w:r"))
                    {
                        lastRunInfo = ProcessSingleRun(inVerbatim, lastRunInfo, stRun, inTable);
                    }
                }
                // math paragraph
                else if (run.Name == "m:oMathPara")
                {
                    //math content
                    ProcessMath(GetNode(run, "./m:oMath"));
                    _tex.AddTextNL(Config.Instance.LaTeXTags.Breaks.Line);
                }
                // math content
                else if (run.Name == "m:oMath")
                {
                    ProcessMath(run);
                }
            }
            // apply style end for standard paragraphs
            if (!inTable && !inVerbatim)
            {
                TextParaStyleEnd(GetNode(paraNode, "./w:pPr"));
            }
        }

        private RunInfo ProcessSingleRun(bool inVerbatim, RunInfo prevRunInfo, XmlNode run, bool inTable)
        {
            var currentRunInfo = new RunInfo(prevRunInfo);
            // if it is not verbatim then process breaks and styles
            if (!inVerbatim)
            {
                if (CountNodes(run, "./w:br") > 0)
                {
                    // page break
                    if (GetString(run, @"./w:br/@w:type") == "page")
                    {
                        _tex.AddTextNL(Config.Instance.LaTeXTags.Breaks.Page);
                    }
                    else
                    {
                        // line break
                        _tex.AddTextNL(Config.Instance.LaTeXTags.Breaks.Line);
                    }
                }
                // tab
                if (CountNodes(run, "./w:tab") > 0)
                {
                    _tex.AddText(@"\ \ \ \ ");
                }
                // apply run level style
                TextRunStyleStart(GetNode(run, "./w:rPr"));
            }
            else
            {
                // for verbatims put a simple newline
                if (CountNodes(run, "./w:br") > 0)
                {
                    _tex.AddNL();
                }
            }

            string tmpFld = GetString(run, @"./w:fldChar[@w:fldCharType='begin' or @w:fldCharType='separate' or @w:fldCharType='end']/@w:fldCharType");
            
            // store the last crossref field command (begin or separator or end)
            if (!string.IsNullOrEmpty(tmpFld))
            {
                currentRunInfo.CFC = tmpFld;
            }

            //// query the name of the bookmark
            string currentBookmarkName = GetString(run, "./w:instrText");

            // if set
            if (!string.IsNullOrEmpty(currentBookmarkName))
            {
                // it is a crossref
                if (currentBookmarkName.StartsWith(" REF "))
                {
                    currentRunInfo.InstrTextType = InstrTextTypeEnum.Reference;
                }
                else
                {
                    // it is not a cross ref
                    currentRunInfo.InstrTextType = InstrTextTypeEnum.OtherField;
                }
            }

            // if end then no instrtext 
            if(currentRunInfo.CFC == "end")
            {
                // no field found
                currentRunInfo.InstrTextType = InstrTextTypeEnum.None;
            }

            // if we are in the exact bookmark reference in the begin part of the CFC then this is a reference
            if (currentRunInfo.CFC == "begin" && currentRunInfo.InstrTextType == InstrTextTypeEnum.Reference && !string.IsNullOrEmpty(currentBookmarkName))
            {
                ProcessReference(currentBookmarkName);
            }

            // else standard text or object
            else
            if (
                // no CFC at all
                (currentRunInfo.CFC == null)
                // if we are after a separate cfc an not reference field was here then we are interested in the text
                || (currentRunInfo.CFC == "separate" && (currentRunInfo.InstrTextType == InstrTextTypeEnum.None || currentRunInfo.InstrTextType == InstrTextTypeEnum.OtherField))
                // or if we are after the fields
                || (currentRunInfo.CFC == "end")
                )
            {
                string text = GetString(run, "./w:t");
                // if standard text
                if (!string.IsNullOrEmpty(text))
                {
                    TextRun(text, inVerbatim);
                }
                else if (GetNode(run, "./w:drawing") != null)
                {
                    // image
                    ProcessDrawing(GetNode(run, "./w:drawing"), inTable);
                }
                else if (GetNode(run, "./w:object") != null)
                {
                    // image
                    ProcessObject(GetNode(run, "./w:object"), inTable);
                }
                else if (GetNode(run, "./w:pict") != null)
                {
                    // textbox
                    ProcessPict(GetNode(run, "./w:pict"));
                }
            }

            // apply styles if not verbatim
            if (!inVerbatim)
            {
                TextRunStyleEnd(GetNode(run, "./w:rPr"));
            }
            return currentRunInfo;
        }

        /// <summary>
        /// Process text run
        /// </summary>
        /// <param name="t"></param>
        /// <param name="inVerbatim"></param>
        private void TextRun(string t, bool inVerbatim)
        {
            if (t == null)
                return;

            // normal
            if (!inVerbatim)
            {
                _tex.AddText(_texingFn.TeXizeText(t));
            }
            else
            {
                // verbatim
                _tex.AddVerbatim(_texingFn.VerbatimizeText(t));
            }
        }
    }

    enum InstrTextTypeEnum
    {
        None,
        Reference = 1,
        OtherField = 2
    }

    /// <summary>
    /// Information about a run
    /// </summary>
    class RunInfo
    {
        public RunInfo()
        {
            CFC = null;
            InstrTextType = InstrTextTypeEnum.None;
        }

        public RunInfo(RunInfo other)
        {
            CFC = other.CFC;
            InstrTextType = other.InstrTextType;
        }

        public string CFC { get; set; }
        public InstrTextTypeEnum InstrTextType { get; set; }
    }
}

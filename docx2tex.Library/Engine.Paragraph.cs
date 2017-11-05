using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace docx2tex.Library
{
    partial class Engine
    {

        /// <summary>
        /// Process a paragraph
        /// </summary>
        /// <param name="paragraphNode"></param>
        /// <param name="prevNode"></param>
        /// <param name="nextNode"></param>
        private void ProcessParagraph(XmlNode paragraphNode, XmlNode prevNode, XmlNode nextNode, bool inTable, bool drawNewLine)
        {
            // list settings of the current paragraph
            int? currentNumId = GetInt(paragraphNode, "./w:pPr/w:numPr/w:numId/@w:val");
            int? currentLevel = GetInt(paragraphNode, "./w:pPr/w:numPr/w:ilvl/@w:val");
            bool isList = currentNumId.HasValue && currentLevel.HasValue;
            ListTypeEnum currentType = _numberingFn.GetNumberingStyle(currentNumId, currentLevel);

            int? prevNumId = null;
            int? prevLevel = null;
            ListTypeEnum prevType = ListTypeEnum.None;
            int? nextNumId = null;
            int? nextLevel = null;
            ListTypeEnum nextType = ListTypeEnum.None;

            // process list data if we are in a list
            if (isList)
            {
                // list settings of the previous paragraph
                prevNumId = GetInt(prevNode, "./w:pPr/w:numPr/w:numId/@w:val");
                prevLevel = GetInt(prevNode, "./w:pPr/w:numPr/w:ilvl/@w:val");
                prevType = _numberingFn.GetNumberingStyle(prevNumId, prevLevel);

                // list settings of the next paragraph
                nextNumId = GetInt(nextNode, "./w:pPr/w:numPr/w:numId/@w:val");
                nextLevel = GetInt(nextNode, "./w:pPr/w:numPr/w:ilvl/@w:val");
                nextType = _numberingFn.GetNumberingStyle(nextNumId, nextLevel);
            }

            // if it is a list
            if (isList)
            {
                ListControl listBegin = _numberingFn.ProcessBeforeListItem(currentNumId.Value, currentLevel.Value, currentType, prevNumId, prevLevel, nextNumId, nextLevel);

                // some numbered
                if (listBegin.ListType == ListTypeEnum.Numbered)
                {
                    switch (listBegin.NumberedCounterType)
                    {
                        // simple numbered begins
                        case NumberedCounterTypeEnum.None:
                            _tex.AddStartTag(TagEnum.Enumerate);
                            _tex.AddNL();
                            break;
                        // a new numbered begins
                        case NumberedCounterTypeEnum.NewCounter:
                            if (Config.Instance.LaTeXTags.AllowContinuousLists.Value)
                            {
                                _tex.AddTextNL(@"\newcounter{numberedCnt" + listBegin.Numbering + "}");
                            }
                            _tex.AddStartTag(TagEnum.Enumerate);
                            _tex.AddNL();
                            break;
                        // a numbered loaded
                        case NumberedCounterTypeEnum.LoadCounter:
                            _tex.AddStartTag(TagEnum.Enumerate);
                            _tex.AddNL();
                            if (Config.Instance.LaTeXTags.AllowContinuousLists.Value)
                            {
                                _tex.AddTextNL(@"\setcounter{enumi}{\thenumberedCnt" + listBegin.Numbering + "}");
                            }
                            break;
                    }
                }
                else if (listBegin.ListType == ListTypeEnum.Bulleted)
                {
                    // bulleted list begins
                    _tex.AddStartTag(TagEnum.Itemize);
                    _tex.AddNL();
                }

                //list item
                _tex.AddText(@"\item ");
            }

            // this will process the real content of the paragraph
            ProcessParagraphContent(paragraphNode, prevNode, nextNode, drawNewLine&true, inTable|false, isList);

            // in case of list
            if (isList)
            {
                List<ListControl> listEnd = _numberingFn.ProcessAfterListItem(currentNumId.Value, currentLevel.Value, currentType, prevNumId, prevLevel, nextNumId, nextLevel);

                // rollback the ended lists
                foreach (var token in listEnd)
                {
                    // if a numbered list found
                    if (token.ListType == ListTypeEnum.Numbered)
                    {
                        // save counter of next use
                        if (token.NumberedCounterType == NumberedCounterTypeEnum.SaveCounter)
                        {
                            if (Config.Instance.LaTeXTags.AllowContinuousLists.Value)
                            {
                                _tex.AddTextNL("\\setcounter{numberedCnt" + token.Numbering + "}{\\theenumi}");
                            }
                        }
                        _tex.AddEndTag(TagEnum.Enumerate);
                        _tex.AddNL();
                    }
                    else if (token.ListType == ListTypeEnum.Bulleted)
                    {
                        // bulleted ended
                        _tex.AddEndTag(TagEnum.Itemize);
                        _tex.AddNL();
                    }
                }
            }
        }

        private string RESOLVED_SECTION = string.Empty;
        private string RESOLVED_SUBSECTION = string.Empty;
        private string RESOLVED_SUBSUBSECTION = string.Empty;
        private string RESOLVED_VERBATIM = string.Empty;

        private void CacheResolvedStyles()
        {
            RESOLVED_SECTION = _stylingFn.ResolveParaStyle("section");
            RESOLVED_SUBSECTION = _stylingFn.ResolveParaStyle("subsection");
            RESOLVED_SUBSUBSECTION = _stylingFn.ResolveParaStyle("subsubsection");
            RESOLVED_VERBATIM = _stylingFn.ResolveParaStyle("verbatim");
        }

        /// <summary>
        /// Process the paragraph's real content
        /// </summary>
        /// <param name="paraNode"></param>
        /// <param name="prevNode"></param>
        /// <param name="nextNode"></param>
        /// <param name="drawNewLine"></param>
        /// <param name="inTable"></param>
        /// <param name="isList"></param>
        private void ProcessParagraphContent(XmlNode paraNode, XmlNode prevNode, XmlNode nextNode, bool drawNewLine, bool inTable, bool isList)
        {
            string paraStyle = GetLowerString(paraNode, @"./w:pPr/w:pStyle/@w:val");

            // if a heading found then process it
            if (paraStyle == RESOLVED_SECTION ||
                paraStyle == RESOLVED_SUBSECTION ||
                paraStyle == RESOLVED_SUBSUBSECTION)
            {
                // put sections
                if (paraStyle == RESOLVED_SECTION)
                {
                    _tex.AddText(Config.Instance.LaTeXTags.Section + "{");
                }
                else if (paraStyle == RESOLVED_SUBSECTION)
                {
                    _tex.AddText(Config.Instance.LaTeXTags.SubSection + "{");
                }
                else if (paraStyle == RESOLVED_SUBSUBSECTION)
                {
                    _tex.AddText(Config.Instance.LaTeXTags.SubSubSection + "{");
                }

                // put text
                ParagraphRuns(paraNode, false, false);
                _tex.AddText("}");

                if (Config.Instance.LaTeXTags.PutSectionReferences.Value)
                {
                    // put the reference name
                    if (CountNodes(paraNode, "w:bookmarkStart") > 0)
                    {
                        _tex.AddText(@"\label{section:" + GetString(paraNode, "./w:bookmarkStart/@w:name") + "}");
                    }
                }

                _tex.AddNL();
            }
            else if (paraStyle == RESOLVED_VERBATIM)
            {
                // if verbatim node found

                string prevParaStyle = GetLowerString(prevNode, "./w:pPr/w:pStyle/@w:val");
                string nextParaStyle = GetLowerString(nextNode, "./w:pPr/w:pStyle/@w:val");

                // the previous was also verbatim
                bool wasVerbatim = prevParaStyle == RESOLVED_VERBATIM;
                // the next will be also verbatim
                bool willVerbatim = nextParaStyle == RESOLVED_VERBATIM;

                // the first verbatim is begining
                if (!wasVerbatim)
                {
                    _tex.AddStartTag(TagEnum.Verbatim);
                }
                _tex.AddNL();
                // content
                ParagraphRuns(paraNode, false, true);
                // the last verbatim ends
                if (!willVerbatim)
                {
                    _tex.AddNL();
                    _tex.AddEndTag(TagEnum.Verbatim);
                    _tex.AddNL();
                }
            }
            else if (CountNodes(paraNode, @"./w:fldSimple[starts-with(@w:instr, ' SEQ ')]") > 0)
            {
                // a caption text here
                ListingCaptionRun(paraNode);
                if (drawNewLine)
                {
                    _tex.AddNL();
                }
            }
            else
            {
                // draw NORMAL paragraph runs
                ParagraphRuns(paraNode, inTable, false);

                if (drawNewLine)
                {
                    _tex.AddNL();
                    if (!isList)
                    {
                        _tex.AddNL();
                    }
                }
            }
        }
    }
}

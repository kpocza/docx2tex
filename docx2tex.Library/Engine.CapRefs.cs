using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace docx2tex.Library
{
    partial class Engine
    {
        private Dictionary<string, XmlNode> _bookmarkNodeCache;

        /// <summary>
        /// Cache the bookmark node to each bookmark name
        /// </summary>
        private void CacheBookmarks()
        {
            _bookmarkNodeCache = new Dictionary<string, XmlNode>();

            foreach (XmlNode bookmarkNode in GetNodes(_doc, "//w:p/w:bookmarkStart"))
            {
                string bookmarkName = GetString(bookmarkNode, "@w:name");
                if (bookmarkName != null)
                {
                    _bookmarkNodeCache.Add(bookmarkName, bookmarkNode);
                }
            }
        }

        /// <summary>
        /// Add reference fields
        /// </summary>
        /// <param name="currentBookmarkName"></param>
        private void ProcessReference(string currentBookmarkName)
        {
            string bookmarkRefName = _texingFn.ResolveBookmarkRef(currentBookmarkName);
            
            // if no such bookmark found
            if (!_bookmarkNodeCache.ContainsKey(bookmarkRefName))
            {
                return;
            }

            XmlNode bookmarkRefNode = _bookmarkNodeCache[bookmarkRefName];//GetNode(_doc, string.Format("//w:p/w:bookmarkStart[@w:name='{0}']", bookmarkRefName));

            string refStyle = GetLowerString(bookmarkRefNode, "./preceding-sibling::*[1]/w:pStyle/@w:val");

            string seq = GetString(bookmarkRefNode.ParentNode, "./w:fldSimple[starts-with(@w:instr, ' SEQ ')]/@w:instr");

            // do for sections
            if (refStyle == RESOLVED_SECTION ||
                refStyle == RESOLVED_SUBSECTION ||
                refStyle == RESOLVED_SUBSUBSECTION)
            {
                if (Config.Instance.LaTeXTags.PutSectionReferences.Value)
                {
                    _tex.AddText(@"\ref{section:" + bookmarkRefName + "}.");
                }
            }
            else if (!string.IsNullOrEmpty(seq))
            {
                // do for tables, listings, figures
                if (seq.Contains("SEQ Table"))
                {
                    if (Config.Instance.LaTeXTags.PutTableReferences.Value)
                    {
                        _tex.AddText(@"\ref{table:" + bookmarkRefName + "}.");
                    }
                }
                else if (seq.Contains("SEQ Listing"))
                {
                    if (Config.Instance.LaTeXTags.PutListingReferences.Value)
                    {
                        _tex.AddText(@"\ref{listing:" + bookmarkRefName + "}.");
                    }
                }
                else if (seq.Contains("SEQ Fig"))
                {
                    if (Config.Instance.LaTeXTags.PutFigureReferences.Value)
                    {
                        _tex.AddText(@"\ref{figure:" + bookmarkRefName + "}.");
                    }
                }
                else
                {
                    _tex.AddText("!!!Unresolved reference!!!");
                }
            }
            else
            {
                _tex.AddText("!!!Unresolved reference!!!");
            }
        }

        private void ListingCaptionRun(XmlNode paraNode)
        {
            if (!string.IsNullOrEmpty(GetString(paraNode, "./w:fldSimple[starts-with(@w:instr, ' SEQ Listing ')]/@w:instr")))
            {
                _tex.AddTextNL(@"\begin{figure}[h]");
                _tex.AddText(@"\caption{");
                if (Config.Instance.LaTeXTags.PutListingReferences.Value)
                {
                    _tex.AddText(@"\label{listing:" + GetString(paraNode, "./w:bookmarkStart/@w:name") + "}");
                }
                CaptionText(paraNode);
                _tex.AddTextNL("}");
                _tex.AddTextNL(@"\end{figure}");
            }
        }

        private void CaptionText(XmlNode captionP)
        {
            XmlNodeList captionNodes = null;
            // if bookmarkstart found then the bookmark text is after the last bookmarkEnd node
            if (!string.IsNullOrEmpty(GetString(captionP, @"./w:bookmarkStart/@w:name")))
            {
                captionNodes = GetNodes(captionP, @"./w:bookmarkEnd[last()]/following-sibling::*");
            }
            else
            {
                captionNodes = GetNodes(captionP, @"./w:r");
            }
            foreach (XmlNode captNode in captionNodes)
            {
                _tex.AddText(GetString(captNode, @"./w:t"));
            }
        }

        /// <summary>
        /// Resolve image caption
        /// </summary>
        /// <param name="captionP"></param>
        private void ImageCaption(XmlNode captionP)
        {
            if (!string.IsNullOrEmpty(GetString(captionP, "./w:fldSimple[starts-with(@w:instr, ' SEQ Figure ')]/@w:instr")))
            {
                _tex.AddText(@"\caption{");
                if (Config.Instance.LaTeXTags.PutFigureReferences.Value)
                {
                    string refName = GetString(captionP, "./w:bookmarkStart/@w:name");
                    if (!string.IsNullOrEmpty(refName))
                    {
                        _tex.AddText(@"\label{figure:" + refName + "}");
                    }
                }
                CaptionText(captionP);
                _tex.AddTextNL("}");
            }
        }
    }
}

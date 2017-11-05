using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace docx2tex.Library
{
    partial class Engine
    {
        /// <summary>
        /// Process table
        /// </summary>
        /// <param name="tblNode"></param>
        private void ProcessTable(XmlNode tblNode)
        {
            int numberOfColumns = CountNodes(tblNode, @"./w:tblGrid/w:gridCol");

            _tex.AddStartTag(TagEnum.Table);
            _tex.AddTextNL("[" + Config.Instance.LaTeXTags.TablePlacement + "]");
            if (Config.Instance.LaTeXTags.CenterTables.Value)
            {
                _tex.AddTextNL(@"\centering");
            }
            _tex.AddText(@"\begin{tabular}{|");
            for (int i = 0; i < numberOfColumns; i++)
            {
                _tex.AddText("l|");
            }
            _tex.AddTextNL("}");
            _tex.AddTextNL(@"\hline");

            //rows
            foreach (XmlNode tr in GetNodes(tblNode, "./w:tr"))
            {
                //columns
                foreach (XmlNode tc in GetNodes(tr, "./w:tc"))
                {
                    BulkMainProcessor(tc, true, false);

                    if (tc.NextSibling != null)
                    {
                        _tex.AddText(" & ");
                    }
                }
                _tex.AddTextNL(@" \\");
                _tex.AddTextNL(@"\hline");
            }

            _tex.AddTextNL(@"\end{tabular}");

            XmlNode captionP = tblNode.NextSibling;
            if (!string.IsNullOrEmpty(GetString(captionP, "./w:fldSimple[starts-with(@w:instr, ' SEQ Table ')]/@w:instr")))
            {
                _tex.AddText(@"\caption{");
                if (Config.Instance.LaTeXTags.PutTableReferences.HasValue)
                {
                    _tex.AddText(@"\label{table:" + GetString(captionP, @"./w:bookmarkStart/@w:name") + "}");
                }
                CaptionText(captionP);
                _tex.AddTextNL("}");
            }

            _tex.AddEndTag(TagEnum.Table);
            _tex.AddNL();
        }
    }
}

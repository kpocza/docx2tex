using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace docx2tex.Library
{
    partial class Engine
    {
        /// <summary>
        /// Add an image
        /// </summary>
        /// <param name="xmlNode"></param>
        private void ProcessDrawing(XmlNode xmlNode, bool inTable)
        {
            if (!Config.Instance.LaTeXTags.ProcessFigures.Value)
                return;

            XmlNode blipNode = GetNode(xmlNode, @"./wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip");

            if (blipNode != null)
            {
                // kill any surrounding styles when simplifying
                _tex.AddStyleKiller();

                // if we are in a table then no \begin{figure} allowed
                if (!inTable)
                {
                    // put as figure
                    _tex.AddStartTag(TagEnum.Figure);
                    _tex.AddTextNL("[" + Config.Instance.LaTeXTags.FigurePlacement + "]");
                    if (Config.Instance.LaTeXTags.CenterFigures.Value)
                    {
                        _tex.AddTextNL(@"\centering");
                    }
                }

                // apply width and height
                XmlNode extentNode = GetNode(blipNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode, "./wp:extent");
                string widthHeightStr = _imagingFn.GetWidthAndHeightFromStyle(GetInt(extentNode, "@cx"), GetInt(extentNode, "@cy"));
                _tex.AddText(@"\includegraphics[" + widthHeightStr + "]{");
                // convert and resolve new image path
                _tex.AddTextNL(_imagingFn.ResolveImage(GetString(blipNode, "@r:embed"), _statusInfo) + "}");

                XmlNode captionP = blipNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.NextSibling;
                // add caption
                ImageCaption(captionP);

                // if we are in a table then no \end{figure} allowed
                if (!inTable)
                {
                    _tex.AddEndTag(TagEnum.Figure);
                    _tex.AddNL();
                }
            }
        }

        /// <summary>
        /// Add an image
        /// </summary>
        /// <param name="xmlNode"></param>
        private void ProcessObject(XmlNode xmlNode, bool inTable)
        {
            if (!Config.Instance.LaTeXTags.ProcessFigures.Value)
                return;

            XmlNode imageData = GetNode(xmlNode, "./v:shape/v:imagedata");

            if (imageData != null)
            {
                // kill any surrounding styles when simplifying
                _tex.AddStyleKiller();

                // if we are in a table then no \begin{figure} allowed
                if (!inTable)
                {
                    // put as figure
                    _tex.AddStartTag(TagEnum.Figure);
                    _tex.AddTextNL("[" + Config.Instance.LaTeXTags.FigurePlacement + "]");
                    if (Config.Instance.LaTeXTags.CenterFigures.Value)
                    {
                        _tex.AddTextNL(@"\centering");
                    }
                }

                // apply width and height
                string widthHeightStr = _imagingFn.GetWidthAndHeightFromStyle(GetString(imageData.ParentNode, "@style"));
                _tex.AddText(@"\includegraphics[" + widthHeightStr + "]{");
                // convert and resolve new image path
                _tex.AddTextNL(_imagingFn.ResolveImage(GetString(imageData, "@r:id"), _statusInfo) + "}");

                XmlNode captionP = imageData.ParentNode.ParentNode.ParentNode.ParentNode.NextSibling;
                // add caption
                ImageCaption(captionP);

                // if we are in a table then no \end{figure} allowed
                if (!inTable)
                {
                    _tex.AddEndTag(TagEnum.Figure);
                    _tex.AddNL();
                }
            }
        }

        /// <summary>
        /// Process picture-like objects
        /// 1. textboxes
        /// </summary>
        /// <param name="xmlNode"></param>
        private void ProcessPict(XmlNode xmlNode)
        {
            // loop through all textbox contents and process them as normal content
            // if or if not grouped
            foreach (XmlNode txbxs in GetNodes(xmlNode, ".//v:shape/v:textbox/w:txbxContent"))
            {
                BulkMainProcessor(txbxs, false, true);
            }
        }

        /// <summary>
        /// process a set of paragraphs
        /// </summary>
        /// <param name="par"></param>
        private void BulkMainProcessor(XmlNode par, bool inTable, bool drawNewLine)
        {
            foreach (XmlNode paragraphNode in GetNodes(par, "./*"))
            {
                MainProcessor(paragraphNode, inTable, drawNewLine);
            }
        }
    }
}

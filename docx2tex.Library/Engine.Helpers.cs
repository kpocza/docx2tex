using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace docx2tex.Library
{
    partial class Engine
    {
        private int? GetInt(XmlNode srcNode, string query)
        {
            if (srcNode == null)
                return null;

            XmlNode node = srcNode.SelectSingleNode(query, _nsmgr);
            if (node == null)
                return null;
            return int.Parse(node.Value);
        }

        private string GetString(XmlNode srcNode, string query)
        {
            if (srcNode == null)
                return null;

            XmlNode node = srcNode.SelectSingleNode(query, _nsmgr);
            if (node == null)
                return null;
            return node.InnerText;
        }

        private string GetLowerString(XmlNode srcNode, string query)
        {
            string ret = GetString(srcNode, query);
            if (ret == null)
                return null;
            return ret.ToLower();
        }

        private int CountNodes(XmlNode srcNode, string query)
        {
            if (srcNode == null)
                return 0;

            return srcNode.SelectNodes(query, _nsmgr).Count;
        }

        private XmlNodeList GetNodes(XmlNode srcNode, string query)
        {
            if (srcNode == null)
                return null;

            return srcNode.SelectNodes(query, _nsmgr);
        }

        private XmlNode GetNode(XmlNode srcNode, string query)
        {
            if (srcNode == null)
                return null;

            return srcNode.SelectSingleNode(query, _nsmgr);
        }
    }
}

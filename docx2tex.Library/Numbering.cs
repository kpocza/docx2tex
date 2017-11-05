using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Packaging;
using System.IO;
using System.Xml;

namespace docx2tex.Library
{
    class Numbering
    {
        private ZipPackagePart _numberingPart;
        private XmlDocument _numberingDoc;
        private XmlNamespaceManager _xmlnsMgr;
        private List<ListInfo> _listStyle;
        private Dictionary<uint, string> _visitedFirstLevelNumberings;

        public Numbering(ZipPackagePart numberingPart)
        {
            _numberingPart = numberingPart;

            // if exist then load it
            if (numberingPart != null)
            {
                Stream numberingStream = _numberingPart.GetStream();

                NameTable nt = new NameTable();
                _xmlnsMgr = new XmlNamespaceManager(nt);
                _xmlnsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _numberingDoc = new XmlDocument(nt);
                _numberingDoc.Load(numberingStream);
            }

            _listStyle = new List<ListInfo>();
            _visitedFirstLevelNumberings = new Dictionary<uint, string>();
        }

        public ListControl ProcessBeforeListItem(int currentNumId, int currentLevel, ListTypeEnum currentType,
                            int? previousNumId, int? previousLevel,
                            int? nextNumId, int? nextLevel)
        {
            // this is the first list element
            if (!previousNumId.HasValue || !previousLevel.HasValue)
            {
                ListInfo listInfo = new ListInfo(currentNumId, currentLevel, currentType);
                _listStyle.Add(listInfo);

                string suffix = string.Empty;
                if (currentType == ListTypeEnum.Numbered && currentLevel == 0)
                {
                    if(_visitedFirstLevelNumberings.ContainsKey(listInfo.HashCode))
                    {
                        return new ListControl(ListTypeEnum.Numbered, NumberedCounterTypeEnum.LoadCounter, UniqueString(listInfo.NumId));
                    }
                    else
                    {
                        _visitedFirstLevelNumberings.Add(listInfo.HashCode, string.Empty);
                        return new ListControl(ListTypeEnum.Numbered, NumberedCounterTypeEnum.NewCounter, UniqueString(listInfo.NumId));
                    }
                }

                return new ListControl(currentType, NumberedCounterTypeEnum.None, null);
            }
            else //this is not the first list element
            {
                int listTopNumId = _listStyle[_listStyle.Count - 1].NumId;
                int listTopLevel = _listStyle[_listStyle.Count - 1].Level;

                // the same list continues
                if (listTopNumId == currentNumId)
                {
                    // the same list continues with the same level
                    if (listTopLevel == currentLevel)
                    {
                        //nothing to do
                    }
                    else // the same list continues but with different level
                    {
                        // a new level started
                        if (currentLevel > listTopLevel)
                        {
                            _listStyle.Add(new ListInfo(currentNumId, currentLevel, currentType));
                            return new ListControl(currentType, NumberedCounterTypeEnum.None, null);
                        }
                        else // the previous level ended
                        {
                            //nothing to do
                        }
                    }
                }
                else // there was an other list before this
                {
                    int indexOfPrevious = FindListElement(previousNumId.Value, previousLevel.Value);

                    //a new list started because the previous cannot find in the list
                    if (indexOfPrevious != -1)
                    {
                        _listStyle.Add(new ListInfo(currentNumId, currentLevel, currentType));
                        return new ListControl(currentType, NumberedCounterTypeEnum.None, null);
                    }
                    else // a previously broken list found
                    {
                        // nothing to do
                    }
                }
            }

            return new ListControl(ListTypeEnum.None, NumberedCounterTypeEnum.None, null);
        }

        public List<ListControl> ProcessAfterListItem(int currentNumId, int currentLevel, ListTypeEnum currentType,
                            int? previousNumId, int? previousLevel,
                            int? nextNumId, int? nextLevel)
        {
            //if this is the last list element
            if (!nextNumId.HasValue || !nextLevel.HasValue)
            {
                ListControl suffix = new ListControl(ListTypeEnum.None, NumberedCounterTypeEnum.None, null);
                if (_listStyle.Count > 0)
                {
                    ListInfo listInfo = _listStyle[0];
                    if (listInfo.Style == ListTypeEnum.Numbered && listInfo.Level == 0)
                    {
                        if (_visitedFirstLevelNumberings.ContainsKey(listInfo.HashCode))
                        {
                            suffix.NumberedCounterType = NumberedCounterTypeEnum.SaveCounter;
                            suffix.Numbering = UniqueString(listInfo.NumId);
                        }
                    }
                }

                List<ListControl> ends = GetReverseListTilIndex(0);

                if (ends.Count > 0)
                {
                    var first = ends[0];
                    first.NumberedCounterType = suffix.NumberedCounterType;
                    first.Numbering = suffix.Numbering;
                    ends.RemoveAt(0);
                    ends.Insert(0, first);
                }

                _listStyle.Clear();

                return ends;
            }
            else //a list 
            {
                // the same list continues
                if (currentNumId == nextNumId.Value && currentLevel == nextLevel.Value)
                {
                    //nothing to do
                }
                else // other list encountered
                {
                    int indexOfNext = FindListElement(nextNumId.Value, nextLevel.Value);
                    
                    //if the next list element cannot find, then a unknown new list will start
                    if (indexOfNext == -1)
                    {
                        //nothing to do
                    }
                    else // else end of list
                    {
                        //remove listStyles and sign ends
                        List<ListControl> ends = GetReverseListTilIndex(indexOfNext + 1);
                        _listStyle.RemoveRange(indexOfNext + 1, _listStyle.Count - indexOfNext - 1);
                        return ends;
                    }
                }
            }

            return new List<ListControl>();
        }

        private string UniqueString(int num)
        {
            if (num == 0)
                return string.Empty;

            string val = (num - 1).ToString();
            string ret = string.Empty;
            foreach (char c in val)
            {
                ret += Convert.ToChar(Convert.ToInt32(c) + Convert.ToInt32('A') - Convert.ToInt32('0'));
            }
            return ret;
        }

        private List<ListControl> GetReverseListTilIndex(int index)
        {
            List<ListControl> revList = new List<ListControl>();

            if (_listStyle.Count == 0)
                return revList;
            int current = _listStyle.Count - 1;

            while (current >= index)
            {
                revList.Add(new ListControl(_listStyle[current].Style, NumberedCounterTypeEnum.None, null));
                current--;
            }
            return revList;
        }

        private int FindListElement(int numId, int level)
        {
            if (_listStyle.Count == 0)
                return -1;

            int foundIndex = _listStyle.Count - 1;

            while (foundIndex >= 0)
            {
                if (_listStyle[foundIndex].NumId == numId &&
                    _listStyle[foundIndex].Level == level)
                {
                    break;
                }
                foundIndex--;
            }

            return foundIndex;
        }

        public ListTypeEnum GetNumberingStyle(int? numbering, int? level)
        {
            if (numbering.HasValue && level.HasValue)
            {
                XmlNode node = null;
                node = _numberingDoc.DocumentElement.SelectSingleNode(
                        string.Format("/w:numbering/w:num[@w:numId='{0}']/w:abstractNumId",
                            numbering), _xmlnsMgr);

                // there are some cases when w:numId is bad in document.xml and references to an unknown numId in numbering.xml
                if (node == null)
                {
                    return ListTypeEnum.None;
                }

                int abstractNumbering = int.Parse(node.Attributes["w:val"].Value);

                XmlNode absNode = _numberingDoc.DocumentElement.SelectSingleNode(
                        string.Format("/w:numbering/w:abstractNum[@w:abstractNumId='{0}']/w:lvl[@w:ilvl='{1}']/w:numFmt",
                            abstractNumbering,
                            level),
                        _xmlnsMgr
                    );

                if (absNode.Attributes["w:val"].Value != "bullet")
                {
                    return ListTypeEnum.Numbered;
                }
                return ListTypeEnum.Bulleted;
            }
            return ListTypeEnum.None;
        }
    }

    struct ListInfo
    {
        public ListInfo(int numId, int level, ListTypeEnum style)
        {
            this.NumId = numId;
            this.Level = level;
            this.Style = style;
        }

        public int NumId;
        public int Level;
        public ListTypeEnum Style;

        public uint HashCode
        {
            get { return (uint)NumId * 65536 + (uint)Level; }
        }
    }

    enum ListTypeEnum
    {
        None = 0,
        Numbered = 1,
        Bulleted = 2
    }

    enum NumberedCounterTypeEnum
    {
        None = 0,
        LoadCounter = 1,
        NewCounter = 2,
        SaveCounter = 3
    }

    struct ListControl
    {
        public ListControl(ListTypeEnum listType, NumberedCounterTypeEnum numberedCounterType, string numbering)
        {
            this.ListType = listType;
            this.NumberedCounterType = numberedCounterType;
            this.Numbering = numbering;
        }

        public ListTypeEnum ListType;
        public NumberedCounterTypeEnum NumberedCounterType;
        public string Numbering;
    }
}

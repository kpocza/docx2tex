using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using docx2tex.Library.Data;

namespace docx2tex.Library
{
    partial class Engine
    {
        Dictionary<string, string> _mathTable;

        private void InitMathTables()
        {
            _mathTable = new Dictionary<string, string>();
            foreach (var ent in CodeTable.Instance.MathOnlyTable)
            {
                _mathTable.Add(ent.Key, ent.Value.TeX);
            }

            _mathTable.Add("&", ""); // no alignment
        }

        private void ProcessMath(XmlNode mathNode)
        {
            // begin math
            _tex.AddStartTag(TagEnum.Math);

            ProcessMathNodes(GetNodes(mathNode, "./*"));

            // end math
            _tex.AddEndTag(TagEnum.Math);
        }

        private void ProcessMathNodes(XmlNodeList xmlNodeList)
        {
            foreach (XmlNode node in xmlNodeList)
            {
                // standard text
                switch (node.Name)
                {
                    case "m:r":
                        {
                            string str = GetString(node, "./m:t");

                            // process as a function or standard text
                            string data = string.Empty;
                            switch (str)
                            {
                                // functions
                                case "cos": data = @"\cos "; break;
                                case "sin": data = @"\sin "; break;
                                case "tan": data = @"\tan "; break;
                                case "csc": data = @"\csc "; break;
                                case "sec": data = @"\sec "; break;
                                case "cot": data = @"\cot "; break;
                                case "sinh": data = @"\sinh "; break;
                                case "cosh": data = @"\cosh "; break;
                                case "tanh": data = @"\tanh "; break;
                                case "csch": data = @"csch "; break;
                                case "sech": data = @"sech "; break;
                                case "coth": data = @"\coth "; break;
                                case "lim": data = @"\lim "; break;
                                case "min": data = @"\min "; break;
                                case "max": data = @"\max "; break;
                                case "log": data = @"\log "; break;
                                case "ln": data = @"\ln "; break;
                                default:
                                    {
                                        // sometimes the text is missing, should check
                                        if (str != null)
                                        {
                                            //standard text
                                            foreach (var c in str)
                                            {
                                                string cs = c.ToString();
                                                // special characters
                                                if (_mathTable.ContainsKey(cs))
                                                {
                                                    data += _mathTable[cs];
                                                }
                                                else
                                                {
                                                    // normal characters
                                                    data += cs;
                                                }
                                            }
                                        }
                                    }
                                    break;
                            }


                            _tex.AddText(data);
                        }
                        break;
                    // ( xxx )
                    case "m:d":
                        {
                            string begChar = GetString(node, "./m:dPr/m:begChr/@m:val");
                            string endChar = GetString(node, "./m:dPr/m:endChr/@m:val");

                            //hack
                            //if no begin and end char specified and a nobar fraction comes (a binomial)
                            //then no ()-s should be put
                            if (string.IsNullOrEmpty(begChar) && string.IsNullOrEmpty(endChar) &&
                                CountNodes(node, "./m:e/m:f/m:fPr/m:type[@m:val='noBar']") == 1)
                            {
                                ProcessMathNodes(GetNodes(node, "./m:e/*"));
                                break;
                            }

                            bool isEqArr = CountNodes(node, "./m:e/m:eqArr") > 0;

                            string b = string.Empty;
                            switch (begChar)
                            {
                                case "{":
                                    if (!isEqArr)
                                    {
                                        b = @"\lbrace ";
                                    }
                                    else
                                    {
                                        b = @"\bigg\{ ";
                                    }
                                    break;
                                case "[": b = @"\lbrack "; break;
                                case "〈": b = @"\langle "; break;
                                case "⌊": b = @"\lfloor "; break;
                                case "⌈": b = @"\lceil "; break;
                                case "|": b = @"\vert "; break;
                                case "‖": b = @"\Vert "; break;
                                case "]": b = @"\rbrack "; break;
                                case "(": b = @"("; break;
                                    
                                default: b = "("; break;
                            }
                            if (!string.IsNullOrEmpty(b))
                            {
                                _tex.AddText(b);
                            }

                            ProcessMathNodes(GetNodes(node, "./m:e/*"));

                            string e = string.Empty;
                            switch (endChar)
                            {
                                case "}": e = @"\rbrace "; break;
                                case "]": e = @"\rbrack "; break;
                                case "〉": e = @"\rangle "; break;
                                case "⌋": e = @"\rfloor "; break;
                                case "⌉": e = @"\rceil "; break;
                                case "|": e = @"\vert "; break;
                                case "‖": e = @"\Vert "; break;
                                case "[": e = @"\lbrack "; break;
                                case ")": e = @")"; break;
                                default:
                                    if (string.IsNullOrEmpty(begChar) || begChar == "(")
                                    {
                                        _tex.AddText(")");
                                    }
                                    break;
                            }
                            if (!string.IsNullOrEmpty(e))
                            {
                                _tex.AddText(e);
                            }
                        }
                        break;
                    // box
                    case "m:box":
                        {
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                        }
                        break;
                    // frac or binom
                    case "m:f":
                        {
                            if (CountNodes(node, "./m:fPr/m:type[@m:val='noBar']") == 1)
                            {
                                _tex.AddText(@"\binom{");
                            }
                            else
                            {
                                _tex.AddText(@"\frac{");
                            }
                            //TODO: frac styles: a/b

                            //numerator
                            ProcessMathNodes(GetNodes(node, "./m:num/*"));
                            _tex.AddText("}{");
                            //denominator
                            ProcessMathNodes(GetNodes(node, "./m:den/*"));
                            _tex.AddText("}");
                        }
                        break;
                    // roots
                    case "m:rad":
                        {
                            _tex.AddText(@"\sqrt");
                            // if has child nodes
                            if (CountNodes(node, "./m:deg/*") > 0)
                            {
                                _tex.AddText("[");
                                ProcessMathNodes(GetNodes(node, "./m:deg/*"));
                                _tex.AddText("]");
                            }

                            // under deg
                            _tex.AddText("{");
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                            _tex.AddText("}");
                        }
                        break;
                    // big operators
                    case "m:nary":
                        {
                            XmlNode naryPr = GetNode(node, "./m:naryPr");
                            string op = @"\int";
                            if (CountNodes(naryPr, "./m:chr") == 1)
                            {
                                string chr = GetString(naryPr, "./m:chr/@m:val");

                                switch (chr)
                                {
                                    // different kinds of big operators
                                    case "∑":op = @"\sum";break;
                                    case "∏":op = @"\prod";break;
                                    case "∐":op = @"\coprod";break;
                                    case "⋃":op = @"\bigcup";break;
                                    case "⋂":op = @"\bigcap";break;
                                    case "⋁":op = @"\bigvee";break;
                                    case "⋀":op = @"\bigwedge";break;
                                    case "∬":op = @"\iint";break;
                                    case "∭":op = @"\iiint";break;
                                    case "∮":op = @"\oint";break;
                                    case "∯": op = @"\oint\oint"; break; //TODO
                                    case "∰": op = @"\oint\oint\oint"; break; //TODO
                                    default:
                                        op = "";
                                        break;
                                }
                            }
                            _tex.AddText(op);

                            //TODO: alignment of scripts
                            //subscript
                            if (CountNodes(node, "./m:sub/*") > 0)
                            {
                                _tex.AddText("_{");
                                ProcessMathNodes(GetNodes(node, "./m:sub/*"));
                                _tex.AddText("}");
                            }
                            //superscript
                            if (CountNodes(node, "./m:sup/*") > 0)
                            {
                                _tex.AddText("^{");
                                ProcessMathNodes(GetNodes(node, "./m:sup/*"));
                                _tex.AddText("}");
                            }

                            // main data
                            _tex.AddText("{");
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                            _tex.AddText("}");
                        }
                        break;
                    //superscript
                    case "m:sSup":
                        {
                            // base
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));

                            // sup
                            _tex.AddText("^{");
                            ProcessMathNodes(GetNodes(node, "./m:sup/*"));
                            _tex.AddText("}");
                        }
                        break;
                    //subscript
                    case "m:sSub":
                        {
                            // base
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));

                            // sub
                            _tex.AddText("_{");
                            ProcessMathNodes(GetNodes(node, "./m:sub/*"));
                            _tex.AddText("}");
                        }
                        break;
                    //sub+superscript
                    case "m:sSubSup":
                        {
                            // base
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));

                            // sub
                            _tex.AddText("_{");
                            ProcessMathNodes(GetNodes(node, "./m:sub/*"));
                            _tex.AddText("}");
                            // sup
                            _tex.AddText("^{");
                            ProcessMathNodes(GetNodes(node, "./m:sup/*"));
                            _tex.AddText("}");
                        }
                        break;
                    //sub+superscript
                    case "m:sPre":
                        {
                            //TODO: nicer

                            // sub
                            _tex.AddText("{}_{");
                            ProcessMathNodes(GetNodes(node, "./m:sub/*"));
                            _tex.AddText("}");
                            // sup
                            _tex.AddText("{}^{");
                            ProcessMathNodes(GetNodes(node, "./m:sup/*"));
                            _tex.AddText("}");

                            // base
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                        }
                        break;
                    // equation arrays
                    case "m:eqArr":
                        {
                            _tex.AddNL();
                            _tex.AddTextNL(@"\begin{gathered}");
                            foreach (XmlNode eq in GetNodes(node, "./m:e"))
                            {
                                ProcessMathNodes(GetNodes(eq, "./*"));
                                _tex.AddTextNL(@" \\");
                            }
                            _tex.AddTextNL(@"\end{gathered}");
                        }
                        break;
                    // functions
                    case "m:func":
                        {
                            //function
                            ProcessMathNodes(GetNodes(node, "./m:fName/*"));

                            //data
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                        }
                        break;
                    // bar
                    case "m:bar":
                        {
                            if (CountNodes(node, "./m:barPr/m:pos[@m:val='top']") == 1)
                            {
                                _tex.AddText(@"\overline{");
                            }
                            else
                            {
                                _tex.AddText(@"\underline{");
                            }
                            //base
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                            _tex.AddText("}");
                        }
                        break;
                    // accents
                    case "m:acc":
                        {
                            string acc = GetString(node, "./m:accPr/m:chr/@m:val");

                            string lacc = string.Empty;
                            switch (acc)
                            {
                                case "̇": lacc = @"\dot{"; break;
                                case "̈": lacc = @"\ddot{"; break;
                                case "⃛": lacc = @"\dddot{"; break;
                                case "̌": lacc = @"\check{"; break;
                                case "̆": lacc = @"\breve{"; break;
                                case "̅": lacc = @"\bar{"; break;
                                case "̿": lacc = @"{"; break; //TODO
                                case "⃖": lacc = @"\overleftarrow{"; break;
                                case "⃗": lacc = @"\overrightarrow{"; break;
                                case "⃡": lacc = @"\overleftrightarrow{"; break;
                                case "⃐": lacc = @"{"; break; //TODO
                                case "⃑": lacc = @"{"; break; //TODO
                                default: lacc = @"{"; break;
                            }

                            _tex.AddText(lacc);

                            //base
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));

                            _tex.AddText("}");
                        }
                        break;
                    // groups with brace
                    case "m:groupChr":
                        {
                            string chr = GetString(node, "./m:groupChrPr/m:chr/@m:val");

                            // over the math
                            if (CountNodes(node, "./m:groupChrPr/m:pos[@m:val='top']") == 1)
                            {
                                string put = string.Empty;
                                switch (chr)
                                {
                                    case "\u23DE": put = @"\overbrace{"; break; //overbrace unicode
                                    case "→": put = @"\overrightarrow{"; break;
                                    case "←": put = @"\overleftarrow{"; break;
                                    case "↔": put = @"\overleftrightarrow{"; break;
                                    case "⇒": put = @"!!!DBL!!!\overrightarrow{"; break;
                                    case "⇐": put = @"!!!DBL!!!\overleftarrow{"; break;
                                    case "⇔": put = @"!!!DBL!!!\overleftrightarrow{"; break;

                                    default: put = "{"; break;
                                }
                                _tex.AddText(put);
                            }
                            else
                            {
                                // under the math
                                string put = string.Empty;
                                switch (chr)
                                {
                                    case "→": put = @"\underrightarrow{"; break;
                                    case "←": put = @"\underleftarrow{"; break;
                                    case "↔": put = @"\underleftrightarrow{"; break;
                                    case "⇒": put = @"!!!DBL!!!\underrightarrow{"; break;
                                    case "⇐": put = @"!!!DBL!!!\underleftarrow{"; break;
                                    case "⇔": put = @"!!!DBL!!!\underleftrightarrow{"; break;

                                    default: put = @"\underbrace{"; break;
                                }
                                _tex.AddText(put);
                            }
                            //base
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                            _tex.AddText("}");
                        }
                        break;
                    // overbrace with data
                    case "m:limUpp":
                        {
                            //this contains a groupChr
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                            _tex.AddText(@"^{");
                            ProcessMathNodes(GetNodes(node, "./m:lim/*"));
                            _tex.AddText("}");
                        }
                        break;
                    // underbrace with data
                    case "m:limLow":
                        {
                            //this contains a groupChr
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                            _tex.AddText(@"_{");
                            ProcessMathNodes(GetNodes(node, "./m:lim/*"));
                            _tex.AddText("}");
                        }
                        break;
                    // border
                    case "m:borderBox":
                        {
                            //TODO: BOX
                            //base
                            ProcessMathNodes(GetNodes(node, "./m:e/*"));
                        }
                        break;
                    // matrix
                    case "m:m":
                        {
                            _tex.AddTextNL(@"\begin{matrix}");
                            foreach (XmlNode mr in GetNodes(node, "./m:mr"))
                            {
                                foreach (XmlNode e in GetNodes(mr, "./m:e"))
                                {
                                    ProcessMathNodes(e.ChildNodes);
                                    _tex.AddText(" & ");
                                }
                                _tex.AddTextNL(@"\\");
                            }
                            _tex.AddTextNL(@"\end{matrix}");
                        }
                        break;
                    default:
                        break;
                }
            }
        }
    }
}

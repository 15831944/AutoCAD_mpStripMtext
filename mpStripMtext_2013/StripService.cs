namespace mpStripMtext
{
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;

    public class StripService
    {
        private readonly Document _doc;
        private readonly Transaction _tr;

        public StripService(Document doc, Transaction tr)
        {
            _doc = doc;
            _tr = tr;
        }

        public string StripField(DBObject obj)
        {
            var text = GetTextContents(obj);

            if (obj.ExtensionDictionary != ObjectId.Null
                && !obj.ExtensionDictionary.IsErased
                && !obj.ExtensionDictionary.IsEffectivelyErased
                && obj.ExtensionDictionary.IsValid)
            {
                var d = _tr.GetObject(obj.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
                if (d != null && d.Contains("ACAD_FIELD"))
                {
                    d.Remove("ACAD_FIELD");
                }
            }
            return text;
        }

        public void StripMask(DBObject obj)
        {
            var mText = obj as MText;
            if (mText != null && mText.BackgroundFill)
            {
                mText.BackgroundFill = false;
                return;
            }

            var dimension = obj as Dimension;
            if (dimension != null)
            {
                dimension.Dimtfill = 0;
                return;
            }

            var mLeader = obj as MLeader;
            if (mLeader != null)
            {
                // mLeader.EnableFrameText = false;
                var mTextClone = mLeader.MText;
                mTextClone.BackgroundFill = false;
                mLeader.MText = mTextClone;
            }

            var attributeReference = obj as AttributeReference;
            if (attributeReference != null)
            {
                attributeReference.MTextAttribute.BackgroundFill = false;
            }
        }

        public void StripMAttribute(AttributeReference attributeReference, List<string> formats)
        {
            var database = _doc.Database;
            if (attributeReference.ExtensionDictionary == ObjectId.Null
                || CanRemoveDictionary(attributeReference))
            {
                var symbolString = GetTextContents(attributeReference.MTextAttribute);
                var s = StripFormat(symbolString, formats);
                attributeReference.TextString = s;
                ////var mTextClone = attributeReference.MTextAttribute;
                ////mTextClone.Contents = s;
                ////attributeReference.MTextAttribute = mTextClone;
            }

            if (formats.Contains("W"))
            {
                attributeReference.WidthFactor = 1;
            }

            if (formats.Contains("Q"))
            {
                attributeReference.Oblique = 0;
            }

            if (formats.Contains("F"))
            {
                var textStyleTable = (TextStyleTable)_tr.GetObject(database.TextStyleTableId, OpenMode.ForRead);
                if (textStyleTable.Has("Standard"))
                {
                    attributeReference.TextStyleId = textStyleTable["Standard"];
                }
            }
        }

        // https://knowledge.autodesk.com/support/autocad-lt/learn-explore/caas/CloudHelp/cloudhelp/2019/ENU/AutoCAD-LT/files/GUID-7D8BB40F-5C4E-4AE5-BD75-9ED7112E5967-htm.html

        public string StripFormat(string symbolString, List<string> formatsAsList)
        {
            // var formatsAsList = FormatsToList(formats);

            string Replace(string newStr, Regex pat, string oldStr)
            {
                if (pat.IsMatch(oldStr))
                {
                    return pat.Replace(oldStr, newStr);
                }

                return oldStr;
            }

            ////List<Match> Execute(Regex pat, string s)
            ////{
            ////    var result = new List<Match>();
            ////    if (pat.IsMatch(s))
            ////    {
            ////        foreach (Match match in pat.Matches(s))
            ////        {
            ////            result.Add(match);
            ////        }
            ////        return result;
            ////    }
            ////    else
            ////    {
            ////        return result;
            ////    }
            ////}

            // Replace linefeeds using this format "\n" with the AutoCAD
            // standard format "\P". The "\n" format occurs when text is
            // copied to ACAD from some other application.

            var resultString = Replace("\\P", new Regex("\\n"), symbolString);

            //A format
            string Alignment(string oldStr)
            {
                return Replace(string.Empty, new Regex("\\\\A[012];"), oldStr);
            }

            //B format
            string Tab(string str)
            {
                //var matches = Execute(new Regex(@"\\\\P\\t|[0-9]+;\\t"), str);
                //foreach (var match in matches)
                //{
                //    var origstr = match.Groups[1].Value;
                //    var tempstr = Replace(string.Empty, new Regex(@"\\\\A[012];"), str);
                //    str = str.Replace(origstr, tempstr);
                //}
                return Replace(" ", new Regex("\\t"), str);
            }

            //C format
            string Color(string str)
            {
                return Replace(string.Empty, new Regex("\\\\[Cc]\\d*;"), str);
            }

            //F format
            string Font(string str)
            {
                return Replace(string.Empty, new Regex("\\\\[Ff].*?;"), str);
            }

            //H format
            string Height(string str)
            {
                return Replace(string.Empty, new Regex("\\\\H[0-9]*?[.]?[0-9]*?(x|X)+;"), str);
            }

            //O format
            string Overline(string str)
            {
                return Replace(string.Empty, new Regex("\\\\[Oo]"), str);
            }

            //O format
            string Paragraph(string str)
            {
                return Replace(string.Empty, new Regex("\\\\P"), str);
            }

            //Q format
            string Oblique(string str)
            {
                return Replace(string.Empty, new Regex("\\\\Q[-]?[0-9]*?[.]?[0-9]+;"), str);
            }

            //T format
            string Tracking(string str)
            {
                return Replace(string.Empty, new Regex("\\\\T[0-9]?[.]?[0-9]+;"), str);
            }

            //U format
            string Underline(string str)
            {
                return Replace(string.Empty, new Regex("\\\\[Ll]"), str);
            }

            //W format
            string Width(string str)
            {
                return Replace(string.Empty, new Regex("\\\\W[0-9]?[.]?[0-9]+;"), str);
            }

            //~ format (Z)
            string HardSpace(string str)
            {
                return Replace(" ", new Regex("{\\\\[Ff](.*?)\\\\~}|\\\\~"), str);
            }

            string Braces(string str)
            {
                var noBracesText = Replace(string.Empty, new Regex("[{}]"), str);

                return noBracesText;
                ////var matches = Execute(new Regex("{[^\\\\]+}"), str);
                ////foreach (var match in matches)
                ////{
                ////    var origstr = match.Groups[0].Value;
                ////    var tempstr = Replace(string.Empty, new Regex("[{}]"), str);
                ////    str = str.Replace(origstr, tempstr);
                ////}

                ////var len = str.Length;
                ////if (123 == (int)str[1] && 125 == str[len])
                ////{
                ////    var teststr = str.Substring(2);
                ////    ////(setq teststr (substr teststr 1 (1- (strlen teststr))))
                ////    ////(not (vl-string-search "{" teststr))
                ////    ////(not (vl-string-search "}" teststr))
                ////    str = teststr;
                ////}

                //return str;
            }

            //var slahFlag = $"<{DateTime.Now.Ticks}>";
            //var text = Replace(slahFlag, new Regex("\\\\\\\\"), resultString);

            //var lbrace = $"<L{DateTime.Now.Ticks}>";
            //text = Replace(slahFlag, new Regex("\\\\{"), text);

            //var rbrace = $"<{DateTime.Now.Ticks}R>";
            //text = Replace(slahFlag, new Regex("\\\\}"), text);

            string Apply(string t, string token, Func<string, string> f)
            {
                if (formatsAsList.Contains(token))
                {
                    return f(t);
                }
                else
                {
                    return t;
                }
            }

            var text = resultString;

            text = Apply(text, "A", Alignment);
            text = Apply(text, "B", Tab);
            text = Apply(text, "C", Color);
            text = Apply(text, "F", Font);
            text = Apply(text, "H", Height);
            text = Apply(text, "O", Overline);
            text = Apply(text, "Q", Oblique);
            text = Apply(text, "P", Paragraph);
            text = Apply(text, "T", Tracking);
            text = Apply(text, "U", Underline);
            text = Apply(text, "W", Width);
            text = Apply(text, "Z", HardSpace); // replaced from ~

            //text = Replace("\\\\", new Regex(slahFlag), text);
            text = Braces(text);
            //text = Replace("\\{", new Regex(lbrace), text);
            //text = Replace("\\}", new Regex(rbrace), text);

            return text;
        }

        public void StripColumn(MText mText)
        {
            mText.ColumnType = ColumnType.NoColumns;
        }

        public void StripTableFields(Table table)
        {
            using (var mText = new MText())
            {
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        //Get the cell and its contents
                        Cell cell = table.Cells[r, c];
                        if (!string.IsNullOrEmpty(cell.TextString))
                        {
                            mText.Contents = cell.TextString;
                            StripField(mText);
                            cell.TextString = mText.Contents;
                        }
                    }
                }
            }
        }

        public void StripTable(Table table, List<string> formats)
        {
            ////using (var mText = new MText())
            ////{
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        //Get the cell and its contents
                        Cell cell = table.Cells[r, c];
                        if (!string.IsNullOrEmpty(cell.TextString))
                        {
                            var s = StripFormat(cell.TextString, formats);
                            cell.TextString = s;
                        }
                    }
                }
            ////}
        }

        public void StripMLeader(MLeader mLeader, List<string> formats)
        {
            if (mLeader.ExtensionDictionary == ObjectId.Null
                || CanRemoveDictionary(mLeader))
            {
                var symbolString = GetTextContents(mLeader);
                var s = StripFormat(symbolString, formats);
                var mTextClone = mLeader.MText;
                mTextClone.Contents = s;
                mLeader.MText = mTextClone;
            }
        }

        public string GetTextContents(DBObject obj)
        {
            if (obj is MText mText)
                return mText.Contents;

            if (obj is MLeader mLeader)
                return mLeader.MText.Contents;

            if (obj is Dimension dimension)
                return dimension.DimensionText;

            if (obj is AttributeReference attributeReference &&
                attributeReference.IsMTextAttribute)
                return attributeReference.MTextAttribute.Contents;

            return string.Empty;
        }

        private bool CanRemoveDictionary(DBObject obj)
        {
            try
            {
                if (obj.ExtensionDictionary != ObjectId.Null
                    && !obj.ExtensionDictionary.IsErased
                    && !obj.ExtensionDictionary.IsEffectivelyErased
                    && obj.ExtensionDictionary.IsValid)
                {
                    var d = _tr.GetObject(obj.ExtensionDictionary, OpenMode.ForWrite);
                    d?.Erase();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}

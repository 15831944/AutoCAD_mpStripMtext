namespace mpStripMtext
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;
    using JetBrains.Annotations;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    public class MainCommand
    {
        private const string LangItem = "mpStripMtext";
        private Document _doc;

        [CommandMethod("ModPlus", "mpStripMtext", CommandFlags.UsePickSet)]
        public void Start()
        {
            Statistic.SendCommandStarting(new ModPlusConnector());
            try
            {
                _doc = AcApp.DocumentManager.MdiActiveDocument;
                using (var tr = _doc.TransactionManager.StartTransaction())
                {
                    var objectSet = GetObjectSet(tr);
                    if (!objectSet.IsEmpty)
                    {
                        var win = new StripSettings
                        {
                            LbFormatItems =
                        {
                            ItemsSource = GetStripFormatItems()
                        }
                        };
                        if (win.ShowDialog() == true)
                        {
                            var stripFormatItems = win.LbFormatItems.ItemsSource.Cast<StripFormatItem>().ToList();
                            SaveStripFormatItems(stripFormatItems);
                            var formats = stripFormatItems.Where(i => i.Selected).Select(i => i.Code).ToList();

                            var stripService = new StripService(_doc, tr);

                            if (formats.Contains("D"))
                            {
                                objectSet.Mtextobjlst.ForEach(obj => stripService.StripField(obj));
                                objectSet.Mldrobjlst.ForEach(obj => stripService.StripField(obj));
                                objectSet.Dimobjlst.ForEach(obj => stripService.StripField(obj));
                                objectSet.Mattobjlst.ForEach(obj => stripService.StripField(obj));
                                objectSet.Tableobjlst.ForEach(obj => stripService.StripTableFields(obj));
                            }

                            if (formats.Contains("N"))
                            {
                                objectSet.Mtextobjlst.ForEach(obj => stripService.StripColumn(obj));
                            }

                            if (formats.Contains("M"))
                            {
                                objectSet.Mtextobjlst.ForEach(obj => stripService.StripMask(obj));
                                objectSet.Mldrobjlst.ForEach(obj => stripService.StripMask(obj));
                                objectSet.Dimobjlst.ForEach(obj => stripService.StripMask(obj));
                                objectSet.Mattobjlst.ForEach(obj => stripService.StripMask(obj));
                            }

                            formats = formats.Except(new List<string> { "D", "M", "N" }).ToList();
                            if (formats.Any())
                            {
                                foreach (var mText in objectSet.Mtextobjlst)
                                {
                                    var s = stripService.StripFormat(stripService.GetTextContents(mText), formats);
                                    mText.Contents = s;
                                }

                                foreach (var mLeader in objectSet.Mldrobjlst)
                                {
                                    stripService.StripMLeader(mLeader, formats);
                                }

                                foreach (var dimension in objectSet.Dimobjlst)
                                {
                                    var s = stripService.StripFormat(stripService.GetTextContents(dimension), formats);
                                    dimension.DimensionText = s;
                                }

                                foreach (var attributeReference in objectSet.Mattobjlst)
                                {
                                    stripService.StripMAttribute(attributeReference, formats);
                                }

                                foreach (var table in objectSet.Tableobjlst)
                                {
                                    stripService.StripTable(table, formats);
                                }
                            }
                        }

                        tr.Commit();
                    }
                    else
                    {
                        tr.Abort();
                    }
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private ObjectSet GetObjectSet(Transaction tr)
        {
            try
            {
                var implementSelection = GetImplementSelection();
                if (implementSelection != null)
                {
                    var objectSet = SelectionSetToObjectSet(implementSelection, tr);
                    if (!objectSet.IsEmpty)
                        return objectSet;
                }

                var selection = GetSelection();
                return SelectionSetToObjectSet(selection, tr);
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
                return new ObjectSet();
            }
        }

        [CanBeNull]
        private SelectionSet GetImplementSelection()
        {
            var ed = _doc.Editor;
            var selectionResult = ed.SelectImplied();
            if (selectionResult.Status == PromptStatus.OK && selectionResult.Value.Count > 0)
                return selectionResult.Value;

            return null;
        }

        [CanBeNull]
        private SelectionSet GetSelection()
        {
            var ed = _doc.Editor;

            var promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding = "\n" + Language.GetItem(LangItem, "m1") + ":";
            promptSelectionOptions.AllowDuplicates = false;
            promptSelectionOptions.AllowSubSelections = true;
            promptSelectionOptions.RejectObjectsFromNonCurrentSpace = true;
            promptSelectionOptions.RejectObjectsOnLockedLayers = true;

            var selectionResult = ed.GetSelection(promptSelectionOptions);
            if (selectionResult.Status == PromptStatus.OK)
                return selectionResult.Value;

            return null;
        }

#pragma warning disable SA1515 // Single-line comment should be preceded by blank line
        private ObservableCollection<StripFormatItem> GetStripFormatItems()
        {
            var stripFormatItems = new List<StripFormatItem>
            {
                new StripFormatItem(
                    "A",
                    // Выравнивание (Alignment)
                    Language.GetItem(LangItem, "s1"),
                    // Вертикальное выравнивание. Возможные значения: вниз, по центру, вверх. Вертикальное выравнивание появляется при наличии в тексте дробей
                    Language.GetItem(LangItem, "st1")),
                new StripFormatItem(
                    "B", 
                    // Табуляция (Tabs)
                    Language.GetItem(LangItem, "s2"),
                    Language.GetItem(LangItem, "st2")),
                new StripFormatItem(
                    "C",
                    // Цвет (Color)
                    Language.GetItem(LangItem, "s3"),
                    Language.GetItem(LangItem, "st3")),
                new StripFormatItem(
                    "D", 
                    // Поля (Fields)
                    Language.GetItem(LangItem, "s4"),
                    Language.GetItem(LangItem, "st4")),
                new StripFormatItem(
                    "F", 
                    // Шрифт (Font)
                    Language.GetItem(LangItem, "s5"),
                    Language.GetItem(LangItem, "st5")),
                new StripFormatItem(
                    "H", 
                    // Высота (Height)
                    Language.GetItem(LangItem, "s6"),
                    Language.GetItem(LangItem, "st6")),
                new StripFormatItem(
                    "K", 
                    // Перечеркивание (Strikethrough)
                    Language.GetItem(LangItem, "s7"),
                    Language.GetItem(LangItem, "st7")),
                new StripFormatItem(
                    "L", 
                    // Переводы строк (Linefeeds)
                    Language.GetItem(LangItem, "s8"),
                    Language.GetItem(LangItem, "st8")),
                new StripFormatItem(
                    "M", 
                    // Маска (Mask)
                    Language.GetItem(LangItem, "s9"),
                    Language.GetItem(LangItem, "st9")),
                new StripFormatItem(
                    "N", 
                    // Столбцы (Columns)
                    Language.GetItem(LangItem, "s10"),
                    Language.GetItem(LangItem, "st10")),
                new StripFormatItem(
                    "O", 
                    // Надчеркивание (Overline)
                    Language.GetItem(LangItem, "s11"),
                    Language.GetItem(LangItem, "st11")),
                new StripFormatItem(
                    "P",
                    // Абзац (Paragraph)
                    Language.GetItem(LangItem, "s12"),
                    Language.GetItem(LangItem, "st12")),
                new StripFormatItem(
                    "Q",
                    // Наклон (Oblique)
                    Language.GetItem(LangItem, "s13"),
                    Language.GetItem(LangItem, "st13")),
                new StripFormatItem(
                    "S", 
                    // Дроби (Stacking)
                    Language.GetItem(LangItem, "s14"),
                    Language.GetItem(LangItem, "st14")),
                new StripFormatItem(
                    "T",
                    // Межстрочный интервал (Tracking)
                    Language.GetItem(LangItem, "s15"),
                    Language.GetItem(LangItem, "st15")),
                new StripFormatItem(
                    "U",
                    // Подчеркивание (Underline)
                    Language.GetItem(LangItem, "s16"),
                    Language.GetItem(LangItem, "st16")),
                new StripFormatItem(
                    "W", 
                    // Коэффициент растяжения (Width)
                    Language.GetItem(LangItem, "s17"),
                    Language.GetItem(LangItem, "st17")),
                new StripFormatItem(
                    "Z", 
                    // Неразрывный пробел (Non-breaking space)
                    Language.GetItem(LangItem, "s18"),
                    Language.GetItem(LangItem, "st18")),
                new StripFormatItem(
                    "X", 
                    // Выравнивание абзаца (Paragraph alignment)
                    Language.GetItem(LangItem, "s19"),
                    Language.GetItem(LangItem, "st19"))
            };
            foreach (var stripFormatItem in stripFormatItems)
            {
                stripFormatItem.Selected = bool.TryParse(UserConfigFile.GetValue(LangItem, stripFormatItem.Code), out var b) && b;
            }

            stripFormatItems.Sort((i1, i2) => string.Compare(i1.DisplayName, i2.DisplayName, StringComparison.Ordinal));

            return new ObservableCollection<StripFormatItem>(stripFormatItems);
        }
#pragma warning restore SA1515 // Single-line comment should be preceded by blank line

        private void SaveStripFormatItems(IEnumerable<StripFormatItem> stripFormatItems)
        {
            foreach (var stripFormatItem in stripFormatItems)
            {
                UserConfigFile.SetValue(LangItem, stripFormatItem.Code, stripFormatItem.Selected.ToString(), false);
            }

            UserConfigFile.SaveConfigFile();
        }

        private ObjectSet SelectionSetToObjectSet(SelectionSet selectionSet, Transaction tr)
        {
            var objectSet = new ObjectSet();
            if (selectionSet == null)
                return objectSet;

            var objIds = selectionSet.GetObjectIds();

            foreach (var objId in objIds)
            {
                var obj = tr.GetObject(objId, OpenMode.ForWrite);
                if (!(obj is Entity))
                {
                    obj.Dispose();
                    continue;
                }

                var mText = obj as MText;
                if (mText != null)
                {
                    objectSet.Mtextobjlst.Add(mText);
                    continue;
                }

                var mLeader = obj as MLeader;
                if (mLeader != null && mLeader.ContentType == ContentType.MTextContent)
                {
                    objectSet.Mldrobjlst.Add(mLeader);
                    continue;
                }

                var dimension = obj as Dimension;
                if (dimension != null)
                {
                    objectSet.Dimobjlst.Add(dimension);
                    continue;
                }

                var table = obj as Table;
                if (table != null)
                {
                    objectSet.Tableobjlst.Add(table);
                    continue;
                }

                var blockReference = obj as BlockReference;
                if (blockReference != null && blockReference.AttributeCollection.Count > 0)
                {
                    foreach (ObjectId objectId in blockReference.AttributeCollection)
                    {
                        var attributeReference = tr.GetObject(objectId, OpenMode.ForWrite) as AttributeReference;
                        if (attributeReference == null
                            || !attributeReference.IsMTextAttribute
                            || attributeReference.MTextAttribute == null
                        )
                            continue;
                        var layerTableRecord =
                            tr.GetObject(attributeReference.LayerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layerTableRecord != null && layerTableRecord.IsLocked)
                            continue;
                        objectSet.Mattobjlst.Add(attributeReference);
                    }
                }
            }

            return objectSet;
        }
    }
}

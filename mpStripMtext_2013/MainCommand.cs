namespace mpStripMtext
{
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;
    using ModPlusAPI;
    using ModPlusAPI.Windows;

    public class MainCommand
    {
        private readonly string _langItem = "mpStripMtext";

        [CommandMethod("ModPlus", "mpStripMtext", CommandFlags.UsePickSet)]
        public void Start()
        {
            Statistic.SendCommandStarting(new ModPlusConnector());
            try
            {
                var selection = GetSelection();
                if (selection?.Count > 0)
                {
                    StripSettings win = new StripSettings();
                    win.LbFormatItems.ItemsSource = GetStripFormatItems();
                    if (win.ShowDialog() == true)
                    {
                        var stripFormatItems = win.LbFormatItems.ItemsSource.Cast<StripFormatItem>();
                        SaveStripFormatItems(stripFormatItems);
                    }
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private SelectionSet GetSelection()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            var promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding = "";
            promptSelectionOptions.AllowDuplicates = false;
            promptSelectionOptions.AllowSubSelections = true;
            promptSelectionOptions.RejectObjectsFromNonCurrentSpace = true;
            promptSelectionOptions.RejectObjectsOnLockedLayers = true;

            var selectionResult = ed.GetSelection(promptSelectionOptions);

            PromptSelectionResult selectionRes = ed.SelectImplied();

            // If there's no pickfirst set available...

            if (selectionRes.Status == PromptStatus.Error)
            {
                // ... ask the user to select entities
                PromptSelectionOptions selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to list: ";
                selectionRes = ed.GetSelection(selectionOpts);
            }
            else
            {
                // If there was a pickfirst set, clear it
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            SelectionSet selectionSet = null;
            if (selectionResult.Status == PromptStatus.OK)
            {
                selectionSet = selectionResult.Value;
            }

            return selectionSet;
        }

        private ObservableCollection<StripFormatItem> GetStripFormatItems()
        {
            // todo localization
            ObservableCollection<StripFormatItem> stripFormatItems = new ObservableCollection<StripFormatItem>
            {
                new StripFormatItem("A", "Alignment", ""),
                new StripFormatItem("B", "Tabs", ""),
                new StripFormatItem("C", "Color", ""),
                new StripFormatItem("D", "Fields", ""),
                new StripFormatItem("F", "Font", ""),
                new StripFormatItem("H", "Height", ""),
                new StripFormatItem("L", "Linefeed", ""),
                new StripFormatItem("M", "Background Mask", ""),
                new StripFormatItem("N", "Columns", ""),
                new StripFormatItem("O", "Overline", ""),
                new StripFormatItem("P", "Paragraph", ""),
                new StripFormatItem("Q", "Oblique", ""),
                new StripFormatItem("S", "Stacking", ""),
                new StripFormatItem("T", "Tracking", ""),
                new StripFormatItem("U", "Underline", ""),
                new StripFormatItem("W", "Width", ""),
                new StripFormatItem("Z", "Non-breaking space", ""),
            };
            foreach (var stripFormatItem in stripFormatItems)
            {
                stripFormatItem.Selected = bool.TryParse(UserConfigFile.GetValue(_langItem, stripFormatItem.Code), out var b) && b;
            }

            return stripFormatItems;
        }

        private void SaveStripFormatItems(IEnumerable<StripFormatItem> stripFormatItems)
        {
            foreach (var stripFormatItem in stripFormatItems)
            {
                UserConfigFile.SetValue(_langItem, stripFormatItem.Code, stripFormatItem.Selected.ToString(), false);
            }

            UserConfigFile.SaveConfigFile();
        }
    }
}

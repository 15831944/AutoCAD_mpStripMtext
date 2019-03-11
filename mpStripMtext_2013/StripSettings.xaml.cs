namespace mpStripMtext
{
    using System.Windows;

    public partial class StripSettings
    {
        public StripSettings()
        {
            InitializeComponent();
        }

        private void BtCheckAll_OnClick(object sender, RoutedEventArgs e)
        {
            foreach (var item in LbFormatItems.Items)
            {
                if (item is StripFormatItem stripFormatItem)
                    stripFormatItem.Selected = true;
            }
        }

        private void BtUncheckAll_OnClick(object sender, RoutedEventArgs e)
        {
            foreach (var item in LbFormatItems.Items)
            {
                if (item is StripFormatItem stripFormatItem)
                    stripFormatItem.Selected = false;
            }
        }

        private void BtAccept_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}

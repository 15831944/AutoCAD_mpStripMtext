namespace mpStripMtext
{
    using ModPlusAPI.Mvvm;

    public class StripFormatItem : VmBase
    {
        private bool _selected;

        public StripFormatItem(string code, string displayName, string toolTip)
        {
            Code = code;
            DisplayName = displayName;
            ToolTip = toolTip;
        }

        /// <summary>
        /// Условный код элемента форматирования
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Отображаемое имя
        /// </summary>
        public string DisplayName { get; }

        /// <summary>
        /// ToolTip
        /// </summary>
        public string ToolTip { get; }

        /// <summary>Selected</summary>
        public bool Selected
        {
            get => _selected;
            set
            {
                if (Equals(value, _selected)) 
                    return;
                _selected = value;
                OnPropertyChanged();
            }
        }
    }
}

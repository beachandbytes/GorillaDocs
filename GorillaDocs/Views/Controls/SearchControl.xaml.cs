using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace GorillaDocs.Views.Controls
{
    public partial class SearchControl : UserControl
    {
        public SearchControl() { InitializeComponent(); }

        public string Text { get; set; }

        public void Focus() { SearchInput.Focus(); }

        public static readonly DependencyProperty TextProperty = DependencyProperty.Register("Text", typeof(string), typeof(SearchControl));
    }
}

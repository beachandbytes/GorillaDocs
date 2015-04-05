using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace GorillaDocs.Views.Controls
{
    public partial class CheckboxControl : UserControl
    {
        public CheckboxControl() { InitializeComponent(); }

        public string Label { get; set; }
        public bool IsChecked { get; set; }

        public void Focus() { CheckboxInput.Focus(); }

        public static readonly DependencyProperty LabelProperty = DependencyProperty.Register("Label", typeof(string), typeof(CheckboxControl));
        public static readonly DependencyProperty IsCheckedProperty = DependencyProperty.Register("IsChecked", typeof(bool), typeof(CheckboxControl));
    }
}

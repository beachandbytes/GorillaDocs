using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace GorillaDocs.Views.Controls
{
    public partial class DateControl : UserControl
    {
        public DateControl() { InitializeComponent(); }

        public string Label { get; set; }
        public string Value { get; set; }

        public void Focus() { DateInput.Focus(); }

        public static readonly DependencyProperty LabelProperty = DependencyProperty.Register("Label", typeof(string), typeof(DateControl));
        public static readonly DependencyProperty ValueProperty = DependencyProperty.Register("Value", typeof(DateTime), typeof(DateControl));
    }
}

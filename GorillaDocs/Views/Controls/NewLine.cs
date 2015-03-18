using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace GorillaDocs.Views.Controls
{
    public class NewLine : FrameworkElement
    {
        public NewLine()
        {
            Height = 0;
            var binding = new Binding
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(WrapPanel), 1),
                Path = new PropertyPath("ActualHeight")
            };
            BindingOperations.SetBinding(this, HeightProperty, binding);
        }
    }
}

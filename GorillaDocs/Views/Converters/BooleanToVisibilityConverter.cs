using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace GorillaDocs.Views.Converters
{
    public sealed class BooleanToVisibilityConverter : BooleanConverter<Visibility>
    {
        public BooleanToVisibilityConverter() :
            base(Visibility.Visible, Visibility.Collapsed) { }
    }
}

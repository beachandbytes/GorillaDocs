using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace GorillaDocs.Views.Controls
{
    public partial class TextControl : UserControl
    {
        public TextControl() { InitializeComponent(); }

        public string Label { get; set; }
        public string Text { get; set; }
        public string TextBoxHeight { get; set; }
        public bool AcceptsReturn { get; set; }
        public TextWrapping TextWrapping { get; set; }
        public bool SpellCheckEnabled { get; set; }
        public string PlaceholderText { get; set; }
        public CharacterCasing CharacterCasing { get; set; }

        public void Focus() { TextInput.Focus(); }

        public static readonly DependencyProperty LabelProperty = DependencyProperty.Register("Label", typeof(string), typeof(TextControl));
        public static readonly DependencyProperty TextProperty = DependencyProperty.Register("Text", typeof(string), typeof(TextControl));
        public static readonly DependencyProperty TextBoxHeightProperty = DependencyProperty.Register("TextBoxHeight", typeof(string), typeof(TextControl));
        public static readonly DependencyProperty AcceptsReturnProperty = DependencyProperty.Register("AcceptsReturn", typeof(bool), typeof(TextControl));
        public static readonly DependencyProperty TextWrappingProperty = DependencyProperty.Register("TextWrapping", typeof(TextWrapping), typeof(TextControl));
        public static readonly DependencyProperty SpellCheckEnabledProperty = DependencyProperty.Register("SpellCheckEnabled", typeof(bool), typeof(TextControl));
        public static readonly DependencyProperty PlaceholderTextProperty = DependencyProperty.Register("PlaceholderText", typeof(string), typeof(TextControl));
        public static readonly DependencyProperty CharacterCasingProperty = DependencyProperty.Register("CharacterCasing", typeof(CharacterCasing), typeof(TextControl));
    }
}

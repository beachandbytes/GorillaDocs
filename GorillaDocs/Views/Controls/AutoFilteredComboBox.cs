using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace GorillaDocs.Views.Controls
{
    /// <summary>
    /// http://weblogs.asp.net/okloeten/archive/2007/11/12/5088649.aspx
    ///                 <TextBox Name="EditableTextBox"/>
    ///                <h:AutoFilteredComboBox x:Name="Name" Text="{Binding Path=Contact.FullName}" ItemsSource="{Binding Favourites}" />
    /// </summary>
    public class AutoFilteredComboBox : ComboBox
    {
        private int silenceEvents = 0;

        /// <summary>
        /// Creates a new instance of <see cref="AutoFilteredComboBox" />.
        /// </summary>
        public AutoFilteredComboBox()
        {
            DependencyPropertyDescriptor textProperty = DependencyPropertyDescriptor.FromProperty(ComboBox.TextProperty, typeof(AutoFilteredComboBox));
            textProperty.AddValueChanged(this, this.OnTextChanged);

            this.RegisterIsCaseSensitiveChangeNotification();
        }

        #region IsCaseSensitive Dependency Property
        /// <summary>
        /// The <see cref="DependencyProperty"/> object of the <see cref="IsCaseSensitive" /> dependency property.
        /// </summary>
        public static readonly DependencyProperty IsCaseSensitiveProperty = DependencyProperty.Register("IsCaseSensitive", typeof(bool), typeof(AutoFilteredComboBox), new UIPropertyMetadata(false));

        /// <summary>
        /// Gets or sets the way the combo box treats the case sensitivity of typed text.
        /// </summary>
        /// <value>The way the combo box treats the case sensitivity of typed text.</value>
        [Description("The way the combo box treats the case sensitivity of typed text.")]
        [Category("AutoFiltered ComboBox")]
        [DefaultValue(true)]
        public bool IsCaseSensitive
        {
            [System.Diagnostics.DebuggerStepThrough]
            get
            {
                return (bool)this.GetValue(IsCaseSensitiveProperty);
            }
            [System.Diagnostics.DebuggerStepThrough]
            set
            {
                this.SetValue(IsCaseSensitiveProperty, value);
            }
        }

        protected virtual void OnIsCaseSensitiveChanged(object sender, EventArgs e)
        {
            if (this.IsCaseSensitive)
                this.IsTextSearchEnabled = false;

            this.RefreshFilter();
        }

        private void RegisterIsCaseSensitiveChangeNotification()
        {
            DependencyPropertyDescriptor.FromProperty(IsCaseSensitiveProperty, typeof(AutoFilteredComboBox)).AddValueChanged(this, this.OnIsCaseSensitiveChanged);
        }
        #endregion

        #region DropDownOnFocus Dependency Property
        /// <summary>
        /// The <see cref="DependencyProperty"/> object of the <see cref="DropDownOnFocus" /> dependency property.
        /// </summary>
        public static readonly DependencyProperty DropDownOnFocusProperty = DependencyProperty.Register("DropDownOnFocus", typeof(bool), typeof(AutoFilteredComboBox), new UIPropertyMetadata(true));

        /// <summary>
        /// Gets or sets the way the combo box behaves when it receives focus.
        /// </summary>
        /// <value>The way the combo box behaves when it receives focus.</value>
        [Description("The way the combo box behaves when it receives focus.")]
        [Category("AutoFiltered ComboBox")]
        [DefaultValue(true)]
        public bool DropDownOnFocus
        {
            [System.Diagnostics.DebuggerStepThrough]
            get
            {
                return (bool)this.GetValue(DropDownOnFocusProperty);
            }
            [System.Diagnostics.DebuggerStepThrough]
            set
            {
                this.SetValue(DropDownOnFocusProperty, value);
            }
        }
        #endregion

        #region | Handle selection |
        /// <summary>
        /// Called when <see cref="ComboBox.ApplyTemplate()"/> is called.
        /// </summary>
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            this.EditableTextBox.SelectionChanged += this.EditableTextBox_SelectionChanged;
        }

        /// <summary>
        /// Gets the text box in charge of the editable portion of the combo box.
        /// </summary>
        protected TextBox EditableTextBox
        {
            get
            {
                return ((TextBox)base.GetTemplateChild("PART_EditableTextBox"));
            }
        }

        private int start = 0, length = 0;

        private void EditableTextBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            if (this.silenceEvents == 0)
            {
                this.start = ((TextBox)(e.OriginalSource)).SelectionStart;
                this.length = ((TextBox)(e.OriginalSource)).SelectionLength;

                this.RefreshFilter();
            }
        }
        #endregion

        #region | Handle focus |
        /// <summary>
        /// Invoked whenever an unhandled <see cref="UIElement.GotFocus" /> event
        /// reaches this element in its route.
        /// </summary>
        /// <param name="e">The <see cref="RoutedEventArgs" /> that contains the event data.</param>
        protected override void OnGotFocus(RoutedEventArgs e)
        {
            base.OnGotFocus(e);

            if (this.ItemsSource != null && this.DropDownOnFocus)
                this.IsDropDownOpen = true;
        }
        #endregion

        #region | Handle filtering |
        private void RefreshFilter()
        {
            if (this.ItemsSource != null)
            {
                ICollectionView view = CollectionViewSource.GetDefaultView(this.ItemsSource);
                view.Refresh();
                this.IsDropDownOpen = true;
            }
        }

        private bool FilterPredicate(object value)
        {
            // We don't like nulls.
            if (value == null)
                return false;

            // If there is no text, there's no reason to filter.
            if (this.Text.Length == 0)
                return true;

            string prefix = this.Text;

            // If the end of the text is selected, do not mind it.
            if (this.length > 0 && this.start + this.length == this.Text.Length)
            {
                prefix = prefix.Substring(0, this.start);
            }

            return value.ToString().StartsWith(prefix, !this.IsCaseSensitive, CultureInfo.CurrentCulture);
        }
        #endregion

        /// <summary>
        /// Called when the source of an item in a selector changes.
        /// </summary>
        /// <param name="oldValue">Old value of the source.</param>
        /// <param name="newValue">New value of the source.</param>
        protected override void OnItemsSourceChanged(System.Collections.IEnumerable oldValue, System.Collections.IEnumerable newValue)
        {
            if (newValue != null)
            {
                ICollectionView view = CollectionViewSource.GetDefaultView(newValue);
                view.Filter += this.FilterPredicate;
            }

            if (oldValue != null)
            {
                ICollectionView view = CollectionViewSource.GetDefaultView(oldValue);
                view.Filter -= this.FilterPredicate;
            }

            base.OnItemsSourceChanged(oldValue, newValue);
        }

        private void OnTextChanged(object sender, EventArgs e)
        {
            if (!this.IsTextSearchEnabled && this.silenceEvents == 0)
            {
                this.RefreshFilter();

                // Manually simulate the automatic selection that would have been
                // available if the IsTextSearchEnabled dependency property was set.
                if (this.Text.Length > 0)
                {
                    foreach (object item in CollectionViewSource.GetDefaultView(this.ItemsSource))
                    {
                        int text = item.ToString().Length, prefix = this.Text.Length;
                        this.SelectedItem = item;

                        this.silenceEvents++;
                        this.EditableTextBox.Text = item.ToString();
                        this.EditableTextBox.Select(prefix, text - prefix);
                        this.silenceEvents--;
                        break;
                    }
                }
            }
        }
    }
}

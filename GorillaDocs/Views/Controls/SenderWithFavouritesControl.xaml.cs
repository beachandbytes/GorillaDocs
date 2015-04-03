using GorillaDocs.Models;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace GorillaDocs.Views.Controls
{
    public partial class ContactWithFavouritesControl : UserControl
    {
        public ContactWithFavouritesControl() { InitializeComponent(); }

        // View Logic: When editable, the Favourite was not selected when the name was typed in. This fixes that.
        void Name_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Name.IsEditable)
                foreach (Contact contact in Name.Items)
                    if (contact.FullName == Name.Text)
                        Name.SelectedItem = contact;
        }

        // View Logic: The bound 'Contact' is never null. Consequently the ComboBox is still 'half linked' to the Favourite when X pressed. This removes the 'half linked'
        void Clear_Click(object sender, RoutedEventArgs e)
        {
            Name.SelectedIndex = -1;
        }
    }
}

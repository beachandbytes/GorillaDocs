using GorillaDocs.Models;
using GorillaDocs.ViewModels;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace GorillaDocs.Views.Controls
{
    public partial class SenderWithFavouritesControl : UserControl
    {
        public SenderWithFavouritesControl() { InitializeComponent(); }

        public Contact Contact
        {
            get { return _Contact; }
            set
            {
                _Contact = value;
                BindData();
            }
        }
        public Outlook Outlook
        {
            get { return _Outlook; }
            set
            {
                _Outlook = value;
                BindData();

            }
        }
        public Favourites Favourites
        {
            get { return _Favourites; }
            set
            {
                _Favourites = value;
                BindData();
            }
        }

        Favourites _Favourites;
        Outlook _Outlook;
        Contact _Contact;

        public static DependencyProperty ContactProperty = DependencyProperty.Register("Contact", typeof(Contact), typeof(SenderWithFavouritesControl));
        public static DependencyProperty OutlookProperty = DependencyProperty.Register("Outlook", typeof(Outlook), typeof(SenderWithFavouritesControl));
        public static DependencyProperty FavouritesProperty = DependencyProperty.Register("Favourites", typeof(Favourites), typeof(SenderWithFavouritesControl));

        void BindData()
        {
            if (Contact != null && Outlook != null && Favourites != null)
                DataContext = new SenderViewModel(Contact, Outlook, Favourites);
        }

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

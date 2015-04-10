//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Interactivity;
//using System.Windows.Media;

//namespace GorillaDocs.Views
//{
//    public class PersistGroupExpandedStateBehavior : Behavior<Expander>
//    {
//        #region Static Fields

//        public static readonly DependencyProperty GroupNameProperty = DependencyProperty.Register(
//            "GroupName",
//            typeof(object),
//            typeof(PersistGroupExpandedStateBehavior),
//            new PropertyMetadata(default(object)));

//        static readonly DependencyProperty ExpandedStateStoreProperty =
//            DependencyProperty.RegisterAttached(
//                "ExpandedStateStore",
//                typeof(IDictionary<object, bool>),
//                typeof(PersistGroupExpandedStateBehavior),
//                new PropertyMetadata(default(IDictionary<object, bool>)));

//        #endregion

//        #region Public Properties

//        public object GroupName
//        {
//            get { return (object)GetValue(GroupNameProperty); }
//            set { SetValue(GroupNameProperty, value); }
//        }

//        #endregion

//        #region Methods

//        protected override void OnAttached()
//        {
//            base.OnAttached();

//            bool? expanded = GetExpandedState();

//            if (expanded != null)
//                AssociatedObject.IsExpanded = expanded.Value;

//            AssociatedObject.Expanded += OnExpanded;
//            AssociatedObject.Collapsed += OnCollapsed;
//        }

//        protected override void OnDetaching()
//        {
//            AssociatedObject.Expanded -= OnExpanded;
//            AssociatedObject.Collapsed -= OnCollapsed;

//            base.OnDetaching();
//        }

//        ItemsControl FindItemsControl()
//        {
//            DependencyObject current = AssociatedObject;

//            while (current != null && !(current is ItemsControl))
//                current = VisualTreeHelper.GetParent(current);

//            if (current == null)
//                return null;

//            return current as ItemsControl;
//        }

//        bool? GetExpandedState()
//        {
//            var dict = GetExpandedStateStore();

//            if (!dict.ContainsKey(GroupName))
//                return null;

//            return dict[GroupName];
//        }

//        IDictionary<object, bool> GetExpandedStateStore()
//        {
//            ItemsControl itemsControl = FindItemsControl();

//            if (itemsControl == null)
//                throw new Exception("Behavior needs to be attached to an Expander that is contained inside an ItemsControl");

//            var dict = (IDictionary<object, bool>)itemsControl.GetValue(ExpandedStateStoreProperty);

//            if (dict == null)
//            {
//                dict = new Dictionary<object, bool>();
//                var states = ExpanderStates.Get();
//                foreach (dynamic group in itemsControl.Items.Groups)
//                    dict.Add(group.Name, states.Get(group.Name));
//                itemsControl.SetValue(ExpandedStateStoreProperty, dict);
//            }

//            return dict;
//        }

//        void OnCollapsed(object sender, RoutedEventArgs e)
//        {
//            SetExpanded(false);
//        }

//        void OnExpanded(object sender, RoutedEventArgs e)
//        {
//            SetExpanded(true);
//        }

//        void SetExpanded(bool expanded)
//        {
//            var dict = GetExpandedStateStore();
//            dict[GroupName] = expanded;

//            var states = ExpanderStates.Get();
//            states.Set(GroupName.ToString(), expanded);
//        }

//        #endregion
//    }
//}

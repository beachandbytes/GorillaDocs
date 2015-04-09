﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Windows;
using System.Windows.Controls;


namespace GorillaDocs.Views.Behaviours
{
    public class DataGridColumnsBehavior
    {
        public static readonly DependencyProperty BindableColumnsProperty =
            DependencyProperty.RegisterAttached("BindableColumns",
                                                typeof(ObservableCollection<DataGridColumn>),
                                                typeof(DataGridColumnsBehavior),
                                                new UIPropertyMetadata(null, BindableColumnsPropertyChanged));
        static void BindableColumnsPropertyChanged(DependencyObject source, DependencyPropertyChangedEventArgs e)
        {
            DataGrid dataGrid = source as DataGrid;
            ObservableCollection<DataGridColumn> columns = e.NewValue as ObservableCollection<DataGridColumn>;
            dataGrid.Columns.Clear();
            if (columns == null)
                return;
            foreach (DataGridColumn column in columns)
                dataGrid.Columns.Add(column);
            columns.CollectionChanged += (sender, e2) =>
            {
                NotifyCollectionChangedEventArgs ne = e2 as NotifyCollectionChangedEventArgs;
                if (ne.Action == NotifyCollectionChangedAction.Reset)
                {
                    dataGrid.Columns.Clear();
                    foreach (DataGridColumn column in ne.NewItems)
                        dataGrid.Columns.Add(column);
                }
                else if (ne.Action == NotifyCollectionChangedAction.Add)
                    foreach (DataGridColumn column in ne.NewItems)
                        dataGrid.Columns.Add(column);
                else if (ne.Action == NotifyCollectionChangedAction.Move)
                    dataGrid.Columns.Move(ne.OldStartingIndex, ne.NewStartingIndex);
                else if (ne.Action == NotifyCollectionChangedAction.Remove)
                    foreach (DataGridColumn column in ne.OldItems)
                        dataGrid.Columns.Remove(column);
                else if (ne.Action == NotifyCollectionChangedAction.Replace)
                    dataGrid.Columns[ne.NewStartingIndex] = ne.NewItems[0] as DataGridColumn;
            };
        }
        public static void SetBindableColumns(DependencyObject element, ObservableCollection<DataGridColumn> value)
        {
            element.SetValue(BindableColumnsProperty, value);
        }
        public static ObservableCollection<DataGridColumn> GetBindableColumns(DependencyObject element)
        {
            return (ObservableCollection<DataGridColumn>)element.GetValue(BindableColumnsProperty);
        }
    }
}

using GorillaDocs.ViewModels;
using GorillaDocs.Views;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.Console
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            System.Console.WriteLine(GorillaDocs.Properties.Resources.AddTableView_Caption);

            var viewModel = new AddTableViewModel();
            var view = new AddTableView(viewModel);
            view.ShowDialog();
            //if (view.DialogResult == true)
            //{
            //}
        }
    }
}

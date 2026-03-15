using System.Windows;
using ExcelComparer.WPFUI.Models;
using ExcelComparer.WPFUI.ViewModels;

namespace ExcelComparer.WPFUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow(MainViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }

        private void TreeSummary_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SelectedSummaryItem = e.NewValue as SummaryTreeItem;
            }
        }
    }
}
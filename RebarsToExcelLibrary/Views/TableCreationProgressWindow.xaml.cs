using RebarsToExcel.Models;
using System.Windows;
using System.Windows.Input;

namespace RebarsToExcel.Views
{
    public partial class TableCreationProgressWindow : Window
    {
        public TableCreationProgressWindow(TableManager fileManager)
        {
            InitializeComponent();
            DataContext = fileManager;
        }
        #region МЕТОД ПЕРЕТАСКИВАНИЯ ОКНА

        private void DragWithMouse(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }

        #endregion МЕТОД ПЕРЕТАСКИВАНИЯ ОКНА
    }
}
using RebarsToExcel.ViewModels;
using System.Windows;
using System.Windows.Input;

namespace RebarsToExcel.Views
{
    public partial class AnalyzeWindow : Window
    {
        public AnalyzeWindow(DataManageVM dataManageVM)
        {
            InitializeComponent();
            DataContext = dataManageVM;
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
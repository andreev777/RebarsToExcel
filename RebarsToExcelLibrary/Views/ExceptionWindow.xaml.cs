using System.Windows;
using System.Windows.Input;

namespace RebarsToExcel.Views
{
    public partial class ExceptionWindow : Window
    {
        public ExceptionWindow(string exceptionMessage, string exceptionStackTrace)
        {
            InitializeComponent();
            exceptionMessageTextBlock.Text = exceptionMessage;
            exceptionStrackTraceTextBox.Text = exceptionStackTrace;
        }

        #region МЕТОДЫ ПЕРЕТАСКИВАНИЯ И ЗАКРЫТИЯ ОКНА

        private void DragWithMouse(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                if (WindowState == WindowState.Maximized)
                {
                    Top = 0;
                    WindowState = WindowState.Normal;
                }

                DragMove();
            }
        }

        private void CommandBinding_CanExecute_1(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void CommandBinding_Executed_1(object sender, ExecutedRoutedEventArgs e)
        {
            SystemCommands.CloseWindow(this);
        }

        #endregion МЕТОДЫ ПЕРЕТАСКИВАНИЯ И ЗАКРЫТИЯ ОКНА
    }
}
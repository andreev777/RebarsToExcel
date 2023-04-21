using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RebarsToExcel.Commands;
using RebarsToExcel.Models;
using RebarsToExcel.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace RebarsToExcel.Views
{
    public partial class StartWindow : Window
    {
        private UIApplication _uiapp;
        private DataManageVM _dataManageVM;

        public StartWindow(UIApplication uiapp, DataManageVM dataManageVM)
        {
            InitializeComponent();
            _uiapp = uiapp;
            _dataManageVM = dataManageVM;
            DataContext = _dataManageVM;

            if (_dataManageVM.BarsTotalCount == 0)
            {
                barsTabItem.IsEnabled = false;
                dataTabControl.SelectedIndex = 1;
            }

            if (_dataManageVM.RebarAssembliesTotalCount == 0)
            {
                rebarAssembliesTabItem.IsEnabled = false;
                dataTabControl.SelectedIndex = 0;
            }
        }

        private void barsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var bitmapImage = new BitmapImage();

            bitmapImage.BeginInit();
            bitmapImage.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
            bitmapImage.CacheOption = BitmapCacheOption.None;

            if (_dataManageVM.SelectedBarShapeImagePath == null)
            {
                bitmapImage.UriSource = null;
            }
            else
            {
                bitmapImage.UriSource = new Uri(_dataManageVM.SelectedBarShapeImagePath);
                bitmapImage.EndInit();
            }

            shapeImage.Source = bitmapImage;
        }

        private void selectInModelBarButton_Click(object sender, RoutedEventArgs e)
        {
            var allSelectedBarIds = new List<ElementId>();

            if (barsDataGrid.SelectedItems.Count > 0)
            {
                for (int i = 0; i < barsDataGrid.SelectedItems.Count; i++)
                {
                    Bar selectedBar = barsDataGrid.SelectedItems[i] as Bar;
                    var selectedBarIds = selectedBar.Ids;

                    foreach (var selectedBarId in selectedBarIds)
                    {
                        allSelectedBarIds.Add(selectedBarId);
                    }

                    allSelectedBarIds.Distinct().ToList();
                }
            }
            else
            {
                WarningWindow warningWindow = new WarningWindow("ПРЕДУПРЕЖДЕНИЕ", "Выберите элементы");
                warningWindow.ShowDialog();
                return;
            }

            _uiapp.ActiveUIDocument.Selection.SetElementIds(allSelectedBarIds.ToList());

            WindowState = WindowState.Minimized;
        }

        private void selectInModelRebarAssemblyButton_Click(object sender, RoutedEventArgs e)
        {
            var allSelectedRebarAssemblyIds = new List<ElementId>();

            if (rebarAssembliesDataGrid.SelectedItems.Count > 0)
            {
                for (int i = 0; i < rebarAssembliesDataGrid.SelectedItems.Count; i++)
                {
                    RebarAssembly selectedRebarAssembly = rebarAssembliesDataGrid.SelectedItems[i] as RebarAssembly;
                    var selectedRebarAssemblyIds = selectedRebarAssembly.Ids;

                    foreach (var selectedRebarAssemblyId in selectedRebarAssemblyIds)
                    {
                        allSelectedRebarAssemblyIds.Add(selectedRebarAssemblyId);
                    }

                    allSelectedRebarAssemblyIds.Distinct().ToList();
                }
            }
            else
            {
                WarningWindow warningWindow = new WarningWindow("ПРЕДУПРЕЖДЕНИЕ", "Выберите элементы");
                warningWindow.ShowDialog();
                return;
            }

            _uiapp.ActiveUIDocument.Selection.SetElementIds(allSelectedRebarAssemblyIds.ToList());

            WindowState = WindowState.Minimized;
        }

        private void getBarIdsButton_Click(object sender, RoutedEventArgs e)
        {
            if (barsDataGrid.SelectedItems.Count > 0)
            {
                var allSelectedBarIds = new List<string>();

                for (int i = 0; i < barsDataGrid.SelectedItems.Count; i++)
                {
                    Bar selectedBar = barsDataGrid.SelectedItems[i] as Bar;
                    var selectedBarIds = selectedBar.GetIdsAsString();
                    allSelectedBarIds.Add(selectedBarIds);
                }

                var messageText = string.Join(", ", allSelectedBarIds.Distinct().ToList());

                SelectedIdsWindow selectedIdsWindow = new SelectedIdsWindow(messageText);
                selectedIdsWindow.ShowDialog();
            }
            else
            {
                WarningWindow warningWindow = new WarningWindow("ПРЕДУПРЕЖДЕНИЕ", "Выберите элементы");
                warningWindow.ShowDialog();
            }
        }

        private void getRebarAssemblyIdsButton_Click(object sender, RoutedEventArgs e)
        {
            if (rebarAssembliesDataGrid.SelectedItems.Count > 0)
            {
                var allSelectedRebarAssemblyIds = new List<string>();

                for (int i = 0; i < rebarAssembliesDataGrid.SelectedItems.Count; i++)
                {
                    RebarAssembly selectedRebarAssembly = rebarAssembliesDataGrid.SelectedItems[i] as RebarAssembly;
                    var selectedRebarAssemblyIds = selectedRebarAssembly.GetIdsAsString();
                    allSelectedRebarAssemblyIds.Add(selectedRebarAssemblyIds);
                }

                var messageText = string.Join(", ", allSelectedRebarAssemblyIds.Distinct().ToList());

                SelectedIdsWindow selectedIdsWindow = new SelectedIdsWindow(messageText);
                selectedIdsWindow.ShowDialog();
            }
            else
            {
                WarningWindow warningWindow = new WarningWindow("ПРЕДУПРЕЖДЕНИЕ", "Выберите элементы");
                warningWindow.ShowDialog();
            }
        }

        private void dataGrid_UnselectClick(object o, MouseButtonEventArgs e)
        {
            if (e.OriginalSource != barsDataGrid)
            {
                barsDataGrid.UnselectAll();
            }

            if (e.OriginalSource != rebarAssembliesDataGrid)
            {
                rebarAssembliesDataGrid.UnselectAll();
            }
        }

        private void helpButton_Click(object sender, RoutedEventArgs e)
        {
            HelpWindow helpWindow = new HelpWindow();
            helpWindow.ShowDialog();
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
            RebarsToExcelApp.IsOpened = false;
            SystemCommands.CloseWindow(this);
        }
        #endregion МЕТОДЫ ПЕРЕТАСКИВАНИЯ И ЗАКРЫТИЯ ОКНА
    }
}
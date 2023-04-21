using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RebarsToExcel.Commands;
using RebarsToExcel.Views;
using RebarsToExcel.ViewModels;
using System;
using System.Reflection;
using System.ComponentModel;

namespace RebarsToExcel
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class StartCommand : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Assembly.LoadFrom(@"P:\03_БИБЛИОТЕКА\Revit_5_ПСМ\Скрипты C#\Библиотеки\AtomStyleLibrary\AtomStyleLibrary.dll");
            }
            catch
            {
                WarningWindow warningWindow = new WarningWindow("ОШИБКА", "Ошибка при загрузке библиотеки стилей");
                warningWindow.ShowDialog();

                return Result.Cancelled;
            }

            if (RebarsToExcelApp.IsOpened)
            {
                WarningWindow warningWindow = new WarningWindow("ПРЕДУПРЕЖДЕНИЕ", "Программа уже запущена");
                warningWindow.ShowDialog();

                return Result.Cancelled;
            }

            UIApplication uiapp = commandData.Application;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                DataManageVM dataManageVM = new DataManageVM(doc);

                if (dataManageVM.IsDataEmpty)
                {
                    WarningWindow warningWindow = new WarningWindow("ПРЕДУПРЕЖДЕНИЕ", "Сборочные единицы в модели не найдены");
                    warningWindow.ShowDialog();

                    return Result.Succeeded;
                }

                dataManageVM.BackgroundWorker.RunWorkerCompleted += (sender, e) =>
                {
                    StartWindow startWindow = new StartWindow(uiapp, dataManageVM);
                    startWindow.Show();
                    startWindow.Activate();
                };
            }
            catch (Exception e)
            {
                ExceptionWindow exceptionWindow = new ExceptionWindow(e.Message, e.StackTrace);
                exceptionWindow.ShowDialog();

                return Result.Cancelled;
            }

            return Result.Succeeded;
        }
    }
}
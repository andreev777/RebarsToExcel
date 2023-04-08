using Autodesk.Revit.UI;
using System;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace RebarsToExcel.Commands
{
    public class RebarsToExcelApp : IExternalApplication
    {
        public static bool IsOpened = false;

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "Надстройки АСК";
            string panelName = "Экспорт данных";
            string assemblyName = Assembly.GetExecutingAssembly().Location;
            string commandName = typeof(StartCommand).FullName;

            string toolTip = "Экспорт деталей и сборочных единиц проекта в таблицу Excel";

            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch { }

            var ribbonPanels = application.GetRibbonPanels(tabName);
            var ribbonPanel = ribbonPanels.FirstOrDefault(panel => panel.Name == panelName) ?? application.CreateRibbonPanel(tabName, panelName);

            PushButtonData startCommandButtonData = new PushButtonData("StartCommand", "Арматура", assemblyName, commandName);

            PushButton startCommandButton = ribbonPanel.AddItem(startCommandButtonData) as PushButton;
            startCommandButton.ToolTip = toolTip;

            BitmapImage startCommandButtonLogo = new BitmapImage(new Uri("pack://application:,,,/RebarsToExcel;component/Images/startCommandButtonLogo.png"));
            startCommandButton.LargeImage = startCommandButtonLogo;

            return Result.Succeeded;
        }
    }
}
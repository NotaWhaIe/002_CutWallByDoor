using Autodesk.Revit.UI;

using System;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace CutWallByDoor
{
    public class App : IExternalApplication
    {
        private DeleteLog deleteLogInstance;
        public Result OnStartup(UIControlledApplication application)
        {
            deleteLogInstance = new DeleteLog();
            deleteLogInstance.Initialize(application);

            application.CreateRibbonTab("Архитектура");

            _ = CreateRibbonPanel(application);
            
            return Result.Succeeded;
        }
        public RibbonPanel CreateRibbonPanel(UIControlledApplication application, string tabName = "Архитектура")
        {
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Архитектура");
            AddPushButton(ribbonPanel, "Объединение", Assembly.GetExecutingAssembly().Location, "CutWallByDoor.CommandCutWallByDoor", "Объединит стены с семейством Дверь_Без основы", @"/CutWallByDoor;component/Resources/walldoor.png");
            AddPushButton(ribbonPanel, "Проверка\nмарок ", Assembly.GetExecutingAssembly().Location, "CutWallByDoor.CommandCheckMarkAr", "Проверка марок перемычек, окон и витражей", @"/CutWallByDoor;component/Resources/checkMarkAr32.png") ;
            return ribbonPanel;
        }
        public void AddPushButton(RibbonPanel ribbonPanel, string buttonName, string path, string linkToCommand, string toolTip, string imagePath)
        {
            var buttonData = new PushButtonData(buttonName, buttonName, path, linkToCommand);
            var button = ribbonPanel.AddItem(buttonData) as PushButton;
            button.ToolTip = toolTip;
            button.LargeImage = new BitmapImage(new Uri(imagePath, UriKind.RelativeOrAbsolute));
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            deleteLogInstance?.Shutdown(application);
            return Result.Succeeded;
        }

    }
}

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Configuration.Assemblies;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Resources;
using System.Windows.Interop;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace CutWallByDoor
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            application.CreateRibbonTab("Архитектура");

            _ = CreateRibbonPanel(application);
            
            return Result.Succeeded;
        }

        public RibbonPanel CreateRibbonPanel(UIControlledApplication application, string tabName = "Архитектура")
        {
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Архитектура");
            AddPushButton(ribbonPanel, "Объединение", Assembly.GetExecutingAssembly().Location, "CutWallByDoor.Command", "Объединит стены с семейством Дверь_Без основы");
            return ribbonPanel;
        }

        public void AddPushButton(RibbonPanel ribbonPanel, string buttonName, string path, string linkToCommand, string toolTip)
        {
            var buttonData = new PushButtonData(buttonName, buttonName, path, linkToCommand);
            var button = ribbonPanel.AddItem(buttonData) as PushButton;
            button.ToolTip = toolTip;
            button.LargeImage = (ImageSource)new BitmapImage(new Uri(@"/CutWallByDoor;component/Resources/walldoor.png", UriKind.RelativeOrAbsolute));
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

    }
}

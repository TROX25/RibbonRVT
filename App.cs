#region Namespaces
using System;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using System.IO;
#endregion

namespace Ribbon
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            
                // Create a custom ribbon tab
                String tabName = "Meet the engineer";
                a.CreateRibbonTab(tabName);

                // Create two push buttons
                string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
                PushButtonData button1 = new PushButtonData("Button1", "My Button #1", thisAssemblyPath, "Ribbon.MyTest");
                PushButtonData button2 = new PushButtonData("Button2", "My Button #2", thisAssemblyPath, "Ribbon.MyTest");

                // Create a ribbon panel
                RibbonPanel m_projectPanel = a.CreateRibbonPanel(tabName, "First Panel name");
                RibbonPanel m_projectPanel2 = a.CreateRibbonPanel(tabName, "Second Panel Name");


                // Add the buttons to the panel
                PushButton pb1 = m_projectPanel.AddItem(button1) as PushButton;
                PushButton pb2 = m_projectPanel2.AddItem(button2) as PushButton;

                //Set images for buttons
                pb1.LargeImage = LoadImage("Ribbon;component/Resources/icon.png");
                // pb2.LargeImage = LoadImageFromFile(@"E:\ICONS\ace-of-spades-playing-card-25429.png");
                pb2.LargeImage = LoadImageFromFile(@"E:\ICONS\wrench-and-screwdriver-9453.png");
                return Result.Succeeded;
            
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
        private BitmapImage LoadImage(string relativePath)
        {
            try
            {
                Uri uri = new Uri($"pack://application:,,,/{relativePath}", UriKind.Absolute);
                return new BitmapImage(uri);
            }
            catch
            {
                TaskDialog.Show("Error", $"Failed to load embedded image: {relativePath}");
                return null;
            }
        }

        private BitmapImage LoadImageFromFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(filePath, UriKind.Absolute);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    return bitmap;
                }
                else
                {
                    TaskDialog.Show("Error", $"Image file not found: {filePath}");
                    return null;
                }
            }
            catch
            {
                TaskDialog.Show("Error", $"Failed to load image from file: {filePath}");
                return null;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class MyTest : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("123", "TEST TEST TEST \n TEST TEST TEST");
            return Result.Succeeded;
        }
    }
}

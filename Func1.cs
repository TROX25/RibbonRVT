#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using WPF_1;

#endregion

namespace Ribbon
{
    [Transaction(TransactionMode.Manual)]
    public class Excel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            MainWindow window = new MainWindow(uiapp);
            window.ShowDialog();  // Otwiera okno WPF
            return Result.Succeeded;
        }
    }
}

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
    public class MyTest : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("123", "PEDRO");
            return Result.Succeeded;
        }
    }
}

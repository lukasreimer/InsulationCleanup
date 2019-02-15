// using System;
// using System.Collections.Generic;
// using System.Linq;
// using System.Text;
// using System.Threading.Tasks;

// using Autodesk.Revit.ApplicationServices;
// using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
// using Autodesk.Revit.UI.Selection;

namespace InsulationCleanup
{
    public class ChangeToHostWorkset : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Insulation Cleanup Addin", "Changing Insulation to Host Workset...");
            return Result.Succeeded;
        }
    }
}

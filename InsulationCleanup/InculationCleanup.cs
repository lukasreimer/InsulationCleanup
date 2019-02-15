using System;
using System.Collections.Generic;
// using System.Linq;
// using System.Text;
// using System.Threading.Tasks;

// using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace InsulationCleanup
{

    [Transaction(TransactionMode.Manual)]
    public class ChangeToHostWorkset : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // TaskDialog.Show("Insulation Cleanup Addin", "Changing Insulation to Host Workset...");

                UIApplication uiapp = commandData.Application;
                Document doc = uiapp.ActiveUIDocument.Document;

                // Select Pipe Insulation
                ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations);
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Element> insulations = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

                String prompt = "The roque insulations in the current document are:\n";
                foreach (Element element in insulations)
                {
                    WorksetId insulationWorksetId = element.WorksetId;
                    InsulationLiningBase insulationElement = element as InsulationLiningBase;
                    ElementId hostId = insulationElement.HostElementId;
                    Element hostElement = doc.GetElement(hostId);
                    WorksetId hostWorksetId = hostElement.WorksetId;
                    if (hostWorksetId != insulationWorksetId)
                    {
                        prompt += "Insulation: " + insulationElement.Id + ", " + insulationWorksetId
                               + " - Host: " + hostElement.Id + ", " + hostWorksetId
                               + "\n";
                    }
                }
                TaskDialog.Show("Revit", prompt);

                // Get Workset of Pipe Insulation
                // Select Host of Pipe Insulation
                // Get Workset of Host
                // Compare Worksets
                // Move Insulation to Host Workset
                //Transaction trans = new Transaction(doc);
                //trans.Start("TRANSACTION NAME");
                //...
                //trans.Commit();
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace InsulationCleanup
{

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ChangeToHostWorkset : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get application and document objects
                UIApplication uiapp = commandData.Application;
                Document doc = uiapp.ActiveUIDocument.Document;

                // Select all pipe insulation elements
                var filter = new ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations);
                var collector = new FilteredElementCollector(doc);
                IList<Element> insulationElements = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();
                   
                // Collect all roque insulation elements (insulation workset is not host workset)
                var roqueElements = new List<Tuple<PipeInsulation, Element>>();  // (roque insulation object, target element object)
                foreach (var element in insulationElements)
                {
                    PipeInsulation insulationElement = element as PipeInsulation;
                    Element hostElement = doc.GetElement(insulationElement.HostElementId);
                    if (insulationElement.WorksetId != hostElement.WorksetId)
                    {
                        var roqueTuple = Tuple.Create(insulationElement, hostElement);
                        roqueElements.Add(roqueTuple);
                    }
                }

                String prompt = "The roque insulations in the current document are:\n";
                foreach (var tuple in roqueElements)
                {
                    PipeInsulation roqueElement = tuple.Item1;
                    Element targetElement = tuple.Item2;
                    prompt += $"roque insulation: {roqueElement.Name} #{roqueElement.Id} @ workset: {roqueElement.WorksetId}, "+
                              $"target element: {targetElement.Name} #{targetElement.Id} @ workset: {targetElement.WorksetId}\n";
                }
                TaskDialog.Show("Revit", prompt);

                // Move Insulation to Host Workset
                var trans = new Transaction(doc, "ChangeToHostWorkset");
                trans.Start();
                foreach (Tuple<PipeInsulation, Element> tuple in roqueElements)
                {
                    // do business...
                }
                trans.Commit();
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

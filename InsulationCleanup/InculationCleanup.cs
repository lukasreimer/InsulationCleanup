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
        private String title = "ChangeToHostWorkset Addin";

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

                // Output a report of the identified roque insulation elements
                String prompt = $"There are {roqueElements.Count} roque insulation elements in this document. ";
                if (roqueElements.Count != 0)
                {
                    prompt += "They are:\n\n";
                    foreach (var tuple in roqueElements)
                    {
                        PipeInsulation roqueElement = tuple.Item1;
                        Element targetElement = tuple.Item2;
                        prompt += $"roque insulation: {roqueElement.Name} #{roqueElement.Id} @ workset: {roqueElement.WorksetId}, ";
                        prompt += $"target element: {targetElement.Name} #{targetElement.Id} @ workset: {targetElement.WorksetId}\n";
                    }
                }
                TaskDialog.Show(this.title, prompt);

                // Save the current workset for resetting later
                WorksetTable worksets = doc.GetWorksetTable();
                WorksetId initalWorksetId = worksets.GetActiveWorksetId();

                // Delete roque insulation and create new insulation on correct host workset
                using (var trans = new Transaction(doc))
                {
                    trans.Start("ChangeToHostWorkset");
                    foreach (var tuple in roqueElements)
                    {
                        // Unpack tuple
                        PipeInsulation roqueElement = tuple.Item1;
                        Element targetElement = tuple.Item2;
                        // Get data needed for PipeInsulation "constructor"
                        var insulationTypeId = roqueElement.GetTypeId();
                        var insulationThickness = roqueElement.Thickness;
                        // Delete roque insulation and create new insulation on the correct host
                        doc.Delete(roqueElement.Id);
                        try
                        {
                            worksets.SetActiveWorksetId(targetElement.WorksetId);
                            PipeInsulation.Create(doc, targetElement.Id, insulationTypeId, insulationThickness);
                        }
                        catch (Exception)
                        {
                            // TODO: keep track of failed reincarnations and report after commiting the transaction
                            // ...
                        }
                    }
                    // Reset the active workset
                    worksets.SetActiveWorksetId(initalWorksetId);
                    trans.Commit();
                }
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

using System;
using System.Collections.Generic;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

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

                // TODO: extract into method
                // Select all pipe insulation elements
                var filter = new ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations);
                var collector = new FilteredElementCollector(doc);
                IList<Element> insulationElements = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();
                // Collect all rogue insulation elements (where insulation workset is not host workset)
                var rogueElements = new List<Tuple<PipeInsulation, Element>>();  // (rogue insulation element, hosting element)
                foreach (var element in insulationElements)
                {
                    PipeInsulation insulationElement = element as PipeInsulation;
                    Element hostElement = doc.GetElement(insulationElement.HostElementId);
                    if (insulationElement.WorksetId != hostElement.WorksetId)
                    {
                        var rogueTuple = Tuple.Create(insulationElement, hostElement);
                        rogueElements.Add(rogueTuple);
                    }
                }

                // Show a report of the found rogue insulation elements
                TaskDialogResult result = ShowReport(rogueElements);

                if (result == TaskDialogResult.Yes)
                {
                    // TODO: extract into method
                    // Save the current workset for resetting later
                    WorksetTable worksets = doc.GetWorksetTable();
                    WorksetId initalWorksetId = worksets.GetActiveWorksetId();

                    // Delete rogue insulation and create new insulation on correct host workset
                    using (var trans = new Transaction(doc))
                    {
                        trans.Start("ChangeToHostWorkset");
                        foreach (var tuple in rogueElements)
                        {
                            // Unpack tuple
                            PipeInsulation rogueElement = tuple.Item1;
                            Element targetElement = tuple.Item2;
                            // Get data needed for PipeInsulation "constructor"
                            var insulationTypeId = rogueElement.GetTypeId();
                            var insulationThickness = rogueElement.Thickness;
                            // Delete rogue insulation and create new insulation on the correct host
                            doc.Delete(rogueElement.Id);
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

                        return Result.Succeeded;
                    }
                }
                else  // not TaskDialogResult.Yes
                {
                    return Result.Cancelled;
                }
            }
            catch (Exception e)  // any exception
            {
                message = e.Message;
                return Result.Failed;
            }
        }

        private TaskDialogResult ShowReport(List<Tuple<PipeInsulation, Element>> rogueElements)
        {
            // Setup text components for the dialog
            string title = "Rogue Insulation Report";
            string message = $"There are {rogueElements.Count} rogue insulation elements in this document.";
            string content = "";
            //string footer = "Do you want to move the rogue insulation to the hosts worksets?";
           
            // Populate the detailed content part of the dialog
            if (rogueElements.Count != 0)  // there are rogue elements
            {
                message += "\nDo you want to move these elements to their hosts workset?";
                // iterate through all (rogue insulation, host element) tuples
                foreach (var tuple in rogueElements)
                {
                    // unpack the tuple
                    PipeInsulation rogueElement = tuple.Item1;
                    Element targetElement = tuple.Item2;
                    // create a report detail line
                    // TODO: add line numbers
                    // TODO: figure out the workset names
                    content += $"insulation: {rogueElement.Name} #{rogueElement.Id} @ workset: {rogueElement.WorksetId}, ";
                    content += $"hosted by: {targetElement.Name} #{targetElement.Id} @ workset: {targetElement.WorksetId}\n";
                }
            }
            
            // Configure and show dialog
            TaskDialog dialog = new TaskDialog(title);
            dialog.MainInstruction = message;
            dialog.ExpandedContent = content;
            //dialog.FooterText = footer;
            dialog.AllowCancellation = false;
            if (rogueElements.Count != 0)
            {
                dialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
                dialog.DefaultButton = TaskDialogResult.No;
            }
            else
            {
                dialog.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
                dialog.CommonButtons = TaskDialogCommonButtons.Ok;
                dialog.DefaultButton = TaskDialogResult.Ok;
            }
            // TODO: add report export option
           
            TaskDialogResult result = dialog.Show();
            return result;
        }
    }
}

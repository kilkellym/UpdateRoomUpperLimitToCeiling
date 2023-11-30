#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace UpdateRoomUpperLimitToCeiling
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            int counter = 0;

            // Collect all rooms in the document
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> rooms = collector.OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements();

            using (Transaction trans = new Transaction(doc, "Update All Room Heights"))
            {
                trans.Start();

                foreach (Element elem in rooms)
                {
                    Room room = elem as Room;
                    if (room != null)
                    {
                        // Use the room's LocationPoint as the start point
                        LocationPoint roomLocationPoint = room.Location as LocationPoint;
                        if (roomLocationPoint == null)
                            continue; // Skip rooms without a valid location

                        XYZ startPoint = roomLocationPoint.Point;
                        XYZ endPoint = new XYZ(startPoint.X, startPoint.Y, startPoint.Z + 80); // 10 units above for intersection

                        // Find the Z value of the ceiling intersection
                        double? ceilingIntersectionZ = FindCeilingIntersection(doc, startPoint, endPoint, room.LevelId);
                        if (ceilingIntersectionZ == null)
                            continue; // Skip rooms without a ceiling above

                        // Update the room's Upper Limit parameter
                        Parameter upperLimitParam = room.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET);
                        if (upperLimitParam != null && !upperLimitParam.IsReadOnly)
                        {
                            upperLimitParam.Set(ceilingIntersectionZ.Value);
                        }

                        counter++;
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Complete", $"Updated {counter} rooms.");

            return Result.Succeeded;
        }
        private double? FindCeilingIntersection(Document doc, XYZ startPoint, XYZ endPoint, ElementId roomLevelId)
        {
            Line line = Line.CreateBound(startPoint, endPoint);
            View3D view3D = GetOrCreate3DView(doc);

            ReferenceIntersector intersector = new ReferenceIntersector(view3D);
            intersector.TargetType = FindReferenceTarget.Element;
            intersector.FindReferencesInRevitLinks = true;

            // Set the filter for the ceiling category
            ElementCategoryFilter ceilingFilter = new ElementCategoryFilter(BuiltInCategory.OST_Ceilings);
            intersector.SetFilter(ceilingFilter);

            IList<ReferenceWithContext> intersectedRefs = intersector.Find(startPoint, endPoint.Subtract(startPoint));
            foreach (ReferenceWithContext refWithContext in intersectedRefs)
            {
                Reference reference = refWithContext.GetReference();
                Element intersectedElement = doc.GetElement(reference.ElementId);
                Ceiling ceiling = intersectedElement as Ceiling;

                if (ceiling != null && ceiling.LevelId == roomLevelId)
                {
                    Parameter heightOffsetParam = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
                    if (heightOffsetParam != null && heightOffsetParam.StorageType == StorageType.Double)
                    {
                        // Convert the internal unit (feet) to the desired unit if necessary
                        return heightOffsetParam.AsDouble(); // Returns the value in Revit's internal units (feet)
                    }
                }
            }

            return null;

        }

        private View3D GetOrCreate3DView(Document doc)
        {
            // Try to find an existing 3D view
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> view3DCollection = collector.OfClass(typeof(View3D)).ToElements();

            foreach (Element e in view3DCollection)
            {
                View3D view3D = e as View3D;

                // Check if the 3D view is not a template
                if (view3D != null && !view3D.IsTemplate)
                {
                    return view3D; // Return the first non-template 3D view found
                }
            }

            // If no suitable view is found, create a new one
            try
            {
                // Start a sub-transaction to create a new 3D view
                using (Transaction trans = new Transaction(doc, "Create 3D View"))
                {
                    trans.Start();

                    ViewFamilyType viewFamilyType = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);

                    if (viewFamilyType != null)
                    {
                        View3D newView3D = View3D.CreateIsometric(doc, viewFamilyType.Id);
                        trans.Commit();
                        return newView3D;
                    }
                    else
                    {
                        trans.RollBack();
                        throw new InvalidOperationException("No 3D view family type found.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create 3D view: " + ex.Message);
            }
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}

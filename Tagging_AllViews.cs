#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Controls;
using System.Windows.Documents;

#endregion

namespace Intermediate_Module_02_v2
{
    [Transaction(TransactionMode.Manual)]
    public class Tagging_AllViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
             this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

             this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            FilteredElementCollector collectorViewsInModel = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType();

            View activeView = doc.ActiveView;
           

            1c. Use LINQ to get family symbol by name
            #region FamilySymbols
            // get room tags via LINQ
            FamilySymbol curtwallTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Curtain Wall Tag"))
                .First();
            FamilySymbol doorTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Door Tag"))
                .First();
            FamilySymbol furnituregTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName
                .Equals("M_Furniture Tag")).First();
            FamilySymbol lightingTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName
                .Equals("M_Lighting Fixture Tag")).First();
            FamilySymbol roomTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Room Tag"))
                .First();
            FamilySymbol wallTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Wall Tag"))
                .First();
            FamilySymbol windowTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Window Tag"))
                .First();
            #endregion

            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            // adding key & values (Add.(Key, Value)) to dictionary
            tags.Add("CurtainWalls", curtwallTag);
            tags.Add("Doors", doorTag);
            tags.Add("Furniture", furnituregTag);
            tags.Add("Lighting Fixture", lightingTag);
            tags.Add("Rooms", roomTag);
            tags.Add("Walls", wallTag);
            tags.Add("Windows", windowTag);

            double counter = 0;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert tag");
                foreach (View currentView in collectorViewsInModel)
                {
                    if (currentView.IsTemplate == true)
                        continue;
                        
                    XYZ insertionPoint = GetInsertionForTag2(currentElementInView);
                    string tagCatName = currentElementInView.Category.Name;
                    if (insertionPoint != null)
                    {
                        ViewType currentViewType = currentView.ViewType;
                        if (currentViewType == ViewType.CeilingPlan)
                        {
                            if (tagCatName == "Lighting Fixtures" || tagCatName == "Rooms")
                            {
                                FamilySymbol currentTag = tags[tagCatName];
                                Reference currentElementReference = new Reference(currentElementInView);
                                IndependentTag newTag = IndependentTag.Create(doc, currentTag.Id, currentView.Id, currentElementReference, false, TagOrientation.Horizontal, insertionPoint);
                                counter++;
                            }
                        }
                        if (currentViewType == ViewType.FloorPlan)
                        {
                            if (tagCatName == "Walls" || tagCatName == "Rooms" || tagCatName == "Furniture" || tagCatName == "Doors")
                            {
                                FamilySymbol currentTag = tags[tagCatName];
                                Reference currentElementReference = new Reference(currentElementInView);
                                IndependentTag newTag = IndependentTag.Create(doc, currentTag.Id, currentView.Id, currentElementReference, false, TagOrientation.Horizontal, insertionPoint);
                                counter++;
                            }
                            else if (tagCatName == "Windows")
                            {
                                FamilySymbol currentTag = tags[tagCatName];
                                Reference currentElementReference = new Reference(currentElementInView);
                                XYZ windowInsertionPoint = new XYZ(insertionPoint.X, insertionPoint.Y + 3, insertionPoint.Z);
                                IndependentTag newTag = IndependentTag.Create(doc, currentTag.Id, currentView.Id, currentElementReference, false, TagOrientation.Horizontal, windowInsertionPoint);
                                counter++;
                            }
                        }
                        if (currentViewType == ViewType.Section)
                        {
                            if (tagCatName == "Rooms")
                            {
                                FamilySymbol currentTag = tags[tagCatName];

                                string roomLevel = currentElementInView.LookupParameter("Level").AsString();
                                FilteredElementCollector collectorFloor = new FilteredElementCollector(doc, currentView.Id).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
                                foreach (Element floor in collectorFloor)
                                {
                                    Double floorHeightOffset = floor.LookupParameter("Level").AsDouble() + 3;
                                    Reference currentElementReference = new Reference(currentElementInView);
                                    XYZ roomInsertionPoint = new XYZ(insertionPoint.X, insertionPoint.Y, floorHeightOffset + 5);
                                    IndependentTag newTag = IndependentTag.Create(doc, currentTag.Id, currentView.Id, currentElementReference, false, TagOrientation.Horizontal, roomInsertionPoint);
                                    counter++;
                                }
                            }
                        }
                        if (currentViewType == ViewType.AreaPlan)
                        {
                            if (tagCatName == "Areas")
                            {
                                ViewPlan curAreaPlan = currentView as ViewPlan;
                                Area curArea = currentElementInView as Area;
                                AreaTag curAreaTag = doc.Create.NewAreaTag(curAreaPlan, curArea, new UV(insertionPoint.X, insertionPoint.Y));
                                counter++;
                                curAreaTag.TagHeadPosition = new XYZ(insertionPoint.X, insertionPoint.Y, 0);
                                curAreaTag.HasLeader = false;
                            }
                        }
                        else
                        {
                            continue;
                        }

                    }

                    

                }
                TaskDialog.Show("Tags count", $" Number of tags just placed in almost all model views: {counter}");

                t.Commit();
            }
            return Result.Succeeded;
        }

        private static FilteredElementCollector getAllElementsInView (Document doc, View view)
        {   
            ElementId viewId = view.Id;

            FilteredElementCollector collectorElementsInView = new FilteredElementCollector (doc, viewId);
     
            List<BuiltInCategory> categoriesList = new List<BuiltInCategory>();
            categoriesList.Add(BuiltInCategory.OST_Areas);
            categoriesList.Add(BuiltInCategory.OST_Walls);
            categoriesList.Add(BuiltInCategory.OST_Doors);
            categoriesList.Add(BuiltInCategory.OST_Furniture);
            categoriesList.Add(BuiltInCategory.OST_LightingFixtures);
            categoriesList.Add(BuiltInCategory.OST_Rooms);
            categoriesList.Add(BuiltInCategory.OST_Windows);

            ElementMulticategoryFilter categoriesFilter = new ElementMulticategoryFilter(categoriesList);

            collectorElementsInView.WherePasses(categoriesFilter).WhereElementIsNotElementType();
            return collectorElementsInView;
        }

        private XYZ GetInsertionForTag(FilteredElementCollector coll)
        {
            foreach (Element curElement in coll)
            {
                XYZ insPoint;
                LocationPoint locationPoint;
                LocationCurve locationCurve;

                Location curElementLocation = curElement.Location;

                if (curElementLocation == null)
                    continue;

                locationPoint = curElementLocation as LocationPoint;

                if (locationPoint != null)
                {
                    insPoint = locationPoint.Point;
                    return insPoint;
                }

                else
                {
                    locationCurve = curElementLocation as LocationCurve;
                    Curve currentLocationCurve = locationCurve.Curve;
                    insPoint = GetMidpointBetweenToPoints(currentLocationCurve.GetEndPoint(0), currentLocationCurve.GetEndPoint(1));
                    return insPoint;
                }
            }
            return null;
        }
        private XYZ GetInsertionForTag2(Element elem)
        {

            XYZ insPoint;
            LocationPoint locationPoint;
            LocationCurve locationCurve;
            Location curElementLocation = elem.Location;

            locationPoint = curElementLocation as LocationPoint;

            if (curElementLocation != null && locationPoint != null)
            {
                insPoint = locationPoint.Point;
                return insPoint;
            }

            else if(curElementLocation != null)
            {
                locationCurve = curElementLocation as LocationCurve;
                Curve currentLocationCurve = locationCurve.Curve;
                insPoint = GetMidpointBetweenToPoints(currentLocationCurve.GetEndPoint(0), currentLocationCurve.GetEndPoint(1));
                return insPoint;
            }
            
            return null;
        }

        private XYZ GetMidpointBetweenToPoints(XYZ point1, XYZ point2)
        {
            XYZ midPoint = new XYZ(
                (point1.X + point2.X) / 2,
                (point1.Y + point2.Y) / 2,
                (point1.Z + point2.Z) / 2);
            return midPoint;
        }

        internal static PushButtonData GetButtonData()
        {
             use this method to define the properties for this command in the Revit ribbon
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


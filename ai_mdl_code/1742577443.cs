using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConduitConnector
{
    public class ConduitConnectionHandler
    {
        private readonly Document _doc;

        public ConduitConnectionHandler(Document document)
        {
            _doc = document ?? throw new ArgumentNullException(nameof(document));
        }

        public Result ConnectConduits(ElementId conduit1Id, ElementId conduit2Id)
        {
            string message = string.Empty;

            try
            {
                // Get the conduit elements
                Conduit conduit1 = _doc.GetElement(conduit1Id) as Conduit;
                Conduit conduit2 = _doc.GetElement(conduit2Id) as Conduit;

                if (conduit1 == null || conduit2 == null)
                {
                    TaskDialog.Show("Error", "One or both conduits could not be found.");
                    return Result.Failed;
                }

                // Get the location curves
                LocationCurve locationCurve1 = conduit1.Location as LocationCurve;
                LocationCurve locationCurve2 = conduit2.Location as LocationCurve;

                if (locationCurve1 == null || locationCurve2 == null)
                {
                    TaskDialog.Show("Error", "Could not get location curves for conduits.");
                    return Result.Failed;
                }

                // Get the lines from the curves
                Line line1 = locationCurve1.Curve as Line;
                Line line2 = locationCurve2.Curve as Line;

                if (line1 == null || line2 == null)
                {
                    TaskDialog.Show("Error", "Conduits must be straight lines.");
                    return Result.Failed;
                }

                // Calculate the closest points between the two lines
                ClosestPointResult closestPoints = FindClosestPointsBetweenLines(line1, line2);

                // Get the connectors
                ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
                ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;

                // Find the free connectors
                Connector freeConnector1 = FindFreeConnector(connectors1);
                Connector freeConnector2 = FindFreeConnector(connectors2);

                if (freeConnector1 == null || freeConnector2 == null)
                {
                    TaskDialog.Show("Error", "Could not find free connectors on conduits.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(_doc, "Connect Conduits"))
                {
                    trans.Start();

                    // Determine if we need to extend or trim the conduits
                    // First, adjust conduit1
                    locationCurve1.Curve = Line.CreateBound(line1.GetEndPoint(0), closestPoints.Point1);

                    // Then, adjust conduit2
                    locationCurve2.Curve = Line.CreateBound(line2.GetEndPoint(0), closestPoints.Point2);

                    // Create an elbow fitting to connect the conduits
                    FamilyInstance elbow = _doc.Create.NewElbowFitting(freeConnector1, freeConnector2);

                    if (elbow == null)
                    {
                        TaskDialog.Show("Error", "Failed to create elbow fitting.");
                        trans.RollBack();
                        return Result.Failed;
                    }

                    trans.Commit();
                }

                TaskDialog.Show("Success", "Conduits connected successfully.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }

        private Connector FindFreeConnector(ConnectorSet connectors)
        {
            foreach (Connector connector in connectors)
            {
                if (connector.IsConnected == false)
                {
                    return connector;
                }
            }
            return null;
        }

        private class ClosestPointResult
        {
            public XYZ Point1 { get; set; }
            public XYZ Point2 { get; set; }
            public double Distance { get; set; }
        }

        private ClosestPointResult FindClosestPointsBetweenLines(Line line1, Line line2)
        {
            // Get the direction vectors of the lines
            XYZ dir1 = line1.Direction;
            XYZ dir2 = line2.Direction;

            // Get the origins of the lines
            XYZ origin1 = line1.Origin;
            XYZ origin2 = line2.Origin;

            // Calculate the vector between the origins
            XYZ w0 = origin1 - origin2;

            // Calculate the dot products
            double a = dir1.DotProduct(dir1);
            double b = dir1.DotProduct(dir2);
            double c = dir2.DotProduct(dir2);
            double d = dir1.DotProduct(w0);
            double e = dir2.DotProduct(w0);

            // Calculate the denominators for the parametric equations
            double denominator = a * c - b * b;

            // Initialize parameters
            double sc, tc;

            // Check if lines are parallel (denominator is zero or very close to zero)
            if (Math.Abs(denominator) < 1e-10)
            {
                // Lines are parallel, use any point on line2 and find closest point on line1
                sc = 0.0;
                tc = d / b;
            }
            else
            {
                // Calculate the parameters for the closest points
                sc = (b * e - c * d) / denominator;
                tc = (a * e - b * d) / denominator;
            }

            // Calculate the closest points
            XYZ point1 = origin1 + sc * dir1;
            XYZ point2 = origin2 + tc * dir2;

            // Calculate the distance between the closest points
            double distance = point1.DistanceTo(point2);

            return new ClosestPointResult
            {
                Point1 = point1,
                Point2 = point2,
                Distance = distance
            };
        }

        public static Result ExecuteForConduits(UIDocument uidoc, ElementId conduit1Id, ElementId conduit2Id)
        {
            Document doc = uidoc.Document;
            ConduitConnectionHandler handler = new ConduitConnectionHandler(doc);
            return handler.ConnectConduits(conduit1Id, conduit2Id);
        }
    }

    public class ConduitConnectorCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // For this example, we'll use the conduit IDs provided
                ElementId conduit1Id = new ElementId(7638672);
                ElementId conduit2Id = new ElementId(7724387);

                return ConduitConnectionHandler.ExecuteForConduits(uidoc, conduit1Id, conduit2Id);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
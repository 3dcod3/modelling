using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConnectConduits
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ConnectConduitsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get the conduits by their IDs
                ElementId conduit1Id = new ElementId(7721116);
                ElementId conduit2Id = new ElementId(7721143);

                Conduit conduit1 = doc.GetElement(conduit1Id) as Conduit;
                Conduit conduit2 = doc.GetElement(conduit2Id) as Conduit;

                if (conduit1 == null || conduit2 == null)
                {
                    message = "Could not find the specified conduits.";
                    return Result.Failed;
                }

                // Connect the conduits
                Result result = ConnectConduits(doc, conduit1, conduit2);
                return result;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private Result ConnectConduits(Document doc, Conduit conduit1, Conduit conduit2)
        {
            // Get the location curves of the conduits
            LocationCurve locationCurve1 = conduit1.Location as LocationCurve;
            LocationCurve locationCurve2 = conduit2.Location as LocationCurve;

            if (locationCurve1 == null || locationCurve2 == null)
            {
                return Result.Failed;
            }

            Line line1 = locationCurve1.Curve as Line;
            Line line2 = locationCurve2.Curve as Line;

            if (line1 == null || line2 == null)
            {
                return Result.Failed;
            }

            // Analyze the relationship between the conduits
            XYZ direction1 = line1.Direction;
            XYZ direction2 = line2.Direction;

            // Calculate dot product to determine if they're perpendicular
            double dotProduct = direction1.DotProduct(direction2);
            bool isPerpendicular = Math.Abs(dotProduct) < 0.001;

            // Find the closest points between the two lines
            XYZ closestPoint1, closestPoint2;
            FindClosestPoints(line1, line2, out closestPoint1, out closestPoint2);

            // Start a transaction
            using (Transaction trans = new Transaction(doc, "Connect Conduits"))
            {
                trans.Start();

                try
                {
                    // Adjust conduit endpoints to meet at the closest points
                    AdjustConduitEndpoints(doc, conduit1, conduit2, closestPoint1, closestPoint2);

                    // Get the updated connectors
                    Connector connector1 = GetNearestConnector(conduit1, closestPoint1);
                    Connector connector2 = GetNearestConnector(conduit2, closestPoint2);

                    if (connector1 != null && connector2 != null)
                    {
                        // Create an elbow fitting
                        doc.Create.NewElbowFitting(connector1, connector2);
                        trans.Commit();
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        return Result.Failed;
                    }
                }
                catch (Exception)
                {
                    if (trans.HasStarted())
                    {
                        trans.RollBack();
                    }
                    return Result.Failed;
                }
            }
        }

        private void FindClosestPoints(Line line1, Line line2, out XYZ closestPoint1, out XYZ closestPoint2)
        {
            // Get the direction vectors of the lines
            XYZ v1 = line1.Direction;
            XYZ v2 = line2.Direction;

            // Get the start points of the lines
            XYZ p1 = line1.GetEndPoint(0);
            XYZ p2 = line2.GetEndPoint(0);

            // Vector between start points
            XYZ w0 = p1 - p2;

            // Calculate parameters
            double a = v1.DotProduct(v1);
            double b = v1.DotProduct(v2);
            double c = v2.DotProduct(v2);
            double d = v1.DotProduct(w0);
            double e = v2.DotProduct(w0);

            // Calculate parameters for closest points
            double denominator = a * c - b * b;
            
            // Handle parallel lines
            if (Math.Abs(denominator) < 1e-10)
            {
                // For parallel lines, project one endpoint onto the other line
                double t1 = 0;
                closestPoint1 = p1;
                closestPoint2 = p2 + t1 * v2;
                return;
            }

            double t1 = (b * e - c * d) / denominator;
            double t2 = (a * e - b * d) / denominator;

            // Calculate closest points
            closestPoint1 = p1 + t1 * v1;
            closestPoint2 = p2 + t2 * v2;
        }

        private void AdjustConduitEndpoints(Document doc, Conduit conduit1, Conduit conduit2, XYZ point1, XYZ point2)
        {
            // Adjust the endpoints of the conduits to meet at the closest points
            LocationCurve locationCurve1 = conduit1.Location as LocationCurve;
            LocationCurve locationCurve2 = conduit2.Location as LocationCurve;

            Line line1 = locationCurve1.Curve as Line;
            Line line2 = locationCurve2.Curve as Line;

            // Determine which endpoint of each conduit is closer to the connection point
            double dist1Start = line1.GetEndPoint(0).DistanceTo(point1);
            double dist1End = line1.GetEndPoint(1).DistanceTo(point1);
            double dist2Start = line2.GetEndPoint(0).DistanceTo(point2);
            double dist2End = line2.GetEndPoint(1).DistanceTo(point2);

            XYZ newStart1 = dist1Start <= dist1End ? point1 : line1.GetEndPoint(0);
            XYZ newEnd1 = dist1Start <= dist1End ? line1.GetEndPoint(1) : point1;
            XYZ newStart2 = dist2Start <= dist2End ? point2 : line2.GetEndPoint(0);
            XYZ newEnd2 = dist2Start <= dist2End ? line2.GetEndPoint(1) : point2;

            // Create new lines with adjusted endpoints
            Line newLine1 = Line.CreateBound(newStart1, newEnd1);
            Line newLine2 = Line.CreateBound(newStart2, newEnd2);

            // Update the conduit locations
            locationCurve1.Curve = newLine1;
            locationCurve2.Curve = newLine2;
        }

        private Connector GetNearestConnector(Conduit conduit, XYZ point)
        {
            ConnectorSet connectors = conduit.ConnectorManager.Connectors;
            Connector nearestConnector = null;
            double minDistance = double.MaxValue;

            foreach (Connector connector in connectors)
            {
                double distance = connector.Origin.DistanceTo(point);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    nearestConnector = connector;
                }
            }

            return nearestConnector;
        }
    }
}
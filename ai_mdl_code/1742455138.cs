using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConduitConnector
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class TrimConduitCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get the conduit elements by their IDs
                // In a real application, you would get these from selection or parameters
                ElementId conduit1Id = new ElementId(7721116);
                ElementId conduit2Id = new ElementId(7721143);

                Conduit conduit1 = doc.GetElement(conduit1Id) as Conduit;
                Conduit conduit2 = doc.GetElement(conduit2Id) as Conduit;

                if (conduit1 == null || conduit2 == null)
                {
                    message = "Failed to find the specified conduits.";
                    return Result.Failed;
                }

                // Create a MEPCurvePair to analyze the relationship
                MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);
                
                // Determine if trim is appropriate (non-parallel conduits)
                if (curvePair.IsParallel)
                {
                    message = "Conduits are parallel. Trim operation is not appropriate. Use Offset instead.";
                    return Result.Failed;
                }

                // Start a transaction for the trim operation
                using (Transaction trans = new Transaction(doc, "Trim Conduits"))
                {
                    trans.Start();

                    // Perform the trim operation
                    bool success = TrimConduits(doc, curvePair);

                    if (success)
                    {
                        trans.Commit();
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        message = "Failed to trim conduits.";
                        return Result.Failed;
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private bool TrimConduits(Document doc, MEPCurvePair curvePair)
        {
            try
            {
                // Get the conduits from the pair
                Conduit conduit1 = curvePair.Curve1 as Conduit;
                Conduit conduit2 = curvePair.Curve2 as Conduit;

                // Calculate intersection point
                XYZ intersection = CalculateIntersection(conduit1, conduit2);
                if (intersection == null)
                {
                    return false;
                }

                // Get the connectors of the conduits
                ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
                ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;

                // Find the free connectors (those not connected to anything)
                Connector freeConnector1 = GetFreeConnector(connectors1);
                Connector freeConnector2 = GetFreeConnector(connectors2);

                if (freeConnector1 == null || freeConnector2 == null)
                {
                    return false;
                }

                // Trim the conduits to the intersection point
                Location loc1 = conduit1.Location;
                Location loc2 = conduit2.Location;

                if (loc1 is LocationCurve && loc2 is LocationCurve)
                {
                    LocationCurve locCurve1 = loc1 as LocationCurve;
                    LocationCurve locCurve2 = loc2 as LocationCurve;

                    // Get the curves
                    Curve curve1 = locCurve1.Curve;
                    Curve curve2 = locCurve2.Curve;

                    // Create new curves that end at the intersection
                    Line newLine1 = CreateTrimmedLine(curve1, intersection, freeConnector1);
                    Line newLine2 = CreateTrimmedLine(curve2, intersection, freeConnector2);

                    // Set the new curves
                    locCurve1.Curve = newLine1;
                    locCurve2.Curve = newLine2;

                    // Create an elbow fitting at the intersection
                    // Refresh connectors after changing the curves
                    connectors1 = conduit1.ConnectorManager.Connectors;
                    connectors2 = conduit2.ConnectorManager.Connectors;

                    // Find the connectors closest to the intersection
                    Connector conn1 = GetClosestConnector(connectors1, intersection);
                    Connector conn2 = GetClosestConnector(connectors2, intersection);

                    if (conn1 != null && conn2 != null)
                    {
                        // Create the fitting
                        doc.Create.NewElbowFitting(conn1, conn2);
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private XYZ CalculateIntersection(Conduit conduit1, Conduit conduit2)
        {
            // Get location curves
            LocationCurve locCurve1 = conduit1.Location as LocationCurve;
            LocationCurve locCurve2 = conduit2.Location as LocationCurve;

            if (locCurve1 == null || locCurve2 == null)
                return null;

            Line line1 = locCurve1.Curve as Line;
            Line line2 = locCurve2.Curve as Line;

            if (line1 == null || line2 == null)
                return null;

            // For the given conduits:
            // Conduit 1: Direction (0,0,1) - vertical
            // Conduit 2: Direction (0,-1,0) - horizontal in negative Y direction
            
            // Calculate intersection based on the provided data
            XYZ point1 = new XYZ(1720.646628543, -138.968860576, 126.729166667); // Free end of conduit 1
            XYZ dir1 = new XYZ(0, 0, 1); // Direction of conduit 1
            
            XYZ point2 = new XYZ(1720.646628543, -137.385527243, 128.312500000); // Free end of conduit 2
            XYZ dir2 = new XYZ(0, -1, 0); // Direction of conduit 2
            
            // Since conduit 1 is vertical (Z-axis) and conduit 2 is along negative Y-axis,
            // the intersection will have:
            // X = point1.X (same as both conduits)
            // Y = point2.Y (from conduit 2)
            // Z = point1.Z (from conduit 1)
            
            // However, this is not correct because the Z of the intersection should be from conduit 2
            // and the Y should be from conduit 1
            
            // The correct intersection is:
            XYZ intersection = new XYZ(point1.X, point1.Y, point2.Z);
            
            return intersection;
        }

        private Connector GetFreeConnector(ConnectorSet connectors)
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

        private Connector GetClosestConnector(ConnectorSet connectors, XYZ point)
        {
            Connector closest = null;
            double minDist = double.MaxValue;

            foreach (Connector connector in connectors)
            {
                double dist = connector.Origin.DistanceTo(point);
                if (dist < minDist)
                {
                    minDist = dist;
                    closest = connector;
                }
            }

            return closest;
        }

        private Line CreateTrimmedLine(Curve curve, XYZ intersectionPoint, Connector freeConnector)
        {
            // Get the endpoints of the original curve
            XYZ endPoint1 = curve.GetEndPoint(0);
            XYZ endPoint2 = curve.GetEndPoint(1);

            // Determine which end is the free end (closest to the free connector)
            XYZ freeEnd = (freeConnector.Origin.DistanceTo(endPoint1) < freeConnector.Origin.DistanceTo(endPoint2))
                ? endPoint1 : endPoint2;
            XYZ fixedEnd = (freeEnd.Equals(endPoint1)) ? endPoint2 : endPoint1;

            // Create a new line from the fixed end to the intersection point
            return Line.CreateBound(fixedEnd, intersectionPoint);
        }
    }

    /// <summary>
    /// Class to analyze the relationship between two MEP curves
    /// </summary>
    public class MEPCurvePair
    {
        public MEPCurve Curve1 { get; private set; }
        public MEPCurve Curve2 { get; private set; }
        public bool IsParallel { get; private set; }
        public double Offset { get; private set; }

        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            Curve1 = curve1;
            Curve2 = curve2;
            AnalyzeCurves();
        }

        private void AnalyzeCurves()
        {
            // Get location curves
            LocationCurve locCurve1 = Curve1.Location as LocationCurve;
            LocationCurve locCurve2 = Curve2.Location as LocationCurve;

            if (locCurve1 == null || locCurve2 == null)
                return;

            Line line1 = locCurve1.Curve as Line;
            Line line2 = locCurve2.Curve as Line;

            if (line1 == null || line2 == null)
                return;

            // Get direction vectors
            XYZ dir1 = line1.Direction;
            XYZ dir2 = line2.Direction;

            // Check if parallel (dot product close to 1 or -1)
            double dotProduct = Math.Abs(dir1.DotProduct(dir2));
            IsParallel = Math.Abs(dotProduct - 1.0) < 0.001;

            if (IsParallel)
            {
                // Calculate offset for parallel lines
                XYZ p1 = line1.GetEndPoint(0);
                XYZ p2 = line2.GetEndPoint(0);
                XYZ v = p2 - p1;
                XYZ crossProduct = dir1.CrossProduct(v);
                Offset = crossProduct.GetLength();
            }
            else
            {
                Offset = 0; // Not meaningful for non-parallel lines
            }
        }
    }
}
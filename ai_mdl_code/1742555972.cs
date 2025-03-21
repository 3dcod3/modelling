using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Linq;
using System.Collections.Generic;

/// <summary>
/// Command to connect two conduits using a trim operation
/// </summary>
public class ConnectConduitsWithTrim
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            // Get the conduit IDs from the input
            ElementId conduitId1 = new ElementId(7721116);
            ElementId conduitId2 = new ElementId(7721143);
            
            // Get the conduit elements
            Conduit conduit1 = doc.GetElement(conduitId1) as Conduit;
            Conduit conduit2 = doc.GetElement(conduitId2) as Conduit;
            
            if (conduit1 == null || conduit2 == null)
            {
                message = "One or both of the specified elements are not conduits.";
                return Result.Failed;
            }
            
            // Start a transaction
            using (Transaction trans = new Transaction(doc, "Connect Conduits with Trim"))
            {
                trans.Start();
                
                // Create a MEPCurvePair to analyze the relationship between the conduits
                MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);
                
                // Get the location curves of the conduits
                LocationCurve locCurve1 = conduit1.Location as LocationCurve;
                LocationCurve locCurve2 = conduit2.Location as LocationCurve;
                
                if (locCurve1 == null || locCurve2 == null)
                {
                    message = "Could not get location curves for the conduits.";
                    trans.RollBack();
                    return Result.Failed;
                }
                
                // Get the curves
                Line line1 = locCurve1.Curve as Line;
                Line line2 = locCurve2.Curve as Line;
                
                if (line1 == null || line2 == null)
                {
                    message = "Conduits must have linear geometry for trim operation.";
                    trans.RollBack();
                    return Result.Failed;
                }
                
                // Get the directions of the conduits
                XYZ direction1 = line1.Direction;
                XYZ direction2 = line2.Direction;
                
                // Check if the conduits are perpendicular (for trim operation)
                double dotProduct = direction1.DotProduct(direction2);
                bool isPerpendicular = Math.Abs(dotProduct) < 0.001; // Close to zero means perpendicular
                
                if (!isPerpendicular)
                {
                    message = "Conduits must be perpendicular for a trim operation.";
                    trans.RollBack();
                    return Result.Failed;
                }
                
                // Find the intersection point
                XYZ intersectionPoint = FindIntersectionPoint(line1, line2);
                
                if (intersectionPoint == null)
                {
                    message = "Could not find intersection point between conduits.";
                    trans.RollBack();
                    return Result.Failed;
                }
                
                // Trim the conduits to the intersection point
                bool success = TrimConduits(doc, conduit1, conduit2, intersectionPoint);
                
                if (!success)
                {
                    message = "Failed to trim conduits.";
                    trans.RollBack();
                    return Result.Failed;
                }
                
                // Connect the conduits with an elbow fitting
                ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
                ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
                
                Connector connector1 = FindClosestConnector(connectors1, intersectionPoint);
                Connector connector2 = FindClosestConnector(connectors2, intersectionPoint);
                
                if (connector1 == null || connector2 == null)
                {
                    message = "Could not find connectors to connect.";
                    trans.RollBack();
                    return Result.Failed;
                }
                
                // Create an elbow fitting to connect the conduits
                FamilyInstance fitting = doc.Create.NewElbowFitting(connector1, connector2);
                
                if (fitting == null)
                {
                    message = "Failed to create elbow fitting.";
                    trans.RollBack();
                    return Result.Failed;
                }
                
                trans.Commit();
                return Result.Succeeded;
            }
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    /// <summary>
    /// Finds the intersection point between two lines
    /// </summary>
    private XYZ FindIntersectionPoint(Line line1, Line line2)
    {
        // Get the start and end points of the lines
        XYZ p1 = line1.GetEndPoint(0);
        XYZ p2 = line1.GetEndPoint(1);
        XYZ p3 = line2.GetEndPoint(0);
        XYZ p4 = line2.GetEndPoint(1);
        
        // Get the directions
        XYZ v1 = line1.Direction;
        XYZ v2 = line2.Direction;
        
        // For perpendicular lines, we can find the intersection by projecting one line onto the other
        // First, create a plane using the first line
        XYZ normal = v1.CrossProduct(XYZ.BasisZ);
        Plane plane = Plane.CreateByNormalAndOrigin(normal, p1);
        
        // Project the second line onto this plane
        XYZ projectedPoint = plane.ProjectOnto(p3);
        
        // Check if the projected point is on the first line
        IntersectionResult result = line1.Project(projectedPoint);
        
        if (result != null && result.Parameter >= 0 && result.Parameter <= 1)
        {
            return result.XYZPoint;
        }
        
        // If we couldn't find an intersection, try a different approach
        // Find the closest point between the two lines
        XYZ closestPoint1, closestPoint2;
        double param1, param2;
        
        bool success = line1.GetClosestPointTo(line2, out closestPoint1, out closestPoint2, out param1, out param2);
        
        if (success && closestPoint1.DistanceTo(closestPoint2) < 0.001)
        {
            // The lines are close enough to consider them intersecting
            return closestPoint1;
        }
        
        return null;
    }
    
    /// <summary>
    /// Trims the conduits to the intersection point
    /// </summary>
    private bool TrimConduits(Document doc, Conduit conduit1, Conduit conduit2, XYZ intersectionPoint)
    {
        try
        {
            // Get the location curves
            LocationCurve locCurve1 = conduit1.Location as LocationCurve;
            LocationCurve locCurve2 = conduit2.Location as LocationCurve;
            
            // Get the current curves
            Line line1 = locCurve1.Curve as Line;
            Line line2 = locCurve2.Curve as Line;
            
            // Determine which end of each conduit is closer to the intersection point
            XYZ p1Start = line1.GetEndPoint(0);
            XYZ p1End = line1.GetEndPoint(1);
            XYZ p2Start = line2.GetEndPoint(0);
            XYZ p2End = line2.GetEndPoint(1);
            
            bool trimStart1 = p1Start.DistanceTo(intersectionPoint) < p1End.DistanceTo(intersectionPoint);
            bool trimStart2 = p2Start.DistanceTo(intersectionPoint) < p2End.DistanceTo(intersectionPoint);
            
            // Create new curves that end at the intersection point
            Line newLine1, newLine2;
            
            if (trimStart1)
            {
                newLine1 = Line.CreateBound(intersectionPoint, p1End);
            }
            else
            {
                newLine1 = Line.CreateBound(p1Start, intersectionPoint);
            }
            
            if (trimStart2)
            {
                newLine2 = Line.CreateBound(intersectionPoint, p2End);
            }
            else
            {
                newLine2 = Line.CreateBound(p2Start, intersectionPoint);
            }
            
            // Update the conduit locations
            locCurve1.Curve = newLine1;
            locCurve2.Curve = newLine2;
            
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }
    
    /// <summary>
    /// Finds the connector closest to a given point
    /// </summary>
    private Connector FindClosestConnector(ConnectorSet connectors, XYZ point)
    {
        Connector closestConnector = null;
        double minDistance = double.MaxValue;
        
        foreach (Connector connector in connectors)
        {
            double distance = connector.Origin.DistanceTo(point);
            if (distance < minDistance)
            {
                minDistance = distance;
                closestConnector = connector;
            }
        }
        
        return closestConnector;
    }
    
    /// <summary>
    /// Helper class to analyze the relationship between two MEP curves
    /// </summary>
    private class MEPCurvePair
    {
        private MEPCurve _curve1;
        private MEPCurve _curve2;
        private bool _isParallel;
        private double _offset;
        
        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            _curve1 = curve1;
            _curve2 = curve2;
            AnalyzeRelationship();
        }
        
        public bool IsParallel => _isParallel;
        public double Offset => _offset;
        
        private void AnalyzeRelationship()
        {
            // Get the location curves
            LocationCurve locCurve1 = _curve1.Location as LocationCurve;
            LocationCurve locCurve2 = _curve2.Location as LocationCurve;
            
            if (locCurve1 == null || locCurve2 == null)
            {
                _isParallel = false;
                _offset = 0;
                return;
            }
            
            // Get the curves
            Line line1 = locCurve1.Curve as Line;
            Line line2 = locCurve2.Curve as Line;
            
            if (line1 == null || line2 == null)
            {
                _isParallel = false;
                _offset = 0;
                return;
            }
            
            // Get the directions
            XYZ direction1 = line1.Direction;
            XYZ direction2 = line2.Direction;
            
            // Check if the lines are parallel
            double dotProduct = Math.Abs(direction1.DotProduct(direction2));
            _isParallel = Math.Abs(dotProduct - 1.0) < 0.001; // Close to 1 means parallel
            
            if (_isParallel)
            {
                // Calculate the offset between the parallel lines
                XYZ p1 = line1.GetEndPoint(0);
                XYZ p2 = line2.GetEndPoint(0);
                XYZ v = p2 - p1;
                
                // Project v onto the direction to get the component along the direction
                XYZ projectedV = direction1.Multiply(v.DotProduct(direction1));
                
                // The offset is the component perpendicular to the direction
                XYZ offsetVector = v - projectedV;
                _offset = offsetVector.GetLength();
            }
            else
            {
                _offset = 0;
            }
        }
    }
}
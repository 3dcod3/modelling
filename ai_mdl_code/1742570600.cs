using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public class TrimConduitCommand
{
    public Result Execute(UIDocument uidoc, Document doc)
    {
        try
        {
            // Get the conduit data from the input
            ElementId conduitId1 = new ElementId(7638672);
            ElementId conduitId2 = new ElementId(7724387);
            
            // Get the conduit elements
            Conduit conduit1 = doc.GetElement(conduitId1) as Conduit;
            Conduit conduit2 = doc.GetElement(conduitId2) as Conduit;
            
            if (conduit1 == null || conduit2 == null)
            {
                TaskDialog.Show("Error", "One or both conduits could not be found.");
                return Result.Failed;
            }
            
            // Create MEPCurvePair to analyze the relationship
            MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);
            
            // Check if the conduits are parallel
            if (curvePair.IsParallel)
            {
                TaskDialog.Show("Operation Error", "The selected conduits are parallel. Use OffsetBuilder instead of TrimBuilder.");
                return Result.Failed;
            }
            
            // Use TrimBuilder to connect the conduits
            using (Transaction transaction = new Transaction(doc, "Connect Conduits with Trim"))
            {
                try
                {
                    transaction.Start();
                    
                    // Create a TrimBuilder and execute the trim operation
                    TrimBuilder trimBuilder = new TrimBuilder(doc);
                    trimBuilder.Execute(conduit1, conduit2);
                    
                    transaction.Commit();
                    TaskDialog.Show("Success", "Conduits connected successfully using trim operation.");
                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
                    return Result.Failed;
                }
            }
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
            return Result.Failed;
        }
    }
}

public class TrimBuilder
{
    private Document _doc;
    
    public TrimBuilder(Document doc)
    {
        _doc = doc;
    }
    
    public void Execute(Conduit conduit1, Conduit conduit2)
    {
        // Get location curves
        LocationCurve locCurve1 = conduit1.Location as LocationCurve;
        LocationCurve locCurve2 = conduit2.Location as LocationCurve;
        
        if (locCurve1 == null || locCurve2 == null)
        {
            throw new InvalidOperationException("Conduits do not have valid location curves.");
        }
        
        // Get the curves
        Line line1 = locCurve1.Curve as Line;
        Line line2 = locCurve2.Curve as Line;
        
        if (line1 == null || line2 == null)
        {
            throw new InvalidOperationException("Conduits must have linear geometry for trim operation.");
        }
        
        // Calculate the intersection point
        XYZ intersection = CalculateIntersectionPoint(line1, line2);
        
        if (intersection == null)
        {
            throw new InvalidOperationException("Could not find intersection point between conduits.");
        }
        
        // Trim the conduits to the intersection point
        TrimConduitToPoint(conduit1, intersection);
        TrimConduitToPoint(conduit2, intersection);
        
        // Create a fitting at the intersection
        CreateFittingAtIntersection(conduit1, conduit2, intersection);
    }
    
    private XYZ CalculateIntersectionPoint(Line line1, Line line2)
    {
        // Get the direction vectors
        XYZ dir1 = line1.Direction;
        XYZ dir2 = line2.Direction;
        
        // Get the origin points
        XYZ origin1 = line1.Origin;
        XYZ origin2 = line2.Origin;
        
        // Check if lines are parallel (should not happen as we checked with MEPCurvePair)
        if (dir1.CrossProduct(dir2).GetLength() < 0.001)
        {
            return null;
        }
        
        // Calculate the closest points between the two lines
        // This is a simplified approach - in real-world scenarios, you might need more robust calculations
        XYZ v = origin2 - origin1;
        double c1 = v.DotProduct(dir1);
        double c2 = v.DotProduct(dir2);
        double d1 = dir1.DotProduct(dir2);
        
        double denominator = 1 - d1 * d1;
        double t1 = (c1 - c2 * d1) / denominator;
        
        // Calculate point on first line
        XYZ point1 = origin1 + t1 * dir1;
        
        // For simplicity, we'll use this point as the intersection
        // In a more robust implementation, you would calculate the actual intersection
        return point1;
    }
    
    private void TrimConduitToPoint(Conduit conduit, XYZ point)
    {
        LocationCurve locCurve = conduit.Location as LocationCurve;
        Line line = locCurve.Curve as Line;
        
        // Determine which end of the conduit is closer to the intersection point
        XYZ endPoint0 = line.GetEndPoint(0);
        XYZ endPoint1 = line.GetEndPoint(1);
        
        double dist0 = endPoint0.DistanceTo(point);
        double dist1 = endPoint1.DistanceTo(point);
        
        // Create a new line from the farther endpoint to the intersection point
        Line newLine;
        if (dist0 > dist1)
        {
            newLine = Line.CreateBound(endPoint0, point);
        }
        else
        {
            newLine = Line.CreateBound(point, endPoint1);
        }
        
        // Update the conduit's location curve
        locCurve.Curve = newLine;
    }
    
    private void CreateFittingAtIntersection(Conduit conduit1, Conduit conduit2, XYZ intersection)
    {
        // Get the connectors at the intersection point
        Connector connector1 = GetConnectorAtPoint(conduit1, intersection);
        Connector connector2 = GetConnectorAtPoint(conduit2, intersection);
        
        if (connector1 == null || connector2 == null)
        {
            throw new InvalidOperationException("Could not find connectors at the intersection point.");
        }
        
        // Create an elbow fitting to connect the two conduits
        NewElbowFitting(connector1, connector2);
    }
    
    private Connector GetConnectorAtPoint(Conduit conduit, XYZ point)
    {
        ConnectorSet connectors = conduit.ConnectorManager.Connectors;
        
        foreach (Connector connector in connectors)
        {
            if (connector.Origin.DistanceTo(point) < 0.01) // Small tolerance for point comparison
            {
                return connector;
            }
        }
        
        return null;
    }
    
    private void NewElbowFitting(Connector connector1, Connector connector2)
    {
        // Connect the two connectors to create an elbow fitting
        connector1.ConnectTo(connector2);
        
        // The Revit API will automatically create the appropriate fitting
        // based on the connector types and positions
    }
}

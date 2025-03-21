using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// Command to connect two conduits using a trim operation
/// </summary>
public class ConnectConduitsWithTrimCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            // Get the conduit elements by their IDs
            ElementId conduitId1 = new ElementId(7721116);
            ElementId conduitId2 = new ElementId(7721143);

            Conduit conduit1 = doc.GetElement(conduitId1) as Conduit;
            Conduit conduit2 = doc.GetElement(conduitId2) as Conduit;

            if (conduit1 == null || conduit2 == null)
            {
                message = "One or both conduits could not be found.";
                return Result.Failed;
            }

            // Create a MEPCurvePair to analyze the relationship between the conduits
            MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);
            
            // Check if TrimBuilder is appropriate (non-parallel conduits with no offset)
            if (!curvePair.IsParallel && curvePair.Offset.IsAlmostEqualTo(XYZ.Zero))
            {
                // Use TrimBuilder to connect the conduits
                TrimBuilder builder = new TrimBuilder(doc);
                bool success = builder.Connect(conduit1, conduit2);
                
                if (success)
                {
                    return Result.Succeeded;
                }
                else
                {
                    message = "Failed to connect conduits with trim operation.";
                    return Result.Failed;
                }
            }
            else
            {
                message = "Conduits are not suitable for trim operation. Consider using Kick or Offset instead.";
                return Result.Failed;
            }
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}

/// <summary>
/// Builder class for connecting conduits using a trim operation
/// </summary>
public class TrimBuilder
{
    private Document _doc;

    public TrimBuilder(Document doc)
    {
        _doc = doc;
    }

    /// <summary>
    /// Connect two conduits using a trim operation
    /// </summary>
    /// <param name="conduit1">First conduit</param>
    /// <param name="conduit2">Second conduit</param>
    /// <returns>True if connection was successful</returns>
    public bool Connect(Conduit conduit1, Conduit conduit2)
    {
        using (Transaction t = new Transaction(_doc, "Connect Conduits with Trim"))
        {
            try
            {
                t.Start();

                // Get the location curves of the conduits
                LocationCurve locCurve1 = conduit1.Location as LocationCurve;
                LocationCurve locCurve2 = conduit2.Location as LocationCurve;

                if (locCurve1 == null || locCurve2 == null)
                {
                    return false;
                }

                Line line1 = locCurve1.Curve as Line;
                Line line2 = locCurve2.Curve as Line;

                if (line1 == null || line2 == null)
                {
                    return false;
                }

                // Calculate intersection point
                XYZ intersection = CalculateIntersectionPoint(line1, line2);
                
                if (intersection == null)
                {
                    return false;
                }

                // Trim the first conduit to the intersection point
                locCurve1.Curve = Line.CreateBound(line1.GetEndPoint(0), intersection);
                
                // Trim the second conduit to the intersection point
                locCurve2.Curve = Line.CreateBound(line2.GetEndPoint(0), intersection);

                // Get the connectors at the trimmed ends
                Connector connector1 = GetConnectorClosestTo(conduit1, intersection);
                Connector connector2 = GetConnectorClosestTo(conduit2, intersection);

                if (connector1 == null || connector2 == null)
                {
                    return false;
                }

                // Create an elbow fitting to connect the conduits
                _doc.Create.NewElbowFitting(connector1, connector2);

                t.Commit();
                return true;
            }
            catch (Exception ex)
            {
                if (t.HasStarted())
                {
                    t.RollBack();
                }
                TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// Calculate the intersection point of two lines in 3D space
    /// </summary>
    private XYZ CalculateIntersectionPoint(Line line1, Line line2)
    {
        // Get the direction vectors of the lines
        XYZ dir1 = line1.Direction;
        XYZ dir2 = line2.Direction;
        
        // Get the start points of the lines
        XYZ start1 = line1.GetEndPoint(0);
        XYZ start2 = line2.GetEndPoint(0);
        
        // Calculate the closest points between the two lines
        XYZ closestPoint1, closestPoint2;
        double param1, param2;
        
        // Check if lines are parallel
        if (dir1.CrossProduct(dir2).IsAlmostEqualTo(XYZ.Zero))
        {
            return null;
        }
        
        // Find the parameters for the closest points
        bool success = GetClosestPointsBetweenLines(
            start1, dir1, start2, dir2, 
            out param1, out param2, 
            out closestPoint1, out closestPoint2);
            
        if (!success)
        {
            return null;
        }
        
        // If the closest points are the same, we have an intersection
        if (closestPoint1.DistanceTo(closestPoint2) < 0.001)
        {
            return closestPoint1;
        }
        
        // Otherwise, use the midpoint between the closest points
        return (closestPoint1 + closestPoint2) / 2;
    }

    /// <summary>
    /// Calculate the closest points between two lines
    /// </summary>
    private bool GetClosestPointsBetweenLines(
        XYZ start1, XYZ dir1, 
        XYZ start2, XYZ dir2, 
        out double param1, out double param2, 
        out XYZ closestPoint1, out XYZ closestPoint2)
    {
        param1 = param2 = 0;
        closestPoint1 = closestPoint2 = XYZ.Zero;
        
        // Calculate coefficients for the system of equations
        double a = dir1.DotProduct(dir1);
        double b = dir1.DotProduct(dir2);
        double c = dir2.DotProduct(dir2);
        double d = dir1.DotProduct(start1 - start2);
        double e = dir2.DotProduct(start1 - start2);
        
        // Calculate determinant
        double det = a * c - b * b;
        
        // If determinant is zero, lines are parallel
        if (Math.Abs(det) < 1e-10)
        {
            return false;
        }
        
        // Calculate parameters
        param1 = (b * e - c * d) / det;
        param2 = (a * e - b * d) / det;
        
        // Calculate closest points
        closestPoint1 = start1 + param1 * dir1;
        closestPoint2 = start2 + param2 * dir2;
        
        return true;
    }

    /// <summary>
    /// Get the connector closest to a specific point
    /// </summary>
    private Connector GetConnectorClosestTo(Conduit conduit, XYZ point)
    {
        ConnectorSet connectors = conduit.ConnectorManager.Connectors;
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
}

/// <summary>
/// Helper class to analyze the relationship between two MEP curves
/// </summary>
public class MEPCurvePair
{
    private MEPCurve _curve1;
    private MEPCurve _curve2;
    private XYZ _offset;
    private bool _isParallel;

    public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
    {
        _curve1 = curve1;
        _curve2 = curve2;
        AnalyzeCurves();
    }

    public MEPCurve Curve1 => _curve1;
    public MEPCurve Curve2 => _curve2;
    public XYZ Offset => _offset;
    public bool IsParallel => _isParallel;

    private void AnalyzeCurves()
    {
        LocationCurve locCurve1 = _curve1.Location as LocationCurve;
        LocationCurve locCurve2 = _curve2.Location as LocationCurve;

        if (locCurve1 == null || locCurve2 == null)
        {
            _isParallel = false;
            _offset = XYZ.Zero;
            return;
        }

        Line line1 = locCurve1.Curve as Line;
        Line line2 = locCurve2.Curve as Line;

        if (line1 == null || line2 == null)
        {
            _isParallel = false;
            _offset = XYZ.Zero;
            return;
        }

        // Check if lines are parallel
        XYZ dir1 = line1.Direction;
        XYZ dir2 = line2.Direction;
        
        // Lines are parallel if their cross product is almost zero
        _isParallel = dir1.CrossProduct(dir2).GetLength() < 0.001;
        
        if (_isParallel)
        {
            // Calculate offset between parallel lines
            XYZ start1 = line1.GetEndPoint(0);
            XYZ start2 = line2.GetEndPoint(0);
            
            // Project start2 onto line1 to find the offset vector
            XYZ projectedPoint = start1 + dir1.DotProduct(start2 - start1) * dir1;
            _offset = start2 - projectedPoint;
        }
        else
        {
            _offset = XYZ.Zero;
        }
    }
}
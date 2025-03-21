using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ConduitConnector : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            // Get the two conduits by their element IDs
            ElementId conduitId1 = new ElementId(7638672);
            ElementId conduitId2 = new ElementId(7724387);

            Conduit conduit1 = doc.GetElement(conduitId1) as Conduit;
            Conduit conduit2 = doc.GetElement(conduitId2) as Conduit;

            if (conduit1 == null || conduit2 == null)
            {
                message = "Failed to find the specified conduits.";
                return Result.Failed;
            }

            // Get the free ends of the conduits
            XYZ freeEnd1 = new XYZ(1710.933383058, -259.456787969, 129.812500000);
            XYZ freeEnd2 = new XYZ(1720.646628543, -261.009697791, 130.121394041);

            // Get the directions of the conduits
            XYZ direction1 = new XYZ(-1.000000000, 0.000000000, 0.000000000);
            XYZ direction2 = new XYZ(0.000000000, 0.980785151, -0.195090973);

            // Analyze the conduit relationship
            bool isParallel = AreVectorsParallel(direction1, direction2);
            double angle = GetAngleBetweenVectors(direction1, direction2);

            using (Transaction trans = new Transaction(doc, "Connect Conduits"))
            {
                trans.Start();

                // Find the connectors at the free ends
                Connector connector1 = FindConnectorAtPoint(conduit1, freeEnd1);
                Connector connector2 = FindConnectorAtPoint(conduit2, freeEnd2);

                if (connector1 == null || connector2 == null)
                {
                    message = "Failed to find connectors at the specified points.";
                    trans.RollBack();
                    return Result.Failed;
                }

                // Determine the appropriate connection method based on geometry
                if (isParallel)
                {
                    // Use OffsetBuilder for parallel conduits
                    // Calculate the offset distance between the conduits
                    double offset = CalculateOffset(freeEnd1, freeEnd2, direction1);
                    
                    // Create a new conduit to connect the two existing conduits
                    XYZ midPoint = (freeEnd1 + freeEnd2) / 2;
                    Conduit connectingConduit = CreateConduit(doc, conduit1, freeEnd1, midPoint);
                    
                    // Create fittings at both ends
                    Connector newConnector1 = FindConnectorAtPoint(connectingConduit, freeEnd1);
                    Connector newConnector2 = FindConnectorAtPoint(connectingConduit, midPoint);
                    
                    doc.Create.NewElbowFitting(connector1, newConnector1);
                    doc.Create.NewElbowFitting(connector2, newConnector2);
                }
                else if (angle < Math.PI / 4) // Less than 45 degrees
                {
                    // Use TrimBuilder for non-parallel conduits with small angle
                    // Extend both conduits to find intersection point
                    XYZ intersectionPoint = FindIntersectionPoint(freeEnd1, direction1, freeEnd2, direction2);
                    
                    if (intersectionPoint != null)
                    {
                        // Create new conduits that extend to the intersection point
                        Conduit extendedConduit1 = CreateConduit(doc, conduit1, freeEnd1, intersectionPoint);
                        Conduit extendedConduit2 = CreateConduit(doc, conduit2, freeEnd2, intersectionPoint);
                        
                        // Create a fitting at the intersection
                        Connector extConnector1 = FindConnectorAtPoint(extendedConduit1, intersectionPoint);
                        Connector extConnector2 = FindConnectorAtPoint(extendedConduit2, intersectionPoint);
                        
                        doc.Create.NewElbowFitting(connector1, extConnector1);
                        doc.Create.NewElbowFitting(connector2, extConnector2);
                    }
                    else
                    {
                        message = "Could not find intersection point between conduits.";
                        trans.RollBack();
                        return Result.Failed;
                    }
                }
                else
                {
                    // Use KickBuilder for non-parallel conduits with large angle
                    // Create an intermediate point for the kick
                    XYZ kickPoint = CalculateKickPoint(freeEnd1, direction1, freeEnd2, direction2);
                    
                    // Create two new conduits to form the kick
                    Conduit kickConduit1 = CreateConduit(doc, conduit1, freeEnd1, kickPoint);
                    Conduit kickConduit2 = CreateConduit(doc, conduit2, freeEnd2, kickPoint);
                    
                    // Create fittings
                    Connector kickConnector1 = FindConnectorAtPoint(kickConduit1, kickPoint);
                    Connector kickConnector2 = FindConnectorAtPoint(kickConduit2, kickPoint);
                    
                    doc.Create.NewElbowFitting(connector1, FindConnectorAtPoint(kickConduit1, freeEnd1));
                    doc.Create.NewElbowFitting(connector2, FindConnectorAtPoint(kickConduit2, freeEnd2));
                    doc.Create.NewElbowFitting(kickConnector1, kickConnector2);
                }

                trans.Commit();
                return Result.Succeeded;
            }
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            return Result.Failed;
        }
    }

    private Conduit CreateConduit(Document doc, Conduit templateConduit, XYZ startPoint, XYZ endPoint)
    {
        // Get the conduit type and level from the template conduit
        ElementId conduitTypeId = templateConduit.GetTypeId();
        ElementId levelId = templateConduit.LevelId;
        
        // Create a new conduit with the same properties
        return Conduit.Create(doc, conduitTypeId, startPoint, endPoint, levelId);
    }

    private Connector FindConnectorAtPoint(Conduit conduit, XYZ point)
    {
        ConnectorSet connectors = conduit.ConnectorManager.Connectors;
        
        foreach (Connector connector in connectors)
        {
            if (connector.Origin.IsAlmostEqualTo(point, 0.001))
            {
                return connector;
            }
        }
        
        return null;
    }

    private bool AreVectorsParallel(XYZ vector1, XYZ vector2)
    {
        // Normalize the vectors
        XYZ v1 = vector1.Normalize();
        XYZ v2 = vector2.Normalize();
        
        // Check if they are parallel (dot product close to 1 or -1)
        double dotProduct = Math.Abs(v1.DotProduct(v2));
        return dotProduct > 0.99; // Allow for small numerical errors
    }

    private double GetAngleBetweenVectors(XYZ vector1, XYZ vector2)
    {
        // Normalize the vectors
        XYZ v1 = vector1.Normalize();
        XYZ v2 = vector2.Normalize();
        
        // Calculate the angle between them
        double dotProduct = v1.DotProduct(v2);
        return Math.Acos(Math.Min(Math.Max(dotProduct, -1.0), 1.0));
    }

    private double CalculateOffset(XYZ point1, XYZ point2, XYZ direction)
    {
        // Calculate the vector between the points
        XYZ pointVector = point2 - point1;
        
        // Project this vector onto the normal of the direction
        XYZ normalVector = direction.CrossProduct(XYZ.BasisZ).Normalize();
        double offset = Math.Abs(pointVector.DotProduct(normalVector));
        
        return offset;
    }

    private XYZ FindIntersectionPoint(XYZ point1, XYZ direction1, XYZ point2, XYZ direction2)
    {
        // This is a simplified 3D line intersection calculation
        // In reality, 3D lines might not intersect exactly, so we find the closest point
        
        // Create a plane containing line 1 and perpendicular to the XY plane
        XYZ normal = direction1.CrossProduct(XYZ.BasisZ).Normalize();
        Plane plane = Plane.CreateByNormalAndOrigin(normal, point1);
        
        // Find where line 2 intersects this plane
        double t;
        if (plane.Project(point2, out XYZ projectedPoint, out t))
        {
            // Check if the projected point is on line 1
            XYZ vectorToPoint = projectedPoint - point1;
            double dotProduct = vectorToPoint.DotProduct(direction1);
            
            if (dotProduct > 0) // Point is in the direction of the line
            {
                return point1 + direction1.Normalize() * dotProduct;
            }
        }
        
        return null;
    }

    private XYZ CalculateKickPoint(XYZ point1, XYZ direction1, XYZ point2, XYZ direction2)
    {
        // Calculate a point that's halfway between the two conduits
        // and offset in a direction that allows for a smooth connection
        
        // Find the midpoint between the two endpoints
        XYZ midPoint = (point1 + point2) / 2;
        
        // Calculate a direction that's perpendicular to both conduit directions
        XYZ perpDirection = direction1.CrossProduct(direction2).Normalize();
        
        // If the cross product is too small, use a default direction
        if (perpDirection.GetLength() < 0.001)
        {
            perpDirection = XYZ.BasisZ;
        }
        
        // Offset the midpoint in the perpendicular direction
        double offsetDistance = point1.DistanceTo(point2) / 2;
        return midPoint + perpDirection * offsetDistance;
    }
}
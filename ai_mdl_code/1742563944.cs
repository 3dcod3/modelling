using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public class ConduitConnector
{
    public void ConnectConduits(Document doc, XYZ point1, XYZ direction1, XYZ point2, XYZ direction2, string conduitType)
    {
        try
        {
            using (Transaction transaction = new Transaction(doc, "Connect Conduits"))
            {
                transaction.Start();

                // Get the conduit type
                ConduitType type = GetConduitType(doc, conduitType);
                if (type == null)
                {
                    TaskDialog.Show("Error", "Conduit type not found.");
                    return;
                }

                // Get the level
                Level level = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault();

                if (level == null)
                {
                    TaskDialog.Show("Error", "No level found in the document.");
                    return;
                }

                // Create the conduits
                Conduit conduit1 = CreateConduit(doc, point1, direction1, type.Id, level.Id);
                Conduit conduit2 = CreateConduit(doc, point2, direction2, type.Id, level.Id);

                // Analyze the relationship between the conduits
                bool isParallel = AreVectorsParallel(direction1, direction2);
                double offset = CalculateOffset(point1, direction1, point2, direction2);

                // Select the appropriate builder based on the relationship
                if (isParallel && offset > 0.001)
                {
                    // Use OffsetBuilder for parallel conduits with offset
                    CreateOffsetConnection(doc, conduit1, conduit2);
                }
                else if (!isParallel && offset > 0.001)
                {
                    // Use KickBuilder for non-parallel conduits with offset
                    CreateKickConnection(doc, conduit1, conduit2);
                }
                else
                {
                    // Use TrimBuilder for non-parallel conduits with no offset
                    CreateTrimConnection(doc, conduit1, conduit2);
                }

                transaction.Commit();
                TaskDialog.Show("Success", "Conduits connected successfully.");
            }
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
        }
    }

    private Conduit CreateConduit(Document doc, XYZ startPoint, XYZ direction, ElementId typeId, ElementId levelId)
    {
        // Normalize the direction vector
        XYZ normalizedDirection = direction.Normalize();
        
        // Create a point 10 feet away in the direction
        XYZ endPoint = startPoint.Add(normalizedDirection.Multiply(10.0));
        
        // Create the conduit
        return Conduit.Create(doc, typeId, startPoint, endPoint, levelId);
    }

    private ConduitType GetConduitType(Document doc, string typeName)
    {
        // Find the conduit type by name
        return new FilteredElementCollector(doc)
            .OfClass(typeof(ConduitType))
            .Cast<ConduitType>()
            .FirstOrDefault(t => t.Name.Contains(typeName));
    }

    private bool AreVectorsParallel(XYZ v1, XYZ v2)
    {
        // Normalize the vectors
        XYZ norm1 = v1.Normalize();
        XYZ norm2 = v2.Normalize();
        
        // Check if they are parallel (or anti-parallel)
        double dotProduct = Math.Abs(norm1.DotProduct(norm2));
        return Math.Abs(dotProduct - 1.0) < 0.001;
    }

    private double CalculateOffset(XYZ point1, XYZ direction1, XYZ point2, XYZ direction2)
    {
        // For parallel lines, calculate the perpendicular distance
        if (AreVectorsParallel(direction1, direction2))
        {
            XYZ v = point2.Subtract(point1);
            XYZ normDir = direction1.Normalize();
            XYZ projection = normDir.Multiply(v.DotProduct(normDir));
            return v.Subtract(projection).GetLength();
        }
        
        // For non-parallel lines, find the closest points
        XYZ v1 = direction1.Normalize();
        XYZ v2 = direction2.Normalize();
        XYZ w0 = point1.Subtract(point2);
        
        double a = v1.DotProduct(v1);
        double b = v1.DotProduct(v2);
        double c = v2.DotProduct(v2);
        double d = v1.DotProduct(w0);
        double e = v2.DotProduct(w0);
        
        double denominator = a * c - b * b;
        
        // If lines are parallel (or nearly so), return a large value
        if (Math.Abs(denominator) < 0.001)
            return 1000.0;
            
        double sc = (b * e - c * d) / denominator;
        double tc = (a * e - b * d) / denominator;
        
        XYZ closestPoint1 = point1.Add(v1.Multiply(sc));
        XYZ closestPoint2 = point2.Add(v2.Multiply(tc));
        
        return closestPoint1.DistanceTo(closestPoint2);
    }

    private void CreateOffsetConnection(Document doc, Conduit conduit1, Conduit conduit2)
    {
        // Get the connectors
        ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
        ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
        
        // Find the closest connectors
        Connector connector1 = FindClosestConnector(connectors1, connectors2);
        Connector connector2 = FindClosestConnector(connectors2, connectors1);
        
        if (connector1 != null && connector2 != null)
        {
            // Create an elbow fitting
            doc.Create.NewElbowFitting(connector1, connector2);
        }
    }

    private void CreateKickConnection(Document doc, Conduit conduit1, Conduit conduit2)
    {
        // Get the connectors
        ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
        ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
        
        // Find the closest connectors
        Connector connector1 = FindClosestConnector(connectors1, connectors2);
        Connector connector2 = FindClosestConnector(connectors2, connectors1);
        
        if (connector1 != null && connector2 != null)
        {
            // Create an elbow fitting
            doc.Create.NewElbowFitting(connector1, connector2);
        }
    }

    private void CreateTrimConnection(Document doc, Conduit conduit1, Conduit conduit2)
    {
        // Get the connectors
        ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
        ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
        
        // Find the closest connectors
        Connector connector1 = FindClosestConnector(connectors1, connectors2);
        Connector connector2 = FindClosestConnector(connectors2, connectors1);
        
        if (connector1 != null && connector2 != null)
        {
            // Create an elbow fitting
            doc.Create.NewElbowFitting(connector1, connector2);
        }
    }

    private Connector FindClosestConnector(ConnectorSet fromConnectors, ConnectorSet toConnectors)
    {
        Connector closestConnector = null;
        double minDistance = double.MaxValue;
        
        foreach (Connector fromConnector in fromConnectors)
        {
            foreach (Connector toConnector in toConnectors)
            {
                double distance = fromConnector.Origin.DistanceTo(toConnector.Origin);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    closestConnector = fromConnector;
                }
            }
        }
        
        return closestConnector;
    }

    // Main method to execute the connection
    public void Execute(Document doc)
    {
        // Parse the conduit data
        XYZ point1 = new XYZ(1710.9333830575522, -261.29012130235617, 129.8125000000008);
        XYZ direction1 = new XYZ(-1.0, 0.0, 0.0);
        
        XYZ point2 = new XYZ(1718.8132952099709, -262.84303096586518, 130.12139483877701);
        XYZ direction2 = new XYZ(0.0, 0.98078505063771826, -0.19509147712180691);
        
        // Connect the conduits
        ConnectConduits(doc, point1, direction1, point2, direction2, "EMT");
    }
}
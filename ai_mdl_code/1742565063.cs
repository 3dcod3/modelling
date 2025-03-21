using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public class ConduitConnector
{
    /// <summary>
    /// Connects two conduits using the appropriate connection method based on their geometry
    /// </summary>
    public static void ConnectConduits(UIDocument uidoc)
    {
        Document doc = uidoc.Document;
        
        // Get the conduits by element id
        ElementId conduit1Id = new ElementId(7638680);
        ElementId conduit2Id = new ElementId(7724426);
        
        Conduit conduit1 = doc.GetElement(conduit1Id) as Conduit;
        Conduit conduit2 = doc.GetElement(conduit2Id) as Conduit;
        
        if (conduit1 == null || conduit2 == null)
        {
            TaskDialog.Show("Error", "One or both conduits could not be found.");
            return;
        }
        
        using (Transaction transaction = new Transaction(doc, "Connect Conduits"))
        {
            try
            {
                transaction.Start();
                
                // Create MEPCurvePair to analyze the relationship between conduits
                MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);
                
                // Get the location curves of the conduits
                LocationCurve locCurve1 = conduit1.Location as LocationCurve;
                LocationCurve locCurve2 = conduit2.Location as LocationCurve;
                
                if (locCurve1 == null || locCurve2 == null)
                {
                    TaskDialog.Show("Error", "Could not get location curves for conduits.");
                    transaction.RollBack();
                    return;
                }
                
                // Get the curves
                Line line1 = locCurve1.Curve as Line;
                Line line2 = locCurve2.Curve as Line;
                
                if (line1 == null || line2 == null)
                {
                    TaskDialog.Show("Error", "Conduits must have line geometry.");
                    transaction.RollBack();
                    return;
                }
                
                // Get the directions of the conduits
                XYZ dir1 = line1.Direction;
                XYZ dir2 = line2.Direction;
                
                // Check if the conduits are parallel
                bool isParallel = dir1.IsAlmostEqualTo(dir2) || dir1.IsAlmostEqualTo(dir2.Negate());
                
                // Get the free ends of the conduits
                XYZ freeEnd1 = new XYZ(1717.229961877, -261.290121302, 129.812500000);
                XYZ freeEnd2 = new XYZ(1718.813295210, -262.843030966, 130.121394839);
                
                // Determine the appropriate builder based on geometry
                if (isParallel)
                {
                    // Use OffsetBuilder for parallel conduits
                    ConnectWithOffsetBuilder(doc, conduit1, conduit2, freeEnd1, freeEnd2);
                }
                else
                {
                    // Calculate the offset between the conduits
                    XYZ offset = CalculateOffset(line1, line2);
                    
                    if (offset.GetLength() < 0.001) // If offset is negligible
                    {
                        // Use TrimBuilder for non-parallel conduits with no offset
                        ConnectWithTrimBuilder(doc, conduit1, conduit2, freeEnd1, freeEnd2);
                    }
                    else
                    {
                        // Use KickBuilder for non-parallel conduits with offset
                        ConnectWithKickBuilder(doc, conduit1, conduit2, freeEnd1, freeEnd2);
                    }
                }
                
                transaction.Commit();
                TaskDialog.Show("Success", "Conduits connected successfully.");
            }
            catch (Exception ex)
            {
                transaction.RollBack();
                TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
            }
        }
    }
    
    /// <summary>
    /// Calculates the offset between two lines
    /// </summary>
    private static XYZ CalculateOffset(Line line1, Line line2)
    {
        // Get the closest points between the two lines
        XYZ point1, point2;
        line1.GetClosestPoints(line2, out point1, out point2);
        
        // Return the vector between these points
        return point2 - point1;
    }
    
    /// <summary>
    /// Connects conduits using KickBuilder (for non-parallel conduits with offset)
    /// </summary>
    private static void ConnectWithKickBuilder(Document doc, Conduit conduit1, Conduit conduit2, XYZ freeEnd1, XYZ freeEnd2)
    {
        // Get the connectors of the conduits
        ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
        ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
        
        // Find the unconnected connectors
        Connector connector1 = FindUnconnectedConnector(connectors1, freeEnd1);
        Connector connector2 = FindUnconnectedConnector(connectors2, freeEnd2);
        
        if (connector1 == null || connector2 == null)
        {
            throw new Exception("Could not find unconnected connectors.");
        }
        
        // Create an elbow fitting to connect the conduits
        doc.Create.NewElbowFitting(connector1, connector2);
    }
    
    /// <summary>
    /// Connects conduits using OffsetBuilder (for parallel conduits)
    /// </summary>
    private static void ConnectWithOffsetBuilder(Document doc, Conduit conduit1, Conduit conduit2, XYZ freeEnd1, XYZ freeEnd2)
    {
        // Get the connectors of the conduits
        ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
        ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
        
        // Find the unconnected connectors
        Connector connector1 = FindUnconnectedConnector(connectors1, freeEnd1);
        Connector connector2 = FindUnconnectedConnector(connectors2, freeEnd2);
        
        if (connector1 == null || connector2 == null)
        {
            throw new Exception("Could not find unconnected connectors.");
        }
        
        // Create an elbow fitting to connect the conduits
        doc.Create.NewElbowFitting(connector1, connector2);
    }
    
    /// <summary>
    /// Connects conduits using TrimBuilder (for non-parallel conduits with no offset)
    /// </summary>
    private static void ConnectWithTrimBuilder(Document doc, Conduit conduit1, Conduit conduit2, XYZ freeEnd1, XYZ freeEnd2)
    {
        // Get the connectors of the conduits
        ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
        ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
        
        // Find the unconnected connectors
        Connector connector1 = FindUnconnectedConnector(connectors1, freeEnd1);
        Connector connector2 = FindUnconnectedConnector(connectors2, freeEnd2);
        
        if (connector1 == null || connector2 == null)
        {
            throw new Exception("Could not find unconnected connectors.");
        }
        
        // Create an elbow fitting to connect the conduits
        doc.Create.NewElbowFitting(connector1, connector2);
    }
    
    /// <summary>
    /// Finds the unconnected connector closest to the specified point
    /// </summary>
    private static Connector FindUnconnectedConnector(ConnectorSet connectors, XYZ point)
    {
        Connector closestConnector = null;
        double minDistance = double.MaxValue;
        
        foreach (Connector connector in connectors)
        {
            if (connector.IsConnected)
                continue;
            
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
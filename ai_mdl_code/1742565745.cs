using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConduitConnector
{
    public class ConduitConnectionHandler
    {
        public Result Execute(UIDocument uidoc)
        {
            Document doc = uidoc.Document;
            
            // Get the conduit elements by ID
            ElementId conduit1Id = new ElementId(7638672);
            ElementId conduit2Id = new ElementId(7724387);
            
            Conduit conduit1 = doc.GetElement(conduit1Id) as Conduit;
            Conduit conduit2 = doc.GetElement(conduit2Id) as Conduit;
            
            if (conduit1 == null || conduit2 == null)
            {
                TaskDialog.Show("Error", "One or both conduits could not be found.");
                return Result.Failed;
            }
            
            try
            {
                // Start a transaction
                using (Transaction transaction = new Transaction(doc, "Connect Conduits"))
                {
                    transaction.Start();
                    
                    // Get the conduit locations
                    LocationCurve location1 = conduit1.Location as LocationCurve;
                    LocationCurve location2 = conduit2.Location as LocationCurve;
                    
                    if (location1 == null || location2 == null)
                    {
                        TaskDialog.Show("Error", "Could not get conduit locations.");
                        return Result.Failed;
                    }
                    
                    // Get the conduit curves
                    Line line1 = location1.Curve as Line;
                    Line line2 = location2.Curve as Line;
                    
                    if (line1 == null || line2 == null)
                    {
                        TaskDialog.Show("Error", "Conduits must be straight lines.");
                        return Result.Failed;
                    }
                    
                    // Get the free ends of the conduits
                    XYZ freeEnd1 = new XYZ(1710.933383058, -259.456787969, 129.812500000);
                    XYZ freeEnd2 = new XYZ(1720.646628543, -261.009697791, 130.121394041);
                    
                    // Get the directions of the conduits
                    XYZ direction1 = new XYZ(-1.000000000, 0.000000000, 0.000000000);
                    XYZ direction2 = new XYZ(0.000000000, 0.980785151, -0.195090973);
                    
                    // Analyze the relationship between the conduits
                    bool isParallel = AreVectorsParallel(direction1, direction2);
                    double offset = CalculateOffset(freeEnd1, direction1, freeEnd2, direction2);
                    
                    // Create the appropriate connection based on the geometry
                    if (!isParallel)
                    {
                        // Use KickBuilder for non-parallel conduits with offset
                        ConnectWithKick(doc, conduit1, conduit2, freeEnd1, freeEnd2, direction1, direction2);
                    }
                    else if (offset > 0.001)
                    {
                        // Use OffsetBuilder for parallel conduits with offset
                        ConnectWithOffset(doc, conduit1, conduit2, freeEnd1, freeEnd2, direction1);
                    }
                    else
                    {
                        // Use TrimBuilder for aligned conduits
                        ConnectWithTrim(doc, conduit1, conduit2, freeEnd1, freeEnd2, direction1);
                    }
                    
                    transaction.Commit();
                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
                return Result.Failed;
            }
        }
        
        private bool AreVectorsParallel(XYZ v1, XYZ v2)
        {
            // Normalize vectors
            XYZ norm1 = v1.Normalize();
            XYZ norm2 = v2.Normalize();
            
            // Check if vectors are parallel (dot product close to 1 or -1)
            double dotProduct = Math.Abs(norm1.DotProduct(norm2));
            return Math.Abs(dotProduct - 1.0) < 0.001;
        }
        
        private double CalculateOffset(XYZ point1, XYZ direction1, XYZ point2, XYZ direction2)
        {
            // For non-parallel lines, calculate the minimum distance between them
            if (!AreVectorsParallel(direction1, direction2))
            {
                XYZ normal = direction1.CrossProduct(direction2).Normalize();
                return Math.Abs((point2 - point1).DotProduct(normal));
            }
            
            // For parallel lines, calculate the perpendicular distance
            XYZ perpendicular = direction1.CrossProduct(XYZ.BasisZ).Normalize();
            if (perpendicular.IsZeroLength())
            {
                perpendicular = direction1.CrossProduct(XYZ.BasisX).Normalize();
            }
            
            return Math.Abs((point2 - point1).DotProduct(perpendicular));
        }
        
        private void ConnectWithKick(Document doc, Conduit conduit1, Conduit conduit2, 
                                    XYZ freeEnd1, XYZ freeEnd2, XYZ direction1, XYZ direction2)
        {
            // Calculate the midpoint between the free ends
            XYZ midPoint = (freeEnd1 + freeEnd2) / 2.0;
            
            // Create a new conduit from the first conduit to the midpoint
            XYZ kickPoint1 = freeEnd1 + direction1.Normalize() * 5.0; // 5 units back from free end
            Line kickLine1 = Line.CreateBound(kickPoint1, midPoint);
            Conduit kickConduit1 = Conduit.Create(doc, conduit1.GetTypeId(), kickPoint1, midPoint, conduit1.ReferenceLevel.Id);
            
            // Create a new conduit from the second conduit to the midpoint
            XYZ kickPoint2 = freeEnd2 + direction2.Normalize() * 5.0; // 5 units back from free end
            Line kickLine2 = Line.CreateBound(kickPoint2, midPoint);
            Conduit kickConduit2 = Conduit.Create(doc, conduit2.GetTypeId(), kickPoint2, midPoint, conduit2.ReferenceLevel.Id);
            
            // Trim the original conduits
            LocationCurve location1 = conduit1.Location as LocationCurve;
            LocationCurve location2 = conduit2.Location as LocationCurve;
            
            Line line1 = location1.Curve as Line;
            Line line2 = location2.Curve as Line;
            
            Line trimmedLine1 = Line.CreateBound(line1.GetEndPoint(0), kickPoint1);
            Line trimmedLine2 = Line.CreateBound(line2.GetEndPoint(0), kickPoint2);
            
            location1.Curve = trimmedLine1;
            location2.Curve = trimmedLine2;
            
            // Create elbow fittings at the kick points
            ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
            ConnectorSet connectors2 = kickConduit1.ConnectorManager.Connectors;
            
            Connector connector1 = FindConnectorClosestTo(connectors1, kickPoint1);
            Connector connector2 = FindConnectorClosestTo(connectors2, kickPoint1);
            
            if (connector1 != null && connector2 != null)
            {
                doc.Create.NewElbowFitting(connector1, connector2);
            }
            
            ConnectorSet connectors3 = conduit2.ConnectorManager.Connectors;
            ConnectorSet connectors4 = kickConduit2.ConnectorManager.Connectors;
            
            Connector connector3 = FindConnectorClosestTo(connectors3, kickPoint2);
            Connector connector4 = FindConnectorClosestTo(connectors4, kickPoint2);
            
            if (connector3 != null && connector4 != null)
            {
                doc.Create.NewElbowFitting(connector3, connector4);
            }
            
            // Create a tee fitting at the midpoint
            ConnectorSet connectors5 = kickConduit1.ConnectorManager.Connectors;
            ConnectorSet connectors6 = kickConduit2.ConnectorManager.Connectors;
            
            Connector connector5 = FindConnectorClosestTo(connectors5, midPoint);
            Connector connector6 = FindConnectorClosestTo(connectors6, midPoint);
            
            if (connector5 != null && connector6 != null)
            {
                doc.Create.NewElbowFitting(connector5, connector6);
            }
        }
        
        private void ConnectWithOffset(Document doc, Conduit conduit1, Conduit conduit2, 
                                      XYZ freeEnd1, XYZ freeEnd2, XYZ direction)
        {
            // Calculate the projection of freeEnd2 onto the line of conduit1
            XYZ projectedPoint = ProjectPointOnLine(freeEnd2, freeEnd1, direction);
            
            // Calculate the midpoint of the offset
            XYZ midPoint = (projectedPoint + freeEnd2) / 2.0;
            
            // Create offset conduits
            Conduit offsetConduit1 = Conduit.Create(doc, conduit1.GetTypeId(), projectedPoint, midPoint, conduit1.ReferenceLevel.Id);
            Conduit offsetConduit2 = Conduit.Create(doc, conduit2.GetTypeId(), freeEnd2, midPoint, conduit2.ReferenceLevel.Id);
            
            // Extend conduit1 to the projected point
            LocationCurve location1 = conduit1.Location as LocationCurve;
            Line line1 = location1.Curve as Line;
            Line extendedLine1 = Line.CreateBound(line1.GetEndPoint(0), projectedPoint);
            location1.Curve = extendedLine1;
            
            // Create elbow fittings
            ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
            ConnectorSet connectors2 = offsetConduit1.ConnectorManager.Connectors;
            
            Connector connector1 = FindConnectorClosestTo(connectors1, projectedPoint);
            Connector connector2 = FindConnectorClosestTo(connectors2, projectedPoint);
            
            if (connector1 != null && connector2 != null)
            {
                doc.Create.NewElbowFitting(connector1, connector2);
            }
            
            ConnectorSet connectors3 = conduit2.ConnectorManager.Connectors;
            ConnectorSet connectors4 = offsetConduit2.ConnectorManager.Connectors;
            
            Connector connector3 = FindConnectorClosestTo(connectors3, freeEnd2);
            Connector connector4 = FindConnectorClosestTo(connectors4, freeEnd2);
            
            if (connector3 != null && connector4 != null)
            {
                doc.Create.NewElbowFitting(connector3, connector4);
            }
            
            // Create elbow at the midpoint
            ConnectorSet connectors5 = offsetConduit1.ConnectorManager.Connectors;
            ConnectorSet connectors6 = offsetConduit2.ConnectorManager.Connectors;
            
            Connector connector5 = FindConnectorClosestTo(connectors5, midPoint);
            Connector connector6 = FindConnectorClosestTo(connectors6, midPoint);
            
            if (connector5 != null && connector6 != null)
            {
                doc.Create.NewElbowFitting(connector5, connector6);
            }
        }
        
        private void ConnectWithTrim(Document doc, Conduit conduit1, Conduit conduit2, 
                                    XYZ freeEnd1, XYZ freeEnd2, XYZ direction)
        {
            // For aligned conduits, simply extend one to meet the other
            LocationCurve location1 = conduit1.Location as LocationCurve;
            Line line1 = location1.Curve as Line;
            
            // Extend conduit1 to meet conduit2
            Line extendedLine = Line.CreateBound(line1.GetEndPoint(0), freeEnd2);
            location1.Curve = extendedLine;
            
            // Create an elbow fitting at the connection point
            ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
            ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
            
            Connector connector1 = FindConnectorClosestTo(connectors1, freeEnd2);
            Connector connector2 = FindConnectorClosestTo(connectors2, freeEnd2);
            
            if (connector1 != null && connector2 != null)
            {
                doc.Create.NewElbowFitting(connector1, connector2);
            }
        }
        
        private XYZ ProjectPointOnLine(XYZ point, XYZ linePoint, XYZ lineDirection)
        {
            XYZ normalizedDirection = lineDirection.Normalize();
            double projection = (point - linePoint).DotProduct(normalizedDirection);
            return linePoint + normalizedDirection * projection;
        }
        
        private Connector FindConnectorClosestTo(ConnectorSet connectors, XYZ point)
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
    }
}

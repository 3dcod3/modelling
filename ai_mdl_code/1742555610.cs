using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ConnectConduits
{
    public class ConduitConnector
    {
        public Result Execute(UIDocument uidoc)
        {
            Document doc = uidoc.Document;
            
            try
            {
                // Get the conduit elements by their IDs
                ElementId conduit1Id = new ElementId(7721116);
                ElementId conduit2Id = new ElementId(7721143);
                
                Conduit conduit1 = doc.GetElement(conduit1Id) as Conduit;
                Conduit conduit2 = doc.GetElement(conduit2Id) as Conduit;
                
                if (conduit1 == null || conduit2 == null)
                {
                    TaskDialog.Show("Error", "One or both conduits could not be found.");
                    return Result.Failed;
                }
                
                // Start a transaction
                using (Transaction trans = new Transaction(doc, "Connect Conduits with Elbow"))
                {
                    trans.Start();
                    
                    // Get the connectors from each conduit
                    ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
                    ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;
                    
                    // Find the free connectors (the ones not already connected)
                    Connector freeConnector1 = GetFreeConnector(connectors1);
                    Connector freeConnector2 = GetFreeConnector(connectors2);
                    
                    if (freeConnector1 == null || freeConnector2 == null)
                    {
                        TaskDialog.Show("Error", "Could not find free connectors on the conduits.");
                        trans.RollBack();
                        return Result.Failed;
                    }
                    
                    // Check if the conduits need to be extended to meet
                    XYZ point1 = freeConnector1.Origin;
                    XYZ point2 = freeConnector2.Origin;
                    
                    // Calculate the intersection point
                    XYZ intersectionPoint = CalculateIntersectionPoint(
                        point1, 
                        new XYZ(point1.X, point1.Y, point2.Z), 
                        point2, 
                        new XYZ(point1.X, point2.Y, point2.Z)
                    );
                    
                    // Extend conduits if needed
                    if (!ArePointsClose(point1, intersectionPoint) && !ArePointsClose(point2, intersectionPoint))
                    {
                        // Extend first conduit to intersection point
                        Line line1 = (conduit1.Location as LocationCurve).Curve as Line;
                        XYZ otherEnd1 = line1.GetEndPoint(line1.GetEndPoint(0).IsAlmostEqualTo(point1) ? 1 : 0);
                        
                        // Create a new line with the intersection point
                        Line newLine1 = Line.CreateBound(otherEnd1, new XYZ(intersectionPoint.X, intersectionPoint.Y, intersectionPoint.Z));
                        (conduit1.Location as LocationCurve).Curve = newLine1;
                        
                        // Extend second conduit to intersection point
                        Line line2 = (conduit2.Location as LocationCurve).Curve as Line;
                        XYZ otherEnd2 = line2.GetEndPoint(line2.GetEndPoint(0).IsAlmostEqualTo(point2) ? 1 : 0);
                        
                        // Create a new line with the intersection point
                        Line newLine2 = Line.CreateBound(otherEnd2, new XYZ(intersectionPoint.X, intersectionPoint.Y, intersectionPoint.Z));
                        (conduit2.Location as LocationCurve).Curve = newLine2;
                        
                        // Update the connectors after extending
                        connectors1 = conduit1.ConnectorManager.Connectors;
                        connectors2 = conduit2.ConnectorManager.Connectors;
                        
                        freeConnector1 = GetFreeConnector(connectors1);
                        freeConnector2 = GetFreeConnector(connectors2);
                    }
                    
                    // Create the elbow fitting
                    doc.Create.NewElbowFitting(freeConnector1, freeConnector2);
                    
                    trans.Commit();
                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
                return Result.Failed;
            }
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
        
        private XYZ CalculateIntersectionPoint(XYZ line1Start, XYZ line1End, XYZ line2Start, XYZ line2End)
        {
            // For perpendicular conduits, one vertical and one horizontal
            // The intersection point will have:
            // X from the vertical conduit
            // Y from the horizontal conduit
            // Z from the intersection of the two
            
            return new XYZ(line1Start.X, line2Start.Y, line2Start.Z);
        }
        
        private bool ArePointsClose(XYZ point1, XYZ point2, double tolerance = 0.001)
        {
            return point1.DistanceTo(point2) < tolerance;
        }
    }
}
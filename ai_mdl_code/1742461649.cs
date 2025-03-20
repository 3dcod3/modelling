using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConduitConnector
{
    public class ConduitConnectionCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            try
            {
                // Get the conduit with ID 7721116 (from the input)
                ElementId conduitId = new ElementId(7721116);
                Conduit firstConduit = doc.GetElement(conduitId) as Conduit;
                
                if (firstConduit == null)
                {
                    TaskDialog.Show("Error", "Conduit with ID 7721116 not found.");
                    return Result.Failed;
                }
                
                // Ask user to select another conduit to connect to
                Reference secondConduitRef = null;
                try
                {
                    secondConduitRef = uidoc.Selection.PickObject(ObjectType.Element, 
                        new ConduitSelectionFilter(), "Select a conduit to connect to");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }
                
                Conduit secondConduit = doc.GetElement(secondConduitRef.ElementId) as Conduit;
                
                if (secondConduit == null)
                {
                    TaskDialog.Show("Error", "Selected element is not a conduit.");
                    return Result.Failed;
                }
                
                // Create MEPCurvePair to analyze the relationship between conduits
                MEPCurvePair conduitPair = new MEPCurvePair(firstConduit, secondConduit);
                
                // Start transaction
                using (Transaction trans = new Transaction(doc, "Connect Conduits"))
                {
                    trans.Start();
                    
                    try
                    {
                        // Choose appropriate builder based on conduit geometry
                        if (conduitPair.IsParallel)
                        {
                            // Use OffsetBuilder for parallel conduits
                            OffsetBuilder builder = new OffsetBuilder(doc);
                            builder.Connect(firstConduit, secondConduit);
                        }
                        else if (conduitPair.HasOffset)
                        {
                            // Use KickBuilder for non-parallel conduits with offset
                            KickBuilder builder = new KickBuilder(doc);
                            builder.Connect(firstConduit, secondConduit);
                        }
                        else
                        {
                            // Use TrimBuilder for non-parallel conduits with no offset
                            TrimBuilder builder = new TrimBuilder(doc);
                            builder.Connect(firstConduit, secondConduit);
                        }
                        
                        trans.Commit();
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
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
    
    // Helper class for MEP curve pair analysis
    public class MEPCurvePair
    {
        private MEPCurve _curve1;
        private MEPCurve _curve2;
        private bool _isParallel;
        private bool _hasOffset;
        
        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            _curve1 = curve1;
            _curve2 = curve2;
            AnalyzeCurves();
        }
        
        public bool IsParallel => _isParallel;
        public bool HasOffset => _hasOffset;
        
        private void AnalyzeCurves()
        {
            // Get curve locations
            LocationCurve loc1 = _curve1.Location as LocationCurve;
            LocationCurve loc2 = _curve2.Location as LocationCurve;
            
            if (loc1 == null || loc2 == null)
                throw new InvalidOperationException("One of the MEP curves does not have a valid location.");
                
            Line line1 = loc1.Curve as Line;
            Line line2 = loc2.Curve as Line;
            
            if (line1 == null || line2 == null)
                throw new InvalidOperationException("One of the MEP curves is not a straight line.");
            
            // Check if lines are parallel
            XYZ dir1 = line1.Direction.Normalize();
            XYZ dir2 = line2.Direction.Normalize();
            
            // If dot product is close to 1 or -1, lines are parallel
            double dotProduct = Math.Abs(dir1.DotProduct(dir2));
            _isParallel = Math.Abs(dotProduct - 1.0) < 0.001;
            
            // Check if there's an offset between the lines
            if (_isParallel)
            {
                // For parallel lines, check if they're colinear
                XYZ vec = line2.Origin - line1.Origin;
                XYZ crossProduct = vec.CrossProduct(dir1);
                _hasOffset = crossProduct.GetLength() > 0.001;
            }
            else
            {
                // For non-parallel lines, check if they intersect
                XYZ closestPt1, closestPt2;
                GetClosestPoints(line1, line2, out closestPt1, out closestPt2);
                _hasOffset = closestPt1.DistanceTo(closestPt2) > 0.001;
            }
        }
        
        private void GetClosestPoints(Line line1, Line line2, out XYZ point1, out XYZ point2)
        {
            XYZ dir1 = line1.Direction;
            XYZ dir2 = line2.Direction;
            XYZ origin1 = line1.Origin;
            XYZ origin2 = line2.Origin;
            
            // Calculate parameters for closest points
            XYZ w0 = origin1 - origin2;
            double a = dir1.DotProduct(dir1);
            double b = dir1.DotProduct(dir2);
            double c = dir2.DotProduct(dir2);
            double d = dir1.DotProduct(w0);
            double e = dir2.DotProduct(w0);
            
            double denominator = a * c - b * b;
            
            // Handle parallel case
            if (Math.Abs(denominator) < 0.001)
            {
                double t1 = 0;
                double t2 = (b > c ? d / b : e / c);
                
                point1 = origin1 + t1 * dir1;
                point2 = origin2 + t2 * dir2;
            }
            else
            {
                double t1 = (b * e - c * d) / denominator;
                double t2 = (a * e - b * d) / denominator;
                
                point1 = origin1 + t1 * dir1;
                point2 = origin2 + t2 * dir2;
            }
        }
    }
    
    // Builder classes for different connection types
    public class KickBuilder
    {
        private Document _doc;
        
        public KickBuilder(Document doc)
        {
            _doc = doc;
        }
        
        public void Connect(Conduit conduit1, Conduit conduit2)
        {
            // Implementation for kick connection (non-parallel conduits with offset)
            LocationCurve loc1 = conduit1.Location as LocationCurve;
            LocationCurve loc2 = conduit2.Location as LocationCurve;
            
            Line line1 = loc1.Curve as Line;
            Line line2 = loc2.Curve as Line;
            
            // Find closest points between the two lines
            XYZ closestPt1, closestPt2;
            GetClosestPoints(line1, line2, out closestPt1, out closestPt2);
            
            // Adjust the end point of the first conduit to create a kick
            loc1.Curve = Line.CreateBound(line1.Origin, closestPt1);
            
            // Create an elbow fitting at the connection point if needed
            ConnectWithFitting(conduit1, conduit2, closestPt1, closestPt2);
        }
        
        private void GetClosestPoints(Line line1, Line line2, out XYZ point1, out XYZ point2)
        {
            XYZ dir1 = line1.Direction;
            XYZ dir2 = line2.Direction;
            XYZ origin1 = line1.Origin;
            XYZ origin2 = line2.Origin;
            
            // Calculate parameters for closest points
            XYZ w0 = origin1 - origin2;
            double a = dir1.DotProduct(dir1);
            double b = dir1.DotProduct(dir2);
            double c = dir2.DotProduct(dir2);
            double d = dir1.DotProduct(w0);
            double e = dir2.DotProduct(w0);
            
            double denominator = a * c - b * b;
            
            // Handle parallel case
            if (Math.Abs(denominator) < 0.001)
            {
                double t1 = 0;
                double t2 = (b > c ? d / b : e / c);
                
                point1 = origin1 + t1 * dir1;
                point2 = origin2 + t2 * dir2;
            }
            else
            {
                double t1 = (b * e - c * d) / denominator;
                double t2 = (a * e - b * d) / denominator;
                
                point1 = origin1 + t1 * dir1;
                point2 = origin2 + t2 * dir2;
            }
        }
        
        private void ConnectWithFitting(Conduit conduit1, Conduit conduit2, XYZ point1, XYZ point2)
        {
            // Get connectors at the ends of the conduits
            Connector connector1 = GetConnectorClosestTo(conduit1, point1);
            Connector connector2 = GetConnectorClosestTo(conduit2, point2);
            
            if (connector1 != null && connector2 != null)
            {
                // Create a fitting between the connectors
                _doc.Create.NewElbowFitting(connector1, connector2);
            }
        }
        
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
    
    public class OffsetBuilder
    {
        private Document _doc;
        
        public OffsetBuilder(Document doc)
        {
            _doc = doc;
        }
        
        public void Connect(Conduit conduit1, Conduit conduit2)
        {
            // Implementation for offset connection (parallel conduits)
            LocationCurve loc1 = conduit1.Location as LocationCurve;
            LocationCurve loc2 = conduit2.Location as LocationCurve;
            
            Line line1 = loc1.Curve as Line;
            Line line2 = loc2.Curve as Line;
            
            // Project the free end of conduit1 onto the line of conduit2
            XYZ freeEnd = new XYZ(1720.646628543, -138.968860576, 126.729166667); // From input
            XYZ direction = line1.Direction;
            
            // Create a new point that's projected onto conduit2's line
            XYZ projectedPoint = ProjectPointOntoLine(freeEnd, line2);
            
            // Create a new segment from the free end to the projected point
            loc1.Curve = Line.CreateBound(line1.Origin, projectedPoint);
            
            // Create an elbow fitting at the connection point if needed
            ConnectWithFitting(conduit1, conduit2, projectedPoint);
        }
        
        private XYZ ProjectPointOntoLine(XYZ point, Line line)
        {
            XYZ lineOrigin = line.Origin;
            XYZ lineDirection = line.Direction.Normalize();
            
            // Vector from line origin to point
            XYZ v = point - lineOrigin;
            
            // Project v onto the line direction
            double projectionLength = v.DotProduct(lineDirection);
            
            // Calculate the projected point
            XYZ projectedPoint = lineOrigin + projectionLength * lineDirection;
            
            return projectedPoint;
        }
        
        private void ConnectWithFitting(Conduit conduit1, Conduit conduit2, XYZ connectionPoint)
        {
            // Get connectors at the ends of the conduits
            Connector connector1 = GetConnectorClosestTo(conduit1, connectionPoint);
            Connector connector2 = GetConnectorClosestTo(conduit2, connectionPoint);
            
            if (connector1 != null && connector2 != null)
            {
                // Create a fitting between the connectors
                _doc.Create.NewTeeFitting(connector1, connector2);
            }
        }
        
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
    
    public class TrimBuilder
    {
        private Document _doc;
        
        public TrimBuilder(Document doc)
        {
            _doc = doc;
        }
        
        public void Connect(Conduit conduit1, Conduit conduit2)
        {
            // Implementation for trim connection (non-parallel conduits with no offset)
            LocationCurve loc1 = conduit1.Location as LocationCurve;
            LocationCurve loc2 = conduit2.Location as LocationCurve;
            
            Line line1 = loc1.Curve as Line;
            Line line2 = loc2.Curve as Line;
            
            // Find intersection point
            XYZ intersectionPoint = FindIntersection(line1, line2);
            
            if (intersectionPoint != null)
            {
                // Trim both conduits to the intersection point
                loc1.Curve = Line.CreateBound(line1.Origin, intersectionPoint);
                loc2.Curve = Line.CreateBound(line2.Origin, intersectionPoint);
                
                // Create a fitting at the intersection point
                ConnectWithFitting(conduit1, conduit2, intersectionPoint);
            }
        }
        
        private XYZ FindIntersection(Line line1, Line line2)
        {
            // This is a simplified intersection finder for non-parallel lines
            // In a real implementation, you would need more robust intersection logic
            
            XYZ dir1 = line1.Direction.Normalize();
            XYZ dir2 = line2.Direction.Normalize();
            XYZ origin1 = line1.Origin;
            XYZ origin2 = line2.Origin;
            
            // Check if lines are parallel
            double dotProduct = Math.Abs(dir1.DotProduct(dir2));
            if (Math.Abs(dotProduct - 1.0) < 0.001)
                return null; // Parallel lines don't intersect
                
            // Find closest points between the lines
            XYZ closestPt1, closestPt2;
            GetClosestPoints(line1, line2, out closestPt1, out closestPt2);
            
            // If closest points are very close, consider it an intersection
            if (closestPt1.DistanceTo(closestPt2) < 0.001)
                return closestPt1;
                
            return null;
        }
        
        private void GetClosestPoints(Line line1, Line line2, out XYZ point1, out XYZ point2)
        {
            XYZ dir1 = line1.Direction;
            XYZ dir2 = line2.Direction;
            XYZ origin1 = line1.Origin;
            XYZ origin2 = line2.Origin;
            
            // Calculate parameters for closest points
            XYZ w0 = origin1 - origin2;
            double a = dir1.DotProduct(dir1);
            double b = dir1.DotProduct(dir2);
            double c = dir2.DotProduct(dir2);
            double d = dir1.DotProduct(w0);
            double e = dir2.DotProduct(w0);
            
            double denominator = a * c - b * b;
            
            // Handle parallel case
            if (Math.Abs(denominator) < 0.001)
            {
                double t1 = 0;
                double t2 = (b > c ? d / b : e / c);
                
                point1 = origin1 + t1 * dir1;
                point2 = origin2 + t2 * dir2;
            }
            else
            {
                double t1 = (b * e - c * d) / denominator;
                double t2 = (a * e - b * d) / denominator;
                
                point1 = origin1 + t1 * dir1;
                point2 = origin2 + t2 * dir2;
            }
        }
        
        private void ConnectWithFitting(Conduit conduit1, Conduit conduit2, XYZ intersectionPoint)
        {
            // Get connectors at the ends of the conduits
            Connector connector1 = GetConnectorClosestTo(conduit1, intersectionPoint);
            Connector connector2 = GetConnectorClosestTo(conduit2, intersectionPoint);
            
            if (connector1 != null && connector2 != null)
            {
                // Create a fitting between the connectors
                _doc.Create.NewElbowFitting(connector1, connector2);
            }
        }
        
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
    
    // Helper class for conduit selection
    public class ConduitSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Conduit;
        }
        
        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
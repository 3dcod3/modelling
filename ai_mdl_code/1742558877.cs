using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConduitConnector
{
    /// <summary>
    /// Represents a pair of MEP curves for analysis
    /// </summary>
    public class MEPCurvePair
    {
        private readonly MEPCurve _curve1;
        private readonly MEPCurve _curve2;
        private readonly Line _line1;
        private readonly Line _line2;
        private readonly double _tolerance = 0.001;

        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            _curve1 = curve1 ?? throw new ArgumentNullException(nameof(curve1));
            _curve2 = curve2 ?? throw new ArgumentNullException(nameof(curve2));
            
            // Get curve locations
            LocationCurve locationCurve1 = _curve1.Location as LocationCurve;
            LocationCurve locationCurve2 = _curve2.Location as LocationCurve;
            
            if (locationCurve1 == null || locationCurve2 == null)
                throw new InvalidOperationException("One or both MEP curves do not have valid location curves.");
            
            _line1 = locationCurve1.Curve as Line;
            _line2 = locationCurve2.Curve as Line;
            
            if (_line1 == null || _line2 == null)
                throw new InvalidOperationException("One or both MEP curves are not straight lines.");
        }

        public MEPCurve Curve1 => _curve1;
        public MEPCurve Curve2 => _curve2;
        public Line Line1 => _line1;
        public Line Line2 => _line2;

        /// <summary>
        /// Determines if the two curves are parallel
        /// </summary>
        public bool IsParallel
        {
            get
            {
                XYZ dir1 = _line1.Direction;
                XYZ dir2 = _line2.Direction;
                
                // Check if directions are parallel or anti-parallel
                double dotProduct = Math.Abs(dir1.DotProduct(dir2));
                return Math.Abs(dotProduct - 1.0) < _tolerance;
            }
        }

        /// <summary>
        /// Calculates the offset distance between parallel lines
        /// </summary>
        public double Offset
        {
            get
            {
                if (!IsParallel)
                    return 0;
                
                // For parallel lines, find the perpendicular distance
                XYZ p1 = _line1.GetEndPoint(0);
                XYZ p2 = _line2.GetEndPoint(0);
                XYZ dir = _line1.Direction;
                
                // Project vector p1p2 onto the direction to get the parallel component
                XYZ p1p2 = p2 - p1;
                XYZ parallelComponent = dir.Multiply(p1p2.DotProduct(dir));
                
                // The perpendicular component gives us the offset
                XYZ perpendicularComponent = p1p2 - parallelComponent;
                return perpendicularComponent.GetLength();
            }
        }

        /// <summary>
        /// Determines if the lines intersect in 3D space
        /// </summary>
        public bool LinesIntersect
        {
            get
            {
                if (IsParallel)
                    return false;
                
                // Find closest points between skew lines
                XYZ p1 = _line1.GetEndPoint(0);
                XYZ dir1 = _line1.Direction;
                XYZ p2 = _line2.GetEndPoint(0);
                XYZ dir2 = _line2.Direction;
                
                // Calculate parameters for closest points
                XYZ p1p2 = p2 - p1;
                double d1d1 = dir1.DotProduct(dir1);
                double d2d2 = dir2.DotProduct(dir2);
                double d1d2 = dir1.DotProduct(dir2);
                double d1p1p2 = dir1.DotProduct(p1p2);
                double d2p1p2 = dir2.DotProduct(p1p2);
                
                double denominator = d1d1 * d2d2 - d1d2 * d1d2;
                if (Math.Abs(denominator) < _tolerance)
                    return false;
                
                double t1 = (d1d2 * d2p1p2 - d2d2 * d1p1p2) / denominator;
                double t2 = (d1d1 * d2p1p2 - d1d2 * d1p1p2) / denominator;
                
                // Calculate closest points
                XYZ closestPoint1 = p1 + dir1.Multiply(t1);
                XYZ closestPoint2 = p2 + dir2.Multiply(t2);
                
                // Check if closest points are within tolerance
                double distance = closestPoint1.DistanceTo(closestPoint2);
                return distance < _tolerance;
            }
        }
    }

    /// <summary>
    /// Base class for all connection builders
    /// </summary>
    public abstract class ConnectionBuilder
    {
        protected readonly Document _doc;
        
        public ConnectionBuilder(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }
        
        public abstract bool CanConnect(MEPCurvePair curvePair);
        public abstract bool Connect(MEPCurvePair curvePair);
    }

    /// <summary>
    /// Builder for connecting conduits that need to be trimmed to meet at a point
    /// </summary>
    public class TrimBuilder : ConnectionBuilder
    {
        public TrimBuilder(Document doc) : base(doc) { }
        
        public override bool CanConnect(MEPCurvePair curvePair)
        {
            // Can connect if lines are not parallel and have no offset
            return !curvePair.IsParallel && curvePair.LinesIntersect;
        }
        
        public override bool Connect(MEPCurvePair curvePair)
        {
            if (!CanConnect(curvePair))
                return false;
            
            try
            {
                // Find intersection point
                Line line1 = curvePair.Line1;
                Line line2 = curvePair.Line2;
                
                // Calculate intersection point
                XYZ p1 = line1.GetEndPoint(0);
                XYZ dir1 = line1.Direction;
                XYZ p2 = line2.GetEndPoint(0);
                XYZ dir2 = line2.Direction;
                
                // Calculate parameters for closest points
                XYZ p1p2 = p2 - p1;
                double d1d1 = dir1.DotProduct(dir1);
                double d2d2 = dir2.DotProduct(dir2);
                double d1d2 = dir1.DotProduct(dir2);
                double d1p1p2 = dir1.DotProduct(p1p2);
                double d2p1p2 = dir2.DotProduct(p1p2);
                
                double denominator = d1d1 * d2d2 - d1d2 * d1d2;
                double t1 = (d1d2 * d2p1p2 - d2d2 * d1p1p2) / denominator;
                
                XYZ intersectionPoint = p1 + dir1.Multiply(t1);
                
                // Trim both conduits to the intersection point
                LocationCurve locationCurve1 = curvePair.Curve1.Location as LocationCurve;
                LocationCurve locationCurve2 = curvePair.Curve2.Location as LocationCurve;
                
                // Determine which end of each line is closer to the intersection
                double dist1Start = line1.GetEndPoint(0).DistanceTo(intersectionPoint);
                double dist1End = line1.GetEndPoint(1).DistanceTo(intersectionPoint);
                double dist2Start = line2.GetEndPoint(0).DistanceTo(intersectionPoint);
                double dist2End = line2.GetEndPoint(1).DistanceTo(intersectionPoint);
                
                // Create new lines that end at the intersection point
                Line newLine1, newLine2;
                
                if (dist1Start < dist1End)
                {
                    newLine1 = Line.CreateBound(intersectionPoint, line1.GetEndPoint(1));
                }
                else
                {
                    newLine1 = Line.CreateBound(line1.GetEndPoint(0), intersectionPoint);
                }
                
                if (dist2Start < dist2End)
                {
                    newLine2 = Line.CreateBound(intersectionPoint, line2.GetEndPoint(1));
                }
                else
                {
                    newLine2 = Line.CreateBound(line2.GetEndPoint(0), intersectionPoint);
                }
                
                // Update the conduit locations
                locationCurve1.Curve = newLine1;
                locationCurve2.Curve = newLine2;
                
                // Create an elbow fitting at the intersection
                Connector connector1 = FindConnectorClosestTo(curvePair.Curve1, intersectionPoint);
                Connector connector2 = FindConnectorClosestTo(curvePair.Curve2, intersectionPoint);
                
                if (connector1 != null && connector2 != null)
                {
                    _doc.Create.NewElbowFitting(connector1, connector2);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
                return false;
            }
        }
        
        private Connector FindConnectorClosestTo(MEPCurve curve, XYZ point)
        {
            ConnectorSet connectors = curve.ConnectorManager.Connectors;
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
    /// Builder for connecting parallel conduits with an offset
    /// </summary>
    public class OffsetBuilder : ConnectionBuilder
    {
        public OffsetBuilder(Document doc) : base(doc) { }
        
        public override bool CanConnect(MEPCurvePair curvePair)
        {
            // Can connect if lines are parallel and have an offset
            return curvePair.IsParallel && curvePair.Offset > 0;
        }
        
        public override bool Connect(MEPCurvePair curvePair)
        {
            if (!CanConnect(curvePair))
                return false;
            
            try
            {
                // For parallel lines with offset, we need to create a perpendicular connection
                Line line1 = curvePair.Line1;
                Line line2 = curvePair.Line2;
                
                // Find the closest points between the two lines
                XYZ p1 = line1.GetEndPoint(0);
                XYZ dir1 = line1.Direction;
                XYZ p2 = line2.GetEndPoint(0);
                
                // Project p2 onto line1 to find the closest point on line1
                XYZ p1p2 = p2 - p1;
                double t = p1p2.DotProduct(dir1);
                XYZ closestPointOnLine1 = p1 + dir1.Multiply(t);
                
                // Create a new conduit connecting the two lines
                Conduit connectingConduit = Conduit.Create(_doc, curvePair.Curve1.GetTypeId(), 
                    closestPointOnLine1, p2, curvePair.Curve1.ReferenceLevel.Id);
                
                // Create elbow fittings at both ends of the connecting conduit
                Connector connector1 = FindConnectorClosestTo(curvePair.Curve1, closestPointOnLine1);
                Connector connector2 = FindConnectorClosestTo(connectingConduit, closestPointOnLine1);
                Connector connector3 = FindConnectorClosestTo(connectingConduit, p2);
                Connector connector4 = FindConnectorClosestTo(curvePair.Curve2, p2);
                
                if (connector1 != null && connector2 != null)
                {
                    _doc.Create.NewElbowFitting(connector1, connector2);
                }
                
                if (connector3 != null && connector4 != null)
                {
                    _doc.Create.NewElbowFitting(connector3, connector4);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
                return false;
            }
        }
        
        private Connector FindConnectorClosestTo(MEPCurve curve, XYZ point)
        {
            ConnectorSet connectors = curve.ConnectorManager.Connectors;
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
    /// Builder for connecting non-parallel conduits with an offset
    /// </summary>
    public class KickBuilder : ConnectionBuilder
    {
        public KickBuilder(Document doc) : base(doc) { }
        
        public override bool CanConnect(MEPCurvePair curvePair)
        {
            // Can connect if lines are not parallel and do not intersect
            return !curvePair.IsParallel && !curvePair.LinesIntersect;
        }
        
        public override bool Connect(MEPCurvePair curvePair)
        {
            if (!CanConnect(curvePair))
                return false;
            
            try
            {
                // For non-parallel lines that don't intersect, we need to create a kick connection
                Line line1 = curvePair.Line1;
                Line line2 = curvePair.Line2;
                
                // Find closest points between the two lines
                XYZ p1 = line1.GetEndPoint(0);
                XYZ dir1 = line1.Direction;
                XYZ p2 = line2.GetEndPoint(0);
                XYZ dir2 = line2.Direction;
                
                // Calculate parameters for closest points
                XYZ p1p2 = p2 - p1;
                double d1d1 = dir1.DotProduct(dir1);
                double d2d2 = dir2.DotProduct(dir2);
                double d1d2 = dir1.DotProduct(dir2);
                double d1p1p2 = dir1.DotProduct(p1p2);
                double d2p1p2 = dir2.DotProduct(p1p2);
                
                double denominator = d1d1 * d2d2 - d1d2 * d1d2;
                double t1 = (d1d2 * d2p1p2 - d2d2 * d1p1p2) / denominator;
                double t2 = (d1d1 * d2p1p2 - d1d2 * d1p1p2) / denominator;
                
                XYZ closestPoint1 = p1 + dir1.Multiply(t1);
                XYZ closestPoint2 = p2 + dir2.Multiply(t2);
                
                // Create a new conduit connecting the closest points
                Conduit connectingConduit = Conduit.Create(_doc, curvePair.Curve1.GetTypeId(), 
                    closestPoint1, closestPoint2, curvePair.Curve1.ReferenceLevel.Id);
                
                // Trim the original conduits to the closest points
                LocationCurve locationCurve1 = curvePair.Curve1.Location as LocationCurve;
                LocationCurve locationCurve2 = curvePair.Curve2.Location as LocationCurve;
                
                // Determine which end of each line is closer to the closest point
                double dist1Start = line1.GetEndPoint(0).DistanceTo(closestPoint1);
                double dist1End = line1.GetEndPoint(1).DistanceTo(closestPoint1);
                double dist2Start = line2.GetEndPoint(0).DistanceTo(closestPoint2);
                double dist2End = line2.GetEndPoint(1).DistanceTo(closestPoint2);
                
                // Create new lines that end at the closest points
                Line newLine1, newLine2;
                
                if (dist1Start < dist1End)
                {
                    newLine1 = Line.CreateBound(closestPoint1, line1.GetEndPoint(1));
                }
                else
                {
                    newLine1 = Line.CreateBound(line1.GetEndPoint(0), closestPoint1);
                }
                
                if (dist2Start < dist2End)
                {
                    newLine2 = Line.CreateBound(closestPoint2, line2.GetEndPoint(1));
                }
                else
                {
                    newLine2 = Line.CreateBound(line2.GetEndPoint(0), closestPoint2);
                }
                
                // Update the conduit locations
                locationCurve1.Curve = newLine1;
                locationCurve2.Curve = newLine2;
                
                // Create elbow fittings at both ends of the connecting conduit
                Connector connector1 = FindConnectorClosestTo(curvePair.Curve1, closestPoint1);
                Connector connector2 = FindConnectorClosestTo(connectingConduit, closestPoint1);
                Connector connector3 = FindConnectorClosestTo(connectingConduit, closestPoint2);
                Connector connector4 = FindConnectorClosestTo(curvePair.Curve2, closestPoint2);
                
                if (connector1 != null && connector2 != null)
                {
                    _doc.Create.NewElbowFitting(connector1, connector2);
                }
                
                if (connector3 != null && connector4 != null)
                {
                    _doc.Create.NewElbowFitting(connector3, connector4);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
                return false;
            }
        }
        
        private Connector FindConnectorClosestTo(MEPCurve curve, XYZ point)
        {
            ConnectorSet connectors = curve.ConnectorManager.Connectors;
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
    /// Main class for connecting conduits with a straight line
    /// </summary>
    public class ConduitConnector
    {
        public static Result ConnectConduits(UIApplication uiApp)
        {
            Document doc = uiApp.ActiveUIDocument.Document;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            
            try
            {
                // Get selected conduits
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                
                if (selectedIds.Count != 2)
                {
                    TaskDialog.Show("Error", "Please select exactly two conduits to connect.");
                    return Result.Failed;
                }
                
                // Get the selected conduits
                List<Conduit> selectedConduits = new List<Conduit>();
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is Conduit conduit)
                    {
                        selectedConduits.Add(conduit);
                    }
                }
                
                if (selectedConduits.Count != 2)
                {
                    TaskDialog.Show("Error", "Please select exactly two conduits to connect.");
                    return Result.Failed;
                }
                
                // Create MEPCurvePair to analyze the relationship
                MEPCurvePair curvePair = new MEPCurvePair(selectedConduits[0], selectedConduits[1]);
                
                // Create the appropriate builder based on geometry
                ConnectionBuilder builder = null;
                
                if (curvePair.IsParallel && curvePair.Offset > 0)
                {
                    builder = new OffsetBuilder(doc);
                }
                else if (!curvePair.IsParallel && curvePair.LinesIntersect)
                {
                    builder = new TrimBuilder(doc);
                }
                else if (!curvePair.IsParallel && !curvePair.LinesIntersect)
                {
                    builder = new KickBuilder(doc);
                }
                
                if (builder == null)
                {
                    TaskDialog.Show("Error", "Could not determine appropriate connection method for the selected conduits.");
                    return Result.Failed;
                }
                
                // Start transaction and connect the conduits
                using (Transaction t = new Transaction(doc, "Connect Conduits"))
                {
                    try
                    {
                        t.Start();
                        
                        bool success = builder.Connect(curvePair);
                        
                        if (success)
                        {
                            t.Commit();
                            return Result.Succeeded;
                        }
                        else
                        {
                            t.RollBack();
                            TaskDialog.Show("Error", "Failed to connect the conduits.");
                            return Result.Failed;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (t.HasStarted())
                        {
                            t.RollBack();
                        }
                        TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
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
}
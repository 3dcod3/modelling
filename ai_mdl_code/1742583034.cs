using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MEPTools
{
    /// <summary>
    /// Builder class for trimming and connecting conduits
    /// </summary>
    public class TrimBuilder
    {
        private readonly Document _doc;
        private Conduit _conduit1;
        private Conduit _conduit2;
        private XYZ _freeEnd1;
        private XYZ _freeEnd2;
        private XYZ _direction1;
        private XYZ _direction2;
        private XYZ _intersectionPoint;

        /// <summary>
        /// Constructor for TrimBuilder
        /// </summary>
        /// <param name="document">Active Revit document</param>
        public TrimBuilder(Document document)
        {
            _doc = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Sets the conduits to be connected
        /// </summary>
        /// <param name="conduit1">First conduit</param>
        /// <param name="conduit2">Second conduit</param>
        /// <returns>This builder instance for method chaining</returns>
        public TrimBuilder SetConduits(Conduit conduit1, Conduit conduit2)
        {
            _conduit1 = conduit1 ?? throw new ArgumentNullException(nameof(conduit1));
            _conduit2 = conduit2 ?? throw new ArgumentNullException(nameof(conduit2));
            
            // Get the location curves
            LocationCurve locationCurve1 = _conduit1.Location as LocationCurve;
            LocationCurve locationCurve2 = _conduit2.Location as LocationCurve;
            
            if (locationCurve1 == null || locationCurve2 == null)
                throw new InvalidOperationException("One or both conduits do not have valid location curves.");
            
            // Get the curves
            Line line1 = locationCurve1.Curve as Line;
            Line line2 = locationCurve2.Curve as Line;
            
            if (line1 == null || line2 == null)
                throw new InvalidOperationException("One or both conduits do not have valid line geometry.");
            
            // Store directions and free ends
            _direction1 = line1.Direction;
            _direction2 = line2.Direction;
            
            // Check if the conduits are parallel
            if (_direction1.IsAlmostEqualTo(_direction2) || _direction1.IsAlmostEqualTo(_direction2.Negate()))
                throw new InvalidOperationException("Cannot trim parallel conduits. Use OffsetBuilder instead.");
            
            // Store the free ends
            _freeEnd1 = line1.GetEndPoint(1);
            _freeEnd2 = line2.GetEndPoint(0);
            
            return this;
        }

        /// <summary>
        /// Sets the conduits using element IDs
        /// </summary>
        /// <param name="conduitId1">ID of first conduit</param>
        /// <param name="conduitId2">ID of second conduit</param>
        /// <returns>This builder instance for method chaining</returns>
        public TrimBuilder SetConduits(ElementId conduitId1, ElementId conduitId2)
        {
            Conduit conduit1 = _doc.GetElement(conduitId1) as Conduit;
            Conduit conduit2 = _doc.GetElement(conduitId2) as Conduit;
            
            if (conduit1 == null)
                throw new ArgumentException("Element is not a conduit", nameof(conduitId1));
            
            if (conduit2 == null)
                throw new ArgumentException("Element is not a conduit", nameof(conduitId2));
            
            return SetConduits(conduit1, conduit2);
        }

        /// <summary>
        /// Sets the conduits using the provided free ends and directions
        /// </summary>
        /// <param name="freeEnd1">Free end of first conduit</param>
        /// <param name="direction1">Direction of first conduit</param>
        /// <param name="freeEnd2">Free end of second conduit</param>
        /// <param name="direction2">Direction of second conduit</param>
        /// <returns>This builder instance for method chaining</returns>
        public TrimBuilder SetConduitGeometry(XYZ freeEnd1, XYZ direction1, XYZ freeEnd2, XYZ direction2)
        {
            _freeEnd1 = freeEnd1 ?? throw new ArgumentNullException(nameof(freeEnd1));
            _direction1 = direction1 ?? throw new ArgumentNullException(nameof(direction1));
            _freeEnd2 = freeEnd2 ?? throw new ArgumentNullException(nameof(freeEnd2));
            _direction2 = direction2 ?? throw new ArgumentNullException(nameof(direction2));
            
            // Check if the conduits are parallel
            if (_direction1.IsAlmostEqualTo(_direction2) || _direction1.IsAlmostEqualTo(_direction2.Negate()))
                throw new InvalidOperationException("Cannot trim parallel conduits. Use OffsetBuilder instead.");
            
            // Find the conduits in the document
            FilteredElementCollector collector = new FilteredElementCollector(_doc)
                .OfClass(typeof(Conduit));
            
            foreach (Conduit conduit in collector)
            {
                LocationCurve locationCurve = conduit.Location as LocationCurve;
                if (locationCurve == null) continue;
                
                Line line = locationCurve.Curve as Line;
                if (line == null) continue;
                
                if (line.GetEndPoint(1).IsAlmostEqualTo(_freeEnd1) && 
                    line.Direction.IsAlmostEqualTo(_direction1))
                {
                    _conduit1 = conduit;
                }
                else if (line.GetEndPoint(0).IsAlmostEqualTo(_freeEnd2) && 
                         line.Direction.IsAlmostEqualTo(_direction2))
                {
                    _conduit2 = conduit;
                }
                
                if (_conduit1 != null && _conduit2 != null) break;
            }
            
            if (_conduit1 == null || _conduit2 == null)
                throw new InvalidOperationException("Could not find conduits with the specified geometry.");
            
            return this;
        }

        /// <summary>
        /// Calculates the intersection point between the two conduits
        /// </summary>
        /// <returns>The intersection point</returns>
        private XYZ CalculateIntersectionPoint()
        {
            // Create infinite lines from the conduits
            Line infiniteLine1 = Line.CreateUnbound(_freeEnd1, _direction1.Negate());
            Line infiniteLine2 = Line.CreateUnbound(_freeEnd2, _direction2);
            
            // Find the closest points between the lines
            XYZ closestPoint1, closestPoint2;
            infiniteLine1.GetClosestPoints(infiniteLine2, out closestPoint1, out closestPoint2);
            
            // If the lines don't intersect exactly, use the midpoint between the closest points
            if (!closestPoint1.IsAlmostEqualTo(closestPoint2))
            {
                return (closestPoint1 + closestPoint2) / 2;
            }
            
            return closestPoint1;
        }

        /// <summary>
        /// Executes the trim operation to connect the conduits
        /// </summary>
        /// <returns>True if successful, false otherwise</returns>
        public bool Execute()
        {
            if (_conduit1 == null || _conduit2 == null)
                throw new InvalidOperationException("Conduits must be set before executing.");
            
            try
            {
                using (Transaction trans = new Transaction(_doc, "Connect Conduits with Trim"))
                {
                    trans.Start();
                    
                    // Calculate the intersection point
                    _intersectionPoint = CalculateIntersectionPoint();
                    
                    // Get the location curves
                    LocationCurve locationCurve1 = _conduit1.Location as LocationCurve;
                    LocationCurve locationCurve2 = _conduit2.Location as LocationCurve;
                    
                    // Get the original curves
                    Line line1 = locationCurve1.Curve as Line;
                    Line line2 = locationCurve2.Curve as Line;
                    
                    // Create new trimmed lines
                    Line trimmedLine1 = Line.CreateBound(line1.GetEndPoint(0), _intersectionPoint);
                    Line trimmedLine2 = Line.CreateBound(_intersectionPoint, line2.GetEndPoint(1));
                    
                    // Update the conduit geometry
                    locationCurve1.Curve = trimmedLine1;
                    locationCurve2.Curve = trimmedLine2;
                    
                    // Get the connectors
                    Connector connector1 = GetConnectorClosestTo(_conduit1, _intersectionPoint);
                    Connector connector2 = GetConnectorClosestTo(_conduit2, _intersectionPoint);
                    
                    if (connector1 == null || connector2 == null)
                    {
                        trans.RollBack();
                        return false;
                    }
                    
                    // Connect the conduits
                    connector1.ConnectTo(connector2);
                    
                    trans.Commit();
                    return true;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to connect conduits: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets the connector closest to the specified point
        /// </summary>
        /// <param name="conduit">The conduit</param>
        /// <param name="point">The point</param>
        /// <returns>The closest connector</returns>
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
    /// External command to connect conduits with trim
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ConnectConduitsWithTrimCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            try
            {
                // Example with the provided conduit data
                XYZ freeEnd1 = new XYZ(1710.933383058, -259.456787969, 129.812500000);
                XYZ direction1 = new XYZ(-1.000000000, 0.000000000, 0.000000000);
                
                XYZ freeEnd2 = new XYZ(1720.646628543, -261.009697791, 130.121394041);
                XYZ direction2 = new XYZ(0.000000000, 0.980785151, -0.195090973);
                
                TrimBuilder builder = new TrimBuilder(doc)
                    .SetConduitGeometry(freeEnd1, direction1, freeEnd2, direction2);
                
                bool success = builder.Execute();
                
                return success ? Result.Succeeded : Result.Failed;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitConduitConnector
{
    /// <summary>
    /// Builder class for trimming and connecting conduits with an elbow fitting
    /// </summary>
    public class TrimBuilder
    {
        private readonly Document _doc;

        public TrimBuilder(Document doc)
        {
            _doc = doc;
        }

        /// <summary>
        /// Connect two conduits by trimming them to their intersection point and adding an elbow fitting
        /// </summary>
        /// <param name="conduitId1">ElementId of first conduit</param>
        /// <param name="conduitId2">ElementId of second conduit</param>
        /// <returns>True if connection was successful, false otherwise</returns>
        public bool ConnectConduits(ElementId conduitId1, ElementId conduitId2)
        {
            // Get the conduit elements
            MEPCurve conduit1 = _doc.GetElement(conduitId1) as MEPCurve;
            MEPCurve conduit2 = _doc.GetElement(conduitId2) as MEPCurve;

            if (conduit1 == null || conduit2 == null)
            {
                TaskDialog.Show("Error", "One or both elements are not valid MEPCurves.");
                return false;
            }

            try
            {
                // Create MEPCurvePair to analyze the relationship between conduits
                MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);
                
                // Check if conduits are parallel - we can't trim parallel conduits
                if (curvePair.IsParallel)
                {
                    TaskDialog.Show("Error", "Cannot trim parallel conduits. Use OffsetBuilder instead.");
                    return false;
                }

                // Find intersection point
                XYZ intersectionPoint = FindIntersectionPoint(conduit1, conduit2);
                if (intersectionPoint == null)
                {
                    TaskDialog.Show("Error", "Could not find valid intersection point between conduits.");
                    return false;
                }

                using (Transaction trans = new Transaction(_doc, "Trim and Connect Conduits"))
                {
                    trans.Start();

                    // Trim both conduits to the intersection point
                    bool trimSuccess = TrimConduitsToIntersection(conduit1, conduit2, intersectionPoint);
                    if (!trimSuccess)
                    {
                        trans.RollBack();
                        return false;
                    }

                    // Create elbow fitting at the intersection
                    FamilyInstance elbow = CreateElbowFitting(conduit1, conduit2);
                    if (elbow == null)
                    {
                        trans.RollBack();
                        return false;
                    }

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
        /// Find the intersection point between two non-parallel conduits
        /// </summary>
        private XYZ FindIntersectionPoint(MEPCurve conduit1, MEPCurve conduit2)
        {
            // Get the location curves of both conduits
            Curve curve1 = (conduit1.Location as LocationCurve).Curve;
            Curve curve2 = (conduit2.Location as LocationCurve).Curve;

            // Get line representations
            Line line1 = curve1 as Line;
            Line line2 = curve2 as Line;

            if (line1 == null || line2 == null)
            {
                return null; // Only straight conduits supported
            }

            // Get the direction vectors
            XYZ dir1 = line1.Direction;
            XYZ dir2 = line2.Direction;

            // Check if lines are not parallel
            if (dir1.CrossProduct(dir2).GetLength() < 0.001)
            {
                return null; // Lines are parallel or nearly parallel
            }

            // Find closest points between the two lines
            XYZ point1, point2;
            double param1, param2;
            line1.GetEndPoint(0).DistanceTo(line2.GetEndPoint(0));
            
            // Calculate closest points between the two lines
            bool result = line1.ComputeClosestPoints(line2, out point1, out point2, out param1, out param2);
            
            if (!result || point1.DistanceTo(point2) > 0.01)
            {
                // Lines don't intersect closely enough
                return null;
            }

            // Return the midpoint between the closest points as the intersection
            return (point1 + point2) * 0.5;
        }

        /// <summary>
        /// Trim both conduits to end at the intersection point
        /// </summary>
        private bool TrimConduitsToIntersection(MEPCurve conduit1, MEPCurve conduit2, XYZ intersectionPoint)
        {
            try
            {
                // Get location curves
                LocationCurve locationCurve1 = conduit1.Location as LocationCurve;
                LocationCurve locationCurve2 = conduit2.Location as LocationCurve;

                // Get original curves
                Curve curve1 = locationCurve1.Curve;
                Curve curve2 = locationCurve2.Curve;

                // Get line representations
                Line line1 = curve1 as Line;
                Line line2 = curve2 as Line;

                if (line1 == null || line2 == null)
                {
                    return false;
                }

                // Determine which end of each conduit to trim
                XYZ start1 = line1.GetEndPoint(0);
                XYZ end1 = line1.GetEndPoint(1);
                XYZ start2 = line2.GetEndPoint(0);
                XYZ end2 = line2.GetEndPoint(1);

                // Calculate distances to determine which end to trim
                bool trimStart1 = start1.DistanceTo(intersectionPoint) < end1.DistanceTo(intersectionPoint);
                bool trimStart2 = start2.DistanceTo(intersectionPoint) < end2.DistanceTo(intersectionPoint);

                // Create new trimmed lines
                Line newLine1, newLine2;
                
                if (trimStart1)
                {
                    newLine1 = Line.CreateBound(intersectionPoint, end1);
                }
                else
                {
                    newLine1 = Line.CreateBound(start1, intersectionPoint);
                }

                if (trimStart2)
                {
                    newLine2 = Line.CreateBound(intersectionPoint, end2);
                }
                else
                {
                    newLine2 = Line.CreateBound(start2, intersectionPoint);
                }

                // Update the conduit geometry
                locationCurve1.Curve = newLine1;
                locationCurve2.Curve = newLine2;

                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to trim conduits: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Create an elbow fitting to connect the two conduits
        /// </summary>
        private FamilyInstance CreateElbowFitting(MEPCurve conduit1, MEPCurve conduit2)
        {
            try
            {
                // Get the connectors from both conduits
                ConnectorSet connectors1 = conduit1.ConnectorManager.Connectors;
                ConnectorSet connectors2 = conduit2.ConnectorManager.Connectors;

                // Find the unconnected connectors closest to each other
                Connector connector1 = FindUnconnectedConnector(connectors1);
                Connector connector2 = FindUnconnectedConnector(connectors2);

                if (connector1 == null || connector2 == null)
                {
                    TaskDialog.Show("Error", "Could not find unconnected connectors.");
                    return null;
                }

                // Create the elbow fitting
                return _doc.Create.NewElbowFitting(connector1, connector2);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create elbow fitting: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Find an unconnected connector from a connector set
        /// </summary>
        private Connector FindUnconnectedConnector(ConnectorSet connectorSet)
        {
            foreach (Connector connector in connectorSet)
            {
                if (connector.IsConnected == false)
                {
                    return connector;
                }
            }
            return null;
        }
    }

    /// <summary>
    /// Helper class to analyze the relationship between two MEP curves
    /// </summary>
    public class MEPCurvePair
    {
        public MEPCurve Curve1 { get; }
        public MEPCurve Curve2 { get; }
        public bool IsParallel { get; private set; }
        public double Offset { get; private set; }

        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            Curve1 = curve1;
            Curve2 = curve2;
            AnalyzeCurves();
        }

        private void AnalyzeCurves()
        {
            // Get location curves
            LocationCurve locationCurve1 = Curve1.Location as LocationCurve;
            LocationCurve locationCurve2 = Curve2.Location as LocationCurve;

            if (locationCurve1 == null || locationCurve2 == null)
            {
                throw new InvalidOperationException("One or both MEPCurves do not have valid location curves.");
            }

            // Get line representations
            Line line1 = locationCurve1.Curve as Line;
            Line line2 = locationCurve2.Curve as Line;

            if (line1 == null || line2 == null)
            {
                throw new InvalidOperationException("One or both MEPCurves are not straight lines.");
            }

            // Get direction vectors
            XYZ dir1 = line1.Direction;
            XYZ dir2 = line2.Direction;

            // Check if lines are parallel (cross product is zero)
            double crossProductLength = dir1.CrossProduct(dir2).GetLength();
            IsParallel = crossProductLength < 0.001;

            // Calculate offset for parallel lines
            if (IsParallel)
            {
                // Project a point from line2 onto line1 to find the offset
                XYZ point2 = line2.GetEndPoint(0);
                XYZ point1 = line1.GetEndPoint(0);
                XYZ vector = point2 - point1;
                
                // Remove the component of vector that's parallel to dir1
                double dot = vector.DotProduct(dir1);
                XYZ projection = dir1.Multiply(dot);
                XYZ perpendicular = vector - projection;
                
                // The length of the perpendicular component is the offset
                Offset = perpendicular.GetLength();
            }
            else
            {
                Offset = 0.0;
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Electrical;

namespace RevitConduitConnector
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ConduitConnector : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get the conduit elements (in a real application, you would select these or get them by ID)
                ElementId conduitId1 = new ElementId(7877720);
                ElementId conduitId2 = new ElementId(7877969);

                Conduit conduit1 = doc.GetElement(conduitId1) as Conduit;
                Conduit conduit2 = doc.GetElement(conduitId2) as Conduit;

                if (conduit1 == null || conduit2 == null)
                {
                    message = "One or both conduits could not be found.";
                    return Result.Failed;
                }

                // Create a MEPCurvePair to analyze the relationship between the conduits
                MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);
                
                // Start a transaction
                using (Transaction trans = new Transaction(doc, "Connect Conduits"))
                {
                    trans.Start();
                    
                    // Choose the appropriate builder based on the geometry
                    if (curvePair.IsParallel)
                    {
                        // Use OffsetBuilder for parallel conduits
                        OffsetBuilder builder = new OffsetBuilder(doc);
                        builder.Connect(conduit1, conduit2);
                    }
                    else if (curvePair.HasOffset)
                    {
                        // Use KickBuilder for non-parallel conduits with offset
                        KickBuilder builder = new KickBuilder(doc);
                        builder.Connect(conduit1, conduit2);
                    }
                    else
                    {
                        // Use TrimBuilder for non-parallel conduits with no offset
                        TrimBuilder builder = new TrimBuilder(doc);
                        builder.Connect(conduit1, conduit2);
                    }
                    
                    trans.Commit();
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    // Class to analyze the relationship between two MEP curves
    public class MEPCurvePair
    {
        private MEPCurve _curve1;
        private MEPCurve _curve2;
        private bool _isParallel;
        private bool _hasOffset;
        private double _offset;

        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            _curve1 = curve1;
            _curve2 = curve2;
            AnalyzeRelationship();
        }

        public bool IsParallel => _isParallel;
        public bool HasOffset => _hasOffset;
        public double Offset => _offset;

        private void AnalyzeRelationship()
        {
            // Get the curves from the MEP elements
            Curve curve1 = (_curve1.Location as LocationCurve).Curve;
            Curve curve2 = (_curve2.Location as LocationCurve).Curve;

            // Get the direction vectors
            XYZ direction1 = GetDirection(curve1);
            XYZ direction2 = GetDirection(curve2);

            // Check if the curves are parallel (dot product of normalized vectors close to 1 or -1)
            double dotProduct = direction1.DotProduct(direction2);
            _isParallel = Math.Abs(Math.Abs(dotProduct) - 1.0) < 0.001;

            // Calculate the offset
            if (_isParallel)
            {
                // For parallel lines, calculate the perpendicular distance
                XYZ point1 = curve1.GetEndPoint(0);
                XYZ point2 = curve2.GetEndPoint(0);
                XYZ vector = point2 - point1;
                
                // Project vector onto direction to get parallel component
                XYZ parallelComponent = direction1.Multiply(vector.DotProduct(direction1));
                
                // Subtract to get perpendicular component
                XYZ perpendicularComponent = vector - parallelComponent;
                
                _offset = perpendicularComponent.GetLength();
                _hasOffset = _offset > 0.001; // Threshold for considering an offset
            }
            else
            {
                // For non-parallel lines, check if they intersect
                // This is a simplified approach - in a real application, you would need more robust intersection detection
                XYZ point1 = curve1.GetEndPoint(0);
                XYZ point2 = curve2.GetEndPoint(0);
                
                // Calculate the minimum distance between the lines
                // This is a simplified calculation and may not work for all cases
                XYZ v1 = direction1;
                XYZ v2 = direction2;
                XYZ w0 = point1 - point2;
                
                double a = v1.DotProduct(v1);
                double b = v1.DotProduct(v2);
                double c = v2.DotProduct(v2);
                double d = v1.DotProduct(w0);
                double e = v2.DotProduct(w0);
                
                double denominator = a * c - b * b;
                
                // If lines are not parallel, denominator should not be zero
                if (Math.Abs(denominator) > 0.001)
                {
                    double sc = (b * e - c * d) / denominator;
                    double tc = (a * e - b * d) / denominator;
                    
                    XYZ closestPoint1 = point1 + v1.Multiply(sc);
                    XYZ closestPoint2 = point2 + v2.Multiply(tc);
                    
                    _offset = closestPoint1.DistanceTo(closestPoint2);
                    _hasOffset = _offset > 0.001; // Threshold for considering an offset
                }
                else
                {
                    // Fallback for near-parallel lines
                    _offset = 0;
                    _hasOffset = false;
                }
            }
        }

        private XYZ GetDirection(Curve curve)
        {
            XYZ direction;
            if (curve is Line line)
            {
                direction = line.Direction;
            }
            else
            {
                // For non-line curves, approximate direction using endpoints
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                direction = (end - start).Normalize();
            }
            return direction;
        }
    }

    // Builder for creating a kick connection between non-parallel conduits with offset
    public class KickBuilder
    {
        private Document _doc;

        public KickBuilder(Document doc)
        {
            _doc = doc;
        }

        public void Connect(Conduit conduit1, Conduit conduit2)
        {
            try
            {
                // Get the location curves
                LocationCurve locationCurve1 = conduit1.Location as LocationCurve;
                LocationCurve locationCurve2 = conduit2.Location as LocationCurve;
                
                Curve curve1 = locationCurve1.Curve;
                Curve curve2 = locationCurve2.Curve;
                
                // Get the direction vectors
                XYZ direction1 = GetDirection(curve1);
                XYZ direction2 = GetDirection(curve2);
                
                // Get the endpoints
                XYZ start1 = curve1.GetEndPoint(0);
                XYZ end1 = curve1.GetEndPoint(1);
                XYZ start2 = curve2.GetEndPoint(0);
                XYZ end2 = curve2.GetEndPoint(1);
                
                // Determine which endpoints to connect (using the free ends provided)
                XYZ freeEnd1 = new XYZ(1721.527970000, 93.732615629, 114.492187500);
                XYZ freeEnd2 = new XYZ(1721.527970000, 91.242775874, 114.348196280);
                
                // Find the closest endpoints to the free ends
                XYZ connectEnd1 = (start1.DistanceTo(freeEnd1) < end1.DistanceTo(freeEnd1)) ? start1 : end1;
                XYZ connectEnd2 = (start2.DistanceTo(freeEnd2) < end2.DistanceTo(freeEnd2)) ? start2 : end2;
                
                // Calculate the midpoint for the kick
                XYZ midPoint = (connectEnd1 + connectEnd2) / 2;
                
                // Create the kick segments
                // First segment: from conduit1 to midpoint
                Line kickSegment1 = Line.CreateBound(connectEnd1, midPoint);
                
                // Second segment: from midpoint to conduit2
                Line kickSegment2 = Line.CreateBound(midPoint, connectEnd2);
                
                // Create the conduit fittings
                // Get the conduit type
                ElementId conduitTypeId = conduit1.GetTypeId();
                
                // Create the kick segments as new conduits
                Conduit kickConduit1 = Conduit.Create(_doc, conduitTypeId, connectEnd1, midPoint, conduit1.ReferenceLevel.Id);
                Conduit kickConduit2 = Conduit.Create(_doc, conduitTypeId, midPoint, connectEnd2, conduit1.ReferenceLevel.Id);
                
                // Copy parameters from original conduits
                CopyParameters(conduit1, kickConduit1);
                CopyParameters(conduit2, kickConduit2);
                
                // Create elbows at the connection points
                ConnectWithElbow(conduit1, kickConduit1, connectEnd1);
                ConnectWithElbow(kickConduit1, kickConduit2, midPoint);
                ConnectWithElbow(kickConduit2, conduit2, connectEnd2);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create kick connection: {ex.Message}");
            }
        }

        private XYZ GetDirection(Curve curve)
        {
            if (curve is Line line)
            {
                return line.Direction;
            }
            else
            {
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                return (end - start).Normalize();
            }
        }

        private void CopyParameters(Conduit source, Conduit target)
        {
            // Copy relevant parameters from source to target
            // This is a simplified version - in a real application, you would copy all relevant parameters
            target.Diameter = source.Diameter;
            target.LookupParameter("Trade")?.Set(source.LookupParameter("Trade")?.AsString());
        }

        private void ConnectWithElbow(Conduit conduit1, Conduit conduit2, XYZ connectionPoint)
        {
            // In a real application, you would use the Revit API to create an elbow fitting
            // This is a placeholder for that functionality
            // Example: FamilyInstance elbow = _doc.Create.NewElbowFitting(conduit1, conduit2, connectionPoint);
        }
    }

    // Builder for creating an offset connection between parallel conduits
    public class OffsetBuilder
    {
        private Document _doc;

        public OffsetBuilder(Document doc)
        {
            _doc = doc;
        }

        public void Connect(Conduit conduit1, Conduit conduit2)
        {
            try
            {
                // Implementation for connecting parallel conduits with an offset
                // Similar to KickBuilder but with different geometry calculations
                TaskDialog.Show("Info", "OffsetBuilder: Connecting parallel conduits");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create offset connection: {ex.Message}");
            }
        }
    }

    // Builder for creating a trim connection between non-parallel conduits with no offset
    public class TrimBuilder
    {
        private Document _doc;

        public TrimBuilder(Document doc)
        {
            _doc = doc;
        }

        public void Connect(Conduit conduit1, Conduit conduit2)
        {
            try
            {
                // Implementation for connecting non-parallel conduits that intersect
                // Similar to KickBuilder but with different geometry calculations
                TaskDialog.Show("Info", "TrimBuilder: Connecting intersecting conduits");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create trim connection: {ex.Message}");
            }
        }
    }
}
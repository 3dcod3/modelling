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
                // Get the two conduits from the input data
                Conduit conduit1 = null;
                Conduit conduit2 = null;

                // Find conduits by ID (you would need to adapt this to your actual selection method)
                ElementId id1 = new ElementId(7638672);
                ElementId id2 = new ElementId(7724387);

                Element elem1 = doc.GetElement(id1);
                Element elem2 = doc.GetElement(id2);

                if (elem1 is Conduit && elem2 is Conduit)
                {
                    conduit1 = elem1 as Conduit;
                    conduit2 = elem2 as Conduit;
                }
                else
                {
                    TaskDialog.Show("Error", "Selected elements are not conduits.");
                    return Result.Failed;
                }

                // Create MEPCurvePair to analyze the relationship
                MEPCurvePair conduitPair = new MEPCurvePair(conduit1, conduit2);

                // Start transaction
                using (Transaction trans = new Transaction(doc, "Connect Conduits"))
                {
                    trans.Start();

                    // Select and execute the appropriate builder based on geometry
                    bool success = ConnectConduits(doc, conduitPair);

                    if (success)
                        trans.Commit();
                    else
                        trans.RollBack();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }

        private bool ConnectConduits(Document doc, MEPCurvePair conduitPair)
        {
            try
            {
                // Select the appropriate builder based on geometry analysis
                if (conduitPair.IsParallel)
                {
                    // Use OffsetBuilder for parallel conduits
                    OffsetBuilder builder = new OffsetBuilder(doc);
                    return builder.Build(conduitPair);
                }
                else if (conduitPair.HasOffset)
                {
                    // Use KickBuilder for non-parallel conduits with offset
                    KickBuilder builder = new KickBuilder(doc);
                    return builder.Build(conduitPair);
                }
                else
                {
                    // Use TrimBuilder for non-parallel conduits with no offset
                    TrimBuilder builder = new TrimBuilder(doc);
                    return builder.Build(conduitPair);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Connection Error", $"Failed to connect conduits: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// Analyzes the relationship between two MEPCurves
    /// </summary>
    public class MEPCurvePair
    {
        private readonly MEPCurve _curve1;
        private readonly MEPCurve _curve2;
        private readonly double _parallelTolerance = 0.001;
        private readonly double _offsetTolerance = 0.001;

        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            _curve1 = curve1 ?? throw new ArgumentNullException(nameof(curve1));
            _curve2 = curve2 ?? throw new ArgumentNullException(nameof(curve2));
            
            AnalyzeGeometry();
        }

        public MEPCurve Curve1 => _curve1;
        public MEPCurve Curve2 => _curve2;
        public bool IsParallel { get; private set; }
        public bool HasOffset { get; private set; }
        public double Offset { get; private set; }
        public XYZ IntersectionPoint { get; private set; }

        private void AnalyzeGeometry()
        {
            // Get the line representations of the curves
            Line line1 = (_curve1.Location as LocationCurve)?.Curve as Line;
            Line line2 = (_curve2.Location as LocationCurve)?.Curve as Line;

            if (line1 == null || line2 == null)
                throw new InvalidOperationException("One or both MEPCurves do not have a valid line representation.");

            // Check if lines are parallel
            XYZ dir1 = line1.Direction;
            XYZ dir2 = line2.Direction;
            
            // Normalize directions
            dir1 = dir1.Normalize();
            dir2 = dir2.Normalize();
            
            // Check if directions are parallel or anti-parallel
            double dotProduct = dir1.DotProduct(dir2);
            IsParallel = Math.Abs(Math.Abs(dotProduct) - 1.0) < _parallelTolerance;

            if (IsParallel)
            {
                // Calculate offset for parallel lines
                XYZ p1 = line1.GetEndPoint(0);
                XYZ p2 = line2.GetEndPoint(0);
                XYZ v = p2 - p1;
                
                // Project v onto dir1 to get the component along the line
                XYZ projV = dir1.Multiply(v.DotProduct(dir1));
                
                // The offset is the perpendicular component
                XYZ perpV = v - projV;
                Offset = perpV.GetLength();
                HasOffset = Offset > _offsetTolerance;
            }
            else
            {
                // For non-parallel lines, find the closest points between them
                XYZ closestPoint1, closestPoint2;
                line1.GetClosestPoints(line2, out closestPoint1, out closestPoint2);
                
                // Calculate the distance between the closest points
                Offset = closestPoint1.DistanceTo(closestPoint2);
                HasOffset = Offset > _offsetTolerance;
                
                // If lines intersect, store the intersection point
                if (!HasOffset)
                {
                    IntersectionPoint = closestPoint1;
                }
            }
        }
    }

    /// <summary>
    /// Base class for all builders
    /// </summary>
    public abstract class BuilderBase
    {
        protected readonly Document _doc;

        public BuilderBase(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public abstract bool Build(MEPCurvePair curvePair);

        protected FamilyInstance CreateElbowFitting(MEPCurve curve1, MEPCurve curve2)
        {
            try
            {
                // Get the connector manager for each curve
                ConnectorManager cm1 = curve1.ConnectorManager;
                ConnectorManager cm2 = curve2.ConnectorManager;

                // Find the closest connectors
                Connector con1 = FindClosestConnector(cm1, curve2);
                Connector con2 = FindClosestConnector(cm2, curve1);

                if (con1 == null || con2 == null)
                    throw new InvalidOperationException("Could not find valid connectors.");

                // Create the elbow fitting
                return con1.ConnectWith(con2);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Fitting Error", $"Failed to create elbow fitting: {ex.Message}");
                return null;
            }
        }

        protected Connector FindClosestConnector(ConnectorManager cm, MEPCurve targetCurve)
        {
            Connector closestConnector = null;
            double minDistance = double.MaxValue;

            // Get the line representation of the target curve
            Line targetLine = (targetCurve.Location as LocationCurve)?.Curve as Line;
            if (targetLine == null)
                return null;

            // Find the closest connector to the target curve
            foreach (Connector con in cm.Connectors)
            {
                if (con.IsConnected)
                    continue;

                double distance = targetLine.Distance(con.Origin);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    closestConnector = con;
                }
            }

            return closestConnector;
        }
    }

    /// <summary>
    /// Builder for creating kick connections between non-parallel conduits with offset
    /// </summary>
    public class KickBuilder : BuilderBase
    {
        public KickBuilder(Document doc) : base(doc) { }

        public override bool Build(MEPCurvePair curvePair)
        {
            try
            {
                MEPCurve curve1 = curvePair.Curve1;
                MEPCurve curve2 = curvePair.Curve2;

                // Get the line representations
                Line line1 = (curve1.Location as LocationCurve)?.Curve as Line;
                Line line2 = (curve2.Location as LocationCurve)?.Curve as Line;

                if (line1 == null || line2 == null)
                    return false;

                // Get closest points between the lines
                XYZ closestPoint1, closestPoint2;
                line1.GetClosestPoints(line2, out closestPoint1, out closestPoint2);

                // Create a new conduit to make the kick connection
                // First, find the midpoint between the closest points
                XYZ midPoint = (closestPoint1 + closestPoint2) * 0.5;

                // Create a new conduit from curve1's closest point to the midpoint
                Conduit kickConduit1 = Conduit.Create(_doc, curve1.GetTypeId(), closestPoint1, midPoint, curve1.ReferenceLevel.Id);

                // Create a new conduit from midpoint to curve2's closest point
                Conduit kickConduit2 = Conduit.Create(_doc, curve2.GetTypeId(), midPoint, closestPoint2, curve2.ReferenceLevel.Id);

                // Trim the original conduits to the closest points
                (curve1.Location as LocationCurve).Curve = Line.CreateBound(line1.GetEndPoint(0), closestPoint1);
                (curve2.Location as LocationCurve).Curve = Line.CreateBound(closestPoint2, line2.GetEndPoint(1));

                // Create elbow fittings to connect the conduits
                CreateElbowFitting(curve1, kickConduit1);
                CreateElbowFitting(kickConduit1, kickConduit2);
                CreateElbowFitting(kickConduit2, curve2);

                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Kick Connection Error", $"Failed to create kick connection: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// Builder for creating offset connections between parallel conduits
    /// </summary>
    public class OffsetBuilder : BuilderBase
    {
        public OffsetBuilder(Document doc) : base(doc) { }

        public override bool Build(MEPCurvePair curvePair)
        {
            try
            {
                MEPCurve curve1 = curvePair.Curve1;
                MEPCurve curve2 = curvePair.Curve2;

                // Get the line representations
                Line line1 = (curve1.Location as LocationCurve)?.Curve as Line;
                Line line2 = (curve2.Location as LocationCurve)?.Curve as Line;

                if (line1 == null || line2 == null)
                    return false;

                // Get the direction of the lines
                XYZ dir1 = line1.Direction.Normalize();
                
                // Project curve2's start point onto curve1's line
                XYZ p1 = line1.GetEndPoint(0);
                XYZ p2 = line2.GetEndPoint(0);
                XYZ v = p2 - p1;
                XYZ projV = dir1.Multiply(v.DotProduct(dir1));
                XYZ projPoint = p1 + projV;
                
                // Create a perpendicular line from curve1 to curve2
                XYZ perpDir = (p2 - projPoint).Normalize();
                
                // Create two points for the offset conduit
                XYZ offsetPoint1 = projPoint;
                XYZ offsetPoint2 = p2;
                
                // Create the offset conduit
                Conduit offsetConduit = Conduit.Create(_doc, curve1.GetTypeId(), offsetPoint1, offsetPoint2, curve1.ReferenceLevel.Id);
                
                // Trim curve1 to the offset point
                (curve1.Location as LocationCurve).Curve = Line.CreateBound(line1.GetEndPoint(0), offsetPoint1);
                
                // Create elbow fittings to connect the conduits
                CreateElbowFitting(curve1, offsetConduit);
                CreateElbowFitting(offsetConduit, curve2);
                
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Offset Connection Error", $"Failed to create offset connection: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// Builder for creating trim connections between non-parallel conduits with no offset
    /// </summary>
    public class TrimBuilder : BuilderBase
    {
        public TrimBuilder(Document doc) : base(doc) { }

        public override bool Build(MEPCurvePair curvePair)
        {
            try
            {
                MEPCurve curve1 = curvePair.Curve1;
                MEPCurve curve2 = curvePair.Curve2;

                // Get the line representations
                Line line1 = (curve1.Location as LocationCurve)?.Curve as Line;
                Line line2 = (curve2.Location as LocationCurve)?.Curve as Line;

                if (line1 == null || line2 == null)
                    return false;

                // If the curves have an intersection point, trim them to that point
                if (curvePair.IntersectionPoint != null)
                {
                    // Trim curve1 to the intersection point
                    (curve1.Location as LocationCurve).Curve = Line.CreateBound(line1.GetEndPoint(0), curvePair.IntersectionPoint);
                    
                    // Trim curve2 to the intersection point
                    (curve2.Location as LocationCurve).Curve = Line.CreateBound(curvePair.IntersectionPoint, line2.GetEndPoint(1));
                    
                    // Create elbow fitting to connect the conduits
                    CreateElbowFitting(curve1, curve2);
                    
                    return true;
                }
                else
                {
                    // If there's no intersection, find the closest points and create a connection
                    XYZ closestPoint1, closestPoint2;
                    line1.GetClosestPoints(line2, out closestPoint1, out closestPoint2);
                    
                    // Trim curve1 to the closest point
                    (curve1.Location as LocationCurve).Curve = Line.CreateBound(line1.GetEndPoint(0), closestPoint1);
                    
                    // Trim curve2 to the closest point
                    (curve2.Location as LocationCurve).Curve = Line.CreateBound(closestPoint2, line2.GetEndPoint(1));
                    
                    // Create a small conduit to connect the closest points if needed
                    if (closestPoint1.DistanceTo(closestPoint2) > 0.001)
                    {
                        Conduit connectingConduit = Conduit.Create(_doc, curve1.GetTypeId(), closestPoint1, closestPoint2, curve1.ReferenceLevel.Id);
                        CreateElbowFitting(curve1, connectingConduit);
                        CreateElbowFitting(connectingConduit, curve2);
                    }
                    else
                    {
                        CreateElbowFitting(curve1, curve2);
                    }
                    
                    return true;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Trim Connection Error", $"Failed to create trim connection: {ex.Message}");
                return false;
            }
        }
    }
}
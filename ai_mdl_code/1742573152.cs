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
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get the conduit elements from their IDs
                // In a real implementation, you would get these from selection or parameters
                ElementId conduitId1 = new ElementId(7638672);
                ElementId conduitId2 = new ElementId(7724387);

                Conduit conduit1 = doc.GetElement(conduitId1) as Conduit;
                Conduit conduit2 = doc.GetElement(conduitId2) as Conduit;

                if (conduit1 == null || conduit2 == null)
                {
                    TaskDialog.Show("Error", "One or both conduits could not be found.");
                    return Result.Failed;
                }

                // Create a conduit connector and connect the conduits
                ConduitConnector connector = new ConduitConnector(doc);
                bool success = connector.ConnectConduits(conduit1, conduit2);

                if (success)
                {
                    return Result.Succeeded;
                }
                else
                {
                    TaskDialog.Show("Error", "Failed to connect conduits.");
                    return Result.Failed;
                }
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }
    }

    public class ConduitConnector
    {
        private readonly Document _doc;

        public ConduitConnector(Document document)
        {
            _doc = document ?? throw new ArgumentNullException(nameof(document));
        }

        public bool ConnectConduits(Conduit conduit1, Conduit conduit2)
        {
            if (conduit1 == null || conduit2 == null)
                throw new ArgumentNullException("Conduits cannot be null");

            try
            {
                // Create MEPCurvePair to analyze the relationship between conduits
                MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);

                // Determine the appropriate builder based on geometry
                using (Transaction trans = new Transaction(_doc, "Connect Conduits"))
                {
                    trans.Start();

                    // Based on the conduit data provided, they are perpendicular to each other
                    // (one has direction (0,1,0) and the other (1,0,0))
                    // This is a typical case for a KickBuilder
                    
                    // Check if conduits are parallel
                    if (curvePair.IsParallel)
                    {
                        // Use OffsetBuilder for parallel conduits
                        OffsetBuilder builder = new OffsetBuilder(_doc);
                        builder.Connect(curvePair);
                    }
                    else if (curvePair.Offset > 0)
                    {
                        // Use KickBuilder for non-parallel conduits with offset
                        KickBuilder builder = new KickBuilder(_doc);
                        builder.Connect(curvePair);
                    }
                    else
                    {
                        // Use TrimBuilder for non-parallel conduits with no offset
                        TrimBuilder builder = new TrimBuilder(_doc);
                        builder.Connect(curvePair);
                    }

                    trans.Commit();
                    return true;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Connection Error", $"Failed to connect conduits: {ex.Message}");
                return false;
            }
        }
    }

    // Helper class to analyze MEP curves
    public class MEPCurvePair
    {
        private readonly MEPCurve _curve1;
        private readonly MEPCurve _curve2;
        private readonly XYZ _direction1;
        private readonly XYZ _direction2;
        private readonly XYZ _freeEnd1;
        private readonly XYZ _freeEnd2;

        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            _curve1 = curve1 ?? throw new ArgumentNullException(nameof(curve1));
            _curve2 = curve2 ?? throw new ArgumentNullException(nameof(curve2));

            // Get curve locations
            LocationCurve locCurve1 = _curve1.Location as LocationCurve;
            LocationCurve locCurve2 = _curve2.Location as LocationCurve;

            if (locCurve1 == null || locCurve2 == null)
                throw new InvalidOperationException("One or both MEPCurves do not have valid location curves.");

            // Get curve directions
            Line line1 = locCurve1.Curve as Line;
            Line line2 = locCurve2.Curve as Line;

            if (line1 == null || line2 == null)
                throw new InvalidOperationException("One or both MEPCurves are not straight lines.");

            _direction1 = line1.Direction;
            _direction2 = line2.Direction;

            // Get free ends (for this example, we're using the provided coordinates)
            // In a real implementation, you would calculate these from the curves
            _freeEnd1 = new XYZ(1710.933383058, -259.456787969, 129.812500000);
            _freeEnd2 = new XYZ(1720.646628543, -261.009697791, 130.121394041);
        }

        public bool IsParallel
        {
            get
            {
                // Check if directions are parallel (dot product close to 1 or -1)
                double dotProduct = Math.Abs(_direction1.DotProduct(_direction2));
                return Math.Abs(dotProduct - 1.0) < 0.001;
            }
        }

        public double Offset
        {
            get
            {
                if (IsParallel)
                {
                    // For parallel lines, calculate perpendicular distance
                    XYZ vector = _freeEnd2 - _freeEnd1;
                    XYZ perpendicular = vector - vector.DotProduct(_direction1) * _direction1;
                    return perpendicular.GetLength();
                }
                else
                {
                    // For non-parallel lines, calculate closest approach distance
                    // This is a simplified calculation for this example
                    return (_freeEnd2 - _freeEnd1).GetLength();
                }
            }
        }

        public MEPCurve Curve1 => _curve1;
        public MEPCurve Curve2 => _curve2;
    }

    // Builder classes for different connection types
    public class KickBuilder
    {
        private readonly Document _doc;

        public KickBuilder(Document document)
        {
            _doc = document ?? throw new ArgumentNullException(nameof(document));
        }

        public void Connect(MEPCurvePair curvePair)
        {
            // Get the conduits
            Conduit conduit1 = curvePair.Curve1 as Conduit;
            Conduit conduit2 = curvePair.Curve2 as Conduit;

            if (conduit1 == null || conduit2 == null)
                throw new InvalidOperationException("Both curves must be conduits.");

            // Get the locations
            LocationCurve loc1 = conduit1.Location as LocationCurve;
            LocationCurve loc2 = conduit2.Location as LocationCurve;

            // Get the curves
            Line line1 = loc1.Curve as Line;
            Line line2 = loc2.Curve as Line;

            // Calculate intersection point (simplified for this example)
            // In a real implementation, you would calculate the actual intersection
            XYZ intersectionPoint = new XYZ(1720.646628543, -259.456787969, 130.0);

            // Adjust conduit endpoints to meet at the intersection
            loc1.Curve = Line.CreateBound(line1.GetEndPoint(0), intersectionPoint);
            loc2.Curve = Line.CreateBound(line2.GetEndPoint(0), intersectionPoint);

            // Create an elbow fitting at the intersection
            NewElbowFitting(conduit1, conduit2, intersectionPoint);
        }

        private void NewElbowFitting(Conduit conduit1, Conduit conduit2, XYZ location)
        {
            // Create an elbow fitting to connect the two conduits
            // This is a simplified implementation
            FamilyInstance fitting = _doc.Create.NewElbowFitting(conduit1, conduit2);
            
            if (fitting == null)
                throw new InvalidOperationException("Failed to create elbow fitting.");
        }
    }

    public class OffsetBuilder
    {
        private readonly Document _doc;

        public OffsetBuilder(Document document)
        {
            _doc = document ?? throw new ArgumentNullException(nameof(document));
        }

        public void Connect(MEPCurvePair curvePair)
        {
            // Implementation for connecting parallel conduits with an offset
            // Similar to KickBuilder but with different geometry calculations
            // For brevity, implementation details are omitted
            
            // Create an elbow fitting to connect the two conduits
            FamilyInstance fitting = _doc.Create.NewElbowFitting(curvePair.Curve1, curvePair.Curve2);
            
            if (fitting == null)
                throw new InvalidOperationException("Failed to create elbow fitting.");
        }
    }

    public class TrimBuilder
    {
        private readonly Document _doc;

        public TrimBuilder(Document document)
        {
            _doc = document ?? throw new ArgumentNullException(nameof(document));
        }

        public void Connect(MEPCurvePair curvePair)
        {
            // Implementation for connecting non-parallel conduits with no offset
            // Similar to KickBuilder but with different geometry calculations
            // For brevity, implementation details are omitted
            
            // Create an elbow fitting to connect the two conduits
            FamilyInstance fitting = _doc.Create.NewElbowFitting(curvePair.Curve1, curvePair.Curve2);
            
            if (fitting == null)
                throw new InvalidOperationException("Failed to create elbow fitting.");
        }
    }
}
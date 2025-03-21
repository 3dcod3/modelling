using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ConduitConnector
{
    [Transaction(TransactionMode.Manual)]
    public class KickConnectionCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            try
            {
                // Select the first conduit
                Reference firstConduitRef = uidoc.Selection.PickObject(ObjectType.Element, 
                    new MEPCurveFilter("Select first conduit"));
                Conduit firstConduit = doc.GetElement(firstConduitRef) as Conduit;
                
                if (firstConduit == null)
                {
                    message = "The selected element is not a conduit.";
                    return Result.Failed;
                }
                
                // Select the second conduit
                Reference secondConduitRef = uidoc.Selection.PickObject(ObjectType.Element, 
                    new MEPCurveFilter("Select second conduit"));
                Conduit secondConduit = doc.GetElement(secondConduitRef) as Conduit;
                
                if (secondConduit == null)
                {
                    message = "The selected element is not a conduit.";
                    return Result.Failed;
                }
                
                // Create MEPCurvePair to analyze the relationship between conduits
                MEPCurvePair conduitPair = new MEPCurvePair(firstConduit, secondConduit);
                
                // Check if the conduits are suitable for a kick connection
                if (conduitPair.IsParallel)
                {
                    TaskDialog.Show("Error", "Selected conduits are parallel. Kick connections require non-parallel conduits.");
                    return Result.Failed;
                }
                
                // Create a kick connection
                using (Transaction trans = new Transaction(doc, "Create Kick Connection"))
                {
                    trans.Start();
                    
                    try
                    {
                        // Create a kick builder and execute the connection
                        KickBuilder builder = new KickBuilder(doc);
                        builder.SourceCurve = firstConduit;
                        builder.TargetCurve = secondConduit;
                        
                        // Set the kick distance (offset)
                        double kickDistance = conduitPair.Offset > 0 ? conduitPair.Offset : 12.0; // Default to 12 inches if no offset
                        builder.KickDistance = kickDistance;
                        
                        // Execute the kick connection
                        bool result = builder.Execute();
                        
                        if (!result)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create kick connection.");
                            return Result.Failed;
                        }
                        
                        trans.Commit();
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        message = ex.Message;
                        return Result.Failed;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
    
    /// <summary>
    /// Filter for selecting MEP curves (conduits, pipes, ducts, etc.)
    /// </summary>
    public class MEPCurveFilter : ISelectionFilter
    {
        private string _prompt;
        
        public MEPCurveFilter(string prompt)
        {
            _prompt = prompt;
        }
        
        public bool AllowElement(Element elem)
        {
            return elem is MEPCurve;
        }
        
        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
    
    /// <summary>
    /// Class to analyze the relationship between two MEP curves
    /// </summary>
    public class MEPCurvePair
    {
        private MEPCurve _curve1;
        private MEPCurve _curve2;
        private bool _isParallel;
        private double _offset;
        
        public MEPCurvePair(MEPCurve curve1, MEPCurve curve2)
        {
            _curve1 = curve1;
            _curve2 = curve2;
            AnalyzeRelationship();
        }
        
        public MEPCurve Curve1 => _curve1;
        public MEPCurve Curve2 => _curve2;
        public bool IsParallel => _isParallel;
        public double Offset => _offset;
        
        private void AnalyzeRelationship()
        {
            // Get the curves from the MEP curves
            Curve c1 = GetCurveFromMEPCurve(_curve1);
            Curve c2 = GetCurveFromMEPCurve(_curve2);
            
            if (c1 is Line && c2 is Line)
            {
                Line line1 = c1 as Line;
                Line line2 = c2 as Line;
                
                // Check if the lines are parallel
                XYZ dir1 = line1.Direction;
                XYZ dir2 = line2.Direction;
                
                // Normalize directions
                dir1 = dir1.Normalize();
                dir2 = dir2.Normalize();
                
                // Check if directions are parallel or anti-parallel
                double dotProduct = dir1.DotProduct(dir2);
                _isParallel = Math.Abs(Math.Abs(dotProduct) - 1.0) < 0.001;
                
                if (_isParallel)
                {
                    // Calculate the offset between parallel lines
                    XYZ p1 = line1.GetEndPoint(0);
                    XYZ v = dir1.CrossProduct(dir1.CrossProduct(dir2));
                    v = v.Normalize();
                    
                    // Project a point from line1 onto line2
                    XYZ p2 = line2.GetEndPoint(0);
                    XYZ vec = p2 - p1;
                    _offset = Math.Abs(vec.DotProduct(v));
                }
                else
                {
                    // For non-parallel lines, find the closest distance between them
                    XYZ p1 = line1.GetEndPoint(0);
                    XYZ p2 = line2.GetEndPoint(0);
                    XYZ v1 = line1.Direction;
                    XYZ v2 = line2.Direction;
                    
                    // Calculate the closest points between the two lines
                    XYZ crossProduct = v1.CrossProduct(v2);
                    double denominator = crossProduct.DotProduct(crossProduct);
                    
                    if (denominator < 0.001)
                    {
                        // Lines are nearly parallel, use point-to-line distance
                        XYZ vec = p2 - p1;
                        XYZ crossVec = v1.CrossProduct(vec);
                        _offset = crossVec.GetLength();
                    }
                    else
                    {
                        // Calculate the parameters for the closest points
                        XYZ p2MinusP1 = p2 - p1;
                        double t1 = (p2MinusP1.CrossProduct(v2)).DotProduct(crossProduct) / denominator;
                        double t2 = (p2MinusP1.CrossProduct(v1)).DotProduct(crossProduct) / denominator;
                        
                        // Calculate the closest points
                        XYZ closestPoint1 = p1 + t1 * v1;
                        XYZ closestPoint2 = p2 + t2 * v2;
                        
                        // Calculate the distance between the closest points
                        _offset = closestPoint1.DistanceTo(closestPoint2);
                    }
                }
            }
            else
            {
                // For non-line curves, use a simplified approach
                _isParallel = false;
                _offset = 0;
            }
        }
        
        private Curve GetCurveFromMEPCurve(MEPCurve mepCurve)
        {
            LocationCurve locationCurve = mepCurve.Location as LocationCurve;
            return locationCurve?.Curve;
        }
    }
    
    /// <summary>
    /// Builder class for creating kick connections between conduits
    /// </summary>
    public class KickBuilder
    {
        private Document _doc;
        private MEPCurve _sourceCurve;
        private MEPCurve _targetCurve;
        private double _kickDistance = 12.0; // Default kick distance (1 foot)
        
        public KickBuilder(Document doc)
        {
            _doc = doc;
        }
        
        public MEPCurve SourceCurve
        {
            get { return _sourceCurve; }
            set { _sourceCurve = value; }
        }
        
        public MEPCurve TargetCurve
        {
            get { return _targetCurve; }
            set { _targetCurve = value; }
        }
        
        public double KickDistance
        {
            get { return _kickDistance; }
            set { _kickDistance = value; }
        }
        
        public bool Execute()
        {
            if (_sourceCurve == null || _targetCurve == null)
            {
                return false;
            }
            
            try
            {
                // Get the curves from the MEP curves
                LocationCurve sourceLocation = _sourceCurve.Location as LocationCurve;
                LocationCurve targetLocation = _targetCurve.Location as LocationCurve;
                
                if (sourceLocation == null || targetLocation == null)
                {
                    return false;
                }
                
                Curve sourceCurve = sourceLocation.Curve;
                Curve targetCurve = targetLocation.Curve;
                
                if (!(sourceCurve is Line) || !(targetCurve is Line))
                {
                    return false; // Only support straight conduits for now
                }
                
                Line sourceLine = sourceCurve as Line;
                Line targetLine = targetLine as Line;
                
                // Calculate intersection point (in 3D space or projected to a plane)
                XYZ sourceDir = sourceLine.Direction;
                XYZ targetDir = targetLine.Direction;
                
                // Get the source and target points
                XYZ sourceStart = sourceLine.GetEndPoint(0);
                XYZ sourceEnd = sourceLine.GetEndPoint(1);
                XYZ targetStart = targetLine.GetEndPoint(0);
                XYZ targetEnd = targetLine.GetEndPoint(1);
                
                // Find the closest points between the two lines
                XYZ closestPointOnSource, closestPointOnTarget;
                FindClosestPoints(sourceLine, targetLine, out closestPointOnSource, out closestPointOnTarget);
                
                // Calculate the kick direction (perpendicular to source line)
                XYZ kickDir = sourceDir.CrossProduct(XYZ.BasisZ);
                if (kickDir.GetLength() < 0.001)
                {
                    // If source is vertical, use a different approach
                    kickDir = sourceDir.CrossProduct(XYZ.BasisX);
                }
                kickDir = kickDir.Normalize();
                
                // Determine which direction to kick based on the target line position
                XYZ vectorToTarget = closestPointOnTarget - closestPointOnSource;
                if (vectorToTarget.DotProduct(kickDir) < 0)
                {
                    kickDir = -kickDir;
                }
                
                // Calculate the kick point on the source line
                XYZ kickPoint = closestPointOnSource;
                
                // Create a new conduit for the kick
                Conduit sourceCopy = ElementTransformUtils.CopyElement(
                    _doc, _sourceCurve.Id, XYZ.Zero).FirstOrDefault() as Conduit;
                
                if (sourceCopy == null)
                {
                    return false;
                }
                
                // Trim the source conduit at the kick point
                LocationCurve sourceCopyLocation = sourceCopy.Location as LocationCurve;
                Line newSourceLine = Line.CreateBound(sourceStart, kickPoint);
                sourceCopyLocation.Curve = newSourceLine;
                
                // Create the horizontal segment of the kick
                XYZ kickEnd = kickPoint + kickDir * _kickDistance;
                Conduit horizontalSegment = Conduit.Create(
                    _doc, 
                    _sourceCurve.GetTypeId(), 
                    _sourceCurve.LevelId, 
                    kickPoint, 
                    kickEnd);
                
                // Copy properties from source to new conduit
                CopyConduitProperties(_sourceCurve, horizontalSegment);
                
                // Create the final segment connecting to the target
                XYZ finalEnd = closestPointOnTarget;
                Conduit finalSegment = Conduit.Create(
                    _doc, 
                    _sourceCurve.GetTypeId(), 
                    _sourceCurve.LevelId, 
                    kickEnd, 
                    finalEnd);
                
                // Copy properties from source to new conduit
                CopyConduitProperties(_sourceCurve, finalSegment);
                
                // Create elbow fittings at the connection points
                CreateElbowFitting(_doc, sourceCopy, horizontalSegment);
                CreateElbowFitting(_doc, horizontalSegment, finalSegment);
                CreateElbowFitting(_doc, finalSegment, _targetCurve);
                
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        
        private void FindClosestPoints(Line line1, Line line2, out XYZ point1, out XYZ point2)
        {
            XYZ p1 = line1.GetEndPoint(0);
            XYZ p2 = line2.GetEndPoint(0);
            XYZ v1 = line1.Direction;
            XYZ v2 = line2.Direction;
            
            // Calculate the closest points between the two lines
            XYZ crossProduct = v1.CrossProduct(v2);
            double denominator = crossProduct.DotProduct(crossProduct);
            
            if (denominator < 0.001)
            {
                // Lines are nearly parallel, use endpoints
                point1 = line1.GetEndPoint(1);
                point2 = line2.GetEndPoint(0);
                return;
            }
            
            // Calculate the parameters for the closest points
            XYZ p2MinusP1 = p2 - p1;
            double t1 = (p2MinusP1.CrossProduct(v2)).DotProduct(crossProduct) / denominator;
            double t2 = (p2MinusP1.CrossProduct(v1)).DotProduct(crossProduct) / denominator;
            
            // Clamp parameters to line segments
            t1 = Math.Max(0, Math.Min(1, t1));
            t2 = Math.Max(0, Math.Min(1, t2));
            
            // Calculate the closest points
            point1 = p1 + t1 * v1;
            point2 = p2 + t2 * v2;
        }
        
        private void CopyConduitProperties(Conduit source, Conduit target)
        {
            // Copy diameter, material, etc.
            Parameter diamParam = source.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
            if (diamParam != null)
            {
                Parameter targetDiamParam = target.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                if (targetDiamParam != null)
                {
                    targetDiamParam.Set(diamParam.AsDouble());
                }
            }
            
            // Copy other parameters as needed
        }
        
        private void CreateElbowFitting(Document doc, MEPCurve curve1, MEPCurve curve2)
        {
            // Get connectors from the curves
            Connector connector1 = GetNearestConnector(curve1, curve2);
            Connector connector2 = GetNearestConnector(curve2, curve1);
            
            if (connector1 != null && connector2 != null)
            {
                // Create the elbow fitting
                doc.Create.NewElbowFitting(connector1, connector2);
            }
        }
        
        private Connector GetNearestConnector(MEPCurve curve, MEPCurve otherCurve)
        {
            ConnectorSet connectors = curve.ConnectorManager.Connectors;
            Connector nearestConnector = null;
            double minDistance = double.MaxValue;
            
            foreach (Connector connector in connectors)
            {
                // Get the nearest connector from the other curve
                ConnectorSet otherConnectors = otherCurve.ConnectorManager.Connectors;
                foreach (Connector otherConnector in otherConnectors)
                {
                    double distance = connector.Origin.DistanceTo(otherConnector.Origin);
                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        nearestConnector = connector;
                    }
                }
            }
            
            return nearestConnector;
        }
    }
}
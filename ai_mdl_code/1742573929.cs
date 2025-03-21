using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
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
                // Get the conduit elements by their IDs
                ElementId conduitId1 = new ElementId(7638672);
                ElementId conduitId2 = new ElementId(7724387);

                Conduit conduit1 = doc.GetElement(conduitId1) as Conduit;
                Conduit conduit2 = doc.GetElement(conduitId2) as Conduit;

                if (conduit1 == null || conduit2 == null)
                {
                    TaskDialog.Show("Error", "One or both conduits could not be found.");
                    return Result.Failed;
                }

                // Create a MEPCurvePair to analyze the relationship between the conduits
                MEPCurvePair curvePair = new MEPCurvePair(conduit1, conduit2);
                
                // Connect the conduits using the appropriate builder
                using (Transaction trans = new Transaction(doc, "Connect Conduits"))
                {
                    trans.Start();
                    
                    bool success = ConnectConduits(doc, curvePair);
                    
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
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }

        private bool ConnectConduits(Document doc, MEPCurvePair curvePair)
        {
            // Extract the conduits from the pair
            Conduit conduit1 = curvePair.FirstCurve as Conduit;
            Conduit conduit2 = curvePair.SecondCurve as Conduit;

            if (conduit1 == null || conduit2 == null)
                return false;

            // Get the location curves
            LocationCurve locCurve1 = conduit1.Location as LocationCurve;
            LocationCurve locCurve2 = conduit2.Location as LocationCurve;

            if (locCurve1 == null || locCurve2 == null)
                return false;

            // Get the curves
            Curve curve1 = locCurve1.Curve;
            Curve curve2 = locCurve2.Curve;

            // Get the directions
            XYZ dir1 = (curve1.GetEndPoint(1) - curve1.GetEndPoint(0)).Normalize();
            XYZ dir2 = (curve2.GetEndPoint(1) - curve2.GetEndPoint(0)).Normalize();

            // Determine if the conduits are parallel
            bool isParallel = dir1.IsAlmostEqualTo(dir2) || dir1.IsAlmostEqualTo(dir2.Negate());

            // Get the free ends of the conduits
            XYZ freeEnd1 = new XYZ(1719.063295210, -259.456787969, 129.812500000);
            XYZ freeEnd2 = new XYZ(1720.646628543, -261.009697791, 130.121394041);

            // Calculate the offset between the conduits
            XYZ offset = freeEnd2 - freeEnd1;
            double offsetDistance = offset.GetLength();

            // Determine which builder to use based on geometry
            if (isParallel && offsetDistance > 0.01)
            {
                // Use OffsetBuilder for parallel conduits with an offset
                return ConnectWithOffsetBuilder(doc, conduit1, conduit2, freeEnd1, freeEnd2);
            }
            else if (!isParallel && offsetDistance > 0.01)
            {
                // Use KickBuilder for non-parallel conduits with an offset
                return ConnectWithKickBuilder(doc, conduit1, conduit2, freeEnd1, freeEnd2);
            }
            else if (!isParallel && offsetDistance <= 0.01)
            {
                // Use TrimBuilder for non-parallel conduits with no significant offset
                return ConnectWithTrimBuilder(doc, conduit1, conduit2, freeEnd1, freeEnd2);
            }
            else
            {
                // Default to KickBuilder if we can't determine the appropriate builder
                return ConnectWithKickBuilder(doc, conduit1, conduit2, freeEnd1, freeEnd2);
            }
        }

        private bool ConnectWithKickBuilder(Document doc, Conduit conduit1, Conduit conduit2, XYZ freeEnd1, XYZ freeEnd2)
        {
            try
            {
                // Create a KickBuilder for the conduits
                KickBuilder builder = new KickBuilder(doc);
                
                // Set the conduits to connect
                builder.AddCurve(conduit1);
                builder.AddCurve(conduit2);
                
                // Set the connection points
                builder.SetConnectionPoint(conduit1, freeEnd1);
                builder.SetConnectionPoint(conduit2, freeEnd2);
                
                // Build the connection
                builder.Build();
                
                // Create an elbow fitting at the connection
                NewElbowFitting(doc, conduit1, conduit2);
                
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to connect with KickBuilder: {ex.Message}");
                return false;
            }
        }

        private bool ConnectWithOffsetBuilder(Document doc, Conduit conduit1, Conduit conduit2, XYZ freeEnd1, XYZ freeEnd2)
        {
            try
            {
                // Create an OffsetBuilder for the conduits
                OffsetBuilder builder = new OffsetBuilder(doc);
                
                // Set the conduits to connect
                builder.AddCurve(conduit1);
                builder.AddCurve(conduit2);
                
                // Set the connection points
                builder.SetConnectionPoint(conduit1, freeEnd1);
                builder.SetConnectionPoint(conduit2, freeEnd2);
                
                // Build the connection
                builder.Build();
                
                // Create an elbow fitting at the connection
                NewElbowFitting(doc, conduit1, conduit2);
                
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to connect with OffsetBuilder: {ex.Message}");
                return false;
            }
        }

        private bool ConnectWithTrimBuilder(Document doc, Conduit conduit1, Conduit conduit2, XYZ freeEnd1, XYZ freeEnd2)
        {
            try
            {
                // Create a TrimBuilder for the conduits
                TrimBuilder builder = new TrimBuilder(doc);
                
                // Set the conduits to connect
                builder.AddCurve(conduit1);
                builder.AddCurve(conduit2);
                
                // Set the connection points
                builder.SetConnectionPoint(conduit1, freeEnd1);
                builder.SetConnectionPoint(conduit2, freeEnd2);
                
                // Build the connection
                builder.Build();
                
                // Create an elbow fitting at the connection
                NewElbowFitting(doc, conduit1, conduit2);
                
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to connect with TrimBuilder: {ex.Message}");
                return false;
            }
        }

        private void NewElbowFitting(Document doc, Conduit conduit1, Conduit conduit2)
        {
            try
            {
                // Get the connector manager for each conduit
                ConnectorManager cm1 = conduit1.ConnectorManager;
                ConnectorManager cm2 = conduit2.ConnectorManager;
                
                if (cm1 == null || cm2 == null)
                    return;
                
                // Find the unconnected connectors
                Connector conn1 = null;
                Connector conn2 = null;
                
                foreach (Connector c in cm1.Connectors)
                {
                    if (c.IsConnected == false)
                    {
                        conn1 = c;
                        break;
                    }
                }
                
                foreach (Connector c in cm2.Connectors)
                {
                    if (c.IsConnected == false)
                    {
                        conn2 = c;
                        break;
                    }
                }
                
                if (conn1 == null || conn2 == null)
                    return;
                
                // Create the elbow fitting
                doc.Create.NewElbowFitting(conn1, conn2);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create elbow fitting: {ex.Message}");
            }
        }
    }
}
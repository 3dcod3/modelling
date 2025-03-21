using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public class ConduitConnector
{
    private readonly Document _doc;
    private readonly Conduit _conduit1;
    private readonly Conduit _conduit2;
    
    public ConduitConnector(Document document, Conduit conduit1, Conduit conduit2)
    {
        _doc = document ?? throw new ArgumentNullException(nameof(document));
        _conduit1 = conduit1 ?? throw new ArgumentNullException(nameof(conduit1));
        _conduit2 = conduit2 ?? throw new ArgumentNullException(nameof(conduit2));
    }
    
    public Result Execute()
    {
        string message = string.Empty;
        
        try
        {
            // Get the connectors from both conduits
            ConnectorSet connectors1 = _conduit1.MEPModel.ConnectorManager.Connectors;
            ConnectorSet connectors2 = _conduit2.MEPModel.ConnectorManager.Connectors;
            
            // Find the free connectors (the ones that need to be connected)
            Connector freeConnector1 = GetFreeConnector(connectors1);
            Connector freeConnector2 = GetFreeConnector(connectors2);
            
            if (freeConnector1 == null || freeConnector2 == null)
            {
                TaskDialog.Show("Error", "Could not find free connectors on both conduits.");
                return Result.Failed;
            }
            
            // Calculate the angle between the conduits
            double angle = CalculateAngleBetweenVectors(
                freeConnector1.CoordinateSystem.BasisZ, 
                freeConnector2.CoordinateSystem.BasisZ);
            
            using (Transaction trans = new Transaction(_doc, "Connect Conduits with Elbow"))
            {
                trans.Start();
                
                // Create an elbow fitting between the two conduits
                FamilyInstance elbow = _doc.Create.NewElbowFitting(freeConnector1, freeConnector2);
                
                if (elbow != null)
                {
                    trans.Commit();
                    return Result.Succeeded;
                }
                else
                {
                    trans.RollBack();
                    TaskDialog.Show("Error", "Failed to create elbow fitting.");
                    return Result.Failed;
                }
            }
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            TaskDialog.Show("Error", message);
            return Result.Failed;
        }
    }
    
    private Connector GetFreeConnector(ConnectorSet connectors)
    {
        foreach (Connector connector in connectors)
        {
            // A free connector is one that is not connected to anything
            if (connector.IsConnected == false)
            {
                return connector;
            }
        }
        return null;
    }
    
    private double CalculateAngleBetweenVectors(XYZ vector1, XYZ vector2)
    {
        // Calculate the dot product and magnitudes
        double dotProduct = vector1.DotProduct(vector2);
        double magnitude1 = vector1.GetLength();
        double magnitude2 = vector2.GetLength();
        
        // Calculate the angle in radians
        double angleInRadians = Math.Acos(dotProduct / (magnitude1 * magnitude2));
        
        // Convert to degrees
        double angleInDegrees = angleInRadians * (180.0 / Math.PI);
        
        return angleInDegrees;
    }
    
    // Main method to be called from external command
    public static Result ConnectConduits(Document doc, ElementId conduitId1, ElementId conduitId2)
    {
        // Get the conduit elements from their IDs
        Conduit conduit1 = doc.GetElement(conduitId1) as Conduit;
        Conduit conduit2 = doc.GetElement(conduitId2) as Conduit;
        
        if (conduit1 == null || conduit2 == null)
        {
            TaskDialog.Show("Error", "One or both conduits could not be found.");
            return Result.Failed;
        }
        
        ConduitConnector connector = new ConduitConnector(doc, conduit1, conduit2);
        return connector.Execute();
    }
}

// Example usage in an external command:
public class ConnectConduitsCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        // For this example, we'll assume the conduit IDs are provided
        // In a real application, you might get these from user selection
        ElementId conduitId1 = new ElementId(7638682); // ID of first conduit
        ElementId conduitId2 = new ElementId(7644644); // ID of second conduit
        
        return ConduitConnector.ConnectConduits(doc, conduitId1, conduitId2);
    }
}

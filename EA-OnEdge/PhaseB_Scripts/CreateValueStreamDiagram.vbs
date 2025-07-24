Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Value Stream Diagram
Sub CreateValueStreamDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE VALUE STREAM DIAGRAM"
    Session.Output "======================================="

    ' Find the root model "Model"
    Dim rootModel As EA.Package
    Set rootModel = Nothing
    Dim i
    For i = 0 To Repository.Models.Count - 1
        Dim tempModel As EA.Package
        Set tempModel = Repository.Models.GetAt(i)
        If tempModel.Name = "Model" Then
            Set rootModel = tempModel
            Exit For
        End If
    Next

    If rootModel Is Nothing Then
        Session.Output "Error: Root model 'Model' not found. Listing available models:"
        For i = 0 To Repository.Models.Count - 1
            Set tempModel = Repository.Models.GetAt(i)
            Session.Output "Model " & i & ": " & tempModel.Name
        Next
        Exit Sub
    End If
    Session.Output "Root model found: " & rootModel.Name

    ' Find the "Architecture Development Method" package
    Dim admPackage As EA.Package
    Set admPackage = findPackageByName(rootModel, "Architecture Development Method")

    If admPackage Is Nothing Then
        Session.Output "Error: Architecture Development Method package not found. Listing available packages:"
        listAllPackages rootModel, 0
        Exit Sub
    End If
    Session.Output "Found package: " & admPackage.Name

    ' Find or create the "Phase B" package
    Dim phaseBPackage As EA.Package
    Set phaseBPackage = findPackageByName(admPackage, "Phase B")
    If phaseBPackage Is Nothing Then
        Set phaseBPackage = admPackage.Packages.AddNew("Phase B", "Package")
        phaseBPackage.Update
        Session.Output "Created package: Phase B"
    End If
    Session.Output "Found package: " & phaseBPackage.Name

    ' Check for existing Value Stream Diagram
    Dim valueStreamDiagram As EA.Diagram
    Set valueStreamDiagram = Nothing
    For i = 0 To phaseBPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseBPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Value Stream" Then
            Set valueStreamDiagram = diag
            Exit For
        End If
    Next

    ' Create Value Stream Diagram if it doesn't exist
    If valueStreamDiagram Is Nothing Then
        Set valueStreamDiagram = phaseBPackage.Diagrams.AddNew("OnEdge Value Stream", "Activity")
        If Not valueStreamDiagram Is Nothing Then
            valueStreamDiagram.Notes = "Value Stream Diagram for OnEdge AI Intelligence Service: SMB Onboarding to Insight Delivery. Participants: SMB Business Managers, IT Teams, Customer Support."
            valueStreamDiagram.Update
            Session.Output "Created diagram: OnEdge Value Stream"
        Else
            Session.Output "Error: Failed to create OnEdge Value Stream diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Value Stream"
    End If

    ' Define elements (value stream stages) with names, stereotypes, colors, and positions
    Dim elements(5)
    Dim stereotypes(5)
    Dim elementTypes(5)
    Dim elementObjects(5)
    Dim elementColors(5)
    Dim elementSizes(5)
    Dim elementCount
    elementCount = 0

    ' Value Stream Stages (Blue)
    elements(0) = "Discovery": stereotypes(0) = "ValueStreamStage": elementTypes(0) = "Activity": elementColors(0) = RGB(0, 102, 204): elementSizes(0) = "l=50;r=250;t=100;b=50"
    elementCount = elementCount + 1
    elements(1) = "Onboarding": stereotypes(1) = "ValueStreamStage": elementTypes(1) = "Activity": elementColors(1) = RGB(0, 102, 204): elementSizes(1) = "l=300;r=500;t=100;b=50"
    elementCount = elementCount + 1
    elements(2) = "Integration": stereotypes(2) = "ValueStreamStage": elementTypes(2) = "Activity": elementColors(2) = RGB(0, 102, 204): elementSizes(2) = "l=550;r=750;t=100;b=50"
    elementCount = elementCount + 1
    elements(3) = "App Selection": stereotypes(3) = "ValueStreamStage": elementTypes(3) = "Activity": elementColors(3) = RGB(0, 102, 204): elementSizes(3) = "l=800;r=1000;t=100;b=50"
    elementCount = elementCount + 1
    elements(4) = "Insight Delivery": stereotypes(4) = "ValueStreamStage": elementTypes(4) = "Activity": elementColors(4) = RGB(0, 102, 204): elementSizes(4) = "l=1050;r=1250;t=100;b=50"
    elementCount = elementCount + 1
    elements(5) = "Support": stereotypes(5) = "ValueStreamStage": elementTypes(5) = "Activity": elementColors(5) = RGB(0, 102, 204): elementSizes(5) = "l=1300;r=1500;t=100;b=50"
    elementCount = elementCount + 1

    ' Create elements in the package
    For i = 0 To elementCount - 1
        Dim elem As EA.Element
        Set elem = phaseBPackage.Elements.AddNew(elements(i), elementTypes(i))
        If Not elem Is Nothing Then
            elem.Stereotype = stereotypes(i)
            elem.Update
            Set elementObjects(i) = elem
            Session.Output "Created element: " & elements(i)
        Else
            Session.Output "Error: Failed to create element: " & elements(i)
            hasError = True
            Exit Sub
        End If
    Next

    ' Add elements to the diagram with positioning and colors
    Dim diagObjects As EA.Collection
    Set diagObjects = valueStreamDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations to represent value stream flow (solid lines)
    Dim connectors As EA.Collection
    Set connectors = phaseBPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = valueStreamDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Discovery flows to Onboarding
    Set conn = connectors.AddNew("", "ControlFlow")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(1).ElementID
    conn.Name = "Flows To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Discovery -> Onboarding (Flows To)"

    ' Onboarding flows to Integration
    Set conn = connectors.AddNew("", "ControlFlow")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Flows To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Onboarding -> Integration (Flows To)"

    ' Integration flows to App Selection
    Set conn = connectors.AddNew("", "ControlFlow")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Flows To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Integration -> App Selection (Flows To)"

    ' App Selection flows to Insight Delivery
    Set conn = connectors.AddNew("", "ControlFlow")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Flows To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: App Selection -> Insight Delivery (Flows To)"

    ' Insight Delivery flows to Support
    Set conn = connectors.AddNew("", "ControlFlow")
    conn.SupplierID = elementObjects(4).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Flows To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Insight Delivery -> Support (Flows To)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseBPackage.PackageID
    valueStreamDiagram.Update
    Repository.ReloadDiagram valueStreamDiagram.DiagramID
    Session.Output "OnEdge Value Stream created successfully under MODEL > Architecture Development Method > Phase B"
    Session.Output "Done!"
End Sub

' Recursive function to find a package by name
Function findPackageByName(parentPackage, targetName)
    Dim i
    Dim pkg
    Set findPackageByName = Nothing
    For i = 0 To parentPackage.Packages.Count - 1
        Set pkg = parentPackage.Packages.GetAt(i)
        Session.Output "Checking package: " & pkg.Name
        If pkg.Name = targetName Then
            Set findPackageByName = pkg
            Exit For
        End If
        Dim subPackage
        Set subPackage = findPackageByName(pkg, targetName)
        If Not subPackage Is Nothing Then
            Set findPackageByName = subPackage
            Exit For
        End If
    Next
End Function

' Function to list all packages and sub-packages for debugging
Sub listAllPackages(parentPackage, level)
    Dim i
    Dim pkg
    For i = 0 To parentPackage.Packages.Count - 1
        Set pkg = parentPackage.Packages.GetAt(i)
        Session.Output String(level * 2, " ") & "Package " & i & ": " & pkg.Name
        listAllPackages pkg, level + 1
    Next
End Sub

' Execute the main function
CreateValueStreamDiagram
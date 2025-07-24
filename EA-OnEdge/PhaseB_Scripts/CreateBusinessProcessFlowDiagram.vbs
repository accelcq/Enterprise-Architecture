Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Business Process Flow Diagram
Sub CreateBusinessProcessFlowDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE BUSINESS PROCESS FLOW DIAGRAM"
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

    ' Check for existing Business Process Flow Diagram
    Dim processDiagram As EA.Diagram
    Set processDiagram = Nothing
    For i = 0 To phaseBPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseBPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Business Process Flow" Then
            Set processDiagram = diag
            Exit For
        End If
    Next

    ' Create Business Process Flow Diagram if it doesn't exist
    If processDiagram Is Nothing Then
        Set processDiagram = phaseBPackage.Diagrams.AddNew("OnEdge Business Process Flow", "Activity")
        If Not processDiagram Is Nothing Then
            processDiagram.Notes = "Business Process Flow Diagram for OnEdge AI Intelligence Service"
            processDiagram.Update
            Session.Output "Created diagram: OnEdge Business Process Flow"
        Else
            Session.Output "Error: Failed to create OnEdge Business Process Flow diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Business Process Flow"
    End If

    ' Define elements (processes) with names, stereotypes, colors, and positions
    Dim elements(3)
    Dim stereotypes(3)
    Dim elementTypes(3)
    Dim elementObjects(3)
    Dim elementColors(3)
    Dim elementSizes(3)
    Dim elementCount
    elementCount = 0

    ' Processes (Blue, consistent with supporting capabilities)
    elements(0) = "AI Inference Process": stereotypes(0) = "Process": elementTypes(0) = "Activity": elementColors(0) = RGB(0, 102, 204): elementSizes(0) = "l=50;r=250;t=100;b=50"
    elementCount = elementCount + 1
    elements(1) = "Payment Processing Process": stereotypes(1) = "Process": elementTypes(1) = "Activity": elementColors(1) = RGB(0, 102, 204): elementSizes(1) = "l=300;r=500;t=100;b=50"
    elementCount = elementCount + 1
    elements(2) = "Compliance Monitoring Process": stereotypes(2) = "Process": elementTypes(2) = "Activity": elementColors(2) = RGB(0, 102, 204): elementSizes(2) = "l=550;r=750;t=100;b=50"
    elementCount = elementCount + 1
    elements(3) = "Marketplace App Deployment Process": stereotypes(3) = "Process": elementTypes(3) = "Activity": elementColors(3) = RGB(0, 102, 204): elementSizes(3) = "l=800;r=1000;t=100;b=50"
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
    Set diagObjects = processDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations to represent process flow (solid lines for sequential flow)
    Dim connectors As EA.Collection
    Set connectors = phaseBPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = processDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' AI Inference Process flows to Marketplace App Deployment Process
    Set conn = connectors.AddNew("", "ControlFlow")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Flows To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: AI Inference Process -> Marketplace App Deployment Process (Flows To)"

    ' Payment Processing Process flows to Marketplace App Deployment Process
    Set conn = connectors.AddNew("", "ControlFlow")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Flows To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Payment Processing Process -> Marketplace App Deployment Process (Flows To)"

    ' Compliance Monitoring Process flows to Marketplace App Deployment Process
    Set conn = connectors.AddNew("", "ControlFlow")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Flows To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Compliance Monitoring Process -> Marketplace App Deployment Process (Flows To)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseBPackage.PackageID
    processDiagram.Update
    Repository.ReloadDiagram processDiagram.DiagramID
    Session.Output "OnEdge Business Process Flow created successfully under MODEL > Architecture Development Method > Phase B"
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
CreateBusinessProcessFlowDiagram
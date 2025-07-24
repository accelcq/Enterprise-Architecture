Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Solution Concept Diagram
Sub CreateSolutionConceptDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE SOLUTION CONCEPT DIAGRAM"
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

    ' Find or create the "Phase A" package
    Dim phaseAPackage As EA.Package
    Set phaseAPackage = findPackageByName(admPackage, "Phase A")
    If phaseAPackage Is Nothing Then
        Set phaseAPackage = admPackage.Packages.AddNew("Phase A", "Package")
        phaseAPackage.Update
        Session.Output "Created package: Phase A"
    End If
    Session.Output "Found package: " & phaseAPackage.Name

    ' Check for existing Solution Concept diagram
    Dim solutionDiagram As EA.Diagram
    Set solutionDiagram = Nothing
    For i = 0 To phaseAPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseAPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Solution Concept" Then
            Set solutionDiagram = diag
            Exit For
        End If
    Next

    ' Create Solution Concept diagram if it doesn't exist
    If solutionDiagram Is Nothing Then
        Set solutionDiagram = phaseAPackage.Diagrams.AddNew("OnEdge Solution Concept", "Class")
        If Not solutionDiagram Is Nothing Then
            solutionDiagram.Notes = "Solution Concept Diagram for OnEdge AI Intelligence Service"
            solutionDiagram.Update
            Session.Output "Created diagram: OnEdge Solution Concept"
        Else
            Session.Output "Error: Failed to create OnEdge Solution Concept diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Solution Concept"
    End If

    ' Define elements (components) with names, stereotypes, colors, and positions
    Dim elements(5)
    Dim stereotypes(5)
    Dim elementTypes(5)
    Dim elementObjects(5)
    Dim elementColors(5)
    Dim elementSizes(5)
    Dim elementCount
    elementCount = 0

    ' Components from OnEdge Solution Concept Diagram
    elements(0) = "SMB Edge Devices": stereotypes(0) = "Component": elementTypes(0) = "Class": elementColors(0) = RGB(255, 204, 102): elementSizes(0) = "l=50;r=200;t=400;b=350" ' Amber for industries
    elementCount = elementCount + 1
    elements(1) = "ONEDGE Micro-Factory": stereotypes(1) = "Component": elementTypes(1) = "Class": elementColors(1) = RGB(153, 102, 255): elementSizes(1) = "l=300;r=500;t=450;b=350" ' Purple for technologies, larger for centrality
    elementCount = elementCount + 1
    elements(2) = "NeoCortex AI": stereotypes(2) = "Component": elementTypes(2) = "Class": elementColors(2) = RGB(153, 102, 255): elementSizes(2) = "l=300;r=450;t=300;b=250" ' Purple for technologies
    elementCount = elementCount + 1
    elements(3) = "KOIN Payment System": stereotypes(3) = "Component": elementTypes(3) = "Class": elementColors(3) = RGB(153, 102, 255): elementSizes(3) = "l=550;r=700;t=400;b=350" ' Purple for technologies
    elementCount = elementCount + 1
    elements(4) = "Starlink Connectivity": stereotypes(4) = "Component": elementTypes(4) = "Class": elementColors(4) = RGB(153, 102, 255): elementSizes(4) = "l=50;r=200;t=200;b=150" ' Purple for technologies
    elementCount = elementCount + 1
    elements(5) = "Cloud": stereotypes(5) = "Component": elementTypes(5) = "Class": elementColors(5) = RGB(102, 204, 255): elementSizes(5) = "l=550;r=700;t=200;b=150" ' Light blue for external systems
    elementCount = elementCount + 1

    ' Create elements in the package
    For i = 0 To elementCount - 1
        Dim elem As EA.Element
        Set elem = phaseAPackage.Elements.AddNew(elements(i), elementTypes(i))
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
    Set diagObjects = solutionDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations based on Solution Concept Diagram
    Dim connectors As EA.Collection
    Set connectors = phaseAPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = solutionDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Data Flow: SMB Edge Devices to Micro-Factory
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(1).ElementID
    conn.Name = "Data Flow"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Flow"

    ' Processes: Micro-Factory to NeoCortex AI
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Processes"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Processes"

    ' Transactions: Micro-Factory to KOIN Payment System
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Transactions"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Transactions"

    ' Connectivity: SMB Edge Devices to Starlink Connectivity
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Connectivity"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Connectivity"

    ' Syncs With: Micro-Factory to Cloud (Dashed line)
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Syncs With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Syncs With (Dashed)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseAPackage.PackageID
    solutionDiagram.Update
    Repository.ReloadDiagram solutionDiagram.DiagramID
    Session.Output "OnEdge Solution Concept created successfully under MODEL > Architecture Development Method > Phase A"
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
CreateSolutionConceptDiagram
Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Architecture Overview Diagram
Sub CreateArchitectureOverviewDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE ARCHITECTURE OVERVIEW DIAGRAM"
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

    ' Check for existing Architecture Overview diagram
    Dim overviewDiagram As EA.Diagram
    Set overviewDiagram = Nothing
    For i = 0 To phaseAPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseAPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Architecture Overview" Then
            Set overviewDiagram = diag
            Exit For
        End If
    Next

    ' Create Architecture Overview diagram if it doesn't exist
    If overviewDiagram Is Nothing Then
        Set overviewDiagram = phaseAPackage.Diagrams.AddNew("OnEdge Architecture Overview", "Class")
        If Not overviewDiagram Is Nothing Then
            overviewDiagram.Notes = "High-Level Architecture Overview Diagram for OnEdge AI Intelligence Service"
            overviewDiagram.Update
            Session.Output "Created diagram: OnEdge Architecture Overview"
        Else
            Session.Output "Error: Failed to create OnEdge Architecture Overview diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Architecture Overview"
    End If

    ' Define elements (goals, components, stakeholders, risks, mitigations) with names, stereotypes, colors, and positions
    Dim elements(9)
    Dim stereotypes(9)
    Dim elementTypes(9)
    Dim elementObjects(9)
    Dim elementColors(9)
    Dim elementSizes(9)
    Dim elementCount
    elementCount = 0

    ' Business Goals (Green, larger size)
    elements(0) = "Real-Time Intelligence": stereotypes(0) = "Goal": elementTypes(0) = "Class": elementColors(0) = RGB(0, 128, 0): elementSizes(0) = "l=50;r=250;t=50;b=0" ' Green
    elementCount = elementCount + 1
    elements(1) = "Affordability": stereotypes(1) = "Goal": elementTypes(1) = "Class": elementColors(1) = RGB(0, 128, 0): elementSizes(1) = "l=300;r=500;t=50;b=0" ' Green
    elementCount = elementCount + 1

    ' Solution Components (Purple, larger size)
    elements(2) = "ONEDGE Micro-Factories": stereotypes(2) = "Component": elementTypes(2) = "Class": elementColors(2) = RGB(153, 102, 255): elementSizes(2) = "l=50;r=250;t=200;b=150" ' Purple
    elementCount = elementCount + 1
    elements(3) = "NeoCortex AI": stereotypes(3) = "Component": elementTypes(3) = "Class": elementColors(3) = RGB(153, 102, 255): elementSizes(3) = "l=300;r=500;t=200;b=150" ' Purple
    elementCount = elementCount + 1
    elements(4) = "Starlink Connectivity": stereotypes(4) = "Component": elementTypes(4) = "Class": elementColors(4) = RGB(153, 102, 255): elementSizes(4) = "l=550;r=750;t=200;b=150" ' Purple
    elementCount = elementCount + 1

    ' Stakeholders (Blue, medium size)
    elements(5) = "CIO/CTO": stereotypes(5) = "Stakeholder": elementTypes(5) = "Class": elementColors(5) = RGB(0, 102, 204): elementSizes(5) = "l=50;r=200;t=350;b=300" ' Blue
    elementCount = elementCount + 1
    elements(6) = "Business Customers": stereotypes(6) = "Stakeholder": elementTypes(6) = "Class": elementColors(6) = RGB(0, 102, 204): elementSizes(6) = "l=250;r=400;t=350;b=300" ' Blue
    elementCount = elementCount + 1

    ' Risks and Mitigations (Red for risk, Green for mitigation, medium size)
    elements(7) = "Data Privacy Concerns": stereotypes(7) = "Risk": elementTypes(7) = "Class": elementColors(7) = RGB(255, 0, 0): elementSizes(7) = "l=450;r=600;t=350;b=300" ' Red
    elementCount = elementCount + 1
    elements(8) = "Compliance Frameworks": stereotypes(8) = "Mitigation": elementTypes(8) = "Class": elementColors(8) = RGB(0, 128, 0): elementSizes(8) = "l=650;r=800;t=350;b=300" ' Green
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
    Set diagObjects = overviewDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations based on Architecture Overview Diagram
    Dim connectors As EA.Collection
    Set connectors = phaseAPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = overviewDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Real-Time Intelligence enabled by NeoCortex AI
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Real-Time Intelligence -> NeoCortex AI (Enabled By)"

    ' Affordability enabled by ONEDGE Micro-Factories
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Affordability -> ONEDGE Micro-Factories (Enabled By)"

    ' ONEDGE Micro-Factories involves Business Customers
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Involves"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: ONEDGE Micro-Factories -> Business Customers (Involves)"

    ' NeoCortex AI involves CIO/CTO
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Involves"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: NeoCortex AI -> CIO/CTO (Involves)"

    ' Starlink Connectivity involves Business Customers
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(4).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Involves"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Starlink Connectivity -> Business Customers (Involves)"

    ' Data Privacy Concerns mitigated by Compliance Frameworks
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(7).ElementID
    conn.ClientID = elementObjects(8).ElementID
    conn.Name = "Mitigated By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Privacy Concerns -> Compliance Frameworks (Mitigated By)"

    ' Data Privacy Concerns impacts CIO/CTO (Dashed)
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(7).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Impacts"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Data Privacy Concerns -> CIO/CTO (Impacts, Dashed)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseAPackage.PackageID
    overviewDiagram.Update
    Repository.ReloadDiagram overviewDiagram.DiagramID
    Session.Output "OnEdge Architecture Overview created successfully under MODEL > Architecture Development Method > Phase A"
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
CreateArchitectureOverviewDiagram
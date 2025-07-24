Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Value Proposition Diagram
Sub CreateValuePropositionDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE VALUE PROPOSITION DIAGRAM"
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

    ' Check for existing Value Proposition diagram
    Dim valueDiagram As EA.Diagram
    Set valueDiagram = Nothing
    For i = 0 To phaseAPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseAPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Value Proposition" Then
            Set valueDiagram = diag
            Exit For
        End If
    Next

    ' Create Value Proposition diagram if it doesn't exist
    If valueDiagram Is Nothing Then
        Set valueDiagram = phaseAPackage.Diagrams.AddNew("OnEdge Value Proposition", "Class")
        If Not valueDiagram Is Nothing Then
            valueDiagram.Notes = "Value Proposition Diagram for OnEdge AI Intelligence Service"
            valueDiagram.Update
            Session.Output "Created diagram: OnEdge Value Proposition"
        Else
            Session.Output "Error: Failed to create OnEdge Value Proposition diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Value Proposition"
    End If

    ' Define elements (value propositions, technologies, industries) with names, stereotypes, colors, and positions
    Dim elements(16)
    Dim stereotypes(16)
    Dim elementTypes(16)
    Dim elementObjects(16)
    Dim elementColors(16)
    Dim elementSizes(16)
    Dim elementCount
    elementCount = 0

    ' Value Propositions (Green/Blue, larger size)
    elements(0) = "Real-Time Intelligence": stereotypes(0) = "ValueProposition": elementTypes(0) = "Class": elementColors(0) = RGB(0, 128, 0): elementSizes(0) = "l=50;r=250;t=50;b=0" ' Green
    elementCount = elementCount + 1
    elements(1) = "Low Latency": stereotypes(1) = "ValueProposition": elementTypes(1) = "Class": elementColors(1) = RGB(0, 128, 0): elementSizes(1) = "l=300;r=500;t=50;b=0" ' Green
    elementCount = elementCount + 1
    elements(2) = "Affordability": stereotypes(2) = "ValueProposition": elementTypes(2) = "Class": elementColors(2) = RGB(0, 102, 204): elementSizes(2) = "l=550;r=750;t=50;b=0" ' Blue
    elementCount = elementCount + 1
    elements(3) = "Data Security": stereotypes(3) = "ValueProposition": elementTypes(3) = "Class": elementColors(3) = RGB(0, 102, 204): elementSizes(3) = "l=800;r=1000;t=50;b=0" ' Blue
    elementCount = elementCount + 1
    elements(4) = "Scalability": stereotypes(4) = "ValueProposition": elementTypes(4) = "Class": elementColors(4) = RGB(0, 128, 0): elementSizes(4) = "l=50;r=250;t=150;b=100" ' Green
    elementCount = elementCount + 1
    elements(5) = "Sustainability": stereotypes(5) = "ValueProposition": elementTypes(5) = "Class": elementColors(5) = RGB(0, 128, 0): elementSizes(5) = "l=300;r=500;t=150;b=100" ' Green
    elementCount = elementCount + 1
    elements(6) = "Customization": stereotypes(6) = "ValueProposition": elementTypes(6) = "Class": elementColors(6) = RGB(0, 102, 204): elementSizes(6) = "l=550;r=750;t=150;b=100" ' Blue
    elementCount = elementCount + 1

    ' Enabling Technologies (Purple)
    elements(7) = "NeoCortex AI": stereotypes(7) = "Component": elementTypes(7) = "Class": elementColors(7) = RGB(153, 102, 255): elementSizes(7) = "l=50;r=200;t=300;b=250" ' Purple
    elementCount = elementCount + 1
    elements(8) = "Starlink Connectivity": stereotypes(8) = "Component": elementTypes(8) = "Class": elementColors(8) = RGB(153, 102, 255): elementSizes(8) = "l=250;r=400;t=300;b=250" ' Purple
    elementCount = elementCount + 1
    elements(9) = "ONEDGE Micro-Factories": stereotypes(9) = "Component": elementTypes(9) = "Class": elementColors(9) = RGB(153, 102, 255): elementSizes(9) = "l=450;r=650;t=300;b=250" ' Purple
    elementCount = elementCount + 1
    elements(10) = "KOIN Payment System": stereotypes(10) = "Component": elementTypes(10) = "Class": elementColors(10) = RGB(153, 102, 255): elementSizes(10) = "l=700;r=850;t=300;b=250" ' Purple
    elementCount = elementCount + 1
    elements(11) = "Developer Marketplace": stereotypes(11) = "Component": elementTypes(11) = "Class": elementColors(11) = RGB(153, 102, 255): elementSizes(11) = "l=900;r=1050;t=300;b=250" ' Purple
    elementCount = elementCount + 1

    ' Industries (Amber)
    elements(12) = "Healthcare": stereotypes(12) = "Business": elementTypes(12) = "Class": elementColors(12) = RGB(255, 204, 102): elementSizes(12) = "l=50;r=200;t=450;b=400" ' Amber
    elementCount = elementCount + 1
    elements(13) = "Manufacturing": stereotypes(13) = "Business": elementTypes(13) = "Class": elementColors(13) = RGB(255, 204, 102): elementSizes(13) = "l=250;r=400;t=450;b=400" ' Amber
    elementCount = elementCount + 1
    elements(14) = "Finance": stereotypes(14) = "Business": elementTypes(14) = "Class": elementColors(14) = RGB(255, 204, 102): elementSizes(14) = "l=450;r=600;t=450;b=400" ' Amber
    elementCount = elementCount + 1
    elements(15) = "Retail": stereotypes(15) = "Business": elementTypes(15) = "Class": elementColors(15) = RGB(255, 204, 102): elementSizes(15) = "l=650;r=800;t=450;b=400" ' Amber
    elementCount = elementCount + 1
    elements(16) = "IoT": stereotypes(16) = "Business": elementTypes(16) = "Class": elementColors(16) = RGB(255, 204, 102): elementSizes(16) = "l=850;r=1000;t=450;b=400" ' Amber
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
    Set diagObjects = valueDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations based on Value Proposition Diagram
    Dim connectors As EA.Collection
    Set connectors = phaseAPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = valueDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Real-Time Intelligence enabled by NeoCortex AI
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(7).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Real-Time Intelligence -> NeoCortex AI (Enabled By)"

    ' Low Latency enabled by Starlink Connectivity
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(8).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Low Latency -> Starlink Connectivity (Enabled By)"

    ' Affordability enabled by ONEDGE Micro-Factories
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(9).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Affordability -> ONEDGE Micro-Factories (Enabled By)"

    ' Affordability enabled by KOIN Payment System
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(10).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Affordability -> KOIN Payment System (Enabled By)"

    ' Data Security enabled by ONEDGE Micro-Factories
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(9).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Security -> ONEDGE Micro-Factories (Enabled By)"

    ' Data Security enabled by KOIN Payment System
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(10).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Security -> KOIN Payment System (Enabled By)"

    ' Scalability enabled by ONEDGE Micro-Factories
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(4).ElementID
    conn.ClientID = elementObjects(9).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Scalability -> ONEDGE Micro-Factories (Enabled By)"

    ' Sustainability enabled by ONEDGE Micro-Factories
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(5).ElementID
    conn.ClientID = elementObjects(9).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Sustainability -> ONEDGE Micro-Factories (Enabled By)"

    ' Customization enabled by Developer Marketplace
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(6).ElementID
    conn.ClientID = elementObjects(11).ElementID
    conn.Name = "Enabled By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Customization -> Developer Marketplace (Enabled By)"

    ' Value Propositions to Industries (Dashed lines for direct impact)
    Dim values, industries
    For values = 0 To 6 ' All value propositions
        For industries = 12 To 16 ' All industries
            Set conn = connectors.AddNew("", "Association")
            conn.SupplierID = elementObjects(values).ElementID
            conn.ClientID = elementObjects(industries).ElementID
            conn.Name = "Delivers Benefit To"
            conn.Update
            Set link = diagLinks.AddNew("", "")
            link.ConnectorID = conn.ConnectorID
            link.LineStyle = 2 ' Dashed line
            link.Update
            Session.Output "Created association: " & elements(values) & " -> " & elements(industries) & " (Delivers Benefit To, Dashed)"
        Next
    Next

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseAPackage.PackageID
    valueDiagram.Update
    Repository.ReloadDiagram valueDiagram.DiagramID
    Session.Output "OnEdge Value Proposition created successfully under MODEL > Architecture Development Method > Phase A"
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
CreateValuePropositionDiagram
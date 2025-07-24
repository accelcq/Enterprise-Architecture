Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Stakeholder Map Diagram
Sub CreateStakeholderMap()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE STAKEHOLDER MAP DIAGRAM"
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

    ' Check for existing Stakeholder Map diagram
    Dim stakeholderDiagram As EA.Diagram
    Set stakeholderDiagram = Nothing
    For i = 0 To phaseAPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseAPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Stakeholder Map" Then
            Set stakeholderDiagram = diag
            Exit For
        End If
    Next

    ' Create Stakeholder Map diagram if it doesn't exist
    If stakeholderDiagram Is Nothing Then
        Set stakeholderDiagram = phaseAPackage.Diagrams.AddNew("OnEdge Stakeholder Map", "Class")
        If Not stakeholderDiagram Is Nothing Then
            stakeholderDiagram.Notes = "Stakeholder Map Diagram for OnEdge AI Intelligence Service"
            stakeholderDiagram.Update
            Session.Output "Created diagram: OnEdge Stakeholder Map"
        Else
            Session.Output "Error: Failed to create OnEdge Stakeholder Map diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Stakeholder Map"
    End If

    ' Define elements (stakeholders, concerns) with names, stereotypes, colors, and positions
    Dim elements(11)
    Dim stereotypes(11)
    Dim elementTypes(11)
    Dim elementObjects(11)
    Dim elementColors(11)
    Dim elementSizes(11)
    Dim elementCount
    elementCount = 0

    ' Stakeholders (Blue, larger size)
    elements(0) = "CIO/CTO": stereotypes(0) = "Stakeholder": elementTypes(0) = "Class": elementColors(0) = RGB(0, 102, 204): elementSizes(0) = "l=50;r=250;t=50;b=0" ' Blue
    elementCount = elementCount + 1
    elements(1) = "Data Privacy Officer": stereotypes(1) = "Stakeholder": elementTypes(1) = "Class": elementColors(1) = RGB(0, 102, 204): elementSizes(1) = "l=300;r=500;t=50;b=0" ' Blue
    elementCount = elementCount + 1
    elements(2) = "Business Customers": stereotypes(2) = "Stakeholder": elementTypes(2) = "Class": elementColors(2) = RGB(0, 102, 204): elementSizes(2) = "l=550;r=750;t=50;b=0" ' Blue
    elementCount = elementCount + 1
    elements(3) = "Solution Architects": stereotypes(3) = "Stakeholder": elementTypes(3) = "Class": elementColors(3) = RGB(0, 102, 204): elementSizes(3) = "l=800;r=1000;t=50;b=0" ' Blue
    elementCount = elementCount + 1
    elements(4) = "Regulatory Bodies": stereotypes(4) = "Stakeholder": elementTypes(4) = "Class": elementColors(4) = RGB(0, 102, 204): elementSizes(4) = "l=50;r=250;t=200;b=150" ' Blue
    elementCount = elementCount + 1
    elements(5) = "Operations Team": stereotypes(5) = "Stakeholder": elementTypes(5) = "Class": elementColors(5) = RGB(0, 102, 204): elementSizes(5) = "l=300;r=500;t=200;b=150" ' Blue
    elementCount = elementCount + 1

    ' Concerns (Orange, smaller size)
    elements(6) = "Data Security": stereotypes(6) = "Concern": elementTypes(6) = "Class": elementColors(6) = RGB(255, 204, 102): elementSizes(6) = "l=50;r=200;t=350;b=300" ' Orange
    elementCount = elementCount + 1
    elements(7) = "Cost Efficiency": stereotypes(7) = "Concern": elementTypes(7) = "Class": elementColors(7) = RGB(255, 204, 102): elementSizes(7) = "l=250;r=400;t=350;b=300" ' Orange
    elementCount = elementCount + 1
    elements(8) = "Scalability": stereotypes(8) = "Concern": elementTypes(8) = "Class": elementColors(8) = RGB(255, 204, 102): elementSizes(8) = "l=450;r=600;t=350;b=300" ' Orange
    elementCount = elementCount + 1
    elements(9) = "Compliance": stereotypes(9) = "Concern": elementTypes(9) = "Class": elementColors(9) = RGB(255, 204, 102): elementSizes(9) = "l=650;r=800;t=350;b=300" ' Orange
    elementCount = elementCount + 1
    elements(10) = "Performance": stereotypes(10) = "Concern": elementTypes(10) = "Class": elementColors(10) = RGB(255, 204, 102): elementSizes(10) = "l=850;r=1000;t=350;b=300" ' Orange
    elementCount = elementCount + 1
    elements(11) = "Usability": stereotypes(11) = "Concern": elementTypes(11) = "Class": elementColors(11) = RGB(255, 204, 102): elementSizes(11) = "l=50;r=200;t=500;b=450" ' Orange
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
    Set diagObjects = stakeholderDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations based on Stakeholder Map Diagram
    Dim connectors As EA.Collection
    Set connectors = phaseAPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = stakeholderDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' CIO/CTO has concern Cost Efficiency
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(7).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: CIO/CTO -> Cost Efficiency (Has Concern)"

    ' CIO/CTO has concern Performance
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(10).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: CIO/CTO -> Performance (Has Concern)"

    ' Data Privacy Officer has concern Data Security
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Privacy Officer -> Data Security (Has Concern)"

    ' Data Privacy Officer has concern Compliance
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(9).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Privacy Officer -> Compliance (Has Concern)"

    ' Business Customers has concern Cost Efficiency
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(7).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Business Customers -> Cost Efficiency (Has Concern)"

    ' Business Customers has concern Usability
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(11).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Business Customers -> Usability (Has Concern)"

    ' Solution Architects has concern Scalability
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(8).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Solution Architects -> Scalability (Has Concern)"

    ' Solution Architects has concern Performance
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(10).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Solution Architects -> Performance (Has Concern)"

    ' Regulatory Bodies has concern Compliance
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(4).ElementID
    conn.ClientID = elementObjects(9).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Regulatory Bodies -> Compliance (Has Concern)"

    ' Operations Team has concern Performance
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(5).ElementID
    conn.ClientID = elementObjects(10).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Operations Team -> Performance (Has Concern)"

    ' Operations Team has concern Scalability
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(5).ElementID
    conn.ClientID = elementObjects(8).ElementID
    conn.Name = "Has Concern"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Operations Team -> Scalability (Has Concern)"

    ' Stakeholder influences (Dashed lines)
    ' CIO/CTO influences Business Customers
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Influences"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: CIO/CTO -> Business Customers (Influences, Dashed)"

    ' Data Privacy Officer influences Regulatory Bodies
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Influences"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Data Privacy Officer -> Regulatory Bodies (Influences, Dashed)"

    ' Solution Architects influences Operations Team
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Influences"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Solution Architects -> Operations Team (Influences, Dashed)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseAPackage.PackageID
    stakeholderDiagram.Update
    Repository.ReloadDiagram stakeholderDiagram.DiagramID
    Session.Output "OnEdge Stakeholder Map created successfully under MODEL > Architecture Development Method > Phase A"
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
CreateStakeholderMap
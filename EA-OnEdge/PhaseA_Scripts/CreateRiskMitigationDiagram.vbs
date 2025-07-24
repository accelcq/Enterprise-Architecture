Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Risk Mitigation Diagram
Sub CreateRiskMitigationDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE RISK MITIGATION DIAGRAM"
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

    ' Check for existing Risk Mitigation diagram
    Dim riskDiagram As EA.Diagram
    Set riskDiagram = Nothing
    For i = 0 To phaseAPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseAPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Risk Mitigation" Then
            Set riskDiagram = diag
            Exit For
        End If
    Next

    ' Create Risk Mitigation diagram if it doesn't exist
    If riskDiagram Is Nothing Then
        Set riskDiagram = phaseAPackage.Diagrams.AddNew("OnEdge Risk Mitigation", "Class")
        If Not riskDiagram Is Nothing Then
            riskDiagram.Notes = "Risk Mitigation Diagram for OnEdge AI Intelligence Service"
            riskDiagram.Update
            Session.Output "Created diagram: OnEdge Risk Mitigation"
        Else
            Session.Output "Error: Failed to create OnEdge Risk Mitigation diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Risk Mitigation"
    End If

    ' Define elements (risks, mitigations, stakeholders/systems) with names, stereotypes, colors, and positions
    Dim elements(15)
    Dim stereotypes(15)
    Dim elementTypes(15)
    Dim elementObjects(15)
    Dim elementColors(15)
    Dim elementSizes(15)
    Dim elementCount
    elementCount = 0

    ' Risks (Red, larger size)
    elements(0) = "Data Privacy Concerns": stereotypes(0) = "Risk": elementTypes(0) = "Class": elementColors(0) = RGB(255, 0, 0): elementSizes(0) = "l=50;r=250;t=50;b=0" ' Red
    elementCount = elementCount + 1
    elements(1) = "High Upfront Costs": stereotypes(1) = "Risk": elementTypes(1) = "Class": elementColors(1) = RGB(255, 0, 0): elementSizes(1) = "l=300;r=500;t=50;b=0" ' Red
    elementCount = elementCount + 1
    elements(2) = "Technical Complexity": stereotypes(2) = "Risk": elementTypes(2) = "Class": elementColors(2) = RGB(255, 0, 0): elementSizes(2) = "l=550;r=750;t=50;b=0" ' Red
    elementCount = elementCount + 1
    elements(3) = "Regulatory Compliance": stereotypes(3) = "Risk": elementTypes(3) = "Class": elementColors(3) = RGB(255, 0, 0): elementSizes(3) = "l=800;r=1000;t=50;b=0" ' Red
    elementCount = elementCount + 1
    elements(4) = "Scalability Challenges": stereotypes(4) = "Risk": elementTypes(4) = "Class": elementColors(4) = RGB(255, 0, 0): elementSizes(4) = "l=50;r=250;t=150;b=100" ' Red
    elementCount = elementCount + 1
    elements(5) = "Adoption Barriers": stereotypes(5) = "Risk": elementTypes(5) = "Class": elementColors(5) = RGB(255, 0, 0): elementSizes(5) = "l=300;r=500;t=150;b=100" ' Red
    elementCount = elementCount + 1

    ' Mitigations (Green, medium size)
    elements(6) = "Compliance Frameworks": stereotypes(6) = "Mitigation": elementTypes(6) = "Class": elementColors(6) = RGB(0, 128, 0): elementSizes(6) = "l=50;r=200;t=300;b=250" ' Green
    elementCount = elementCount + 1
    elements(7) = "Cost-Effective Pricing": stereotypes(7) = "Mitigation": elementTypes(7) = "Class": elementColors(7) = RGB(0, 128, 0): elementSizes(7) = "l=250;r=400;t=300;b=250" ' Green
    elementCount = elementCount + 1
    elements(8) = "Open-Source SDKs": stereotypes(8) = "Mitigation": elementTypes(8) = "Class": elementColors(8) = RGB(0, 128, 0): elementSizes(8) = "l=450;r=600;t=300;b=250" ' Green
    elementCount = elementCount + 1
    elements(9) = "Regulatory Certifications": stereotypes(9) = "Mitigation": elementTypes(9) = "Class": elementColors(9) = RGB(0, 128, 0): elementSizes(9) = "l=650;r=800;t=300;b=250" ' Green
    elementCount = elementCount + 1
    elements(10) = "Scalable Infrastructure": stereotypes(10) = "Mitigation": elementTypes(10) = "Class": elementColors(10) = RGB(0, 128, 0): elementSizes(10) = "l=850;r=1000;t=300;b=250" ' Green
    elementCount = elementCount + 1
    elements(11) = "Training Programs": stereotypes(11) = "Mitigation": elementTypes(11) = "Class": elementColors(11) = RGB(0, 128, 0): elementSizes(11) = "l=50;r=200;t=400;b=350" ' Green
    elementCount = elementCount + 1

    ' Stakeholders/Systems (Blue, smaller size)
    elements(12) = "Data Privacy Officer": stereotypes(12) = "Stakeholder": elementTypes(12) = "Class": elementColors(12) = RGB(0, 102, 204): elementSizes(12) = "l=250;r=400;t=400;b=350" ' Blue
    elementCount = elementCount + 1
    elements(13) = "ONEDGE Micro-Factories": stereotypes(13) = "Component": elementTypes(13) = "Class": elementColors(13) = RGB(0, 102, 204): elementSizes(13) = "l=450;r=600;t=400;b=350" ' Blue
    elementCount = elementCount + 1
    elements(14) = "Business Customers": stereotypes(14) = "Stakeholder": elementTypes(14) = "Class": elementColors(14) = RGB(0, 102, 204): elementSizes(14) = "l=650;r=800;t=400;b=350" ' Blue
    elementCount = elementCount + 1
    elements(15) = "Regulatory Bodies": stereotypes(15) = "Stakeholder": elementTypes(15) = "Class": elementColors(15) = RGB(0, 102, 204): elementSizes(15) = "l=850;r=1000;t=400;b=350" ' Blue
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
    Set diagObjects = riskDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations based on Risk Mitigation Diagram
    Dim connectors As EA.Collection
    Set connectors = phaseAPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = riskDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Data Privacy Concerns mitigated by Compliance Frameworks
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Mitigated By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Privacy Concerns -> Compliance Frameworks (Mitigated By)"

    ' High Upfront Costs mitigated by Cost-Effective Pricing
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(7).ElementID
    conn.Name = "Mitigated By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: High Upfront Costs -> Cost-Effective Pricing (Mitigated By)"

    ' Technical Complexity mitigated by Open-Source SDKs
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(8).ElementID
    conn.Name = "Mitigated By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Technical Complexity -> Open-Source SDKs (Mitigated By)"

    ' Regulatory Compliance mitigated by Regulatory Certifications
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(9).ElementID
    conn.Name = "Mitigated By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Regulatory Compliance -> Regulatory Certifications (Mitigated By)"

    ' Scalability Challenges mitigated by Scalable Infrastructure
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(4).ElementID
    conn.ClientID = elementObjects(10).ElementID
    conn.Name = "Mitigated By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Scalability Challenges -> Scalable Infrastructure (Mitigated By)"

    ' Adoption Barriers mitigated by Training Programs
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(5).ElementID
    conn.ClientID = elementObjects(11).ElementID
    conn.Name = "Mitigated By"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Adoption Barriers -> Training Programs (Mitigated By)"

    ' Risk to Stakeholder/System Impacts (Dashed lines)
    ' Data Privacy Concerns impacts Data Privacy Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(12).ElementID
    conn.Name = "Impacts"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Data Privacy Concerns -> Data Privacy Officer (Impacts, Dashed)"

    ' Data Privacy Concerns impacts ONEDGE Micro-Factories
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(13).ElementID
    conn.Name = "Impacts"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Data Privacy Concerns -> ONEDGE Micro-Factories (Impacts, Dashed)"

    ' High Upfront Costs impacts Business Customers
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(14).ElementID
    conn.Name = "Impacts"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: High Upfront Costs -> Business Customers (Impacts, Dashed)"

    ' Regulatory Compliance impacts Regulatory Bodies
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(15).ElementID
    conn.Name = "Impacts"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Regulatory Compliance -> Regulatory Bodies (Impacts, Dashed)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseAPackage.PackageID
    riskDiagram.Update
    Repository.ReloadDiagram riskDiagram.DiagramID
    Session.Output "OnEdge Risk Mitigation created successfully under MODEL > Architecture Development Method > Phase A"
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
CreateRiskMitigationDiagram
Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Stakeholder Interaction Diagram
Sub CreateStakeholderInteractionDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE STAKEHOLDER INTERACTION DIAGRAM"
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

    ' Check for existing Stakeholder Interaction Diagram
    Dim stakeholderDiagram As EA.Diagram
    Set stakeholderDiagram = Nothing
    For i = 0 To phaseBPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseBPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Stakeholder Interaction" Then
            Set stakeholderDiagram = diag
            Exit For
        End If
    Next

    ' Create Stakeholder Interaction Diagram if it doesn't exist
    If stakeholderDiagram Is Nothing Then
        Set stakeholderDiagram = phaseBPackage.Diagrams.AddNew("OnEdge Stakeholder Interaction", "Activity")
        If Not stakeholderDiagram Is Nothing Then
            stakeholderDiagram.Notes = "Stakeholder Interaction Diagram for OnEdge AI Intelligence Service"
            stakeholderDiagram.Update
            Session.Output "Created diagram: OnEdge Stakeholder Interaction"
        Else
            Session.Output "Error: Failed to create OnEdge Stakeholder Interaction diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Stakeholder Interaction"
    End If

    ' Define elements (stakeholders) with names, stereotypes, colors, and positions
    Dim elements(11)
    Dim stereotypes(11)
    Dim elementTypes(11)
    Dim elementObjects(11)
    Dim elementColors(11)
    Dim elementSizes(11)
    Dim elementCount
    elementCount = 0

    ' Stakeholders (Blue)
    elements(0) = "CIO/CTO": stereotypes(0) = "Stakeholder": elementTypes(0) = "Actor": elementColors(0) = RGB(0, 102, 204): elementSizes(0) = "l=400;r=600;t=50;b=0"
    elementCount = elementCount + 1
    elements(1) = "Data Privacy Officer": stereotypes(1) = "Stakeholder": elementTypes(1) = "Actor": elementColors(1) = RGB(0, 102, 204): elementSizes(1) = "l=50;r=250;t=150;b=100"
    elementCount = elementCount + 1
    elements(2) = "Business Managers": stereotypes(2) = "Stakeholder": elementTypes(2) = "Actor": elementColors(2) = RGB(0, 102, 204): elementSizes(2) = "l=300;r=500;t=150;b=100"
    elementCount = elementCount + 1
    elements(3) = "IT Teams": stereotypes(3) = "Stakeholder": elementTypes(3) = "Actor": elementColors(3) = RGB(0, 102, 204): elementSizes(3) = "l=550;r=750;t=150;b=100"
    elementCount = elementCount + 1
    elements(4) = "Capital Contributors": stereotypes(4) = "Stakeholder": elementTypes(4) = "Actor": elementColors(4) = RGB(0, 102, 204): elementSizes(4) = "l=800;r=1000;t=150;b=100"
    elementCount = elementCount + 1
    elements(5) = "Data Scientists": stereotypes(5) = "Stakeholder": elementTypes(5) = "Actor": elementColors(5) = RGB(0, 102, 204): elementSizes(5) = "l=1050;r=1250;t=150;b=100"
    elementCount = elementCount + 1
    elements(6) = "End Users": stereotypes(6) = "Stakeholder": elementTypes(6) = "Actor": elementColors(6) = RGB(0, 102, 204): elementSizes(6) = "l=50;r=250;t=300;b=250"
    elementCount = elementCount + 1
    elements(7) = "Regulatory Bodies": stereotypes(7) = "Stakeholder": elementTypes(7) = "Actor": elementColors(7) = RGB(0, 102, 204): elementSizes(7) = "l=300;r=500;t=300;b=250"
    elementCount = elementCount + 1
    elements(8) = "Application Developers": stereotypes(8) = "Stakeholder": elementTypes(8) = "Actor": elementColors(8) = RGB(0, 102, 204): elementSizes(8) = "l=550;r=750;t=300;b=250"
    elementCount = elementCount + 1
    elements(9) = "Sustainability Auditors": stereotypes(9) = "Stakeholder": elementTypes(9) = "Actor": elementColors(9) = RGB(0, 102, 204): elementSizes(9) = "l=800;r=1000;t=300;b=250"
    elementCount = elementCount + 1
    elements(10) = "Customer Support Teams": stereotypes(10) = "Stakeholder": elementTypes(10) = "Actor": elementColors(10) = RGB(0, 102, 204): elementSizes(10) = "l=1050;r=1250;t=300;b=250"
    elementCount = elementCount + 1
    elements(11) = "Marketing Teams": stereotypes(11) = "Stakeholder": elementTypes(11) = "Actor": elementColors(11) = RGB(0, 102, 204): elementSizes(11) = "l=1300;r=1500;t=300;b=250"
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
    Set diagObjects = stakeholderDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations for stakeholder interactions
    Dim connectors As EA.Collection
    Set connectors = phaseBPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = stakeholderDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Direct interactions (Solid lines)
    ' CIO/CTO collaborates with Data Privacy Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(1).ElementID
    conn.Name = "Collaborates With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: CIO/CTO -> Data Privacy Officer (Collaborates With)"

    ' CIO/CTO collaborates with Capital Contributors
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Collaborates With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: CIO/CTO -> Capital Contributors (Collaborates With)"

    ' Business Managers collaborate with IT Teams
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Collaborates With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Business Managers -> IT Teams (Collaborates With)"

    ' IT Teams collaborate with Customer Support Teams
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(10).ElementID
    conn.Name = "Collaborates With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: IT Teams -> Customer Support Teams (Collaborates With)"

    ' Data Privacy Officer collaborates with Regulatory Bodies
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(7).ElementID
    conn.Name = "Collaborates With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Privacy Officer -> Regulatory Bodies (Collaborates With)"

    ' Data Scientists collaborate with Application Developers
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(5).ElementID
    conn.ClientID = elementObjects(8).ElementID
    conn.Name = "Collaborates With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Scientists -> Application Developers (Collaborates With)"

    ' Customer Support Teams collaborate with End Users
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(10).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Collaborates With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Customer Support Teams -> End Users (Collaborates With)"

    ' Marketing Teams collaborate with Business Managers
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(11).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Collaborates With"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Marketing Teams -> Business Managers (Collaborates With)"

    ' Influence relationships (Dashed lines)
    ' CIO/CTO influences Business Managers
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Influences"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: CIO/CTO -> Business Managers (Influences, Dashed)"

    ' Capital Contributors influence CIO/CTO
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(4).ElementID
    conn.ClientID = elementObjects(0).ElementID
    conn.Name = "Influences"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Capital Contributors -> CIO/CTO (Influences, Dashed)"

    ' Regulatory Bodies influence Data Privacy Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(7).ElementID
    conn.ClientID = elementObjects(1).ElementID
    conn.Name = "Influences"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Regulatory Bodies -> Data Privacy Officer (Influences, Dashed)"

    ' Business Managers influence End Users
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Influences"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Business Managers -> End Users (Influences, Dashed)"

    ' Sustainability Auditors influence Sustainability Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(9).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Influences"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Sustainability Auditors -> Sustainability Officer (Influences, Dashed)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseBPackage.PackageID
    stakeholderDiagram.Update
    Repository.ReloadDiagram stakeholderDiagram.DiagramID
    Session.Output "OnEdge Stakeholder Interaction created successfully under MODEL > Architecture Development Method > Phase B"
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
CreateStakeholderInteractionDiagram
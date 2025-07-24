Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Organization Structure Diagram
Sub CreateOrganizationStructureDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE ORGANIZATION STRUCTURE DIAGRAM"
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

    ' Find or create the "Phase C" package
    Dim phaseCPackage As EA.Package
    Set phaseCPackage = findPackageByName(admPackage, "Phase C")
    If phaseCPackage Is Nothing Then
        Set phaseCPackage = admPackage.Packages.AddNew("Phase C", "Package")
        phaseCPackage.Update
        Session.Output "Created package: Phase C"
    End If
    Session.Output "Found package: " & phaseCPackage.Name

    ' Check for existing Organization Structure Diagram
    Dim orgDiagram As EA.Diagram
    Set orgDiagram = Nothing
    For i = 0 To phaseCPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseCPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Organization Structure" Then
            Set orgDiagram = diag
            Exit For
        End If
    Next

    ' Create Organization Structure Diagram if it doesn't exist
    If orgDiagram Is Nothing Then
        Set orgDiagram = phaseCPackage.Diagrams.AddNew("OnEdge Organization Structure", "Class")
        If Not orgDiagram Is Nothing Then
            orgDiagram.Notes = "Organization Structure Diagram for OnEdge AI Intelligence Service"
            orgDiagram.Update
            Session.Output "Created diagram: OnEdge Organization Structure"
        Else
            Session.Output "Error: Failed to create OnEdge Organization Structure diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Organization Structure"
    End If

    ' Define elements (roles and governance bodies) with names, stereotypes, colors, and positions
    Dim elements(8)
    Dim stereotypes(8)
    Dim elementTypes(8)
    Dim elementObjects(8)
    Dim elementColors(8)
    Dim elementSizes(8)
    Dim elementCount
    elementCount = 0

    ' Roles (Blue)
    elements(0) = "Chief Intelligence Officer": stereotypes(0) = "Role": elementTypes(0) = "Actor": elementColors(0) = RGB(0, 102, 204): elementSizes(0) = "l=400;r=600;t=50;b=0"
    elementCount = elementCount + 1
    elements(1) = "Data Privacy Officer": stereotypes(1) = "Role": elementTypes(1) = "Actor": elementColors(1) = RGB(0, 102, 204): elementSizes(1) = "l=50;r=250;t=150;b=100"
    elementCount = elementCount + 1
    elements(2) = "Operations Manager": stereotypes(2) = "Role": elementTypes(2) = "Actor": elementColors(2) = RGB(0, 102, 204): elementSizes(2) = "l=300;r=500;t=150;b=100"
    elementCount = elementCount + 1
    elements(3) = "Developer Relations Manager": stereotypes(3) = "Role": elementTypes(3) = "Actor": elementColors(3) = RGB(0, 102, 204): elementSizes(3) = "l=550;r=750;t=150;b=100"
    elementCount = elementCount + 1
    elements(4) = "Customer Success Manager": stereotypes(4) = "Role": elementTypes(4) = "Actor": elementColors(4) = RGB(0, 102, 204): elementSizes(4) = "l=800;r=1000;t=150;b=100"
    elementCount = elementCount + 1
    elements(5) = "Sustainability Officer": stereotypes(5) = "Role": elementTypes(5) = "Actor": elementColors(5) = RGB(0, 102, 204): elementSizes(5) = "l=1050;r=1250;t=150;b=100"
    elementCount = elementCount + 1

    ' Governance Bodies (Orange)
    elements(6) = "Steering Committee": stereotypes(6) = "Governance Body": elementTypes(6) = "Class": elementColors(6) = RGB(255, 204, 102): elementSizes(6) = "l=300;r=500;t=300;b=250"
    elementCount = elementCount + 1
    elements(7) = "Compliance Board": stereotypes(7) = "Governance Body": elementTypes(7) = "Class": elementColors(7) = RGB(255, 204, 102): elementSizes(7) = "l=550;r=750;t=300;b=250"
    elementCount = elementCount + 1
    elements(8) = "Developer Council": stereotypes(8) = "Governance Body": elementTypes(8) = "Class": elementColors(8) = RGB(255, 204, 102): elementSizes(8) = "l=800;r=1000;t=300;b=250"
    elementCount = elementCount + 1

    ' Create elements in the package
    For i = 0 To elementCount - 1
        Dim elem As EA.Element
        Set elem = phaseCPackage.Elements.AddNew(elements(i), elementTypes(i))
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
    Set diagObjects = orgDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations for reporting lines and governance relationships
    Dim connectors As EA.Collection
    Set connectors = phaseCPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = orgDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Reporting lines (Solid lines)
    ' Data Privacy Officer reports to Chief Intelligence Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(0).ElementID
    conn.Name = "Reports To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Data Privacy Officer -> Chief Intelligence Officer (Reports To)"

    ' Operations Manager reports to Chief Intelligence Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(0).ElementID
    conn.Name = "Reports To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Operations Manager -> Chief Intelligence Officer (Reports To)"

    ' Developer Relations Manager reports to Chief Intelligence Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(0).ElementID
    conn.Name = "Reports To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Developer Relations Manager -> Chief Intelligence Officer (Reports To)"

    ' Customer Success Manager reports to Chief Intelligence Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(4).ElementID
    conn.ClientID = elementObjects(0).ElementID
    conn.Name = "Reports To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Customer Success Manager -> Chief Intelligence Officer (Reports To)"

    ' Sustainability Officer reports to Chief Intelligence Officer
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(5).ElementID
    conn.ClientID = elementObjects(0).ElementID
    conn.Name = "Reports To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Sustainability Officer -> Chief Intelligence Officer (Reports To)"

    ' Governance relationships (Dashed lines)
    ' Chief Intelligence Officer participates in Steering Committee
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Participates In"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Chief Intelligence Officer -> Steering Committee (Participates In, Dashed)"

    ' Data Privacy Officer participates in Compliance Board
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(7).ElementID
    conn.Name = "Participates In"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Data Privacy Officer -> Compliance Board (Participates In, Dashed)"

    ' Developer Relations Manager participates in Developer Council
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(8).ElementID
    conn.Name = "Participates In"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Developer Relations Manager -> Developer Council (Participates In, Dashed)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseCPackage.PackageID
    orgDiagram.Update
    Repository.ReloadDiagram orgDiagram.DiagramID
    Session.Output "OnEdge Organization Structure created successfully under MODEL > Architecture Development Method > Phase C"
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
CreateOrganizationStructureDiagram
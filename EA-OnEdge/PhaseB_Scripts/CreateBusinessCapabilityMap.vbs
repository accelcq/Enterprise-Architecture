Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Business Capability Map Diagram
Sub CreateBusinessCapabilityMap()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE BUSINESS CAPABILITY MAP DIAGRAM"
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

    ' Check for existing Business Capability Map diagram
    Dim capabilityDiagram As EA.Diagram
    Set capabilityDiagram = Nothing
    For i = 0 To phaseBPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseBPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Business Capability Map" Then
            Set capabilityDiagram = diag
            Exit For
        End If
    Next

    ' Create Business Capability Map diagram if it doesn't exist
    If capabilityDiagram Is Nothing Then
        Set capabilityDiagram = phaseBPackage.Diagrams.AddNew("OnEdge Business Capability Map", "Class")
        If Not capabilityDiagram Is Nothing Then
            capabilityDiagram.Notes = "Business Capability Map Diagram for OnEdge AI Intelligence Service"
            capabilityDiagram.Update
            Session.Output "Created diagram: OnEdge Business Capability Map"
        Else
            Session.Output "Error: Failed to create OnEdge Business Capability Map diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Business Capability Map"
    End If

    ' Define elements (capabilities) with names, stereotypes, colors, and positions
    Dim elements(9)
    Dim stereotypes(9)
    Dim elementTypes(9)
    Dim elementObjects(9)
    Dim elementColors(9)
    Dim elementSizes(9)
    Dim elementCount
    elementCount = 0

    ' Core Capabilities (Green, larger size)
    elements(0) = "AI Inference Delivery": stereotypes(0) = "Core Capability": elementTypes(0) = "Class": elementColors(0) = RGB(0, 128, 0): elementSizes(0) = "l=50;r=250;t=50;b=0"
    elementCount = elementCount + 1
    elements(1) = "Low-Latency Connectivity": stereotypes(1) = "Core Capability": elementTypes(1) = "Class": elementColors(1) = RGB(0, 128, 0): elementSizes(1) = "l=300;r=500;t=50;b=0"
    elementCount = elementCount + 1
    elements(2) = "Custom AI Application Delivery": stereotypes(2) = "Core Capability": elementTypes(2) = "Class": elementColors(2) = RGB(0, 128, 0): elementSizes(2) = "l=550;r=750;t=50;b=0"
    elementCount = elementCount + 1
    elements(3) = "Secure Data Processing": stereotypes(3) = "Core Capability": elementTypes(3) = "Class": elementColors(3) = RGB(0, 128, 0): elementSizes(3) = "l=800;r=1000;t=50;b=0"
    elementCount = elementCount + 1

    ' Supporting Capabilities (Blue, smaller size)
    elements(4) = "Payment Processing": stereotypes(4) = "Supporting Capability": elementTypes(4) = "Class": elementColors(4) = RGB(0, 102, 204): elementSizes(4) = "l=300;r=450;t=200;b=150"
    elementCount = elementCount + 1
    elements(5) = "Infrastructure Management": stereotypes(5) = "Supporting Capability": elementTypes(5) = "Class": elementColors(5) = RGB(0, 102, 204): elementSizes(5) = "l=50;r=200;t=300;b=250"
    elementCount = elementCount + 1
    elements(6) = "Compliance Management": stereotypes(6) = "Supporting Capability": elementTypes(6) = "Class": elementColors(6) = RGB(0, 102, 204): elementSizes(6) = "l=300;r=450;t=300;b=250"
    elementCount = elementCount + 1
    elements(7) = "Sustainability Monitoring": stereotypes(7) = "Supporting Capability": elementTypes(7) = "Class": elementColors(7) = RGB(0, 102, 204): elementSizes(7) = "l=550;r=700;t=300;b=250"
    elementCount = elementCount + 1
    elements(8) = "Customer Support": stereotypes(8) = "Supporting Capability": elementTypes(8) = "Class": elementColors(8) = RGB(0, 102, 204): elementSizes(8) = "l=800;r=950;t=300;b=250"
    elementCount = elementCount + 1
    elements(9) = "Developer Ecosystem Management": stereotypes(9) = "Supporting Capability": elementTypes(9) = "Class": elementColors(9) = RGB(0, 102, 204): elementSizes(9) = "l=300;r=450;t=400;b=350"
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
    Set diagObjects = capabilityDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations based on Business Capability Map relationships
    Dim connectors As EA.Collection
    Set connectors = phaseBPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = capabilityDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Dependencies (Solid lines)
    ' AI Inference Delivery depends on Infrastructure Management
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Depends On"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: AI Inference Delivery -> Infrastructure Management (Depends On)"

    ' AI Inference Delivery depends on Secure Data Processing
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Depends On"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: AI Inference Delivery -> Secure Data Processing (Depends On)"

    ' Custom AI Application Delivery depends on Developer Ecosystem Management
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(9).ElementID
    conn.Name = "Depends On"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Custom AI Application Delivery -> Developer Ecosystem Management (Depends On)"

    ' Custom AI Application Delivery depends on Secure Data Processing
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Depends On"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Custom AI Application Delivery -> Secure Data Processing (Depends On)"

    ' Support relationships (Dashed lines)
    ' AI Inference Delivery supports Custom AI Application Delivery
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Supports"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: AI Inference Delivery -> Custom AI Application Delivery (Supports, Dashed)"

    ' Payment Processing supports Custom AI Application Delivery
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(4).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Supports"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Payment Processing -> Custom AI Application Delivery (Supports, Dashed)"

    ' Compliance Management supports Secure Data Processing
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(6).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Supports"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Compliance Management -> Secure Data Processing (Supports, Dashed)"

    ' Sustainability Monitoring supports Infrastructure Management
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(7).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Supports"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line
    link.Update
    Session.Output "Created association: Sustainability Monitoring -> Infrastructure Management (Supports, Dashed)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseBPackage.PackageID
    capabilityDiagram.Update
    Repository.ReloadDiagram capabilityDiagram.DiagramID
    Session.Output "OnEdge Business Capability Map created successfully under MODEL > Architecture Development Method > Phase B"
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
CreateBusinessCapabilityMap
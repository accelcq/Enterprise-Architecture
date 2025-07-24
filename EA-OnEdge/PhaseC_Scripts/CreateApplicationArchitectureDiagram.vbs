Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Application Architecture Diagram
Sub CreateApplicationArchitectureDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE APPLICATION ARCHITECTURE DIAGRAM"
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

    ' Check for existing Application Architecture Diagram
    Dim appDiagram As EA.Diagram
    Set appDiagram = Nothing
    For i = 0 To phaseCPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseCPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Application Architecture" Then
            Set appDiagram = diag
            Exit For
        End If
    Next

    ' Create Application Architecture Diagram if it doesn't exist
    If appDiagram Is Nothing Then
        Set appDiagram = phaseCPackage.Diagrams.AddNew("OnEdge Application Architecture", "Component")
        If Not appDiagram Is Nothing Then
            appDiagram.Notes = "Application Architecture Diagram for OnEdge AI Intelligence Service. Integrates with ERP/CRM (SAP, Salesforce), Starlink, Zendesk, and regulatory systems."
            appDiagram.Update
            Session.Output "Created diagram: OnEdge Application Architecture"
        Else
            Session.Output "Error: Failed to create OnEdge Application Architecture diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Application Architecture"
    End If

    ' Define elements (application components) with names, stereotypes, colors, and positions
    Dim elements(7)
    Dim stereotypes(7)
    Dim elementTypes(7)
    Dim elementObjects(7)
    Dim elementColors(7)
    Dim elementSizes(7)
    Dim elementCount
    elementCount = 0

    ' Application Components (Blue)
    elements(0) = "NeoCortex AI Engine": stereotypes(0) = "ApplicationComponent": elementTypes(0) = "Component": elementColors(0) = RGB(0, 102, 204): elementSizes(0) = "l=50;r=250;t=100;b=50"
    elementCount = elementCount + 1
    elements(1) = "Micro-Factory Manager": stereotypes(1) = "ApplicationComponent": elementTypes(1) = "Component": elementColors(1) = RGB(0, 102, 204): elementSizes(1) = "l=300;r=500;t=100;b=50"
    elementCount = elementCount + 1
    elements(2) = "Marketplace Platform": stereotypes(2) = "ApplicationComponent": elementTypes(2) = "Component": elementColors(2) = RGB(0, 102, 204): elementSizes(2) = "l=550;r=750;t=100;b=50"
    elementCount = elementCount + 1
    elements(3) = "KOIN Payment Gateway": stereotypes(3) = "ApplicationComponent": elementTypes(3) = "Component": elementColors(3) = RGB(0, 102, 204): elementSizes(3) = "l=800;r=1000;t=100;b=50"
    elementCount = elementCount + 1
    elements(4) = "Compliance Manager": stereotypes(4) = "ApplicationComponent": elementTypes(4) = "Component": elementColors(4) = RGB(0, 102, 204): elementSizes(4) = "l=50;r=250;t=300;b=250"
    elementCount = elementCount + 1
    elements(5) = "Sustainability Monitor": stereotypes(5) = "ApplicationComponent": elementTypes(5) = "Component": elementColors(5) = RGB(0, 102, 204): elementSizes(5) = "l=300;r=500;t=300;b=250"
    elementCount = elementCount + 1
    elements(6) = "Customer Support Portal": stereotypes(6) = "ApplicationComponent": elementTypes(6) = "Component": elementColors(6) = RGB(0, 102, 204): elementSizes(6) = "l=550;r=750;t=300;b=250"
    elementCount = elementCount + 1
    elements(7) = "Edge Device Interface": stereotypes(7) = "ApplicationComponent": elementTypes(7) = "Component": elementColors(7) = RGB(0, 102, 204): elementSizes(7) = "l=800;r=1000;t=300;b=250"
    elementCount = elementCount + 1

    ' Create elements in the package
    For i = 0 To elementCount - 1
        Dim elem As EA.Element
        Set elem = phaseCPackage.Elements.AddNew(elements(i), elementTypes(i))
        If Not elem Is Nothing Then
            elem.Stereotype = stereotypes(i)
            ' Add functions to element notes based on document
            If elements(i) = "NeoCortex AI Engine" Then
                elem.Notes = "Functions: Data preprocessing, model execution, insight generation. Interfaces: REST/MQTT APIs."
            ElseIf elements(i) = "Micro-Factory Manager" Then
                elem.Notes = "Functions: Resource monitoring, OTA updates, workload balancing, environmental metrics. Interfaces: Ansible, REST APIs."
            ElseIf elements(i) = "Marketplace Platform" Then
                elem.Notes = "Functions: App publication, validation, deployment, revenue-sharing. Interfaces: Web portal, OTA APIs."
            ElseIf elements(i) = "KOIN Payment Gateway" Then
                elem.Notes = "Functions: Transaction validation, blockchain integration, fiat support. Interfaces: KOIN blockchain API, Stripe."
            ElseIf elements(i) = "Compliance Manager" Then
                elem.Notes = "Functions: Audit trail generation, compliance monitoring, report submission. Interfaces: KOIN blockchain, regulatory APIs."
            ElseIf elements(i) = "Sustainability Monitor" Then
                elem.Notes = "Functions: Metrics collection, lifecycle assessments, sustainability reporting. Interfaces: NeoCortex analytics API."
            ElseIf elements(i) = "Customer Support Portal" Then
                elem.Notes = "Functions: Issue tracking, SLA management, feedback loops. Interfaces: REST APIs, Zendesk integration."
            ElseIf elements(i) = "Edge Device Interface" Then
                elem.Notes = "Functions: Data ingestion, insight delivery, device management. Interfaces: REST/MQTT, device SDKs."
            End If
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
    Set diagObjects = appDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations for application interactions
    Dim connectors As EA.Collection
    Set connectors = phaseCPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = appDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Data Processing Flow
    ' Edge Device Interface to NeoCortex AI Engine
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(7).ElementID
    conn.ClientID = elementObjects(0).ElementID
    conn.Name = "Sends Data To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Edge Device Interface -> NeoCortex AI Engine (Sends Data To)"

    ' NeoCortex AI Engine to Edge Device Interface
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(7).ElementID
    conn.Name = "Delivers Insights To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: NeoCortex AI Engine -> Edge Device Interface (Delivers Insights To)"

    ' App Deployment and Usage
    ' Marketplace Platform to Micro-Factory Manager
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(1).ElementID
    conn.Name = "Deploys Apps To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Marketplace Platform -> Micro-Factory Manager (Deploys Apps To)"

    ' NeoCortex AI Engine to Marketplace Platform
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Executes Apps From"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: NeoCortex AI Engine -> Marketplace Platform (Executes Apps From)"

    ' Payment Processing
    ' Marketplace Platform to KOIN Payment Gateway
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(3).ElementID
    conn.Name = "Initiates Transactions To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Marketplace Platform -> KOIN Payment Gateway (Initiates Transactions To)"

    ' Compliance Monitoring
    ' NeoCortex AI Engine to Compliance Manager
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Generates Logs For"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line for secondary relationship
    link.Update
    Session.Output "Created association: NeoCortex AI Engine -> Compliance Manager (Generates Logs For, Dashed)"

    ' Micro-Factory Manager to Compliance Manager
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Generates Logs For"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line for secondary relationship
    link.Update
    Session.Output "Created association: Micro-Factory Manager -> Compliance Manager (Generates Logs For, Dashed)"

    ' Sustainability Tracking
    ' Micro-Factory Manager to Sustainability Monitor
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Provides Metrics To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Micro-Factory Manager -> Sustainability Monitor (Provides Metrics To)"

    ' Customer Support
    ' Edge Device Interface to Customer Support Portal
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(7).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Routes Queries To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Edge Device Interface -> Customer Support Portal (Routes Queries To)"

    ' Marketplace Platform to Customer Support Portal
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(6).ElementID
    conn.Name = "Routes Queries To"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Marketplace Platform -> Customer Support Portal (Routes Queries To)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseCPackage.PackageID
    appDiagram.Update
    Repository.ReloadDiagram appDiagram.DiagramID
    Session.Output "OnEdge Application Architecture created successfully under MODEL > Architecture Development Method > Phase C"
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
CreateApplicationArchitectureDiagram
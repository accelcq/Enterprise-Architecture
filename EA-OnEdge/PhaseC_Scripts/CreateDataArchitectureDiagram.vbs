Option Explicit
!INC Local Scripts.EAConstants-VBScript

' Main function to create the OnEdge Data Architecture Diagram
Sub CreateDataArchitectureDiagram()
    ' Show the script output window
    Repository.EnsureOutputVisible "Script"
   
    Dim hasError
    hasError = False

    Session.Output "VBScript CREATE ONEDGE DATA ARCHITECTURE DIAGRAM"
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

    ' Check for existing Data Architecture Diagram
    Dim dataDiagram As EA.Diagram
    Set dataDiagram = Nothing
    For i = 0 To phaseCPackage.Diagrams.Count - 1
        Dim diag As EA.Diagram
        Set diag = phaseCPackage.Diagrams.GetAt(i)
        If diag.Name = "OnEdge Data Architecture" Then
            Set dataDiagram = diag
            Exit For
        End If
    Next

    ' Create Data Architecture Diagram if it doesn't exist
    If dataDiagram Is Nothing Then
        Set dataDiagram = phaseCPackage.Diagrams.AddNew("OnEdge Data Architecture", "Class")
        If Not dataDiagram Is Nothing Then
            dataDiagram.Notes = "Data Architecture Diagram for OnEdge AI Intelligence Service. Storage: Local AES-256 encrypted on Micro-Factories; Transaction Records and Compliance Logs on KOIN blockchain."
            dataDiagram.Update
            Session.Output "Created diagram: OnEdge Data Architecture"
        Else
            Session.Output "Error: Failed to create OnEdge Data Architecture diagram."
            hasError = True
            Exit Sub
        End If
    Else
        Session.Output "Diagram already exists: OnEdge Data Architecture"
    End If

    ' Define elements (data entities) with names, stereotypes, colors, and positions
    Dim elements(5)
    Dim stereotypes(5)
    Dim elementTypes(5)
    Dim elementObjects(5)
    Dim elementColors(5)
    Dim elementSizes(5)
    Dim elementCount
    elementCount = 0

    ' Data Entities (Blue)
    elements(0) = "Raw Data": stereotypes(0) = "DataEntity": elementTypes(0) = "Class": elementColors(0) = RGB(0, 102, 204): elementSizes(0) = "l=50;r=250;t=100;b=50"
    elementCount = elementCount + 1
    elements(1) = "Processed Data": stereotypes(1) = "DataEntity": elementTypes(1) = "Class": elementColors(1) = RGB(0, 102, 204): elementSizes(1) = "l=300;r=500;t=100;b=50"
    elementCount = elementCount + 1
    elements(2) = "AI Insights": stereotypes(2) = "DataEntity": elementTypes(2) = "Class": elementColors(2) = RGB(0, 102, 204): elementSizes(2) = "l=550;r=750;t=100;b=50"
    elementCount = elementCount + 1
    elements(3) = "Transaction Records": stereotypes(3) = "DataEntity": elementTypes(3) = "Class": elementColors(3) = RGB(0, 102, 204): elementSizes(3) = "l=50;r=250;t=300;b=250"
    elementCount = elementCount + 1
    elements(4) = "Compliance Logs": stereotypes(4) = "DataEntity": elementTypes(4) = "Class": elementColors(4) = RGB(0, 102, 204): elementSizes(4) = "l=300;r=500;t=300;b=250"
    elementCount = elementCount + 1
    elements(5) = "Environmental Metrics": stereotypes(5) = "DataEntity": elementTypes(5) = "Class": elementColors(5) = RGB(0, 102, 204): elementSizes(5) = "l=550;r=750;t=300;b=250"
    elementCount = elementCount + 1

    ' Create elements in the package
    For i = 0 To elementCount - 1
        Dim elem As EA.Element
        Set elem = phaseCPackage.Elements.AddNew(elements(i), elementTypes(i))
        If Not elem Is Nothing Then
            elem.Stereotype = stereotypes(i)
            ' Add attributes to elements based on document
            If elements(i) = "Raw Data" Then
                elem.Notes = "Attributes: Timestamp, Device ID, Data Type, Payload, Source. Storage: AES-256 encrypted local storage, 24-hour retention."
            ElseIf elements(i) = "Processed Data" Then
                elem.Notes = "Attributes: Preprocessed Payload, Metadata, Parent Raw Data ID. Storage: Temporary cache (in-memory/SSD), deleted post-inference."
            ElseIf elements(i) = "AI Insights" Then
                elem.Notes = "Attributes: Insight ID, Result, Confidence Score, Timestamp, Target Device/Application. Storage: Local or secure cloud sync (TLS 1.3)."
            ElseIf elements(i) = "Transaction Records" Then
                elem.Notes = "Attributes: Transaction ID, Amount, Currency, Payment Method, Timestamp, Status, User ID. Storage: KOIN blockchain, local encrypted backups."
            ElseIf elements(i) = "Compliance Logs" Then
                elem.Notes = "Attributes: Log ID, Action, User ID, Timestamp, Compliance Status, Metadata. Storage: KOIN blockchain, local encrypted storage."
            ElseIf elements(i) = "Environmental Metrics" Then
                elem.Notes = "Attributes: Metric ID, Emission Value, Energy Consumption, Timestamp, Micro-Factory ID. Storage: Local, aggregated for ISO 14001."
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
    Set diagObjects = dataDiagram.DiagramObjects
    For i = 0 To elementCount - 1
        Dim diagObj As EA.DiagramObject
        Set diagObj = diagObjects.AddNew(elementSizes(i) & ";", "")
        diagObj.ElementID = elementObjects(i).ElementID
        diagObj.BackgroundColor = elementColors(i)
        diagObj.Update
    Next

    ' Create associations for data relationships
    Dim connectors As EA.Collection
    Set connectors = phaseCPackage.Connectors
    Dim diagLinks As EA.Collection
    Set diagLinks = dataDiagram.DiagramLinks
    Dim conn As EA.Connector
    Dim link As EA.DiagramLink

    ' Raw Data to Processed Data (One-to-One)
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(1).ElementID
    conn.Name = "Transforms To (1:1)"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Raw Data -> Processed Data (Transforms To, 1:1)"

    ' Processed Data to AI Insights (One-to-Many)
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(1).ElementID
    conn.ClientID = elementObjects(2).ElementID
    conn.Name = "Generates (1:0..*)"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Processed Data -> AI Insights (Generates, 1:0..*)"

    ' Transaction Records to Compliance Logs (One-to-Many)
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(3).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Generates Logs (1:0..*)"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Transaction Records -> Compliance Logs (Generates Logs, 1:0..*)"

    ' Raw Data to Compliance Logs (Many-to-One)
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(0).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Logged By (0..*:1)"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line for secondary relationship
    link.Update
    Session.Output "Created association: Raw Data -> Compliance Logs (Logged By, 0..*:1, Dashed)"

    ' AI Insights to Compliance Logs (Many-to-One)
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(2).ElementID
    conn.ClientID = elementObjects(4).ElementID
    conn.Name = "Logged By (0..*:1)"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.LineStyle = 2 ' Dashed line for secondary relationship
    link.Update
    Session.Output "Created association: AI Insights -> Compliance Logs (Logged By, 0..*:1, Dashed)"

    ' Micro-Factory to Environmental Metrics (One-to-Many)
    Set conn = connectors.AddNew("", "Association")
    conn.SupplierID = elementObjects(5).ElementID
    conn.ClientID = elementObjects(5).ElementID
    conn.Name = "Generates (1:0..*)"
    conn.Update
    Set link = diagLinks.AddNew("", "")
    link.ConnectorID = conn.ConnectorID
    link.Update
    Session.Output "Created association: Micro-Factory -> Environmental Metrics (Generates, 1:0..*)"

    Session.Output "Created all associations"

    ' Refresh the model view
    Repository.RefreshModelView phaseCPackage.PackageID
    dataDiagram.Update
    Repository.ReloadDiagram dataDiagram.DiagramID
    Session.Output "OnEdge Data Architecture created successfully under MODEL > Architecture Development Method > Phase C"
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
CreateDataArchitectureDiagram
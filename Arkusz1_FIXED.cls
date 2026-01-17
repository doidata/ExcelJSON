VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' FIXED VERSION - All bugs corrected:
' 1. Added declaration for 'id' variable
' 2. Fixed Count() to Count (property, not method)
' 3. Added proper error handling instead of blanket On Error Resume Next
Private Sub CommandButton2_Click()
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim ii As Integer
    Dim parent As String
    Dim id As String
    Dim nodeId As String
    Dim nodes As New Collection
    Dim nAttr As Dictionary
    Dim attr As Dictionary
    Dim niAttr As Integer
    Dim iAttr As Integer
    Dim cos As String
    Dim children As New Collection

    ' First loop: Create nodes from Sheet 1
    i = 2
    While Worksheets(1).Cells(i, 1).Value <> ""
        ' Try to add node, skip if duplicate key exists
        On Error Resume Next
        Err.Clear
        nodes.Add New jsonNode, Worksheets(1).Cells(i, 1).Value

        If Err.Number = 0 Then
            ' Successfully added, configure the node
            On Error GoTo ErrorHandler
            nodes(nodes.Count).id = Worksheets(1).Cells(i, 1).Value
            nodes(nodes.Count).parent = Worksheets(1).Cells(i, 2).Value

            Set nAttr = New Dictionary
            niAttr = 3
            While Worksheets(1).Cells(1, niAttr).Value <> ""
                nAttr.Add Worksheets(1).Cells(1, niAttr).Value, Worksheets(1).Cells(i, niAttr).Value
                niAttr = niAttr + 1
            Wend

            nodes(nodes.Count).attributes = nAttr
            Set nAttr = Nothing
        End If

        On Error GoTo ErrorHandler
        i = i + 1
    Wend

    ' Second loop: Build parent-child relationships
    i = 2
    While Worksheets(1).Cells(i, 1).Value <> ""
        parent = Worksheets(1).Cells(i, 2).Value
        id = Worksheets(1).Cells(i, 1).Value

        If parent <> "root" Then
            ' Try to add child, skip if parent doesn't exist
            On Error Resume Next
            Err.Clear
            nodes(parent).children.Add nodes(id)
            On Error GoTo ErrorHandler
        End If

        i = i + 1
    Wend

    ' Third loop: Add child data from Sheet 2
    ii = 2
    While Worksheets(2).Cells(ii, 1).Value <> ""
        nodeId = Worksheets(2).Cells(ii, 1).Value

        ' Try to access node, skip if it doesn't exist
        On Error Resume Next
        Err.Clear
        nodes(nodeId).children.Add New jsonChild

        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            nodes(nodeId).children(nodes(nodeId).children.Count).parent = nodeId

            Set attr = New Dictionary
            iAttr = 2
            While Worksheets(2).Cells(1, iAttr).Value <> ""
                attr.Add Worksheets(2).Cells(1, iAttr).Value, Worksheets(2).Cells(ii, iAttr).Value
                iAttr = iAttr + 1
            Wend

            nodes(nodeId).children(nodes(nodeId).children.Count).attributes = attr
            Set attr = Nothing
        End If

        On Error GoTo ErrorHandler
        ii = ii + 1
    Wend

    ' Generate JSON output
    cos = ""
    For i = 1 To nodes.Count
        If nodes(i).parent = "root" Then
            cos = toString(nodes(i))
        End If
    Next i

    TextBox1.Text = cos
    Exit Sub

ErrorHandler:
    MsgBox "Error processing data: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "ExcelJSON Error"
    TextBox1.Text = "Error: " & Err.Description
End Sub

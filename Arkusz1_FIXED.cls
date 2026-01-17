VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' FIXED VERSION - Bug fix: Added declaration for 'id' variable
Private Sub CommandButton2_Click()
    Dim i As Integer
    i = 2
    Dim nodes As New Collection

    While Worksheets(1).Cells(i, 1).Value <> ""
        On Error Resume Next
        nodes.Add New jsonNode, Worksheets(1).Cells(i, 1).Value
        nodes(nodes.Count).id = Worksheets(1).Cells(i, 1).Value
        nodes(nodes.Count).parent = Worksheets(1).Cells(i, 2).Value

        Dim nAttr As New Dictionary
        Dim niAttr As Integer
        niAttr = 3
        While Worksheets(1).Cells(1, niAttr).Value <> ""
            nAttr.Add Worksheets(1).Cells(1, niAttr).Value, Worksheets(1).Cells(i, niAttr).Value
            niAttr = niAttr + 1
        Wend

        nodes(nodes.Count).attributes = nAttr

        Set nAttr = Nothing
        i = i + 1
    Wend

    i = 2
    While Worksheets(1).Cells(i, 1).Value <> ""
        Dim parent As String
        Dim id As String  ' FIX: Added proper declaration for 'id' variable
        parent = Worksheets(1).Cells(i, 2).Value
        id = Worksheets(1).Cells(i, 1).Value
        If parent <> "root" Then
            On Error Resume Next
            nodes(parent).children.Add nodes(id)
        End If

        i = i + 1
    Wend


    Dim ii As Integer
    ii = 2
    Dim children As New Collection

    While Worksheets(2).Cells(ii, 1).Value <> ""
        nodes(Worksheets(2).Cells(ii, 1).Value).children.Add New jsonChild
        nodes(Worksheets(2).Cells(ii, 1).Value).children(nodes(Worksheets(2).Cells(ii, 1).Value).children.Count).parent = Worksheets(2).Cells(ii, 1).Value

        Dim attr As New Dictionary
        Dim iAttr As Integer
        iAttr = 2
        While Worksheets(2).Cells(1, iAttr).Value <> ""
            attr.Add Worksheets(2).Cells(1, iAttr).Value, Worksheets(2).Cells(ii, iAttr).Value
            iAttr = iAttr + 1
        Wend

        nodes(Worksheets(2).Cells(ii, 1).Value).children(nodes(Worksheets(2).Cells(ii, 1).Value).children.Count).attributes = attr

        Set attr = Nothing
        ii = ii + 1
    Wend


    Dim cos As String
    cos = ""
    For i = 1 To nodes.Count
        If nodes(i).parent = "root" Then
            cos = toString(nodes(i))
        End If
    Next i

    TextBox1.Text = cos
End Sub

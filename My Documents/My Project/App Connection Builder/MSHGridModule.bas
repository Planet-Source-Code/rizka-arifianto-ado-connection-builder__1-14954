Attribute VB_Name = "MSHGridModule"
Function MSHFlxFillGrid(ByVal aGrid As MSHFlexGrid, ByVal GrdConnection As ADODB.Connection, ByVal SQLGrid As String) As Long
    Dim rsGrid As ADODB.Recordset
    Dim i As Integer
    'Dim DummyItem As String
    
    On Error GoTo FillError
    aGrid.MousePointer = flexHourglass
    Set rsGrid = New ADODB.Recordset
    rsGrid.Open SQLGrid, GrdConnection, adOpenStatic, adLockReadOnly
    MSHFlxFillGrid = rsGrid.RecordCount
    Set aGrid.DataSource = rsGrid
    Set aGrid.DataSource = Nothing
    Set rsGrid = Nothing
    
    On Error GoTo 0
    aGrid.MousePointer = flexDefault
    Exit Function
FillError:
    MSHFlxFillGrid = -1
    Set aGrid.DataSource = Nothing
    Set rsGrid = Nothing
    aGrid.MousePointer = flexDefault
    On Error GoTo 0
End Function

Sub MSHFlxSetColFitSize(ByVal MSHGrid As MSHFlexGrid, ParamArray GrdSize() As Variant)
    Dim i As Integer
    i = 0
    For Each ColumSize In GrdSize
        MSHGrid.ColWidth(i) = Int((ColumSize * (MSHGrid.Width - 225)) / 100)
        i = i + 1
    Next ColumSize
    If MSHGrid.Cols <= i Then MSHGrid.Cols = i + 1
    MSHGrid.ColWidth(MSHGrid.Cols - 1) = 225
End Sub

Sub MSHFlxSetColSize(ByVal MSHGrid As MSHFlexGrid, ParamArray GrdSize() As Variant)
    Dim i As Integer
    i = 0
    For Each ColumSize In GrdSize
        MSHGrid.ColWidth(i) = ColumSize
        i = i + 1
    Next ColumSize
    If MSHGrid.Cols <= i Then MSHGrid.Cols = i + 1
    MSHGrid.ColWidth(MSHGrid.Cols - 1) = 300
End Sub

Sub MSHFlxSetColCaption(ByVal MSHGrid As MSHFlexGrid, ParamArray GrdCaption() As Variant)
    Dim i As Integer
    i = 0
    For Each Caption In GrdCaption
        MSHGrid.TextMatrix(0, i) = Caption
        MSHGrid.ColAlignmentFixed(i) = flexAlignCenterCenter
        i = i + 1
    Next Caption
    If MSHGrid.Cols <= i Then MSHGrid.Cols = i + 1
    MSHGrid.TextMatrix(0, MSHGrid.Cols - 1) = ""
End Sub

Sub MSHFlxSetColAlign(ByVal MSHGrid As MSHFlexGrid, ParamArray GrdAlign() As Variant)
    Dim i As Integer
    i = 0
    For Each Alignment In GrdAlign
        MSHGrid.ColAlignment(i) = Alignment
        i = i + 1
    Next Alignment
End Sub

Sub MSHFlxColSetting(ByVal MSHGrid As MSHFlexGrid, ByVal ncolIndex As Integer, ByVal GrdCaption$, ByVal Size As Long, ByVal GrdAllign As Long)
    MSHGrid.TextMatrix(MSHGrid.Row, ncolIndex) = GrdCaption
    MSHGrid.ColWidth(ncolIndex) = Size
    MSHGrid.ColAlignmentFixed(ncolIndex) = flexAlignCenterCenter
    MSHGrid.ColAlignment(ncolIndex) = GrdAllign
End Sub

Sub MSHFlxColFixSetting(ByVal MSHGrid As MSHFlexGrid, ByVal ncolIndex As Integer, ByVal GrdCaption$, ByVal Size As Long, ByVal GrdAllign As Long)
    Dim ActualSize As Long
    
    ActualSize = (Size / 100) * MSHGrid.Width
    MSHGrid.Row = 0
    MSHGrid.Col = ncolIndex
    MSHGrid.Text = GrdCaption
    MSHGrid.ColWidth(ncolIndex) = ActualSize
    MSHGrid.ColAlignmentFixed(ncolIndex) = flexAlignCenterCenter
    MSHGrid.ColAlignment(ncolIndex) = GrdAllign
End Sub

Sub MSHFlxHighLight(aGrid As MSHFlexGrid)
    With aGrid
        If .Rows - 1 >= 1 Then
        .Col = 0
        .RowSel = .Row
        .ColSel = .Cols - 1
        '.ColSel = 2
        '.TopRow = .Row
        End If
    End With
End Sub

Sub MSHFlxAddItem(ByVal aGrid As MSHFlexGrid, ByVal StringToAdd As String)
    Dim StrToAdd As String
    
    StrToAdd = Replace(StringToAdd, "|", vbTab)
    With aGrid
        varEmptyRow = Trim(CStr(.Text))
        nRow = .Row
        .AddItem StrToAdd
        If varEmptyRow = "" Then .RemoveItem nRow
        .Row = .Rows - 1
    End With
    'Call MSHFlxHighLight(aGrid)
End Sub

Function MSHFlxRemoveItem(ByVal aGrid As MSHFlexGrid, ByVal Index As Long) As Boolean
    With aGrid
    If .Rows = .FixedRows + 1 Then
        .Clear
        MSHFlxRemoveItem = False
    Else
        .RemoveItem Index
        MSHFlxRemoveItem = True
    End If
    '.Refresh
    End With
    'Call MSHFlxHighLight(aGrid)
End Function

Sub MSHFlxEditItem(ByVal aGrid As MSHFlexGrid, ByVal StrReplacement As String)
    Dim StrEdit As String
    
    StrEdit = Replace(StrReplacement, "|", vbTab)
    aGrid.Clip = StrEdit
    'Call MSHFlxHighLight(aGrid)
End Sub


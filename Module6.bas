Attribute VB_Name = "Module6"
Sub DCDP()
Dim currentTableIndex As Integer
Dim currentColumnIndex As Integer
Dim currentRowIndex As Integer
Dim claimelement As String
Dim claimelements() As String
Dim element As Integer
Dim oTable As Table
Dim oRow As Row
Dim oColumn As Column
Dim Cell1 As cell
If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Sub
    End If
currentTableIndex = ActiveDocument.Range(0, Selection.Tables(1).Range.End).Tables.Count
currentColumnIndex = Selection.Cells(1).ColumnIndex
currentRowIndex = Selection.Cells(1).RowIndex
Set oTable = ActiveDocument.Tables(currentTableIndex)
oTable.Columns(currentColumnIndex).Select
For Each Cell1 In Selection.Cells
If Cell1.Range.Characters(2) = "P" Then
Cell1.Shading.ForegroundPatternColor = RGB(240, 212, 120)
End If
If Cell1.Range.Characters(2) = "C" Then
Cell1.Shading.ForegroundPatternColor = RGB(160, 225, 130)
End If
If Cell1.Range.Characters(2) = "A" Then
Cell1.Shading.ForegroundPatternColor = wdColorWhite
End If
Next Cell1
End Sub

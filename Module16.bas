Attribute VB_Name = "Module16"
Sub claimsScr1()
  ' Declare variables
   
    Dim i As Integer
    
    Dim currentTableIndex As Integer
    Dim currentColumnIndex As Integer
    Dim currentRowIndex As Integer
    Dim oTable As Table
    Dim oRow As Column
    Dim oColumn As Column
    Dim oCell As cell
    Dim rng As Range



Dim charcount As Variant
Dim claims As Variant
Dim claim() As String
Dim claimnum As String
Dim claimnums() As String
claimnum = "1. "
m = 1

Dim clipboard As Object
    Dim clipboardData As Variant
For l = 2 To 50
claimnum = claimnum + "," + CStr(l) + ". "
Next l
claimnums = Split(claimnum, ",")

Set clipboard = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
clipboard.GetFromClipboard
If clipboard.GetFormat(1) Then ' 1 represents text format
    clipboardData = clipboard.GetText
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Clipboard data: " & clipboardData & vbCrLf & "Is this the required data?", vbQuestion + vbYesNo)
    If userResponse = vbNo Then
        ' Exit the subroutine
        Exit Sub
       
    End If
Else
    MsgBox "No text data found in the clipboard."
End If

If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Sub
    End If
    
    
Set oTable = Selection.Tables(1)
     'Set the URL for the search query
      
    
   
         
     Set oCell = Selection.Cells(1)
    
    charcount = oCell.RowIndex
     
      
    claim = Split(Replace(clipboardData, "What is claimed is:" & Chr(10), ""), Chr(10))
   

 For Each claims In claim
 
      
            
      oTable.Rows.Add
      If (Left(claims, 2) + " ") = claimnums(m) Or (Left(claims, 3) + " ") = claimnums(m) Then
      m = m + 1
      i = 0
      oTable.Rows.Add
     oTable.cell(charcount, oCell.ColumnIndex + 1).Merge oTable.cell(charcount, oCell.ColumnIndex)
     oTable.cell(charcount, oCell.ColumnIndex).Range.Text = "Claim " & Format(m, "00")
     oTable.cell(charcount, oCell.ColumnIndex).Range.Bold = True
     oTable.cell(charcount, oCell.ColumnIndex).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
     oTable.cell(charcount, oCell.ColumnIndex).Range.Font.Name = "Arial"
     oTable.cell(charcount, oCell.ColumnIndex).Range.Font.Size = 10
     oTable.cell(charcount, oCell.ColumnIndex).Range.Shading.BackgroundPatternColorIndex = wdGray25
     oTable.cell(charcount, oCell.ColumnIndex).Range.ParagraphFormat.SpaceAfter = 6
     oTable.cell(charcount, oCell.ColumnIndex).Range.ParagraphFormat.SpaceBefore = 6
     oTable.cell(charcount, oCell.ColumnIndex).Range.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
     charcount = charcount + 1
     End If
     oTable.cell(charcount, oCell.ColumnIndex).Select
     
     
     claims = Replace(claims, m & ".", "")
            
     Selection.Text = m & "." & i & " " & Replace(Trim(claims), Chr(13), "")
     i = i + 1
     oTable.cell(charcount, oCell.ColumnIndex).Range.Bold = False
     oTable.cell(charcount, oCell.ColumnIndex).Range.ParagraphFormat.Alignment = wdAlignParagraphJustify
     oTable.cell(charcount, oCell.ColumnIndex).Range.Font.Name = "Arial"
     oTable.cell(charcount, oCell.ColumnIndex).Range.Font.Size = 10
     oTable.cell(charcount, oCell.ColumnIndex).Range.Shading.BackgroundPatternColorIndex = wdNoHighlight
     oTable.cell(charcount, oCell.ColumnIndex).Range.ParagraphFormat.SpaceBefore = 6
     oTable.cell(charcount, oCell.ColumnIndex).Range.ParagraphFormat.SpaceAfter = 6
     oTable.cell(charcount, oCell.ColumnIndex).Range.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
     
        
     charcount = charcount + 1
     Next claims
     

End Sub


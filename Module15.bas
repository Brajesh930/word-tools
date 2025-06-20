Attribute VB_Name = "Module15"
Sub GetDCDPdata()
    Dim tbl As Table
    Dim cell As cell
    Dim rng As Range
    Dim firstLine As String
    Dim cellText As String
    Dim startCell As cell
    Dim cellIndex As Long
    Dim firstChar As String
    Dim output As String
    Dim DataObj As New MSForms.DataObject
    
    ' Initialize the output string
    output = ""
    
    ' Check if the selection is inside a table
    If Selection.Information(wdWithInTable) Then
        ' Set the table to the current table where the cursor is located
        Set tbl = Selection.Tables(1)
        
        ' Set the start cell to the cell where the cursor is currently located
        Set startCell = Selection.Cells(1)
        
        ' Initialize the cell index
        cellIndex = 1
        
        ' Loop through all cells in the current table starting from the current cell
        For Each cell In tbl.Range.Cells
            ' Check if the loop has reached the start cell or is past it
            If cell.RowIndex >= startCell.RowIndex And cell.ColumnIndex >= startCell.ColumnIndex Then
                ' Set the range to the cell
                Set rng = cell.Range
                
                ' Remove any cell markers (like end-of-cell markers)
                rng.End = rng.End - 1
                
                ' Get the text from the cell
                cellText = Trim(rng.Text)
                
                ' Get the first character of the text
                firstChar = Left(cellText, 1)
                
                ' Determine how much text to output based on the first character
                If IsNumeric(firstChar) Then
                    ' If the first character is a number, output only the first word
                    firstLine = Split(cellText, " ")(0)
                ElseIf firstChar Like "[A-Za-z]" Then
                    ' If the first character is an alphabet letter, output the first line
                    If InStr(cellText, vbCr) > 0 Then
                        firstLine = Left(cellText, InStr(cellText, vbCr) - 1)
                    Else
                        firstLine = cellText  ' If there's no line break, output the full text
                    End If
                    
                    ' Check for specific phrases in the first line
                    Select Case Trim(firstLine)
                        Case "Disclosed Completely"
                            firstLine = "DC"
                        Case "Disclosed Partially"
                            firstLine = "DP"
                        Case "NA"
                            firstLine = "NA"
                    End Select
                Else
                    firstLine = cellText  ' Default case, output the full text
                End If
                
                ' Output the result to the Immediate window without any prefix
                Debug.Print firstLine
                
                ' Append the result to the output string with a new line
                output = output & firstLine & vbCrLf
                
                ' Increment the cell index
                cellIndex = cellIndex + 1
            End If
        Next cell
        
        ' Copy the output to the clipboard
        DataObj.SetText output
        DataObj.PutInClipboard
        
    Else
        MsgBox "The cursor is not currently inside a table.", vbExclamation, "Error"
    End If
End Sub

Attribute VB_Name = "Module13"
Sub resultoverview()
    On Error GoTo ErrorHandler
    
    Dim driver As New WebDriver
    Dim url As String
    Dim oTable As Table
    Dim oCell As cell
    
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Sub
    End If
    
    driver.Start "chrome"
    driver.Window.Maximize
    
    Set oTable = Selection.Tables(1)
    
    For Each oCell In Selection.Cells
        oCell.Select
        url = "https://portal.unifiedpatents.com/patents/patent/" + FormatPatentNumber(CStr(Selection.Text))
        driver.Get url
        driver.Wait 5000
        
        oTable.cell(oCell.RowIndex, oCell.ColumnIndex + 1).Select
        Selection.Text = ExtractTitle(driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/h2", timeout:=200).Text)
        
        oTable.cell(oCell.RowIndex, oCell.ColumnIndex + 2).Select
        Selection.Text = Format(CleanAndAddOneDay(driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/div[1]/div/div[2]/div/div[2]/div[1]", timeout:=200).Text), "mmmm dd, yyyy")
        
        oTable.cell(oCell.RowIndex, oCell.ColumnIndex + 3).Select
        Selection.Text = CStr(driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/div[1]/div/div[2]/div/div[3]/div[3]/p", timeout:=200).Text)
    Next oCell
    
    driver.Quit
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: delay in internet " & Err.Description
End Sub
' Helper Functions for formatting patent numbers, extracting titles, and adjusting dates.
Function FormatPatentNumber(patentNumber As String) As String
    Dim countryCode As String
    Dim kindCode As String
    Dim numberPart As String
    Dim i As Integer
    
    ' Extract the country code (first two characters)
    countryCode = Left(patentNumber, 2)
    
    ' Loop to extract the number and kind code
    For i = 3 To Len(patentNumber)
        If Mid(patentNumber, i, 1) Like "[A-Za-z]" Then
            kindCode = Mid(patentNumber, i)
            Exit For
        Else
            numberPart = numberPart & Mid(patentNumber, i, 1)
        End If
    Next i
    
    ' Construct the formatted string
    FormatPatentNumber = countryCode & "-" & numberPart & "-" & kindCode
End Function

Function ExtractTitle(fullText As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim cleanedTitle As String
    
    ' Remove line breaks and extract the title
    fullText = Replace(fullText, vbCr, "")
    fullText = Replace(fullText, vbLf, "")
    startPos = InStr(fullText, " - ")
    endPos = InStr(fullText, "Find Prior Art")
    If Not endPos > 0 Then endPos = InStr(fullText, "Report Error")
    
    If startPos > 0 And endPos > 0 Then
        cleanedTitle = Mid(fullText, startPos + 3, endPos - (startPos + 3))
        cleanedTitle = Trim(cleanedTitle)
    Else
        cleanedTitle = fullText
    End If
    
    ExtractTitle = cleanedTitle
End Function

Function CleanAndAddOneDay(dateString As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim datePart As Date
    Dim result As String
    
    ' Regular expression to match dates in YYYY-MM-DD format
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{4}-\d{2}-\d{2}"
    regex.Global = True
    
    ' Find and adjust each date
    Set matches = regex.Execute(dateString)
    result = ""
    
    For Each match In matches
        datePart = CDate(match.value)
        datePart = datePart + 1
        result = result & Format(datePart, "yyyy-mm-dd") & ", "
    Next match
    
    ' Remove trailing comma
    If Len(result) > 0 Then result = Left(result, Len(result) - 2)
    
    CleanAndAddOneDay = result
End Function

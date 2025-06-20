Attribute VB_Name = "Module11"
Sub ScrapePatentTitlesWithFormatting()
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim driver As New WebDriver
    Dim url As String
    Dim oTable As Table
    Dim rng As Range
    Dim Inventors As String
    Dim baseXPath As String

    
    ' Ensure that the selection is within a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Sub
    End If

    ' Set the URL for the patent query
    url = "https://portal.unifiedpatents.com/patents/patent/" + FormatPatentNumber(CStr(Selection.Text))

    ' Start a new Chrome browser session
    driver.Start "chrome", url
    driver.Get url
    driver.Window.Maximize
    driver.Wait (5000)

    ' Set reference to the first table
    Set oTable = Selection.Tables(1)

    ' Example: Insert title without formatting
    Set rng = oTable.cell(1, 2).Range
    rng.MoveEnd wdCharacter, -1 ' Adjust the range to exclude the last paragraph mark
    rng.Text = ExtractTitle(driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/h2", timeout:=200).Text) ' Insert the title text

    ' Example: Apply formatting before inserting additional data for abstract
    Set rng = oTable.cell(2, 2).Range
    rng.MoveEnd wdCharacter, -1
    rng.Select
    Selection.Collapse wdCollapseEnd ' Collapse to the end of the cell content
    Call FormatTableCell ' Apply formatting only to the newly inserted text
    Selection.TypeText Text:=driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div", timeout:=200).Text
    

    ' Example: Formatting for date publication
    Set rng = oTable.cell(3, 1).Range
    rng.MoveEnd wdCharacter, -1
    rng.Select
    Selection.Collapse wdCollapseEnd ' Collapse to the end of the cell content
    Call FormatTableCell ' Apply formatting only to the newly inserted text
    Selection.TypeText Text:=Format(CleanAndAddOneDay(driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/div[1]/div/div[2]/div/div[2]/div[3]", timeout:=200).Text), "mmmm dd, yyyy")
    
    
    ' Example: Formatting for date application
    Set rng = oTable.cell(4, 1).Range
    rng.MoveEnd wdCharacter, -1
    rng.Select
    Selection.Collapse wdCollapseEnd ' Collapse to the end of the cell content
    Call FormatTableCell ' Apply formatting only to the newly inserted text
    Selection.TypeText Text:=Format(CleanAndAddOneDay(driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/div[1]/div/div[2]/div/div[2]/div[1]", timeout:=200).Text), "mmmm dd, yyyy")
    
    
    ' Example: Formatting for date priority
    Set rng = oTable.cell(2, 3).Range
    rng.MoveEnd wdCharacter, -1
    rng.Select
    Selection.Collapse wdCollapseEnd ' Collapse to the end of the cell content
    Call FormatTableCell ' Apply formatting only to the newly inserted text
    Selection.TypeText Text:=Format(CleanAndAddOneDay(driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/div[1]/div/div[2]/div/div[2]/div[2]", timeout:=200).Text), "mmmm dd, yyyy")
   
    
    ' Example: Formatting for assignee
    Assignee = driver.FindElementByXPath("/html/body/div[1]/div/div[2]/div/main/div/div/div[1]/div/div[2]/div/div[3]/div[3]/p", timeout:=200).Text
    Set rng = oTable.cell(3, 3).Range
    rng.MoveEnd wdCharacter, -1
    rng.Select
    Selection.Collapse wdCollapseEnd ' Collapse to the end of the cell content
    Call FormatTableCell ' Apply formatting only to the newly inserted text
    Selection.TypeText Text:=Assignee
    
    ' Insert additional content with new formatting for cell (4, 3)
    baseXPath = "/html/body/div[1]/div/div[2]/div/main/div/div/div[1]/div/div[2]/div/div[1]/div[5]/"
    Set rng = oTable.cell(4, 3).Range
    rng.MoveEnd wdCharacter, -1 ' Adjust the range to exclude the last paragraph mark
    rng.Select
    Selection.Collapse wdCollapseEnd ' Collapse the selection to the end to ensure formatting applies only to new content
    Call FormatTableCell ' Apply formatting only to the new content
    Selection.TypeText Text:=GetInventorsCommaSeparated(driver, baseXPath, 20) ' Add inventor name

    ' Apply formatting to newly inserted content only
    
    
    ' Close the browser
    driver.Quit
    Exit Sub

ErrorHandler:
    ' Handle any errors
    MsgBox "An error occurred: " & Err.Description
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

Sub FormatTableCell()
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.5)
        .Alignment = wdAlignParagraphJustify
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .ColorIndex = wdBlack
    End With
End Sub
Function GetInventorsCommaSeparated(driver As Object, baseXPath As String, timeout As Integer) As String
    Dim i As Integer
    Dim inventor As String
    Dim Inventors As String
    Dim xpath As String
    

    ' Initialize variables
    Inventors = ""
   

    'If no single inventor is found, loop through multiple inventors using p[1], p[2], etc.
        For i = 1 To 15
            ' Construct the XPath for each inventor (e.g., p[1], p[2], etc.)
            xpath = baseXPath & "p[" & i & "]"
            
            ' Reset the inventor variable
            inventor = ""

            On Error Resume Next
            inventor = driver.FindElementByXPath(xpath, timeout).Text
            
            
            ' If no inventor is found, exit the loop
            If Len(inventor) = 0 Then Exit For
            
            ' Append inventor names to the string, separated by commas
            If Len(Inventors) > 0 Then
                Inventors = Inventors & ", " & inventor
            Else
                Inventors = inventor
            End If
        Next i
   

    ' Return the inventor names as a comma-separated string
    GetInventorsCommaSeparated = Inventors
End Function



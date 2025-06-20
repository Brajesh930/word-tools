Attribute VB_Name = "Module20"
' Main subroutine to scrape patent details from Unified Patents API and insert them into a Word table
Sub ScrapePatentDataPatentCenterAPI()
    Dim url As String
    Dim oTable As Table
    Dim rng As Range
    Dim Inventors As Object, Assignees As Object
    Dim inventorList As String, AssigneeList As String
    Dim inventor As Variant, Assignee As Variant
    Dim httpReq As Object
    Dim jsonResponse As Object
    Dim continueScript As VbMsgBoxResult
    Dim patentNumber As String

    ' Ensure the selection is within a Word table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    patentNumber = Trim(Selection.Text)
    patentNumber = Replace(patentNumber, Chr(13) & Chr(7), "") ' Clean selection text

    ' Construct API URL with formatted patent number from selected text
    url = "https://api.unifiedpatents.com/patents/" & FormatPatentNumber(patentNumber)

    ' Make HTTP GET request to fetch patent details
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    httpReq.Open "GET", url, False
    httpReq.Send

    ' Check if HTTP response status is OK (200)
    If httpReq.Status = 200 Then
        Set jsonResponse = JsonConverter.ParseJson(httpReq.responseText)

        Set oTable = Selection.Tables(1)

        '--- Insert Patent Title ---
        On Error GoTo MissingElement
        Set rng = oTable.cell(1, 2).Range
        rng.MoveEnd wdCharacter, -1
        rng.Text = jsonResponse("_source")("title")

        '--- Insert Patent Abstract ---
        Set rng = oTable.cell(2, 2).Range
        rng.MoveEnd wdCharacter, -1
        rng.Select
        Selection.Collapse wdCollapseEnd
        Call FormatTableCell
        Selection.TypeText Text:=jsonResponse("_source")("abstract")

        '--- Insert Publication Date ---
        Set rng = oTable.cell(4, 3).Range
        rng.MoveEnd wdCharacter, -1
        rng.Select
        Selection.Collapse wdCollapseEnd
        Call FormatTableCell
        Selection.TypeText Text:=FormatJsonDate(CStr(jsonResponse("_source")("publication_date")))

        '--- Insert Priority Date ---
        Set rng = oTable.cell(2, 3).Range
        rng.MoveEnd wdCharacter, -1
        rng.Select
        Selection.Collapse wdCollapseEnd
        Call FormatTableCell
        Selection.TypeText Text:=FormatJsonDate(CStr(jsonResponse("_source")("priority_date")))

        '--- Insert Application Date ---
        Set rng = oTable.cell(3, 3).Range
        rng.MoveEnd wdCharacter, -1
        rng.Select
        Selection.Collapse wdCollapseEnd
        Call FormatTableCell
        Selection.TypeText Text:=FormatJsonDate(CStr(jsonResponse("_source")("application_date")))

        '--- Insert Assignee Names ---
        Set Assignees = jsonResponse("_source")("assignee_current")
        AssigneeList = ""
        For Each Assignee In Assignees
            AssigneeList = AssigneeList & Assignee & ", "
        Next Assignee
        If Len(AssigneeList) > 2 Then AssigneeList = Left(AssigneeList, Len(AssigneeList) - 2)

        Set rng = oTable.cell(3, 1).Range
        rng.MoveEnd wdCharacter, -1
        rng.Select
        Selection.Collapse wdCollapseEnd
        Call FormatTableCell
        Selection.TypeText Text:=AssigneeList

        '--- Insert Inventor Names ---
        Set Inventors = jsonResponse("_source")("inventors")
        inventorList = ""
        For Each inventor In Inventors
            inventorList = inventorList & inventor & ", "
        Next inventor
        If Len(inventorList) > 2 Then inventorList = Left(inventorList, Len(inventorList) - 2)

        Set rng = oTable.cell(4, 1).Range
        rng.MoveEnd wdCharacter, -1
        rng.Select
        Selection.Collapse wdCollapseEnd
        Call FormatTableCell
        Selection.TypeText Text:=inventorList

        MsgBox "Patent data fetched successfully!", vbInformation
        Exit Sub

    Else
        MsgBox "Failed to fetch patent data. HTTP Status: " & httpReq.Status, vbCritical
        Exit Sub
    End If

    Exit Sub

'--- Error Handling: Missing JSON Element ---
MissingElement:
    continueScript = MsgBox("The element '" & Err.Description & "' was not found for patent number: " & patentNumber & _
                            vbCrLf & "Do you want to end the script?", vbYesNo + vbExclamation, "Missing Element")
    If continueScript = vbYes Then
        Exit Sub
    Else
        Resume Next
    End If

End Sub

' Format JSON date string into "mmmm dd, yyyy"
Function FormatJsonDate(dateString As String) As String
    On Error GoTo DateError

    Dim dateOnly As String
    dateOnly = Left(dateString, 10)

    Dim dt As Date
    dt = DateSerial(CInt(Left(dateOnly, 4)), _
                    CInt(Mid(dateOnly, 6, 2)), _
                    CInt(Right(dateOnly, 2)))

    FormatJsonDate = Format(dt, "mmmm dd, yyyy")
    Exit Function

DateError:
    FormatJsonDate = "Invalid Date"
End Function

' Format patent number for API compatibility (e.g., "US123456A1" to "US-123456-A1")
Function FormatPatentNumber(patentNumber As String) As String
    Dim http As Object
    Dim jsonBody As String
    Dim responseText As String
    Dim json As Object

    ' JSON request payload
    jsonBody = "{""publications"": [""" & patentNumber & """]}"

    ' Create and send HTTP POST
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", "https://api.unifiedpatents.com/helpers/transform-publication-numbers", False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send jsonBody

    ' Check if response is OK
    If http.Status = 200 Then
        responseText = http.responseText
        Debug.Print "Response: " & responseText
        Dim cleanedString As String
        
cleanedString = Replace(responseText, "[", "")
cleanedString = Replace(cleanedString, "]", "")
cleanedString = Replace(cleanedString, """", "")  ' Remove quotes

        FormatPatentNumber = cleanedString
        Debug.Print cleanedString
       
    Else
        FormatPatentNumber = patentNumber
    End If
End Function

' Apply standardized formatting to table cell content
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



Attribute VB_Name = "Module21"
' Main subroutine: Scrape patent data from Unified Patents API with robust error handling
Sub ScrapePatentTitlesResultOverView()

    Dim url As String
    Dim oTable As Table
    Dim oCell As cell
    Dim httpReq As Object
    Dim jsonResponse As Object
    Dim Assignees As Object
    Dim AssigneeList As String
    Dim Assignee As Variant
    Dim continueScript As VbMsgBoxResult
    Dim patentNumber As String

    ' Ensure the cursor is within a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor within a table.", vbExclamation
        Exit Sub
    End If

    Set oTable = Selection.Tables(1)

    ' Iterate through each cell in the selection
    For Each oCell In Selection.Cells

        patentNumber = Trim(oCell.Range.Text)
        patentNumber = Replace(patentNumber, Chr(13) & Chr(7), "") ' Clean cell content

        ' Format URL
        url = "https://api.unifiedpatents.com/patents/" & FormatPatentNumber(patentNumber)

        ' HTTP Request
        Set httpReq = CreateObject("MSXML2.XMLHTTP")
        httpReq.Open "GET", url, False
        httpReq.Send

        ' Verify HTTP status
        If httpReq.Status = 200 Then
            Set jsonResponse = JsonConverter.ParseJson(httpReq.responseText)

            ' Insert Title
            On Error GoTo ElementMissing
            oTable.cell(oCell.RowIndex, oCell.ColumnIndex + 1).Range.Text = _
                CStr(jsonResponse("_source")("title"))

            ' Insert Priority Date
            oTable.cell(oCell.RowIndex, oCell.ColumnIndex + 2).Range.Text = _
                FormatJsonDate(CStr(jsonResponse("_source")("priority_date")))

            ' Insert Assignee
            Set Assignees = jsonResponse("_source")("assignee_current")
            AssigneeList = ""
            For Each Assignee In Assignees
                AssigneeList = AssigneeList & Assignee & ", "
            Next Assignee

            If Len(AssigneeList) > 2 Then
                AssigneeList = Left(AssigneeList, Len(AssigneeList) - 2)
            End If

            oTable.cell(oCell.RowIndex, oCell.ColumnIndex + 3).Range.Text = CStr(AssigneeList)
            

            ' Optional: Format each inserted cell (call if needed)
            'Call FormatTableCell

        Else
            continueScript = MsgBox("Failed to fetch data for patent: " & patentNumber & _
                                    ". HTTP Status: " & httpReq.Status & vbCrLf & _
                                    "Do you want to continue?", vbYesNo + vbExclamation)
            If continueScript = vbNo Then Exit Sub
        End If

NextCell:
        ' Continue to next cell
        On Error GoTo 0
    Next oCell

    MsgBox "Patent data fetching completed.", vbInformation
    Exit Sub

'--- Error Handling for Missing Elements ---
ElementMissing:
    continueScript = MsgBox("Missing element in patent data for patent: " & patentNumber & vbCrLf & _
                            "Error: " & Err.Description & vbCrLf & _
                            "Do you want to continue with next patent?", vbYesNo + vbCritical)
    If continueScript = vbYes Then
        Resume NextCell
    Else
        Exit Sub
    End If

End Sub

'--- Function to format JSON date into readable form ---
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

'--- Helper to format patent numbers for API ---
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

'--- Optional Formatting for Table Cells ---
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



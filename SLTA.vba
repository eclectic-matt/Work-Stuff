'Code to split and extract Student-Led Teaching Award Nominations and generate emails and documentation

' GENERATE LISTS OF WINNERS FOR NEWS ITEMS
Sub generateNewsItemList()

intClusterCol = 1
intSchoolCol = 2
intNameCol = 3
intAwardCol = 4

'Start at row 2 to avoid the header
ActiveSheet.Cells(2, intClusterCol).Select

'Clear the current cluster variable
strCurrentCluster = ActiveCell.Value
strCurrentSchool = ActiveCell.Offset(0, 1).Value
'strCurrentName = ActiveCell.Offset(0, 2).Value
'strCurrentAward = ActiveCell.Offset(0, 3).Value
strNewsItem = ""

strNewsItem = "<h2>Complete List of Student-Led Teaching Awards Winners</h2>" & vbNewLine

strNewsItem = strNewsItem & "<h3>" & strCurrentCluster & "</h3>" & vbNewLine & _
            "<h3>" & strCurrentSchool & "</h3><br><table><tr>"

' Loop through all the rows until the cluster column is empty
Do While Not IsEmpty(ActiveCell.Value)
    
    'Are we on a new cluster?
    If (ActiveCell.Value <> strCurrentCluster) Then
        strCurrentCluster = ActiveCell.Value
        strNewsItem = strNewsItem & "</table><h3>" & strCurrentCluster & "</h3>" & vbNewLine
    End If
    
    ' Are we on a new School?
    If (ActiveCell.Offset(0, 1).Value <> strCurrentSchool) Then
        strNewsItem = strNewsItem & "</table><br>" & vbNewLine
        strCurrentSchool = ActiveCell.Offset(0, 1).Value
        strNewsItem = strNewsItem & "<h4>" & strCurrentSchool & "</h4>" & vbNewLine
        strNewsItem = strNewsItem & "<table><tr>"
    End If
    
    strCurrentName = ActiveCell.Offset(0, 2).Value
    strCurrentAward = ActiveCell.Offset(0, 3).Value
    strNewsItem = strNewsItem & "<tr><td><strong>" & strCurrentName & "</strong></td>"
    strNewsItem = strNewsItem & "<td>" & strCurrentAward & "</td>" & "</tr>" & vbNewLine
    ActiveCell.Offset(1, 0).Select
    
Loop
strNewsItem = strNewsItem & "</table>"

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile("N:/Documents/newsItemPT2_" & Format(Now(), "dd.mm.yy h.m.s") & ".txt")
oFile.WriteLine strNewsItem
oFile.Close
Set fso = Nothing
Set oFile = Nothing

End Sub


' GENERATE LISTS OF WINNERS FOR WEBPAGES
Sub generateWinnerWebpage()

intClusterCol = 2
intSchoolCol = 3
intNameCol = 4
intAwardCol = 1

'Start at row 2 to avoid the header
ActiveSheet.Cells(2, intAwardCol).Select

'Clear the current award variable
strCurrentAward = ActiveCell.Value
strWebpage = "<h4><span style='text-decoration: underline;'><strong>" & strCurrentAward & "</strong></span></h4>" & vbNewLine & "<br><ul>"

' Loop through all the rows until the cluster column is empty
Do While Not IsEmpty(ActiveCell.Value)
    
    'Are we on a new award
    If (ActiveCell.Value <> strCurrentAward) Then
        strCurrentAward = ActiveCell.Value
        strWebpage = strWebpage & "</ul><br>"
        strWebpage = strWebpage & "<h4><span style='text-decoration: underline;'><strong>" & strCurrentAward & "</strong></span></h4>" & vbNewLine & "<br><ul>"
    End If
    
    strCurrentName = ActiveCell.Offset(0, 3).Value
    strCurrentSchool = ActiveCell.Offset(0, 2).Value
    strWebpage = strWebpage & "<li>" & strCurrentName & " (" & strCurrentSchool & ")</li>" & vbNewLine
    ActiveCell.Offset(1, 0).Select
    
Loop
strWebpage = strWebpage & "</ul>"

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile("N:/Documents/webpage_" & Format(Now(), "dd.mm.yy h.m.s") & ".txt")
oFile.WriteLine strWebpage
oFile.Close
Set fso = Nothing
Set oFile = Nothing

End Sub


' GENERATE EMAILS FOR THE WINNERS, INCLUDING THEIR NOMINATION COMMENTS
Sub generateSLTAWinnerEmails()

Dim OL As Object
Dim OM As Object
    
Set OL = CreateObject("Outlook.Application")
    
Dim strIntro As String
strIntro = "I am delighted to let you know that you have won a Student-Led Teaching Award for 2016-17:"
Dim strConf As String
strConf = "All winners are invited to the Annual Teaching and Learning Conference to be presented with their certificates." & vbNewLine & _
          "The presentation will take place in the XXXX" & vbNewLine & vbNewLine & _
          "If you are not already booked onto the conference, please reply to indicate if you plan to attend the conference just for the certificate presentation." & vbNewLine & _
          "Anyone unable to attend will receive their certificate in the internal post next week."

Dim strBody As String
strBody = ""

intNameCol = 2
intEmailCol = 3
intGroupCol = 4
intAwardCol = 6
intNomCol = 7
intNumNomCol = 8
intTotalNomCol = 10

'Start at row 2 to avoid the header
ActiveSheet.Cells(2, intNameCol).Select

'Clear the variables that hold the current staff member and their nomination text
strCurrentStaff = ActiveSheet.Cells(2, intNameCol).Value
strCurrentStaff = ActiveCell.Value
strStaffEmail = ActiveCell.Offset(0, intEmailCol - intNameCol)
strTotalNoms = ActiveCell.Offset(0, intTotalNomCol - intNameCol)
strAward = ActiveCell.Offset(0, intAwardCol - intNameCol)

strNominationText = ""
intNomCount = 0

' Loop through all the rows until the name column is empty
Do While Not IsEmpty(ActiveCell.Value)
    
        'Debug.Print "--------"
        'Debug.Print "New staff member - " & strCurrentStaff
        'Debug.Print "--------"
    strNominationText = ""
    
    'Loop through all the nominations for this member of staff
    intNomCount = 0
    Do While (ActiveCell.Value = strCurrentStaff)
        
        intNomCount = intNomCount + 1
        intNomStudents = ActiveCell.Offset(0, intNumNomCol - intNameCol).Value
        strNomText = ActiveCell.Offset(0, intNomCol - intNameCol).Value
        strGrpFlag = ActiveCell.Offset(0, intGroupCol - intNameCol).Value
        If strGrpFlag = "Y" Then
            strGrpFlag = ", GROUP NOMINATION"
        Else
            strGrpFlag = ""
        End If
        strThisNom = "NOMINATION " & intNomCount & " - FROM " & intNomStudents & " STUDENT(S)" & strGrpFlag & ": " & strNomText
            'Debug.Print strThisNom
        strNominationText = strNominationText & vbNewLine & strThisNom
        ActiveCell.Offset(1, 0).Select
    
    Loop

    'Generates email with information
    Set OM = OL.CreateItem(0)
                        
    With OM
        .To = strStaffEmail
        .CC = ""
        .BCC = ""
        .Subject = "Congratulations: 2016-17 Student-Led Teaching Award Winner"
        .Body = "Dear " & strCurrentStaff & _
                vbNewLine & vbNewLine & _
                strIntro & vbNewLine & vbNewLine & _
                "*** Award: " & strAward & " ***" & vbNewLine & vbNewLine & vbNewLine & _
                "*SUPPORTING COMMENTS*" & vbNewLine & _
                "The comments from nominations supporting your award are listed below (from " & strTotalNoms & " students in total):" & vbNewLine & _
                strNominationText & vbNewLine & vbNewLine & vbNewLine & _
                "*CERTIFICATE PRESENTATION*" & vbNewLine & _
                strConf & vbNewLine & vbNewLine & _
                "Congratulations again," & vbNewLine & vbNewLine & vbNewLine & _
                  "YOUR EMAIL SIGNATURE HERE"
        '.Display
        .Send
    End With

    Set OM = Nothing
    
    ' Get things ready for the next member of staff
    strCurrentStaff = ActiveCell.Value
    strStaffEmail = ActiveCell.Offset(0, intEmailCol - intNameCol)
    strTotalNoms = ActiveCell.Offset(0, intTotalNomCol - intNameCol)
    strAward = ActiveCell.Offset(0, intAwardCol - intNameCol)

Loop

Set OL = Nothing

End Sub


' GENERATE DOCUMENTS FOR THE DECISION PANELS, SPLIT BY SCHOOL OF STUDY
Sub generateSLTADecisionDocuments()

' -- DEFINED ROWS AND VALUES
startRow = 2

schoolChecking = ""
nameChecking = ""
NominationNo = 1

'-- DEFINED COLUMNS
schoolCol = 1
nameCol = 2
GroupNomCol = 3
awardCol = 4
nomTextCol = 5
NumNomCol = 6
eduLvlCol = 7


Debug.Print "Firing up...."

'-- SET UP WORKBOOK AND SHEET
Dim thisWb As Workbook
Dim thisWs As Worksheet
Set thisWb = ActiveWorkbook
Set thisWs = thisWb.ActiveSheet

' -- SET UP WORD DOCUMENT FOR OUTPUT
'GoTo skip:
Dim wrdApp As Word.Application
Dim wrdDoc As Word.document
Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = True
Set wrdDoc = wrdApp.Documents.Add

' HEADERS AND FOOTERS AND PAGE SETUP
With wrdDoc
    .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = "Student Led Teaching Nominations 2017 - " & thisWs.Name
    .Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "SLTA Report (generated: " & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ")"
    .PageSetup.Orientation = wdOrientPortrait
End With

'-- SET UP END ROWS AND RANGES
Dim endRow As Integer
Dim totalRange As Range

With thisWs
    endRow = .Cells(.Rows.count, 1).End(xlUp).Row
    'MsgBox ("End Row = " & endRow)
    Set totalRange = .Range(.Cells(2, 1), .Cells(endRow, 9))
End With

'-- PROCESS ROW BY ROW
For i = startRow To endRow Step 1
    Debug.Print "On Row " & i
    ' Process Schools
    schoolVal = thisWs.Cells(i, schoolCol).Value
    If schoolVal <> schoolChecking Then
        ' Add this school as a heading in word doc
        If schoolChecking <> "" Then
                wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak wdPageBreak
        End If
        wrdDoc.Range(wrdDoc.Characters.count - 1).InlineShapes.AddHorizontalLineStandard
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading1)
        wrdDoc.Content.InsertAfter schoolVal
        wrdDoc.Content.InsertParagraphAfter
            Debug.Print schoolVal
        ' Set schoolChecking = schoolVal
        schoolChecking = schoolVal
    End If
    
    ' Process names
    nameVal = thisWs.Cells(i, nameCol).Value
    If nameVal <> nameChecking Then
        ' Add this name as a heading in word doc
        'wrdDoc.Content.InsertBreak wdPageBreak
        wrdDoc.Range(wrdDoc.Characters.count - 1).InlineShapes.AddHorizontalLineStandard
        'wrdDoc.Content.InsertParagraphAfter
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter nameVal
        wrdDoc.Content.InsertParagraphAfter

        ' Set nameChecking = nameVal
        nameChecking = nameVal
        NominationNo = 1
    End If
    
    '--- Process nominations
    
    ' Get values
    eduLevelVal = thisWs.Cells(i, eduLvlCol).Text
        'Debug.Print eduLevelVal
    numNomVal = thisWs.Cells(i, NumNomCol).Text
    AwardVal = thisWs.Cells(i, awardCol).Text
    If AwardVal = "" Then
        AwardVal = "None suggested"
    End If
    nomTextVal = thisWs.Cells(i, nomTextCol).Text
    groupVal = thisWs.Cells(i, GroupNomCol).Text
    If groupVal = "Y" Then
        groupVal = " PART OF A GROUP NOMINATION"
    Else
        groupVal = ""
    End If
    ' Create nomination header and body
    nomHeader = "Nomination " & NominationNo & " (" & eduLevelVal & ", " & numNomVal & " student(s) nominating" & groupVal & ")"
    nomAward = "Suggested Award: " & AwardVal
    'nomBody = "Nomination Text: " & nomTextVal
    nomBody = nomTextVal
    
    ' Insert into Word doc
    wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading3)
    wrdDoc.Content.InsertAfter nomHeader
    wrdDoc.Content.InsertParagraphAfter
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading4)
    wrdDoc.Content.InsertAfter nomAward
    wrdDoc.Content.InsertParagraphAfter
    
    wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleStrong)
    wrdDoc.Content.InsertAfter nomBody
    wrdDoc.Content.InsertParagraphAfter
    'wrdDoc.Content.InsertAfter ""
    wrdDoc.Content.InsertParagraphAfter
    
    NominationNo = NominationNo + 1

Next

Debug.Print "Complete"

' Output report as saved PDF document
wrdDoc.Activate

Filename = "G:\ar\ar_adqe\Shared\Enhancement\Teaching Awards\SLTAs 16-17\DECISION PANEL DOCS\SLTA_Decision_Docs_" & thisWs.Name & " [" & Format(Date, "dd-mm-yy") & "_" & Format(Time, "hh.mm.ss") & "]"
wrdDoc.SaveAs2 Filename:=Filename & ".docx", FileFormat:=wdFormatXMLDocument
wrdDoc.SaveAs2 Filename:=Filename & ".pdf", FileFormat:=wdFormatPDF
wrdDoc.Close (False)

wrdApp.Quit (False)
Set wrdApp = Nothing



skip:
End Sub

' ADDITIONAL AWARDS FOLLOWING ISSUE WITH NOMINATIONS (NOT ALL RECEIVED)
Sub generateAdditionalSLTADecisionDocuments()

' -- DEFINED ROWS AND VALUES
startRow = 2

schoolChecking = ""
nameChecking = ""
NominationNo = 1

'-- DEFINED COLUMNS
schoolCol = 1
nameCol = 2
GroupNomCol = 3
awardCol = 4
nomTextCol = 5
NumNomCol = 6
eduLvlCol = 7
lateFlagCol = 11
awardGivenCol = 12

Debug.Print "Firing up...."

'-- SET UP WORKBOOK AND SHEET
Dim thisWb As Workbook
Dim thisWs As Worksheet
Set thisWb = ActiveWorkbook
Set thisWs = thisWb.ActiveSheet

' -- SET UP WORD DOCUMENT FOR OUTPUT
'GoTo skip:
Dim wrdApp As Word.Application
Dim wrdDoc As Word.document
Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = True
Set wrdDoc = wrdApp.Documents.Add

' HEADERS AND FOOTERS AND PAGE SETUP
With wrdDoc
    .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = "ADDITIONAL Student Led Teaching Nominations 2017 - " & thisWs.Name
    .Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "ADDITIONAL SLTA Report (generated: " & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ")"
    .PageSetup.Orientation = wdOrientPortrait
End With

'-- SET UP END ROWS AND RANGES
Dim endRow As Integer
Dim totalRange As Range

With thisWs
    endRow = .Cells(.Rows.count, 1).End(xlUp).Row
    'MsgBox ("End Row = " & endRow)
    Set totalRange = .Range(.Cells(2, 1), .Cells(endRow, 9))
End With

'-- PROCESS ROW BY ROW
For i = startRow To endRow Step 1
    Debug.Print "On Row " & i
    ' Process Schools
    schoolVal = thisWs.Cells(i, schoolCol).Value
    If schoolVal <> schoolChecking Then
        ' Add this school as a heading in word doc
        If schoolChecking <> "" Then
                wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak wdPageBreak
        End If
        wrdDoc.Range(wrdDoc.Characters.count - 1).InlineShapes.AddHorizontalLineStandard
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading1)
        wrdDoc.Content.InsertAfter schoolVal
        wrdDoc.Content.InsertParagraphAfter
            Debug.Print schoolVal
        ' Set schoolChecking = schoolVal
        schoolChecking = schoolVal
    End If
    
    ' Process names
    nameVal = thisWs.Cells(i, nameCol).Value
    awardGivenVal = thisWs.Cells(i, awardGivenCol).Value
    If nameVal <> nameChecking Then
        ' Add this name as a heading in word doc
        'wrdDoc.Content.InsertBreak wdPageBreak
        wrdDoc.Range(wrdDoc.Characters.count - 1).InlineShapes.AddHorizontalLineStandard
        'wrdDoc.Content.InsertParagraphAfter
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter nameVal
        wrdDoc.Content.InsertParagraphAfter
        If (awardGivenVal <> "") Then
            awardGivenVal = "Award Given by Decision Panel (22 March 2017): " & awardGivenVal
            wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleStrong)
            wrdDoc.Content.InsertAfter awardGivenVal
            wrdDoc.Content.InsertParagraphAfter
        End If

        ' Set nameChecking = nameVal
        nameChecking = nameVal
        NominationNo = 1
    End If
    
    '--- Process nominations
    
    ' Get values
    eduLevelVal = thisWs.Cells(i, eduLvlCol).Text
        'Debug.Print eduLevelVal
    numNomVal = thisWs.Cells(i, NumNomCol).Text
    AwardVal = thisWs.Cells(i, awardCol).Text
    If AwardVal = "" Then
        AwardVal = "None suggested"
    End If
    nomTextVal = thisWs.Cells(i, nomTextCol).Text
    groupVal = thisWs.Cells(i, GroupNomCol).Text
    If groupVal = "Y" Then
        groupVal = " PART OF A GROUP NOMINATION"
    Else
        groupVal = ""
    End If
    lateVal = thisWs.Cells(i, lateFlagCol).Value
    If lateVal = "Y" Then
        lateVal = ", Late Nomination (after 5pm on Fri 17 March)"
    Else
        lateVal = ""
    End If
    ' Create nomination header and body
    nomHeader = "Nomination " & NominationNo & " (" & eduLevelVal & ", " & numNomVal & " student(s) nominating" & groupVal & lateVal & ")"
    nomAward = "Suggested Award: " & AwardVal
    'nomBody = "Nomination Text: " & nomTextVal
    nomBody = nomTextVal
    
    ' Insert into Word doc
    wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading3)
    wrdDoc.Content.InsertAfter nomHeader
    wrdDoc.Content.InsertParagraphAfter
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading4)
    wrdDoc.Content.InsertAfter nomAward
    wrdDoc.Content.InsertParagraphAfter
    
    wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleStrong)
    wrdDoc.Content.InsertAfter nomBody
    wrdDoc.Content.InsertParagraphAfter
    'wrdDoc.Content.InsertAfter ""
    wrdDoc.Content.InsertParagraphAfter
    
    NominationNo = NominationNo + 1

Next

Debug.Print "Complete"

' Output report as saved PDF document
wrdDoc.Activate

Filename = "G:\ar\ar_adqe\Shared\Enhancement\Teaching Awards\SLTAs 16-17\DECISION PANEL DOCS\ADDITIONAL_SLTA_Decision_Docs_" & thisWs.Name & " [" & Format(Date, "dd-mm-yy") & "_" & Format(Time, "hh.mm.ss") & "]"
wrdDoc.SaveAs2 Filename:=Filename & ".docx", FileFormat:=wdFormatXMLDocument
wrdDoc.SaveAs2 Filename:=Filename & ".pdf", FileFormat:=wdFormatPDF
wrdDoc.Close (False)

wrdApp.Quit (False)
Set wrdApp = Nothing



skip:
End Sub

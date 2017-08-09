
Sub insertSTTBuildingBlock(groupNum As Integer)

Dim oInspector As Inspector
Dim oDoc As Word.Document
Dim wordApp As Word.Application
Dim oTemplate As Word.Template
Dim oBuildingBlock As Word.BuildingBlock

Dim sttGrp As String
sttGrp = "stt g" & groupNum

Set oInspector = Application.ActiveInspector

If oInspector.EditorType = olEditorWord Then

    Set oDoc = oInspector.WordEditor
    
    Set wordApp = oDoc.Application
    Set oTemplate = wordApp.Templates(1)
    Set oBuildingBlock = oTemplate.BuildingBlockEntries(sttGrp)
    
    wordApp.Selection.EndOf Unit:=wdStory, Extend:=wdMove
    
    oBuildingBlock.Insert wordApp.Selection.Range, True
    
End If

End Sub

Sub sttBookingAutoEmail()

Dim oMail As Outlook.MailItem
Dim strMessageBody As String
Dim intGroup As Integer

If ActiveExplorer.Selection.count > 1 Then
    m = MsgBox("More than one email selected." & vbNewLine & _
                "This tool currently processes one message at a time." & vbNewLine & _
                "Please try again with only one email highlighted." & vbNewLine & vbNewLine & _
                "Exiting the tool now!", vbCritical, "ERROR - Multiple Emails")
    Exit Sub
End If

With ActiveExplorer.Selection.Item(1)
    
    strMessageBody = .Body
    
    intStartE = InStr(1, .Body, "EMAIL: ") + 6
    intEndE = InStr(1, .Body, "WORK PHONE:")
        intLenE = intEndE - intStartE
        
    strStaffEmail = Mid(.Body, intStartE, intLenE)
    
    strNewSubject = Replace(.Subject, "New form submission: ", "")
    intGroup = CInt(Right(strNewSubject, 1))

End With

strMessageBody = Replace(strMessageBody, "All submission details are stored in the database and you can view the information and download it as a spreadsheet via the WCM.", "")
strMessageBody = Replace(strMessageBody, "This is an automated email.  Please do not reply as your email will not be seen by anyone.", "")

closingComments = "If you have any questions, please email startingtoteach@sussex.ac.uk"
    strMessageBody = Replace(strMessageBody, closingComments, "", vbTextCompare)
closingComments = "Kind regards."
    strMessageBody = Replace(strMessageBody, closingComments, "", vbTextCompare)
closingComments = "Emmy Bastin"
    strMessageBody = Replace(strMessageBody, closingComments, "", vbTextCompare)
closingComments = "Course Coordinator for PGCertHE"
    strMessageBody = Replace(strMessageBody, closingComments, "", vbTextCompare)
closingComments = "--"
    strMessageBody = Replace(strMessageBody, closingComments, "", vbTextCompare)
closingComments = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    strMessageBody = Replace(strMessageBody, closingComments, "", vbTextCompare)

closingComments = "<span style='font-family:""Calibri"",sans-serif'>If you have any questions, please email startingtoteach@sussex.ac.uk<br><br>Kind regards.<br>Emmy Bastin <br>Course Coordinator for PGCertHE <br>--</span>"


Set oMail = Application.CreateItem(olMailItem)

    oMail.BodyFormat = olFormatHTML
    
    oMail.Subject = strNewSubject
    oMail.To = strStaffEmail

    oMail.Body = strMessageBody
    
    oMail.Display
    
    ' INSERTS THE QUICK PART FOR THE NUMBERED STT GROUP!
    Call insertSTTBuildingBlock(intGroup)
    
    oMail.HTMLBody = oMail.HTMLBody & closingComments & vbNewLine & S

    oMail.Display
    Set oMail = Nothing

End Sub

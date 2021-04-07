Const ArrayDim = 18
Private Selected_mail_items As Variant
Private ext As String
Private exportString As String

Public Sub to_CSV()
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    
    ext = ".csv"
    Call Mail_Scrape
    
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Sheets("Sheet1")
    xlSheet.Range("A1:S" & UBound(Selected_mail_items) + 1) = Selected_mail_items
    
    xlWB.SaveAs fileName:="C:\Users\" & Environ("UserName") & "\Desktop\" & fileName(), FileFormat:=xlCSV, CreateBackup:=False
    xlWB.Application.DisplayAlerts = False
    xlWB.Close

    Set olItem = Nothing
    Set obj = Nothing
    Set currentExplorer = Nothing
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    
End Sub

Public Sub Mail_JSON()
    ext = ".json"
    Call Mail_Scrape
    Call Array_To_JSON
    Call WriteFile
End Sub

Public Sub Mail_XML()
    ext = ".xml"
    Call Mail_Scrape
    Call Array_To_XML
    WriteFile
End Sub

Public Sub Save_Attachment()
    Dim olMail As Outlook.MailItem
    Dim objView As Explorer
    Dim MailMetadata As Variant
    Dim olAttachment As Outlook.Attachment
    Dim i As Integer
    
    Set objView = Application.ActiveExplorer

    i = 1
    On Error Resume Next
        MkDir "C:\Users\" & Environ("UserName") & "\Desktop\Attachments"
    On Error GoTo 0
    For Each omail In objView.Selection
        For Each olAttachment In omail.Attachments
            olAttachment.SaveAsFile "C:\Users\" & Environ("UserName") & "\Desktop\Attachments\" & olAttachment.fileName
        Next
    Next omail
End Sub






Private Sub Mail_Scrape()
    Call get_Selected_mail_items
    Call CleanText
End Sub

Private Function fileName() As String
    Dim sFolder As String
    Dim FileDate As String
    Dim UserName As String
    Dim tempArray As Variant
    
    FileDate = Format(Now(), "yyyymmdd")
    UserName = Environ("UserName")
    tempArray = Split(UserName, ".")
    UserName = ""
    
    
    For i = 0 To UBound(tempArray)
        If i = UBound(tempArray) Then
            UserName = UserName & tempArray(i)
        Else
            UserName = UserName & tempArray(i) & "_"
        End If
    Next i
    
    fileName = FileDate & "-" & UserName & "-" & "Mail_Scrape" & ext
    
End Function

Private Sub WriteFile()
    Dim TextFile As Integer
    Dim FilePath As String
    
    
    
    FilePath = "C:\Users\" & Environ("UserName") & "\Desktop\" & fileName
    TextFile = FreeFile
    Open FilePath For Output As TextFile
    Print #TextFile, exportString
    Close TextFile

End Sub


Private Sub CleanText()
    Dim i As Single, j As Single
    Dim myString As String
    
    myString = ""
    
    For i = 1 To UBound(Selected_mail_items)
        For j = 0 To ArrayDim
            Selected_mail_items(i, j) = Replace(Selected_mail_items(i, j), """", "'")
        Next j
    Next i

End Sub

Private Sub Array_To_JSON()
    Dim i As Single, j As Single
    Dim Array_To_JSON As String
    
    Array_To_JSON = ""
    
    For i = 1 To UBound(Selected_mail_items)
        Array_To_JSON = Array_To_JSON & "{" & vbNewLine
        For j = 0 To ArrayDim
            Array_To_JSON = Array_To_JSON & vbTab & """" & Selected_mail_items(0, j) & """: """ & Selected_mail_items(i, j) & """," & vbNewLine
        Next j
        Array_To_JSON = Array_To_JSON & "}," & vbNewLine
    Next i
    exportString = Array_To_JSON
End Sub

Private Sub Array_To_XML()
    Dim i As Single, j As Single
    Dim Array_To_XML As String
    Array_To_XML = ""
    
    For i = 1 To UBound(Selected_mail_items)
        Array_To_XML = Array_To_XML & "<Email Item>" & vbNewLine
        For j = 0 To ArrayDim
            Array_To_XML = Array_To_XML & vbTab & "<" & Selected_mail_items(0, j) & ">" & """" & Selected_mail_items(i, j) & """" & "</" & Selected_mail_items(0, j) & ">" & vbNewLine
        Next j
        Array_To_XML = Array_To_XML & "</Email Item>" & vbNewLine
    Next i
    exportString = Array_To_XML
End Sub
Private Sub get_Selected_mail_items()
    Dim olMail As Outlook.MailItem
    Dim objView As Explorer
    Dim MailMetadata As Variant
    Dim i As Integer
    
    Set objView = Application.ActiveExplorer

    i = 1
    
    For Each omail In objView.Selection
        i = i + 1
    Next omail
    
    ReDim MailMetadata(0 To i - 1, 0 To ArrayDim)
    
    MailMetadata(0, 0) = "To"
    MailMetadata(0, 1) = "CC"
    MailMetadata(0, 2) = "Reply_Recipient_Names"
    MailMetadata(0, 3) = "Sender_Email_Address"
    MailMetadata(0, 4) = "Sender_Name"
    MailMetadata(0, 5) = "Sent_On_Behalf_Of_Name"
    MailMetadata(0, 6) = "Sender_Email_Type"
    MailMetadata(0, 7) = "Sent"
    MailMetadata(0, 8) = "Size"
    MailMetadata(0, 9) = "Unread"
    MailMetadata(0, 10) = "Creation_Time"
    MailMetadata(0, 11) = "Last_Modification_Time"
    MailMetadata(0, 12) = "Sent_On"
    MailMetadata(0, 13) = "Received_Time"
    MailMetadata(0, 14) = "Importance"
    MailMetadata(0, 15) = "Received_By_Name"
    MailMetadata(0, 16) = "Received_On_Behalf_Of_Name"
    MailMetadata(0, 17) = "Subject"
    MailMetadata(0, 18) = "Body"
    
    i = 1
    
    For Each omail In objView.Selection
        MailMetadata(i, 0) = omail.To
        MailMetadata(i, 1) = omail.CC
        MailMetadata(i, 2) = omail.ReplyRecipientNames
        MailMetadata(i, 3) = omail.SenderEmailAddress
        MailMetadata(i, 4) = omail.SenderName
        MailMetadata(i, 5) = omail.SentOnBehalfOfName
        MailMetadata(i, 6) = omail.SenderEmailType
        MailMetadata(i, 7) = omail.Sent
        MailMetadata(i, 8) = omail.Size
        MailMetadata(i, 9) = omail.UnRead
        MailMetadata(i, 10) = omail.CreationTime
        MailMetadata(i, 11) = omail.LastModificationTime
        MailMetadata(i, 12) = omail.SentOn
        MailMetadata(i, 13) = omail.ReceivedTime
        MailMetadata(i, 14) = omail.Importance
        MailMetadata(i, 15) = omail.ReceivedByName
        MailMetadata(i, 16) = omail.ReceivedOnBehalfOfName
        MailMetadata(i, 17) = omail.Subject
        MailMetadata(i, 18) = omail.Body
        i = i + 1
    Next omail
    
    Selected_mail_items = MailMetadata
    
End Sub

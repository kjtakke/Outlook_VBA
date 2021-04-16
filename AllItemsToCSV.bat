Sub AllItemsToCSV()
    Dim MailMetadata() As Variant
    Dim oDeletedItems As Outlook.Folder
    Dim oFolders As Outlook.Folders
    Dim oFolders2 As Outlook.Folders

    Dim oItem As Outlook.Items
    Dim oItem2 As Outlook.Items
    
    Dim k As Double, z As Integer, i As Integer
    Dim h As Integer, f As Integer, g As Integer
    
    Set oFolders = Application.Session.Folders
    k = 1
    f = 0
    
    For z = 1 To oFolders.Count
        Set oFolders2 = oFolders.Item(z).Folders
        For i = 1 To oFolders2.Count
            For j = 1 To oFolders2.Item(i).Items.Count
                Set oItem2 = oFolders2.Item(i).Items
                For g = 1 To oItem2.Count
                    If oItem2.Item(g).Class = olMail Or _
                    oItem2.Item(g).Class = olTask Or _
                    oItem2.Item(g).Class = olMeeting Then

                    k = k + 1
                    End If
                Next
            Next
        Next
    Next
    
    ReDim MailMetadata(0 To k, 0 To 21)
    k = 1
    Set oFolders = Application.Session.Folders
    For z = 1 To oFolders.Count
        Set oFolders2 = oFolders.Item(z).Folders
        
        For i = 1 To oFolders2.Count

            For j = 1 To oFolders2.Item(i).Items.Count
                Set oItem2 = oFolders2.Item(i).Items
            
                For g = 1 To oItem2.Count
                
                On Error Resume Next
                    MailMetadata(k, 0) = oItem2.Item(g).To
                    MailMetadata(k, 1) = oItem2.Item(g).CC
                    MailMetadata(k, 2) = oItem2.Item(g).ReplyRecipientNames
                    MailMetadata(k, 3) = oItem2.Item(g).SenderEmailAddress
                    MailMetadata(k, 4) = oItem2.Item(g).SenderName
                    MailMetadata(k, 5) = oItem2.Item(g).SentOnBehalfOfName
                    MailMetadata(k, 6) = oItem2.Item(g).SenderEmailType
                    MailMetadata(k, 7) = oItem2.Item(g).Sent
                    MailMetadata(k, 8) = oItem2.Item(g).Size
                    MailMetadata(k, 9) = oItem2.Item(g).UnRead
                    MailMetadata(k, 10) = oItem2.Item(g).CreationTime
                    MailMetadata(k, 11) = oItem2.Item(g).LastModificationTime
                    MailMetadata(i, 12) = oItem2.Item(g).SentOn
                    MailMetadata(k, 13) = oItem2.Item(g).ReceivedTime
                    MailMetadata(k, 14) = oItem2.Item(g).Importance
                    MailMetadata(k, 15) = oItem2.Item(g).ReceivedByName
                    MailMetadata(k, 16) = oItem2.Item(g).ReceivedOnBehalfOfName
                    MailMetadata(k, 17) = oItem2.Item(g).Subject
                    MailMetadata(k, 18) = oItem2.Item(g).Body
                    MailMetadata(k, 19) = oItem2.Item(g).MessageClass
                    MailMetadata(k, 20) = oFolders2.Item(z).Items.Item(i).Name
                    MailMetadata(k, 21) = oItem2.Item(z).Name
                
                    k = k + 1
                Next
            Next
        'End If
        Next
    'End If
    Next
    
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
    MailMetadata(0, 19) = "Type"
    MailMetadata(0, 20) = "Mail_Box"
    MailMetadata(0, 21) = "Folder"
    
    Dim FileName As String
    Dim sFolder As String
    Dim FileDate As String
    Dim UserName As String
    Dim tempArray As Variant

    FileDate = Format(Now(), "yymmdd")
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
    FileName = FileDate & "-" & UserName & "-" & "Mail_Scrape" & ext
    
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    
    ext = ".csv"
    
    On Error Resume Next
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    
    Set xlSheet = xlWB.Sheets("Sheet1")
    
    xlSheet.Range("A1:V" & UBound(MailMetadata) + 1) = MailMetadata
    xlWB.SaveAs FileName:="C:\Users\" & Environ("UserName") & "\Desktop\" & FileName, FileFormat:=xlCSV, CreateBackup:=False
    xlWB.Application.DisplayAlerts = False
    xlWB.Close

    Set olItem = Nothing
    Set obj = Nothing
    Set currentExplorer = Nothing
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    
    
End Sub




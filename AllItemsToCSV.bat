Sub AllItemsToCSV()
    Dim MailMetadata() As Variant
    Dim k As Integer
    Dim oDeletedItems As Outlook.Folder
    Dim oFolders As Outlook.Folders
    Dim oFolders2 As Outlook.Folders

    Dim oItem As Outlook.Items
    Dim oItem2 As Outlook.Items
    
    Dim h As Integer, f As Integer
    
    Set oFolders = Application.Session.Folders
    k = 1
    f = 0
    
    
    For Each oItems In oFolders
    'If oItems.Name = Environ("UserName") & "@defence.gov.au" Then
        Set oFolders2 = oItems.Folders
        
        For i = 1 To oFolders2.Count
        'If oFolders2.Item(i).Name = "Inbox" Then
            Set oItem2 = oFolders2.Item(i).Items
            h = oFolders2.Item(i).Items.Count
            f = h + f
            
            For j = 1 To oFolders2.Item(i).Items.Count

            ReDim Preserve MailMetadata(0 To f, 0 To 21)
            
            MailMetadata(k, 0) = oFolders2.Item(i).Items.Item(j).To
            MailMetadata(k, 1) = oFolders2.Item(i).Items.Item(j).CC
            MailMetadata(k, 2) = oFolders2.Item(i).Items.Item(j).ReplyRecipientNames
            MailMetadata(k, 3) = oFolders2.Item(i).Items.Item(j).SenderEmailAddress
            MailMetadata(k, 4) = oFolders2.Item(i).Items.Item(j).SenderName
            MailMetadata(k, 5) = oFolders2.Item(i).Items.Item(j).SentOnBehalfOfName
            MailMetadata(k, 6) = oFolders2.Item(i).Items.Item(j).SenderEmailType
            MailMetadata(k, 7) = oFolders2.Item(i).Items.Item(j).Sent
            MailMetadata(k, 8) = oFolders2.Item(i).Items.Item(j).Size
            MailMetadata(k, 9) = oFolders2.Item(i).Items.Item(j).UnRead
            MailMetadata(k, 10) = oFolders2.Item(i).Items.Item(j).CreationTime
            MailMetadata(k, 11) = oFolders2.Item(i).Items.Item(j).LastModificationTime
            MailMetadata(i, 12) = oFolders2.Item(i).Items.Item(j).SentOn
            MailMetadata(k, 13) = oFolders2.Item(i).Items.Item(j).ReceivedTime
            MailMetadata(k, 14) = oFolders2.Item(i).Items.Item(j).Importance
            MailMetadata(k, 15) = oFolders2.Item(i).Items.Item(j).ReceivedByName
            MailMetadata(k, 16) = oFolders2.Item(i).Items.Item(j).ReceivedOnBehalfOfName
            MailMetadata(k, 17) = oFolders2.Item(i).Items.Item(j).Subject
            MailMetadata(k, 18) = oFolders2.Item(i).Items.Item(j).Body
            MailMetadata(k, 19) = oFolders2.Item(i).Items.Item(j).MessageClass
            MailMetadata(k, 20) = oItems.Name
            MailMetadata(k, 21) = oFolders2.Item(i).Name
            
            k = k + 1
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

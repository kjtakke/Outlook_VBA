'Updated 08 Apr 2021 : 1453

Const ArrayDim = 18
Const FileLocation As String = "Desktop"
Private Selected_mail_items As Variant
Private ext As String
Private exportString As String


Public Sub json_att()
    Dim i As Integer
    Dim TextFile As Integer
    Dim FilePath As String
    Dim FileName As String
    Dim olMail As Outlook.MailItem
    Dim objView As Explorer
    Dim olAttachment As Outlook.Attachment
    
    Set objView = Application.ActiveExplorer

    On Error Resume Next
    

    'ReplyRecipientNames, SenderEmailAddress, SenderName, SentOnBehalfOfName, ReceivedOnBehalfOfName

    exportString = ""
    For Each olMail In objView.Selection
    
        exportString = "{" & vbNewLine & vbTab & _
                            """people"" : {" & vbNewLine & vbTab & vbTab & _
                                """to"" : """ & olMail.To & """," & vbNewLine & vbTab & vbTab & _
                                """cc"" : """ & olMail.CC & """" & vbNewLine & vbTab & _
                            "}," & vbNewLine & vbTab
                            
        exportString = exportString & _
                            """names"" : {" & vbNewLine & vbTab & vbTab & _
                                """ReplyRecipientNames"" : """ & olMail.ReplyRecipientNames & """," & vbNewLine & vbTab & vbTab & _
                                """SenderName"" : """ & olMail.SenderName & """," & vbNewLine & vbTab & vbTab & _
                                """SentOnBehalfOfName"" : """ & olMail.SentOnBehalfOfName & """," & vbNewLine & vbTab & vbTab & _
                                """ReceivedOnBehalfOfName"" : """ & olMail.ReceivedOnBehalfOfName & """," & vbNewLine & vbTab & vbTab & _
                                """ReceivedByName"" : """ & olMail.ReceivedByName & """" & vbNewLine & vbTab & _
                            "}," & vbNewLine & vbTab
                            
        exportString = exportString & _
                            """time"" : {" & vbNewLine & vbTab & vbTab & _
                                """CreationTime"" : """ & olMail.CreationTime & """," & vbNewLine & vbTab & vbTab & _
                                """LastModificationTime"" : """ & olMail.LastModificationTime & """," & vbNewLine & vbTab & vbTab & _
                                """SentOn"" : """ & olMail.SentOn & """," & vbNewLine & vbTab & vbTab & _
                                """ReceivedTime"" : """ & olMail.ReceivedTime & """" & vbNewLine & vbTab & _
                            "}," & vbNewLine & vbTab
                            
        exportString = exportString & _
                            """metadata"" : {" & vbNewLine & vbTab & vbTab & _
                                """SenderEmailType"" : """ & olMail.SenderEmailType & """," & vbNewLine & vbTab & vbTab & _
                                """Size"" : " & olMail.Size & "," & vbNewLine & vbTab & vbTab & _
                                """UnRead"" : " & olMail.UnRead & "," & vbNewLine & vbTab & vbTab & _
                                """Sent"" : " & olMail.Sent & "," & vbNewLine & vbTab & vbTab & _
                                """Importance"" : " & olMail.Importance & vbNewLine & vbTab & _
                            "}," & vbNewLine & vbTab
                            
        exportString = exportString & _
                            """text"" : {" & vbNewLine & vbTab & vbTab & _
                                    """Subject"" : """ & Replace(olMail.Subject, """", "'") & """," & vbNewLine & vbTab & vbTab & _
                                    """Body"" : """ & Replace(olMail.Body, """", "'") & """" & vbNewLine & vbTab & _
                                "}" & vbNewLine & _
                        "}"
                        
        FileName = Format(olMail.SentOn, "yymmdd") & "-" & Format(olMail.ReceivedTime, "hhmmss") & "-" & olMail.SenderName & "-" & Left(olMail.Subject, 15)
        FileName = Replace(FileName, "\", "-")
        FileName = Replace(FileName, "/", "-")
        FileName = Replace(FileName, ".", "-")
        
        Debug.Print (exportString)
        FilePath = "C:\Users\" & Environ("UserName") & "\Desktop\" & FileName & ".txt" 'change to json
        FilePath = File_Exists(FilePath)
        TextFile = FreeFile
        Open FilePath For Output As TextFile
        Print #TextFile, exportString
        Close TextFile

        On Error Resume Next
        For Each olAttachment In omail.Attachments
            MkDir "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\" & FileName
            olAttachment.SaveAsFile File_Exists("C:\Users\" & Environ("UserName") & "\" & FileLocation & "\" & FileName & olAttachment.FileName)
        Next olAttachment
   
    Next olMail
    
End Sub


Public Sub Mail_CSV()
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    
    'Assigns the ext string variable with ".csv"
    ext = ".csv"
    
    'Scrapes and retrievs all mail items in to a Module level 2D Array
    Call Mail_Scrape
    
    'Common Errors are:
    '   Excel not installed
    '   System 32 dll file missing or curupt
    On Error Resume Next
    
    'Create an empty instence of Excel in memory only
    Set xlApp = CreateObject("Excel.Application")
    
    'Add an Excel Workbook to the Empty Excel Instence
    Set xlWB = xlApp.Workbooks.Add
    
    'Assign Sheet1 to xlSheet
    Set xlSheet = xlWB.Sheets("Sheet1")
    
    'Deturmine the array size | convert it in to a Range String | Paste the array to Sheet1
    xlSheet.Range("A1:S" & UBound(Selected_mail_items) + 1) = Selected_mail_items
    
    'Save the Excel Workbook to the users Desktop as a .csv
    xlWB.SaveAs FileName:="C:\Users\" & Environ("UserName") & "\Desktop\" & FileName(), FileFormat:=xlCSV, CreateBackup:=False
    
    'Stop any save/save as notifications
    xlWB.Application.DisplayAlerts = False
    
    'Close the Excel Workbook
    xlWB.Close

    'Clear all the Variables from memory
    Set olItem = Nothing
    Set obj = Nothing
    Set currentExplorer = Nothing
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    
End Sub

Public Sub Mail_JSON()
    'Assigns the ext string variable with ".json"
    ext = ".json"
    
    'Scrapes and retrievs all mail items in to a Module level 2D Array
    Call Mail_Scrape
    
    'Convet the array to json format as a single string
    Call Array_To_JSON
    
    'Writ the string to a text/json file on the users desktop
    Call WriteFile
End Sub

Public Sub Mail_XML()
    'Assigns the ext string variable with ".xml"
    ext = ".xml"
    
    'Scrapes and retrievs all mail items in to a Module level 2D Array
    Call Mail_Scrape
    
    'Convet the array to xml format as a single string
    Call Array_To_XML
    
    'Writ the string to a text/xml file on the users desktop
    WriteFile
End Sub

Public Sub Save_Attachment()
    Dim olMail As Outlook.MailItem
    Dim objView As Explorer
    Dim MailMetadata As Variant
    Dim olAttachment As Outlook.Attachment
    
    'Set the objView Objext to be the users active Outlook window
    Set objView = Application.ActiveExplorer
    
    'Common Errors include:
    '   Lack of memory due to a 32 bit system
    '   File not type recognised/corupted
    '   File to large to export due to a 32 bit system
    On Error Resume Next
    
    'Make a new folder on the users Desktop | Will skip this is it allready exists through the above error handeling
    MkDir "C:\Users\" & Environ("UserName") & "\Desktop\Attachments"
    
    'Loop through each selected mail items
    For Each omail In objView.Selection
    
        'Loop through each attachment in the selected mail item
        For Each olAttachment In omail.Attachments
        
            'Export the mail item to the users desktop
            'this uses the File_Exists function, where is identifies if teh file exists and changes/incruments the file name
            olAttachment.SaveAsFile File_Exists("C:\Users\" & Environ("UserName") & "\Desktop\Attachments\" & olAttachment.FileName)
            
        Next olAttachment
    Next omail
          
End Sub

Private Function File_Exists(fielPath As String) As String
    Dim strFileExists As String
    Dim fileExists As Boolean
    Dim temp_FileName As String, temp_FileName_Placeholder As String
    Dim temp_FileArray As Variant
    Dim temp_FileExt As String
    Dim temp_path As String
    Dim i As Integer
    
    'Look for the item (filePath)
    strFileExists = Dir(fielPath)
    
    'Does the file exist
    If strFileExists <> "" Then
    
        'Breakuup the filepath in to three components
        '   Path | File Name | File Extention
        
        'Split the filepath into an array by "."
        temp_FileArray = Split(strFileExists, ".")
        
        'Extract the File Extention by Concatenating "." * the last item in the temp_FileArray()
        temp_FileExt = "." & temp_FileArray(UBound(temp_FileArray))
        

        'Extract the File Name by through last item in the temp_FileArray()
        temp_FileName = temp_FileArray(0)
        
        'Resplit the filePath this time by "\"
        temp_FileArray = Split(fielPath, "\")
        
        'Initilise the temp_path string variable
        temp_path = ""
        
        'Loop through temp_FileArray() stopping fhort of the last array item
        For i = 0 To UBound(temp_FileArray) - 1
            
            'Concatenating all teh looped temp_FileArray() items
            temp_path = temp_path & temp_FileArray(i) & "\"
            
        Next i
        
        'Initilising the fileExists Boolean Variable which operates as a gate/switch for the below Do While Loop
        fileExists = True
        
        'Initilise the temp_FileName_Placeholder to be reset and ammended each loop
        temp_FileName_Placeholder = temp_FileName
        
        'Initilise the counter (i) to be appended to teh file name
        i = 1
        
        'While fileExists = True rename the variable by concatenating "(" & i & ")"
        Do While fileExists = True
        
            'Increment the temp_FileName_Placeholder by appending temp_FileName & "(" & i & ")"
            temp_FileName_Placeholder = temp_FileName & "(" & i & ")"
            
            
            'Check if teh appended file name exists
            If Dir(temp_path & temp_FileName_Placeholder & temp_FileExt) <> "" Then
            
            'Incrument the counter (i)
            i = i + 1
            
            Else
            
            'Return the new appended fileName
            fielPath = temp_path & temp_FileName_Placeholder & temp_FileExt
            
            'Break teh loop
            fileExists = False
            
            End If
            
        Loop
        
    Else
        
        'File does not exist and return fielPath
        fielPath = fielPath
        
    End If
    
    'Return teh new or same fiel name
    File_Exists = fielPath
End Function



'########Not Visible in Outlook#############
                        
Private Sub Mail_Scrape()
    'Scrapes and retrievs all mail items in to a Module level 2D Array
    Call get_Selected_mail_items
    
    'Replace all " in the body with ' for file formatting standards
    Call CleanText
End Sub

Private Function FileName() As String
    Dim sFolder As String
    Dim FileDate As String
    Dim UserName As String
    Dim tempArray As Variant
    
    'Convert the current date to text YYMMDD
    FileDate = Format(Now(), "yymmdd")
    
    'Convert the users profile name to text
    UserName = Environ("UserName")
    
    'Split the username by "."
    tempArray = Split(UserName, ".")
    
    'Initiate teh UserName String variable to be reformed without a "."
    UserName = ""
    
    'Loop through the User name Array | tempArray()
    For i = 0 To UBound(tempArray)
    
        'If Last item in array then
        If i = UBound(tempArray) Then
        
            'Concatenate UserName with the last array item
            UserName = UserName & tempArray(i)
            
        'Not the last item in the array
        Else
        
            'Concatenate UserName with the current array item and "_"
            UserName = UserName & tempArray(i) & "_"
            
        End If
        
    Next i
    
    'Retutn fileName by concatenating FileDate-UserName-Mail_Scrape.ext
    FileName = FileDate & "-" & UserName & "-" & "Mail_Scrape" & ext
    
End Function

Private Sub WriteFile()
    Dim TextFile As Integer
    Dim FilePath As String

    'Set the File path to the users desktop and append the filename (with extention)
    FilePath = "C:\Users\" & Environ("UserName") & "\Desktop\" & FileName
    
    'Create the file (Writes over teh existing file)
    'To write to a new file run:
    '   FilePath = File_Exists(FilePath)
    TextFile = FreeFile
    Open FilePath For Output As TextFile
    Print #TextFile, exportString
    Close TextFile

End Sub


Private Sub CleanText()
    Dim i As Single, j As Single
    Dim myString As String
    
    'Initilise myString as the cleaned string
    myString = ""
    
    'Loop through all rows (except the header) in the 2D Array | Selected_mail_items()
    For i = 1 To UBound(Selected_mail_items)
        
        'Loop through each column/dimention in the 2D Array | Selected_mail_items()
        For j = 0 To ArrayDim
        
            'Replace " with '
            Selected_mail_items(i, j) = Replace(Selected_mail_items(i, j), """", "'")
            
        Next j
        
    Next i

End Sub

Private Sub Array_To_JSON()
    Dim i As Single, j As Single
    Dim Array_To_JSON As String
    
    'Initilise Array_To_JSON as the final single string in json format
    Array_To_JSON = ""
    
    'Loop through all rows in the 2D Array | Selected_mail_items()
    For i = 1 To UBound(Selected_mail_items)
        
        'Open a json object "{"
        Array_To_JSON = Array_To_JSON & "{" & vbNewLine
        
        'Loop through each column/dimention in the 2D Array | Selected_mail_items()
        For j = 0 To ArrayDim
        
            'Append | TAB | " | current array item header | " | : | " | current array item | " | , | newline(\n equivelant)
            Array_To_JSON = Array_To_JSON & vbTab & """" & Selected_mail_items(0, j) & """: """ & Selected_mail_items(i, j) & """," & vbNewLine
        Next j
        
        'Close a json object "}"
        Array_To_JSON = Array_To_JSON & "}," & vbNewLine
        
    Next i
    
    'Make exportString (Module level string Variable) = Array_To_JSON ready for writng to a file
    exportString = Array_To_JSON
        
End Sub

Private Sub Array_To_XML()
    Dim i As Single, j As Single
    Dim Array_To_XML As String
    
    'Initilise Array_To_XML as the final single string in xml format
    Array_To_XML = ""
    
    'Loop through all rows in the 2D Array | Selected_mail_items()
    For i = 1 To UBound(Selected_mail_items)
    
        'Opep the xml document | <Email Item>
        Array_To_XML = Array_To_XML & "<Email Item>" & vbNewLine
        For j = 0 To ArrayDim
        
            'Create and write an xml object <header> | current item | </header>
            Array_To_XML = Array_To_XML & vbTab & "<" & Selected_mail_items(0, j) & ">" & """" & Selected_mail_items(i, j) & """" & "</" & Selected_mail_items(0, j) & ">" & vbNewLine
        Next j
        
        'Close the xml document | </Email Item>
        Array_To_XML = Array_To_XML & "</Email Item>" & vbNewLine
        
        
    Next i
    
    'Make exportString (Module level string Variable) = Array_To_XML ready for writng to a file
    exportString = Array_To_XML
    
End Sub
Private Sub get_Selected_mail_items()
    Dim olMail As Outlook.MailItem
    Dim objView As Explorer
    Dim MailMetadata As Variant
    Dim i As Integer
    
    'Set the objView Objext to be the users active Outlook window
    Set objView = Application.ActiveExplorer
    
    'Initilis the counter i as 1
    i = 1
    
    'Loop through each selected mail item to get a count to initilise the below 2D array | MailMetadata()
    For Each omail In objView.Selection
        i = i + 1
    Next omail
    
    'initilise the 2D Array | MailMetadata()
    ReDim MailMetadata(0 To i - 1, 0 To ArrayDim)
    
    'Add headders to the 2D Array | MailMetadata(0,?)
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
    
    'Reinitilise that counter (i) to skip the header file
    i = 1

    'Any incompatable mail items are skipped
    On Error GoTo nxt:
    
    'Loop through each selected mail item and add teh metadat to the 2D Array | MailMetadata(?>0,?)
    For Each olMail In objView.Selection
        MailMetadata(i, 0) = olMail.To
        MailMetadata(i, 1) = olMail.CC
        MailMetadata(i, 2) = olMail.ReplyRecipientNames
        MailMetadata(i, 3) = olMail.SenderEmailAddress
        MailMetadata(i, 4) = olMail.SenderName
        MailMetadata(i, 5) = olMail.SentOnBehalfOfName
        MailMetadata(i, 6) = olMail.SenderEmailType
        MailMetadata(i, 7) = olMail.Sent
        MailMetadata(i, 8) = olMail.Size
        MailMetadata(i, 9) = olMail.UnRead
        MailMetadata(i, 10) = olMail.CreationTime
        MailMetadata(i, 11) = olMail.LastModificationTime
        MailMetadata(i, 12) = olMail.SentOn
        MailMetadata(i, 13) = olMail.ReceivedTime
        MailMetadata(i, 14) = olMail.Importance
        MailMetadata(i, 15) = olMail.ReceivedByName
        MailMetadata(i, 16) = olMail.ReceivedOnBehalfOfName
        MailMetadata(i, 17) = olMail.Subject
        MailMetadata(i, 18) = olMail.Body
        
        i = i + 1
        
'Skipped Mail Item
nxt:
    'Reinitilise error to exit the subroutine is errors persist
    On Error GoTo en:
    Next olMail
    
'Persistant erros | Exit Sub
en:

    'Reinitilise error handeler to default
    On Error GoTo 0
    
    'Add MailMetadata array to Selected_mail_items (Module Level Array/Variant Variable)
    Selected_mail_items = MailMetadata
    
End Sub

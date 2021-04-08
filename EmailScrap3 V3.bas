'Updated 08-Apr-2021 | 10:18 PM

'Imporvements:
'   Added a file picker | Was required to load an instence of Excel in memory
'   to use the Excel File Library which included a file picker method that does not exist in Outlook

'Public Subroutines
'   JSON                        |   Saves all selected mail items to a selected folder as a json file and all Attachment to the matching folder
'   CSV                         |   Saves all selected items metadata in a single csv file in a specifiled file location
'   Attachments                 |   Saved all Attachments to a specified file lcation

'Private Subroutines
'   Mail_Scrape                 |   Calls get_Selected_mail_items() then cleans the data with CleanText()
'   CleanText                   |   Cleans an array of data by replacing " with '
'   get_Selected_mail_items     |   Scrapes all the metadata from selected mail items and puts it into a 2D array - Selected_mail_items()

'Private Functions
'   FolderPicker                |   Allows a user to pick a folder location through an Excel Method
'   FileName                    |   Used to clean a file name and path to remove teh "." from the Senders name
'   jsonArray                   |   Used to create json sub arraus from the To and CC Objects
'   File_Exists                 |   Identifies if a file exists, if so it adds a indexed numper at the end of the file name

Const ArrayDim = 18                             ' Number of Columns/Dimentions in the scraped mail metada array
Const FileLocation As String = "Documents"      ' [UNUSED | PLACE HOLDER] to be used in lue of a flolder picker
Private Selected_mail_items As Variant          ' An object that sores all the metada of selected mail items
Private ext As String                           ' Used to store the file extention of an exported json/xml file
Private exportString As String                  ' This is the final string of text to be written to a file
Private filePathPicked As String                ' This is the selected folder location stored as a string/text

Public Sub JSON()
    Dim i As Integer
    Dim TextFile As Integer
    Dim FilePath As String
    Dim FileName As String
    Dim olMail As Outlook.MailItem
    Dim objView As Explorer
    Dim olAttachment As Outlook.Attachment
    Dim FilePathConverter As String
    Set objView = Application.ActiveExplorer
    On Error Resume Next
    filePathPicked = FolderPicker()
    'MkDir "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\"
    exportString = ""
    For Each olMail In objView.Selection
        Dim jsonArrays As Collection
        Set jsonArrays = New Collection
        'Creating json sub Arrays
        jsonArrays.Add Item:=jsonArray(olMail.To, ";")
        jsonArrays.Add Item:=jsonArray(olMail.CC, ";")
        'Creating the main json array
        exportString = "{" & vbNewLine & vbTab & _
                            """people"" : {" & vbNewLine & vbTab & vbTab & _
                                """to"" : """ & jsonArrays(1) & """," & vbNewLine & vbTab & vbTab & _
                                """cc"" : """ & jsonArrays(2) & """" & vbNewLine & vbTab & _
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
        'Create File name
        FileName = Format(olMail.SentOn, "yymmdd") & "-" & Format(olMail.ReceivedTime, "hhmmss") & "-" & olMail.SenderName & "-" & Left(olMail.Subject, 30)
        'Remove reserved characters fron teh file name
        FileName = Replace(FileName, "\", " ")
        FileName = Replace(FileName, "/", " ")
        FileName = Replace(FileName, ".", " ")
        FileName = Replace(FileName, "|", " ")
        FileName = Replace(FileName, "*", " ")
        FileName = Replace(FileName, "*", " ")
        FileName = Replace(FileName, "?", " ")
        FileName = Replace(FileName, ":", " ")
        FileName = Replace(FileName, "<", " ")
        FileName = Replace(FileName, ">", " ")
        'Set the file path
        'FilePath = "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & ".json"
        FilePath = filePathPicked & "\" & FileName & ".json"
        'Insure the file path is unique
        FilePath = File_Exists(FilePath)
        'Write text file (.json)
        TextFile = FreeFile
        Open FilePath For Output As TextFile
        Print #TextFile, exportString
        Close TextFile
        'Extract all Attachments and place into their own folder
        'the folder name matched the wmail item json name
        On Error Resume Next
        For Each olAttachment In olMail.Attachments
        If olAttachment.FileName <> "" Then
            'MkDir "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & "\"
            MkDir filePathPicked & "\" & FileName & "\"
            'FilePathConverter = File_Exists("C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & "\" & olAttachment.FileName)
            FilePathConverter = File_Exists(filePathPicked & "\" & FileName & "\" & olAttachment.FileName)
            olAttachment.SaveAsFile FilePathConverter
        End If
        Next olAttachment
    Next olMail
End Sub

Public Sub CSV()
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    filePathPicked = FolderPicker()
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
    xlWB.SaveAs FileName:=filePathPicked & "\" & FileName(), FileFormat:=xlCSV, CreateBackup:=False
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

Public Sub Attachments()
    Dim olMail As Outlook.MailItem
    Dim objView As Explorer
    Dim MailMetadata As Variant
    Dim olAttachment As Outlook.Attachment
    filePathPicked = FolderPicker()
    'Set the objView Objext to be the users active Outlook window
    Set objView = Application.ActiveExplorer
    'Common Errors include:
    '   Lack of memory due to a 32 bit system
    '   File not type recognised/corupted
    '   File to large to export due to a 32 bit system
    On Error Resume Next
    
    'Make a new folder on the users Desktop | Will skip this is it allready exists through the above error handeling
    'MkDir "C:\Users\" & Environ("UserName") & "\Desktop\Attachments"
    
    'Loop through each selected mail items
    For Each omail In objView.Selection
        'Loop through each attachment in the selected mail item
        For Each olAttachment In omail.Attachments
            'Export the mail item to the users desktop
            'this uses the File_Exists function, where is identifies if teh file exists and changes/incruments the file name
            olAttachment.SaveAsFile File_Exists(filePathPicked & "\" & olAttachment.FileName)
        Next olAttachment
    Next omail
End Sub

'########Not Visible in Outlook#############
                        
Private Function FolderPicker() As String
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = False
    Dim fd As Office.FileDialog
    Set fd = xlApp.Application.FileDialog(msoFileDialogFolderPicker)
    Dim selectedItem As Variant
    If fd.Show = -1 Then
        For Each selectedItem In fd.SelectedItems
            FolderPicker = selectedItem
        Next
    End If
    Set fd = Nothing
        xlApp.Quit
    Set xlApp = Nothing
End Function

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

Private Function jsonArray(str As String, del As String) As String
    Dim tmpArray As Variant
    Dim tmpString As String
    'Split up the string
    tmpArray = Split(str, del)
    tmpString = "[" & vbNewLine & vbTab & tbTab & vbTab & vbTab
    For i = 0 To UBound(tmpArray)
        tmpArray(i) = Trim(tmpArray(i))
        If i = UBound(tmpArray) Then
            tmpString = tmpString & "{""email"":""" & tmpArray(i) & """}" & vbNewLine & vbTab & tbTab & vbTab
        Else
            tmpString = tmpString & "{""email"":""" & tmpArray(i) & """}," & vbNewLine & vbTab & tbTab & vbTab & vbTab
        End If
    Next i
    tmpString = tmpString & "]"
    jsonArray = tmpString
End Function

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

Sub DayView()
    Load_ERP_Calendars
    Dim objView As CalendarView
    Set objView = Application.ActiveExplorer.CurrentView
    On Error Resume Next
    With objView
        .CalendarViewMode = olCalendarViewDay
        .Save
    End With
    On Error GoTo 0
End Sub

Sub ThreeDayView()
    Load_ERP_Calendars
    Dim objView As CalendarView
    Set objView = Application.ActiveExplorer.CurrentView
    On Error Resume Next
    With objView
        .CalendarViewMode = olCalendarViewMultiDay
        .Save
    End With
    On Error GoTo 0
End Sub

Sub WorkWeekView()
    Load_ERP_Calendars
    Dim objView As CalendarView
    Set objView = Application.ActiveExplorer.CurrentView
    
    On Error Resume Next
    With objView
        .CalendarViewMode = olCalendarView5DayWeek
        .Save
    End With
    On Error GoTo 0
End Sub

Sub WeekView()
    Load_ERP_Calendars
    Dim objView As CalendarView
    Set objView = Application.ActiveExplorer.CurrentView
    On Error Resume Next
    With objView
        .CalendarViewMode = olCalendarViewWeek
        .Save
    End With
    On Error GoTo 0
End Sub

Sub MonthView()
    Load_ERP_Calendars
    Dim objView As CalendarView
    Set objView = Application.ActiveExplorer.CurrentView
    On Error Resume Next
    With objView
        .CalendarViewMode = olCalendarViewMonth
        .Save
    End With
    On Error GoTo 0
End Sub

Sub Load_ERP_Calendars()
    
    Dim objPane As Outlook.NavigationPane
    Dim objModule As Outlook.CalendarModule
    Dim objGroup As Outlook.NavigationGroup
    Dim objNavFolder As Outlook.NavigationFolder
    Dim objCalendar As Folder
    Dim objFolder As Folder
    
    Application.Session.GetDefaultFolder(olFolderCalendar).Display
    
    Dim i As Integer
    
    Set Application.ActiveExplorer.CurrentFolder = Session.GetDefaultFolder(olFolderCalendar)
    DoEvents
    
    Set objCalendar = Session.GetDefaultFolder(olFolderCalendar)
    Set objPane = Application.ActiveExplorer.NavigationPane
    Set objModule = objPane.Modules.GetNavigationModule(olModuleCalendar)

    For j = 1 To objModule.NavigationGroups.Count
        For i = 1 To objModule.NavigationGroups.Item(j).NavigationFolders.Count
            Set objNavFolder = objModule.NavigationGroups.Item(j).NavigationFolders.Item(i)
            If InStr(objNavFolder.DisplayName, "ERP") <> 0 Or InStr(objNavFolder.DisplayName, "SAP") <> 0 Then
                On Error Resume Next
                objNavFolder.IsSelected = True
                objNavFolder.IsSideBySide = False
                On Error GoTo 0
            End If
        Next i
    Next j
    
    Set objPane = Nothing
    Set objModule = Nothing
    Set objGroup = Nothing
    Set objNavFolder = Nothing
    Set objCalendar = Nothing
    Set objFolder = Nothing
End Sub

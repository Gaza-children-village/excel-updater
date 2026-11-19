Attribute VB_Name = "GcvLoader"
Option Explicit

' ============================================================================
' GCV Loader - the small, stable add-in that Excel opens at start-up. It opens
' every GCV add-in installed for this user:
'
'   %LOCALAPPDATA%\GCV\ExcelTools\GCV Loader.xlam            (this file)
'   %LOCALAPPDATA%\GCV\ExcelTools\addins\GCV Logger.xlam      (installed add-ins)
'   %LOCALAPPDATA%\GCV\ExcelTools\previous\GCV Logger.xlam    (version before, fallback)
'   %LOCALAPPDATA%\GCV\ExcelTools\pending\GCV Updater.xlam    (downloaded, used at next start)
'
' Add-in file names never change (formulas such as =TO_ENGLISH(A2) keep
' working after updates). GCV Updater installs most updates while Excel runs;
' an update it cannot apply then (its own, or a file another Excel window
' holds) waits in "pending" and is moved into place here, before anything is
' opened. An add-in that does not start (open fails or GcvAddinPing does not
' answer) is replaced by its previous version, and that version is marked
' "bad" so the updater does not install it again.
'
' The add-ins are opened when the first workbook is opened or created, not
' while Excel starts: opening files during start-up makes Excel skip its
' start screen (new blank workbook, templates, recent files).
'
' Keep this module small: the loader itself is updated by hand.
' ============================================================================

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Public Const LOADER_VERSION As String = "2.1.0"
Public Const LOADER_FILE As String = "GCV Loader.xlam"
Public Const SETTINGS_APP As String = "GCV Tools"
Public Const PENDING_SECTION As String = "PendingVersions"     ' file name -> version waiting in pending\
Public Const BAD_SECTION As String = "BadVersions"             ' file name -> "1.2.0;1.3.0"

' Opening order = order of the groups on the ribbon tab. The logger first, so
' the others can log while they start; the updater last. Other add-ins in between.
Private Const FIRST_FILES As String = "GCV Logger.xlam;GCV Student Lookup.xlam;GCV English Names.xlam"
Private Const LAST_FILES As String = "GCV Updater.xlam"

Private Const MB_ICONWARNING As Long = &H30
Private Const MB_RIGHT As Long = &H80000
Private Const MB_RTLREADING As Long = &H100000

Private mLoaded As String           ' "GCV Logger.xlam;GCV Updater.xlam" - opened in this session
Private mStarted As Boolean         ' LoadAddins has run in this session
Private mEvents As LoaderEvents     ' waits for the first workbook
Private mFso As Object
Private mBaseFolder As String

' ----------------------------------------------------------------------------
' Public API (GCV Updater and GCV Logger call these with Application.Run)
' ----------------------------------------------------------------------------

Public Function LoaderVersion() As String
    LoaderVersion = LOADER_VERSION
End Function

Public Function ToolsFolder() As String
    If Len(mBaseFolder) = 0 Then
        mBaseFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\GCV\ExcelTools"
    End If
    ToolsFolder = mBaseFolder
End Function

Public Function AddinsFolder() As String
    AddinsFolder = ToolsFolder() & "\addins"
End Function

Public Function PreviousFolder() As String
    PreviousFolder = ToolsFolder() & "\previous"
End Function

Public Function PendingFolder() As String
    PendingFolder = ToolsFolder() & "\pending"
End Function

' Add-in files opened in this session, ";"-separated.
Public Function LoadedAddins() As String
    LoadedAddins = mLoaded
End Function

' Called by ThisWorkbook.Workbook_Open. A workbook already open (or in
' Protected View): load now. Otherwise wait for the first one, so Excel's
' start screen stays.
Public Sub StartWhenReady()
    If HasOpenWorkbook() Then
        ScheduleLoad
    Else
        Set mEvents = New LoaderEvents
        mEvents.Attach
    End If
End Sub

' LoaderEvents: the first workbook was opened or created.
Public Sub OnFirstWorkbook()
    If Not mEvents Is Nothing Then mEvents.Detach
    Set mEvents = Nothing
    ScheduleLoad
End Sub

' Runs LoadAddins as soon as Excel has finished opening the workbook.
Private Sub ScheduleLoad()
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!GcvLoader.LoadAddins"
End Sub

' Add-ins are not counted in Workbooks.Count.
Private Function HasOpenWorkbook() As Boolean
    On Error Resume Next
    HasOpenWorkbook = (Workbooks.Count > 0) Or (Application.ProtectedViewWindows.Count > 0)
    On Error GoTo 0
End Function

Public Sub LoadAddins()
    Dim name As Variant
    Dim failedNames As String
    Dim active As Window

    On Error GoTo Crashed
    If mStarted Then Exit Sub                             ' already done in this session
    mStarted = True
    Set active = ActiveWindow                             ' keep the secretary's workbook in front
    LoaderLog "INFO", "Loader " & LOADER_VERSION & " started; Excel " & Application.Version & "."
    ApplyPending
    For Each name In OpeningOrder(AddinFileNames(AddinsFolder()))
        If IsOpen(CStr(name)) Then
            ' e.g. a DEV copy opened by hand: two files with one name cannot be open
            LoaderLog "WARN", name & " is already open; the installed copy was not loaded."
        ElseIf StartAddin(CStr(name)) Then
            mLoaded = AddToList(mLoaded, CStr(name))
        Else
            failedNames = failedNames & vbCrLf & "- " & name
        End If
    Next name
    If Not active Is Nothing Then
        On Error Resume Next                              ' the window may have been closed meanwhile
        active.Activate
        On Error GoTo Crashed
    End If
    If Len(failedNames) > 0 Then
        ShowWarning FromCodePoints("062A 0639 0630 0631 0020 062A 0634 063A 064A 0644 0020 0628 0639 0636 0020 0623 062F 0648 0627 062A 0020 0642 0631 064A 0629 0020 0627 0644 0623 0637 0641 0627 0644 003A") & failedNames & vbCrLf & vbCrLf & _
                    FromCodePoints("064A 0645 0643 0646 0643 0020 0645 062A 0627 0628 0639 0629 0020 0627 0644 0639 0645 0644 002E 0020 0623 0628 0644 063A 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 0020 0028 0632 0631 0020 0627 0644 0625 0628 0644 0627 063A 0020 0639 0646 0020 0645 0634 0643 0644 0629 0029 002E")
    End If
    Exit Sub

Crashed:
    LoaderLog "ERROR", "Loader stopped: " & Err.Number & " " & Err.Description
End Sub

' Support macro: forget every "bad version" mark (next update offers them again).
Public Sub GcvLoaderClearBadVersions()
    On Error Resume Next
    DeleteSetting SETTINGS_APP, BAD_SECTION
    On Error GoTo 0
    LoaderLog "INFO", "Bad-version marks cleared."
End Sub

' ----------------------------------------------------------------------------
' Pure functions (tested)
' ----------------------------------------------------------------------------

' "GCV <something>.xlam", not the loader, not a download in progress.
Public Function IsAddinFileName(ByVal fileName As String) As Boolean
    If LCase$(Left$(fileName, 4)) <> "gcv " Then Exit Function
    If LCase$(Right$(fileName, 5)) <> ".xlam" Then Exit Function
    If LCase$(fileName) = LCase$(LOADER_FILE) Then Exit Function
    IsAddinFileName = True
End Function

' FIRST_FILES in their order, then the others alphabetically, then LAST_FILES.
Public Function OpeningOrder(ByVal fileNames As Collection) As Collection
    Dim result As Collection
    Dim middle As Collection
    Dim part As Variant
    Dim name As Variant

    Set result = New Collection
    For Each part In Split(FIRST_FILES, ";")
        If ContainsText(fileNames, CStr(part)) Then result.Add CStr(part)
    Next part
    Set middle = New Collection
    For Each name In fileNames
        If Not InList(FIRST_FILES, CStr(name)) And Not InList(LAST_FILES, CStr(name)) Then InsertSorted middle, CStr(name)
    Next name
    For Each name In middle
        result.Add name
    Next name
    For Each part In Split(LAST_FILES, ";")
        If ContainsText(fileNames, CStr(part)) Then result.Add CStr(part)
    Next part
    Set OpeningOrder = result
End Function

' "a;b" lists (registry values and session lists), compared without case.
Public Function InList(ByVal list As String, ByVal item As String) As Boolean
    InList = InStr(1, ";" & list & ";", ";" & item & ";", vbTextCompare) > 0
End Function

Public Function AddToList(ByVal list As String, ByVal item As String) As String
    If InList(list, item) Then
        AddToList = list
    ElseIf Len(list) = 0 Then
        AddToList = item
    Else
        AddToList = list & ";" & item
    End If
End Function

' ----------------------------------------------------------------------------
' Pending updates
' ----------------------------------------------------------------------------

' Moves every downloaded add-in from pending\ to addins\, keeping the current
' one in previous\. Files that cannot be moved (another Excel window holds
' them) stay pending for the next start.
Private Sub ApplyPending()
    Dim name As Variant
    Dim target As String
    Dim backup As String

    For Each name In AddinFileNames(PendingFolder())
        target = AddinsFolder() & "\" & name
        backup = PreviousFolder() & "\" & name
        On Error GoTo NotNow
        MakeFolder AddinsFolder()
        MakeFolder PreviousFolder()
        If Fso().FileExists(target) Then
            If Fso().FileExists(backup) Then Fso().DeleteFile backup, True
            Fso().MoveFile target, backup
        End If
        Fso().MoveFile PendingFolder() & "\" & name, target
        LoaderLog "INFO", "Update applied: " & name & " " & PendingVersion(CStr(name)) & "."
        GoTo NextFile
NotNow:
        LoaderLog "WARN", "Update of " & name & " not applied now (" & Err.Description & "); trying at next start."
        Resume NextFile
NextFile:
        On Error GoTo 0
    Next name
End Sub

Private Function PendingVersion(ByVal fileName As String) As String
    PendingVersion = GetSetting(SETTINGS_APP, PENDING_SECTION, fileName, vbNullString)
End Function

' ----------------------------------------------------------------------------
' Starting an add-in, with fallback to the previous version
' ----------------------------------------------------------------------------

Private Function StartAddin(ByVal fileName As String) As Boolean
    Dim failure As String
    Dim version As String
    Dim pendingVer As String

    pendingVer = PendingVersion(fileName)
    If TryOpen(AddinsFolder() & "\" & fileName, version, failure) Then
        LoaderLog "INFO", fileName & " " & version & " loaded."
        ClearPending fileName
        StartAddin = True
        Exit Function
    End If

    LoaderLog "ERROR", fileName & " did not start (" & failure & ")."
    If Len(pendingVer) > 0 Then MarkBad fileName, pendingVer
    ClearPending fileName
    If Not Fso().FileExists(PreviousFolder() & "\" & fileName) Then Exit Function

    ' Put the previous version back and start it.
    On Error GoTo NoFallback
    Fso().CopyFile PreviousFolder() & "\" & fileName, AddinsFolder() & "\" & fileName, True
    On Error GoTo 0
    If TryOpen(AddinsFolder() & "\" & fileName, version, failure) Then
        LoaderLog "WARN", fileName & ": previous version " & version & " restored and loaded."
        StartAddin = True
    Else
        LoaderLog "ERROR", fileName & ": previous version did not start either (" & failure & ")."
    End If
    Exit Function

NoFallback:
    LoaderLog "ERROR", fileName & ": previous version could not be restored (" & Err.Description & ")."
End Function

Private Function TryOpen(ByVal path As String, ByRef version As String, ByRef failure As String) As Boolean
    Dim wb As Workbook
    Dim answer As Variant

    On Error GoTo Failed
    Application.EnableEvents = True                    ' the add-in needs its Workbook_Open
    Application.DisplayAlerts = False                  ' a damaged file must not wait for an answer
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Application.DisplayAlerts = True
    answer = Application.Run("'" & Replace(wb.Name, "'", "''") & "'!GcvAddin.GcvAddinPing")
    version = CStr(answer)
    If Len(version) = 0 Then
        failure = "no version from GcvAddinPing"
        GoTo CloseIt
    End If
    TryOpen = True
    Exit Function

Failed:
    failure = Err.Number & " " & Err.Description
    Application.DisplayAlerts = True
CloseIt:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Sub MarkBad(ByVal fileName As String, ByVal version As String)
    SaveSetting SETTINGS_APP, BAD_SECTION, fileName, _
                AddToList(GetSetting(SETTINGS_APP, BAD_SECTION, fileName, vbNullString), version)
    LoaderLog "WARN", fileName & " " & version & " marked as bad; it will not be installed again."
End Sub

Private Sub ClearPending(ByVal fileName As String)
    On Error Resume Next                    ' raises when nothing is pending
    DeleteSetting SETTINGS_APP, PENDING_SECTION, fileName
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Files, log, messages
' ----------------------------------------------------------------------------

Private Function AddinFileNames(ByVal folder As String) As Collection
    Dim result As Collection
    Dim file As Object
    Set result = New Collection
    If Fso().FolderExists(folder) Then
        For Each file In Fso().GetFolder(folder).Files
            If IsAddinFileName(file.Name) Then result.Add file.Name
        Next file
    End If
    Set AddinFileNames = result
End Function

' Add-ins are not part of a For Each over Workbooks, but Workbooks("name") finds them.
Public Function IsOpen(ByVal workbookName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(workbookName)
    On Error GoTo 0
    IsOpen = Not wb Is Nothing
End Function

Private Function ContainsText(ByVal items As Collection, ByVal value As String) As Boolean
    Dim item As Variant
    For Each item In items
        If StrComp(item, value, vbTextCompare) = 0 Then
            ContainsText = True
            Exit Function
        End If
    Next item
End Function

Private Sub InsertSorted(ByVal items As Collection, ByVal value As String)
    Dim i As Long
    For i = 1 To items.Count
        If StrComp(value, items(i), vbTextCompare) < 0 Then
            items.Add value, Before:=i
            Exit Sub
        End If
    Next i
    items.Add value
End Sub

' Same files and line format as GCV Logger (which is not open yet).
Public Sub LoaderLog(ByVal levelName As String, ByVal message As String)
    Dim folder As String
    Dim stream As Object
    Dim i As Long
    Dim code As Long
    Dim clean As String

    Debug.Print "[loader] - " & UCase$(levelName) & ": " & message
    On Error GoTo Done                      ' logging must never stop the loader
    For i = 1 To Len(message)
        code = AscW(Mid$(message, i, 1)) And &HFFFF&
        If code < 32 Then
            clean = clean & " "
        ElseIf code > 126 Then
            clean = clean & "\u" & Right$("000" & Hex$(code), 4)
        Else
            clean = clean & Mid$(message, i, 1)
        End If
    Next i
    folder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\GCV\ExcelLogs\GCV Tools"
    MakeFolder folder
    Set stream = Fso().OpenTextFile(folder & "\" & Format$(Now, "yyyy-mm-dd") & ".log", 8, True, 0)
    stream.WriteLine "[loader] - " & UCase$(levelName) & ": " & Format$(Now, "yyyy-mm-dd") & " - " & _
                     Format$(Now, "hh:nn:ss") & " " & clean
    stream.Close
Done:
End Sub

Private Sub MakeFolder(ByVal path As String)
    Dim parts() As String
    Dim current As String
    Dim i As Long
    parts = Split(path, "\")
    current = parts(0)
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            current = current & "\" & parts(i)
            If Not Fso().FolderExists(current) Then Fso().CreateFolder current
        End If
    Next i
End Sub

Private Function Fso() As Object
    If mFso Is Nothing Then Set mFso = CreateObject("Scripting.FileSystemObject")
    Set Fso = mFso
End Function

Private Sub ShowWarning(ByVal text As String)
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr("GCV Tools"), MB_ICONWARNING Or MB_RIGHT Or MB_RTLREADING
End Sub

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function

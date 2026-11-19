Attribute VB_Name = "Updater"
Option Explicit

' ============================================================================
' GCV Updater add-in. Ribbon button "update tools":
'
'   1. read the catalog (catalog.json in the excel-updater repository): every
'      GCV add-in with its repository and whether it is required
'   2. compare the installed version of each add-in with its newest release
'   3. one question lists the updates and new required add-ins; each new
'      optional add-in is offered separately
'   4. download + check (size, SHA-256) into pending\, then install at once:
'      close the add-in, keep the current file in previous\, put the new one
'      in addins\, open it and check it answers with its version. If it does
'      not, the previous version is put back and the new one marked "bad".
'      The updater's own update, and files another Excel window holds, are
'      installed by GCV Loader at the next Excel start.
'
' Folders: %LOCALAPPDATA%\GCV\ExcelTools\{addins, previous, pending} (see GcvLoader).
' ============================================================================

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Public Const CATALOG_URL As String = _
    "https://raw.githubusercontent.com/Gaza-children-village/excel-updater/main/catalog.json"

Private Const MB_OK As Long = &H0
Private Const MB_YESNO As Long = &H4
Private Const MB_ICONERROR As Long = &H10
Private Const MB_ICONWARNING As Long = &H30
Private Const MB_ICONINFO As Long = &H40
Private Const MB_RIGHT As Long = &H80000
Private Const MB_RTLREADING As Long = &H100000
Private Const ID_YES As Long = 6

Private Const LOG_SOURCE As String = "updater"
Private Const LOADER_FILE As String = "GCV Loader.xlam"
Private Const LOGGER_ID As String = "gcv/logger"
Private Const SETTINGS_APP As String = "GCV Tools"
Private Const PENDING_SECTION As String = "PendingVersions"
Private Const BAD_SECTION As String = "BadVersions"
Private Const TIMEOUT_MS As Long = 6000
Private Const CHANGELOG_MAX As Long = 400

Public Enum InstallResult
    irInstalled = 1
    irAtNextStart = 2
    irFailedKeptOld = 3
    irNeedsLoader = 4
End Enum

Private mHttp As IUpdHttp            ' tests replace the network ...
Private mAnswers As Collection       ' ... the questions (queued Yes/No answers) ...
Private mMessages As Collection      ' ... record every message ...
Private mBaseForTests As String      ' ... use a temporary folder ...
Private mInstalledForTests As Object ' ... fix the installed versions (id -> version) ...
Private mFailPingForTests As Boolean ' ... and simulate an add-in that does not start
Private mReleases As Object          ' newest release per id, for one run

' ----------------------------------------------------------------------------
' Add-in life cycle
' ----------------------------------------------------------------------------

Public Sub AddinOpened()
    On Error Resume Next
    LogInfo LOG_SOURCE, GcvAddin.ADDIN_NAME & " " & GcvAddin.ADDIN_VERSION & " opened."
End Sub

Public Sub AddinClosing()
End Sub

Public Sub AddinSelfCheck()
    Dim probe As Object
    UpdCore.IsValidVersion GcvAddin.ADDIN_VERSION
    Set probe = JsonParse("{}")
    TagPrefixFor "gcv/x"
End Sub

' ----------------------------------------------------------------------------
' Ribbon
' ----------------------------------------------------------------------------

Public Sub Ribbon_UpdateTools(control As IRibbonControl)
    UpdateTools
End Sub

Public Sub Ribbon_About(control As IRibbonControl)
    Notify T_About(), MB_ICONINFO
End Sub

' ----------------------------------------------------------------------------
' The update
' ----------------------------------------------------------------------------

Public Sub UpdateTools()
    Dim catalog As Collection
    Dim installed As Object
    Dim latest As Object
    Dim plan As Collection
    Dim item As Variant
    Dim results As String
    Dim loaderVersion As String
    Dim offline As Boolean
    Dim anyInstalled As Boolean
    Dim outcome As InstallResult

    On Error GoTo Failed
    UpdLog "INFO", "Update started."
    Set mReleases = CreateObject("Scripting.Dictionary")
    Set catalog = LoadCatalog(offline)
    If offline Then
        UpdLog "INFO", "No internet connection; nothing checked."
        Notify T_Offline(), MB_ICONINFO
        Exit Sub
    End If

    Set installed = InstalledVersions(catalog)
    Set latest = LatestVersions(catalog, offline)
    If offline Then
        Notify T_Offline(), MB_ICONINFO
        Exit Sub
    End If

    Set plan = AskOptionalTools(BuildPlan(catalog, installed, latest, BadVersions(catalog)))
    If plan.Count = 0 Then
        UpdLog "INFO", "Everything is up to date."
        Notify T_UpToDate(catalog, installed), MB_ICONINFO
        Exit Sub
    End If
    If Not Ask(T_Plan(plan)) Then
        UpdLog "INFO", "Update declined: " & PlanSummary(plan) & "."
        Exit Sub
    End If

    loaderVersion = CurrentLoaderVersion()
    For Each item In plan
        outcome = InstallOne(item, loaderVersion)
        If outcome = irInstalled Then anyInstalled = True
        results = results & vbCrLf & ChrW$(&H2022) & " " & item("name") & " " & item("new") & ": " & T_Result(outcome)
    Next item
    If anyInstalled And mAnswers Is Nothing Then Application.CalculateFull   ' formulas of updated add-ins
    UpdLog "INFO", "Update finished: " & PlanSummary(plan) & "."
    Notify T_Results(results), MB_ICONINFO
    Exit Sub

Failed:
    UpdLog "ERROR", "Update failed: " & Err.Number & " " & Err.Description
    Notify T_Failed(Err.Description), MB_ICONERROR
End Sub

' ----------------------------------------------------------------------------
' Catalog
' ----------------------------------------------------------------------------
' {"tools": [ {"id": "gcv/logger", "name": "GCV Logger", "file": "GCV Logger.xlam",
'              "repo": "Gaza-children-village/excel-logger", "required": true,
'              "description": "..."}, ... ]}

Private Function LoadCatalog(ByRef offline As Boolean) As Collection
    Dim result As UpdHttpResult
    Dim body() As Byte

    Set result = Http.GetUrl(CATALOG_URL, TIMEOUT_MS)
    If result.IsOffline Then
        offline = True
        Exit Function
    End If
    If Not result.IsOk Then RaiseUpd ueDownload, "Cannot read the add-in catalog: " & result.Describe
    body = result.Body
    Set LoadCatalog = ParseCatalog(Utf8Decode(body))
    UpdLog "DEBUG", "Catalog: " & LoadCatalog.Count & " add-in(s)."
End Function

' Raises ueManifest when an entry is not usable.
Public Function ParseCatalog(ByVal json As String) As Collection
    Dim root As Variant
    Dim entries As Object
    Dim entry As Variant
    Dim tool As Object
    Dim result As Collection
    Dim seen As Object

    JsonParseInto json, root
    If Not IsObject(root) Then RaiseUpd ueManifest, "The catalog is not a JSON object."
    If TypeName(root) <> "Dictionary" Then RaiseUpd ueManifest, "The catalog is not a JSON object."
    Set entries = JsonObj(root, "tools")
    If entries Is Nothing Then RaiseUpd ueManifest, "The catalog has no tools."
    If TypeName(entries) <> "Collection" Then RaiseUpd ueManifest, "The catalog has no tools."

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    For Each entry In entries
        If Not IsObject(entry) Then RaiseUpd ueManifest, "A catalog entry is not an object."
        Set tool = CreateObject("Scripting.Dictionary")
        tool("id") = JsonStr(entry, "id")
        tool("name") = JsonStr(entry, "name")
        tool("file") = JsonStr(entry, "file")
        tool("repo") = JsonStr(entry, "repo")
        tool("required") = JsonBool(entry, "required")
        tool("description") = JsonStr(entry, "description")
        If Not IsValidAddinId(tool("id")) Then RaiseUpd ueManifest, "Catalog: invalid id '" & tool("id") & "'."
        If Not IsAddinFile(tool("file")) Then RaiseUpd ueManifest, "Catalog: invalid file '" & tool("file") & "'."
        If Not IsRepo(tool("repo")) Then RaiseUpd ueManifest, "Catalog: invalid repo '" & tool("repo") & "'."
        If Len(tool("name")) = 0 Then tool("name") = tool("file")
        If seen.Exists(tool("id")) Then RaiseUpd ueManifest, "Catalog: " & tool("id") & " is listed twice."
        seen(tool("id")) = True
        tool("tagPrefix") = TagPrefixFor(tool("id"))
        result.Add tool
    Next entry
    Set ParseCatalog = result
End Function

Public Function TagPrefixFor(ByVal addinId As String) As String
    TagPrefixFor = Mid$(addinId, InStr(addinId, "/") + 1) & "-v"
End Function

Public Function IsValidAddinId(ByVal text As String) As Boolean
    Dim parts() As String
    Dim i As Long
    parts = Split(text, "/")
    If UBound(parts) <> 1 Then Exit Function
    For i = 0 To 1
        If Len(parts(i)) = 0 Then Exit Function
        If parts(i) Like "*[!a-z0-9-]*" Then Exit Function
    Next i
    IsValidAddinId = True
End Function

' "GCV <letters, digits, spaces, dashes>.xlam" - no folders, no dots inside.
Public Function IsAddinFile(ByVal fileName As String) As Boolean
    Dim middle As String
    If Left$(fileName, 4) <> "GCV " Or LCase$(Right$(fileName, 5)) <> ".xlam" Then Exit Function
    If StrComp(fileName, LOADER_FILE, vbTextCompare) = 0 Then Exit Function
    middle = Mid$(fileName, 5, Len(fileName) - 9)
    If Len(middle) = 0 Or Len(fileName) > 64 Then Exit Function
    IsAddinFile = Not (middle Like "*[!A-Za-z0-9 -]*")
End Function

Private Function IsRepo(ByVal text As String) As Boolean
    Dim parts() As String
    parts = Split(text, "/")
    If UBound(parts) <> 1 Then Exit Function
    If Len(parts(0)) = 0 Or Len(parts(1)) = 0 Then Exit Function
    IsRepo = Not (text Like "*[!A-Za-z0-9_./-]*")
End Function

' ----------------------------------------------------------------------------
' Installed and newest versions
' ----------------------------------------------------------------------------

' id -> version for every catalog add-in found on this PC. An installed file
' that is not running counts as "0.0.0", so it is replaced.
Private Function InstalledVersions(ByVal catalog As Collection) As Object
    Dim result As Object
    Dim tool As Variant
    Dim info As String

    If Not mInstalledForTests Is Nothing Then
        Set InstalledVersions = mInstalledForTests
        Exit Function
    End If
    Set result = CreateObject("Scripting.Dictionary")
    For Each tool In catalog
        If tool("id") = GcvAddin.ADDIN_ID Then
            result(tool("id")) = GcvAddin.ADDIN_VERSION
        ElseIf IsWorkbookOpen(tool("file")) Then
            info = RunText("'" & tool("file") & "'!GcvAddin.GcvAddinInfo")
            If Len(info) > 0 Then result(tool("id")) = Split(info, "|")(1) Else result(tool("id")) = "0.0.0"
        ElseIf CreateObject("Scripting.FileSystemObject").FileExists(AddinsFolder() & "\" & tool("file")) Then
            result(tool("id")) = "0.0.0"
        End If
    Next tool
    Set InstalledVersions = result
End Function

' id -> newest stable version; the release objects are kept in mReleases.
Private Function LatestVersions(ByVal catalog As Collection, ByRef offline As Boolean) As Object
    Dim result As Object
    Dim tool As Variant
    Dim response As UpdHttpResult
    Dim body() As Byte
    Dim releases As Object
    Dim release As Object

    Set result = CreateObject("Scripting.Dictionary")
    For Each tool In catalog
        Set response = Http.GetUrl(ReleasesUrl(tool("repo")), TIMEOUT_MS)
        If response.IsOffline Then
            offline = True
            Exit Function
        End If
        If response.IsOk Then
            body = response.Body
            Set releases = JsonParse(Utf8Decode(body))
            If TypeName(releases) = "Collection" Then
                Set release = PickLatestRelease(releases, tool("tagPrefix"))
                If Not release Is Nothing Then
                    result(tool("id")) = ReleaseVersion(release, tool("tagPrefix"))
                    Set mReleases(tool("id")) = release
                End If
            End If
        Else
            UpdLog "WARN", tool("repo") & ": " & response.Describe & "."
        End If
    Next tool
    Set LatestVersions = result
End Function

' file -> "1.2.0;1.3.0" for every catalog add-in (marked by loader or updater).
Private Function BadVersions(ByVal catalog As Collection) As Object
    Dim result As Object
    Dim tool As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each tool In catalog
        result(tool("file")) = GetSetting(SETTINGS_APP, BAD_SECTION, tool("file"), vbNullString)
    Next tool
    Set BadVersions = result
End Function

' ----------------------------------------------------------------------------
' Plan (pure, tested)
' ----------------------------------------------------------------------------

' Items: catalog entry + "current" ("" = not installed), "new", "kind"
' ("update" / "new"). Logger first (the others log through it), the updater
' itself last, the rest in catalog order.
Public Function BuildPlan(ByVal catalog As Collection, ByVal installed As Object, _
                          ByVal latest As Object, ByVal bad As Object) As Collection
    Dim plan As Collection
    Dim tool As Variant
    Dim item As Object
    Dim key As Variant

    Set plan = New Collection
    For Each tool In catalog
        If latest.Exists(tool("id")) Then
            Set item = Nothing
            If installed.Exists(tool("id")) Then
                If CompareVersions(latest(tool("id")), installed(tool("id"))) > 0 Then Set item = NewItem(tool, "update")
            Else
                Set item = NewItem(tool, "new")
            End If
            If Not item Is Nothing Then
                ' (no IIf: it would read installed(id) and so add the missing key)
                If installed.Exists(tool("id")) Then item("current") = installed(tool("id")) Else item("current") = vbNullString
                item("new") = latest(tool("id"))
                If IsBad(bad, tool("file"), item("new")) Then
                    UpdLog "WARN", tool("name") & " " & item("new") & " did not start before; skipped."
                ElseIf tool("id") = LOGGER_ID Then
                    If plan.Count = 0 Then plan.Add item Else plan.Add item, Before:=1
                Else
                    plan.Add item
                End If
            End If
        End If
    Next tool
    ' the updater itself last
    For key = plan.Count To 1 Step -1
        If plan(key)("id") = GcvAddin.ADDIN_ID And key < plan.Count Then
            Set item = plan(key)
            plan.Remove key
            plan.Add item
            Exit For
        End If
    Next key
    Set BuildPlan = plan
End Function

Private Function NewItem(ByVal tool As Object, ByVal kind As String) As Object
    Dim item As Object
    Dim key As Variant
    Set item = CreateObject("Scripting.Dictionary")
    For Each key In tool.Keys
        item(key) = tool(key)
    Next key
    item("kind") = kind
    Set NewItem = item
End Function

Private Function IsBad(ByVal bad As Object, ByVal fileName As String, ByVal version As String) As Boolean
    If Not bad.Exists(fileName) Then Exit Function
    IsBad = InStr(";" & bad(fileName) & ";", ";" & version & ";") > 0
End Function

' New optional add-ins are offered one by one; declined ones leave the plan.
Private Function AskOptionalTools(ByVal plan As Collection) As Collection
    Dim result As Collection
    Dim item As Variant
    Set result = New Collection
    For Each item In plan
        If item("kind") = "new" And Not item("required") Then
            If Ask(T_NewTool(item("name"), item("description"), item("new"))) Then result.Add item
        Else
            result.Add item
        End If
    Next item
    Set AskOptionalTools = result
End Function

Private Function PlanSummary(ByVal plan As Collection) As String
    Dim item As Variant
    For Each item In plan
        If Len(PlanSummary) > 0 Then PlanSummary = PlanSummary & ", "
        PlanSummary = PlanSummary & item("id") & " " & IIf(Len(item("current")) = 0, "(new)", item("current")) & _
                      " -> " & item("new")
    Next item
End Function

' ----------------------------------------------------------------------------
' Installing one add-in
' ----------------------------------------------------------------------------

Private Function InstallOne(ByVal item As Object, ByVal loaderVersion As String) As InstallResult
    Dim release As Object
    Dim manifest As Object

    Set release = mReleases(item("id"))
    Set manifest = FetchToolsManifest(Http, release, item("id"), item("new"))
    If Not LoaderIsEnough(manifest, loaderVersion) Then
        UpdLog "WARN", item("id") & " " & item("new") & " needs GCV Loader " & JsonStr(manifest, "minLoaderVersion") & _
                       " (installed: " & loaderVersion & ")."
        InstallOne = irNeedsLoader
        Exit Function
    End If
    DownloadToolsFile Http, release, manifest, item("new"), PendingFolder(), item("file")
    SaveSetting SETTINGS_APP, PENDING_SECTION, item("file"), item("new")
    UpdLog "INFO", item("id") & " " & item("new") & " downloaded and verified."

    If item("id") = GcvAddin.ADDIN_ID Then
        UpdLog "INFO", "The updater's own update is installed at the next Excel start."
        InstallOne = irAtNextStart
    Else
        InstallOne = SwapAddin(item("file"), item("new"))
    End If
End Function

' Puts pending\<file> in place while Excel runs. Public for tests.
Public Function SwapAddin(ByVal fileName As String, ByVal version As String) As InstallResult
    Dim fso As Object
    Dim current As String
    Dim backup As String
    Dim waiting As String
    Dim wasOpen As Boolean
    Dim hadPrevious As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    current = AddinsFolder() & "\" & fileName
    backup = PreviousFolder() & "\" & fileName
    waiting = PendingFolder() & "\" & fileName

    wasOpen = IsWorkbookOpen(fileName)
    If wasOpen Then CloseAddin fileName

    On Error GoTo Locked
    EnsureFolder AddinsFolder()
    EnsureFolder PreviousFolder()
    hadPrevious = fso.FileExists(current)
    If hadPrevious Then
        If fso.FileExists(backup) Then fso.DeleteFile backup, True
        fso.MoveFile current, backup
    End If
    fso.MoveFile waiting, current
    On Error GoTo 0

    If StartAddin(current, version) Then
        ClearPending fileName
        UpdLog "INFO", fileName & " " & version & " installed and started."
        SwapAddin = irInstalled
        Exit Function
    End If

    ' The new version does not start: put the previous one back.
    UpdLog "ERROR", fileName & " " & version & " did not start; previous version restored."
    MarkBad fileName, version
    ClearPending fileName
    If hadPrevious Then
        fso.CopyFile backup, current, True
        If wasOpen Then StartAddin current, vbNullString
    Else
        fso.DeleteFile current, True
    End If
    SwapAddin = irFailedKeptOld
    Exit Function

Locked:
    ' Another Excel window holds the file: GCV Loader installs it at the next start.
    UpdLog "WARN", fileName & " is in use (" & Err.Description & "); installed at the next Excel start."
    On Error GoTo -1
    On Error Resume Next
    If wasOpen And Not IsWorkbookOpen(fileName) And fso.FileExists(current) Then StartAddin current, vbNullString
    On Error GoTo 0
    SwapAddin = irAtNextStart
End Function

' Opens an add-in and checks it answers with `expectedVersion` ("" = any).
Private Function StartAddin(ByVal path As String, ByVal expectedVersion As String) As Boolean
    Dim wb As Workbook
    Dim answer As String

    If Not mAnswers Is Nothing Then                     ' tests: no Excel files are opened
        StartAddin = Not mFailPingForTests Or Len(expectedVersion) = 0
        Exit Function
    End If
    On Error GoTo Failed
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Application.DisplayAlerts = True
    answer = CStr(Application.Run("'" & Replace(wb.Name, "'", "''") & "'!GcvAddin.GcvAddinPing"))
    StartAddin = (Len(answer) > 0 And (Len(expectedVersion) = 0 Or answer = expectedVersion))
    If Not StartAddin Then wb.Close SaveChanges:=False
    Exit Function
Failed:
    Application.DisplayAlerts = True
    UpdLog "WARN", "Could not start " & path & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

Private Sub CloseAddin(ByVal fileName As String)
    If Not mAnswers Is Nothing Then Exit Sub
    On Error Resume Next
    Workbooks(fileName).Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub MarkBad(ByVal fileName As String, ByVal version As String)
    Dim list As String
    list = GetSetting(SETTINGS_APP, BAD_SECTION, fileName, vbNullString)
    If InStr(";" & list & ";", ";" & version & ";") = 0 Then
        SaveSetting SETTINGS_APP, BAD_SECTION, fileName, IIf(Len(list) = 0, version, list & ";" & version)
    End If
End Sub

Private Sub ClearPending(ByVal fileName As String)
    On Error Resume Next
    DeleteSetting SETTINGS_APP, PENDING_SECTION, fileName
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Folders and other add-ins
' ----------------------------------------------------------------------------

Public Function ToolsFolder() As String
    If Len(mBaseForTests) > 0 Then
        ToolsFolder = mBaseForTests
    Else
        ToolsFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\GCV\ExcelTools"
    End If
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

' "" when GCV Loader is not open (DEV copy opened by hand).
Public Function CurrentLoaderVersion() As String
    If Not mAnswers Is Nothing Then
        CurrentLoaderVersion = "2.0.0"
    ElseIf IsWorkbookOpen(LOADER_FILE) Then
        CurrentLoaderVersion = RunText("'" & LOADER_FILE & "'!GcvLoader.LoaderVersion")
    End If
End Function

Private Function IsWorkbookOpen(ByVal workbookName As String) As Boolean
    Dim wb As Workbook
    If Not mAnswers Is Nothing Then Exit Function
    On Error Resume Next
    Set wb = Workbooks(workbookName)
    On Error GoTo 0
    IsWorkbookOpen = Not wb Is Nothing
End Function

Private Function RunText(ByVal macro As String) As String
    On Error Resume Next
    RunText = CStr(Application.Run(macro))
    If Err.Number <> 0 Then RunText = vbNullString
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Network and messages (replaceable by tests)
' ----------------------------------------------------------------------------

Private Property Get Http() As IUpdHttp
    If mHttp Is Nothing Then Set mHttp = New UpdWinHttp
    Set Http = mHttp
End Property

Public Sub UseTestDoubles(ByVal fakeHttp As IUpdHttp, ByVal answers As Collection, ByVal baseFolder As String, _
                          ByVal installed As Object, ByVal failPing As Boolean)
    Set mHttp = fakeHttp
    Set mAnswers = answers
    Set mMessages = New Collection
    mBaseForTests = baseFolder
    Set mInstalledForTests = installed
    mFailPingForTests = failPing
    Set mReleases = CreateObject("Scripting.Dictionary")
End Sub

Public Sub ResetTestDoubles()
    Set mHttp = Nothing
    Set mAnswers = Nothing
    Set mMessages = Nothing
    mBaseForTests = vbNullString
    Set mInstalledForTests = Nothing
    mFailPingForTests = False
End Sub

Public Function TestMessages() As Collection
    Set TestMessages = mMessages
End Function

Private Function Ask(ByVal text As String) As Boolean
    If Not mAnswers Is Nothing Then
        mMessages.Add text
        If mAnswers.Count > 0 Then
            Ask = mAnswers(1)
            mAnswers.Remove 1
        End If
        Exit Function
    End If
    Ask = (ShowMessage(text, MB_YESNO Or MB_ICONINFO) = ID_YES)
End Function

Private Sub Notify(ByVal text As String, ByVal icon As Long)
    If Not mAnswers Is Nothing Then
        mMessages.Add text
        Exit Sub
    End If
    ShowMessage text, MB_OK Or icon
End Sub

Private Function ShowMessage(ByVal text As String, ByVal flags As Long) As Long
    ShowMessage = MessageBoxW(Application.hWnd, StrPtr(text), StrPtr(T_Title()), _
                              flags Or MB_RIGHT Or MB_RTLREADING)
End Function

' ----------------------------------------------------------------------------
' Texts (Arabic, built from code points)
' ----------------------------------------------------------------------------

Public Function T_Title() As String
    T_Title = FromCodePoints("0623 062F 0648 0627 062A 0020 0642 0631 064A 0629 0020 0627 0644 0623 0637 0641 0627 0644 0020 002D 0020 0627 0644 062A 062D 062F 064A 062B 0627 062A")
End Function

Public Function T_Plan(ByVal plan As Collection) As String
    Dim item As Variant
    T_Plan = FromCodePoints("0627 0644 062A 062D 062F 064A 062B 0627 062A 0020 0627 0644 0645 062A 0627 062D 0629 003A")
    For Each item In plan
        If item("kind") = "new" Then
            T_Plan = T_Plan & vbCrLf & ChrW$(&H2022) & " " & item("name") & " " & item("new") & FromCodePoints("0020 0028 0623 062F 0627 0629 0020 062C 062F 064A 062F 0629 0029")
        Else
            T_Plan = T_Plan & vbCrLf & ChrW$(&H2022) & " " & item("name") & ": " & item("current") & " " & ChrW$(&H2190) & " " & item("new")
        End If
    Next item
    T_Plan = T_Plan & vbCrLf & vbCrLf & FromCodePoints("0647 0644 0020 062A 0631 064A 062F 0020 062A 062B 0628 064A 062A 0647 0627 0020 0627 0644 0622 0646 061F 0020 0627 062D 0641 0638 0020 0639 0645 0644 0643 0020 0623 0648 0644 0627 064B 002E")
End Function

Public Function T_NewTool(ByVal name As String, ByVal description As String, ByVal version As String) As String
    T_NewTool = FromCodePoints("0623 062F 0627 0629 0020 062C 062F 064A 062F 0629 0020 0645 062A 0627 062D 0629 003A 0020") & name & " " & version & vbCrLf & vbCrLf & description & vbCrLf & vbCrLf & _
                FromCodePoints("0647 0644 0020 062A 0631 064A 062F 0020 062A 062B 0628 064A 062A 0647 0627 061F")
End Function

Public Function T_Result(ByVal outcome As InstallResult) As String
    Select Case outcome
        Case irInstalled: T_Result = FromCodePoints("062A 0645 0020 0627 0644 062A 062B 0628 064A 062A")
        Case irAtNextStart: T_Result = FromCodePoints("064A 064F 062B 0628 064E 0651 062A 0020 0639 0646 062F 0020 0641 062A 062D 0020 0045 0078 0063 0065 006C 0020 0641 064A 0020 0627 0644 0645 0631 0629 0020 0627 0644 0642 0627 062F 0645 0629")
        Case irFailedKeptOld: T_Result = FromCodePoints("0644 0645 0020 064A 0639 0645 0644 0020 0627 0644 0625 0635 062F 0627 0631 0020 0627 0644 062C 062F 064A 062F 060C 0020 0648 0628 0642 064A 0020 0627 0644 0625 0635 062F 0627 0631 0020 0627 0644 0633 0627 0628 0642")
        Case irNeedsLoader: T_Result = FromCodePoints("064A 062D 062A 0627 062C 0020 0625 0644 0649 0020 062A 062D 062F 064A 062B 0020 0628 0631 0646 0627 0645 062C 0020 0627 0644 062A 062D 0645 064A 0644 0020 064A 062F 0648 064A 0627 064B 0020 0028 0623 0628 0644 063A 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 0029")
    End Select
End Function

Public Function T_Results(ByVal lines As String) As String
    T_Results = FromCodePoints("0646 062A 064A 062C 0629 0020 0627 0644 062A 062D 062F 064A 062B 003A") & lines
End Function

Public Function T_UpToDate(ByVal catalog As Collection, ByVal installed As Object) As String
    Dim tool As Variant
    T_UpToDate = FromCodePoints("062C 0645 064A 0639 0020 0627 0644 0623 062F 0648 0627 062A 0020 0645 062D 062F 0651 062B 0629 002E")
    For Each tool In catalog
        If installed.Exists(tool("id")) Then
            T_UpToDate = T_UpToDate & vbCrLf & ChrW$(&H2022) & " " & tool("name") & " " & installed(tool("id"))
        End If
    Next tool
End Function

Public Function T_Offline() As String
    T_Offline = FromCodePoints("0644 0627 0020 064A 0648 062C 062F 0020 0627 062A 0635 0627 0644 0020 0628 0627 0644 0625 0646 062A 0631 0646 062A 002E 0020 062D 0627 0648 0644 0020 0645 0631 0629 0020 0623 062E 0631 0649 0020 0639 0646 062F 0020 062A 0648 0641 0631 0020 0627 0644 0627 062A 0635 0627 0644 002E")
End Function

Public Function T_Failed(ByVal details As String) As String
    T_Failed = FromCodePoints("062A 0639 0630 0631 0020 0627 0644 062A 062D 062F 064A 062B 003A") & vbCrLf & details & vbCrLf & vbCrLf & _
               FromCodePoints("0627 0644 0623 062F 0648 0627 062A 0020 0627 0644 0645 062B 0628 062A 0629 0020 0644 0645 0020 062A 062A 063A 064A 0631 060C 0020 0625 0644 0627 0020 0645 0627 0020 0630 064F 0643 0631 0020 0623 0646 0647 0020 062A 0645 0020 062A 062B 0628 064A 062A 0647 002E")
End Function

Public Function T_About() As String
    Dim loaderText As String
    Dim files As String
    Dim fileName As Variant
    Dim info As String

    loaderText = CurrentLoaderVersion()
    If Len(loaderText) = 0 Then loaderText = FromCodePoints("063A 064A 0631 0020 0645 0648 062C 0648 062F")
    T_About = FromCodePoints("0623 062F 0648 0627 062A 0020 0642 0631 064A 0629 0020 0627 0644 0623 0637 0641 0627 0644 0020 0028 0047 0061 007A 0061 0020 0043 0068 0069 006C 0064 0072 0065 006E 0020 0056 0069 006C 006C 0061 0067 0065 0020 0054 006F 006F 006C 0073 0029") & vbCrLf & vbCrLf & _
              "GCV Loader " & loaderText
    If IsWorkbookOpen(LOADER_FILE) Then files = RunText("'" & LOADER_FILE & "'!GcvLoader.LoadedAddins")
    If Len(files) = 0 Then files = GcvAddin.ADDIN_FILE
    For Each fileName In Split(files, ";")
        If StrComp(fileName, GcvAddin.ADDIN_FILE, vbTextCompare) = 0 Then
            info = GcvAddin.GcvAddinInfo()
        Else
            info = RunText("'" & fileName & "'!GcvAddin.GcvAddinInfo")
        End If
        If Len(info) > 0 Then T_About = T_About & vbCrLf & Split(info, "|")(3) & " " & Split(info, "|")(1)
    Next fileName
End Function

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function

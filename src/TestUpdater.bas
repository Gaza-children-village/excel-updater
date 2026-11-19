Attribute VB_Name = "TestUpdater"
Option Explicit

' Tests for GCV Updater. Run RunUpdaterTests from the VBA editor (F5) in the
' DEV copy of GCV Updater. Uses UpdFakeHttp (no network), a temporary folder
' and only test add-in names ("GCV Test Tool.xlam"), whose registry marks are
' removed at the end. No Excel file is opened or closed.
' Do NOT ship this module or UpdFakeHttp.

Private passedCount As Long
Private failedCount As Long
Private currentTest As String

Private Const TOOL_ID As String = "gcv/test-tool"
Private Const TOOL_FILE As String = "GCV Test Tool.xlam"
Private Const TOOL_REPO As String = "test-org/excel-test-tool"
Private Const EXTRA_ID As String = "gcv/test-extra"
Private Const EXTRA_FILE As String = "GCV Test Extra.xlam"
Private Const EXTRA_REPO As String = "test-org/excel-test-extra"
Private Const NEW_CONTENT As String = "new add-in content"
Private Const OLD_CONTENT As String = "old add-in content"

Public Sub RunUpdaterTests()
    passedCount = 0
    failedCount = 0
    Debug.Print String$(60, "-")
    Debug.Print "Updater tests"
    On Error GoTo Crash

    Test_CompareVersions_OrdersNumerically
    Test_Sha256_MatchesKnownVector
    Test_PickLatestRelease_IgnoresDraftsPrereleasesAndOtherTags
    Test_ValidateManifest_RejectsWrongIdVersionAndFile
    Test_IsAddinFile_AcceptsOnlyPlainGcvNames
    Test_ParseCatalog_ReadsTools
    Test_ParseCatalog_RejectsBadEntries
    Test_BuildPlan_UpdatesNewerAndAddsMissingTools
    Test_BuildPlan_LoggerFirstUpdaterLast
    Test_BuildPlan_SkipsBadVersions
    Test_Download_SavesVerifiedFileOnly
    Test_Swap_InstallsAndKeepsPrevious
    Test_Swap_RestoresPreviousWhenNewDoesNotStart
    Test_Update_OfflineSaysSoAndChangesNothing
    Test_Update_UpToDateListsVersions
    Test_Update_InstallsUpdateAndOfferedNewTool
    Test_Update_DeclinedChangesNothing

Finish:
    Updater.ResetTestDoubles
    CleanRegistry
    Debug.Print "Passed: " & passedCount & "   Failed: " & failedCount
    MsgBox "Passed: " & passedCount & vbCrLf & "Failed: " & failedCount & vbCrLf & _
           "Details in the Immediate window (Ctrl+G).", _
           IIf(failedCount = 0, vbInformation, vbCritical), "Updater tests"
    Exit Sub

Crash:
    failedCount = failedCount + 1
    Debug.Print "CRASH in " & currentTest & ": " & Err.Number & " " & Err.Description
    Resume Finish
End Sub

' ----------------------------------------------------------------------------
' Core
' ----------------------------------------------------------------------------

Private Sub Test_CompareVersions_OrdersNumerically()
    StartTest "Versions compare numerically"
    AssertEqual 1, CompareVersions("0.10.0", "0.9.0"), "0.10 > 0.9"
    AssertEqual 0, CompareVersions("v1.2.3", "1.2.3"), "leading v"
    AssertEqual -1, CompareVersions("1.2.3", "1.2.4"), "patch"
End Sub

Private Sub Test_Sha256_MatchesKnownVector()
    Dim data() As Byte
    StartTest "SHA-256 of 'abc'"
    data = AsciiBytes("abc")
    AssertEqual "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", Sha256Hex(data), "abc"
End Sub

Private Sub Test_PickLatestRelease_IgnoresDraftsPrereleasesAndOtherTags()
    StartTest "Newest stable release with the add-in's tag prefix"
    AssertEqual "1.3.1", ReleaseVersion(PickLatestRelease(JsonParse("[" & _
        ReleaseJson("test-tool-v9.0.0", True, False, "x") & "," & _
        ReleaseJson("test-tool-v8.0.0", False, True, "x") & "," & _
        ReleaseJson("other-v9.0.0", False, False, "x") & "," & _
        ReleaseJson("test-tool-v1.3.1", False, False, "x") & "," & _
        ReleaseJson("test-tool-v1.3.0", False, False, "x") & "]"), "test-tool-v"), "test-tool-v"), "1.3.1"
End Sub

Private Sub Test_ValidateManifest_RejectsWrongIdVersionAndFile()
    Dim sha As String
    StartTest "Manifests are checked"
    sha = ShaOf(NEW_CONTENT)
    AssertEqual 0, ManifestError(ManifestJson(TOOL_ID, "1.1.0", sha, Len(NEW_CONTENT))), "good"
    AssertEqual UPD_ERR_BASE + ueManifest, ManifestError(ManifestJson("gcv/other", "1.1.0", sha, 5)), "wrong id"
    AssertEqual UPD_ERR_BASE + ueManifest, ManifestError(ManifestJson(TOOL_ID, "1.2.0", sha, 5)), "wrong version"
    AssertEqual UPD_ERR_BASE + ueManifest, ManifestError(ManifestJson(TOOL_ID, "1.1.0", "abc", 5)), "bad sha"
    AssertEqual UPD_ERR_BASE + ueManifest, ManifestError(ManifestJson(TOOL_ID, "1.1.0", sha, 0)), "bad size"
End Sub

' ----------------------------------------------------------------------------
' Catalog and plan
' ----------------------------------------------------------------------------

Private Sub Test_IsAddinFile_AcceptsOnlyPlainGcvNames()
    StartTest "Catalog file names"
    AssertEqual True, IsAddinFile("GCV Student Lookup.xlam"), "normal"
    AssertEqual True, IsAddinFile("GCV English-Names 2.xlam"), "dash and digit"
    AssertEqual False, IsAddinFile("GCV Loader.xlam"), "loader"
    AssertEqual False, IsAddinFile("..\GCV Evil.xlam"), "folder"
    AssertEqual False, IsAddinFile("GCV a.b.xlam"), "dot inside"
    AssertEqual False, IsAddinFile("Other.xlam"), "no prefix"
    AssertEqual False, IsAddinFile("GCV Tool.xlsm"), "not an add-in"
End Sub

Private Sub Test_ParseCatalog_ReadsTools()
    Dim catalog As Collection
    StartTest "Catalog entries are read"
    Set catalog = ParseCatalog(CatalogJson(True))
    AssertEqual 2, catalog.Count, "two tools"
    AssertEqual TOOL_FILE, catalog(1)("file"), "file"
    AssertEqual "test-tool-v", catalog(1)("tagPrefix"), "tag prefix"
    AssertEqual True, catalog(1)("required"), "required"
    AssertEqual False, catalog(2)("required"), "optional"
End Sub

Private Sub Test_ParseCatalog_RejectsBadEntries()
    StartTest "Bad catalog entries are refused"
    AssertEqual UPD_ERR_BASE + ueManifest, CatalogError("{""tools"": [" & ToolJson("GCV/Upper", TOOL_FILE, TOOL_REPO, True) & "]}"), "id"
    AssertEqual UPD_ERR_BASE + ueManifest, CatalogError("{""tools"": [" & ToolJson(TOOL_ID, "..\x.xlam", TOOL_REPO, True) & "]}"), "file"
    AssertEqual UPD_ERR_BASE + ueManifest, CatalogError("{""tools"": [" & ToolJson(TOOL_ID, TOOL_FILE, "no-slash", True) & "]}"), "repo"
    AssertEqual UPD_ERR_BASE + ueManifest, CatalogError("{""tools"": [" & ToolJson(TOOL_ID, TOOL_FILE, TOOL_REPO, True) & "," & _
                                                        ToolJson(TOOL_ID, EXTRA_FILE, TOOL_REPO, True) & "]}"), "twice"
    AssertEqual UPD_ERR_BASE + ueManifest, CatalogError("[]"), "not an object"
End Sub

Private Sub Test_BuildPlan_UpdatesNewerAndAddsMissingTools()
    Dim plan As Collection
    StartTest "Plan: newer versions are updated, missing tools added, current ones left"
    Set plan = BuildPlan(ParseCatalog(CatalogJson(True)), Dict(TOOL_ID, "1.0.0"), Dict(TOOL_ID, "1.1.0", EXTRA_ID, "0.1.0"), Dict())
    AssertEqual 2, plan.Count, "two items"
    AssertEqual "update", plan(1)("kind"), "update"
    AssertEqual "1.0.0", plan(1)("current"), "from"
    AssertEqual "1.1.0", plan(1)("new"), "to"
    AssertEqual "new", plan(2)("kind"), "new tool"
    AssertEqual "", plan(2)("current"), "not installed"
    Set plan = BuildPlan(ParseCatalog(CatalogJson(True)), Dict(TOOL_ID, "1.1.0", EXTRA_ID, "0.1.0"), _
                         Dict(TOOL_ID, "1.1.0", EXTRA_ID, "0.1.0"), Dict())
    AssertEqual 0, plan.Count, "up to date"
End Sub

Private Sub Test_BuildPlan_LoggerFirstUpdaterLast()
    Dim plan As Collection
    Dim json As String
    StartTest "Plan: logger first, updater last"
    json = "{""tools"": [" & ToolJson("gcv/updater", "GCV Updater.xlam", "o/excel-updater", True) & "," & _
           ToolJson(TOOL_ID, TOOL_FILE, TOOL_REPO, True) & "," & _
           ToolJson("gcv/logger", "GCV Logger.xlam", "o/excel-logger", True) & "]}"
    Set plan = BuildPlan(ParseCatalog(json), Dict(), Dict("gcv/updater", "2.0.0", TOOL_ID, "1.0.0", "gcv/logger", "1.0.0"), Dict())
    AssertEqual "gcv/logger,gcv/test-tool,gcv/updater", PlanIds(plan), "order"
End Sub

Private Sub Test_BuildPlan_SkipsBadVersions()
    Dim plan As Collection
    StartTest "Plan: a version that did not start before is skipped"
    Set plan = BuildPlan(ParseCatalog(CatalogJson(True)), Dict(TOOL_ID, "1.0.0", EXTRA_ID, "0.1.0"), _
                         Dict(TOOL_ID, "1.1.0", EXTRA_ID, "0.1.0"), Dict(TOOL_FILE, "1.0.5;1.1.0"))
    AssertEqual 0, plan.Count, "skipped"
End Sub

' ----------------------------------------------------------------------------
' Download and swap (files only)
' ----------------------------------------------------------------------------

Private Sub Test_Download_SavesVerifiedFileOnly()
    Dim fake As UpdFakeHttp
    Dim folder As String
    Dim n As Long
    StartTest "Download: verified file saved; wrong checksum saves nothing"
    folder = TempFolder("download")
    Set fake = FakeRepo(ShaOf(NEW_CONTENT))
    DownloadToolsFile fake, ToolRelease(), JsonParse(ManifestJson(TOOL_ID, "1.1.0", ShaOf(NEW_CONTENT), Len(NEW_CONTENT))), _
                      "1.1.0", folder, TOOL_FILE
    AssertEqual NEW_CONTENT, ReadAscii(folder & "\" & TOOL_FILE), "saved"
    AssertEqual False, FileExists(folder & "\" & TOOL_FILE & ".part"), "no .part"
    On Error Resume Next
    DownloadToolsFile fake, ToolRelease(), JsonParse(ManifestJson(TOOL_ID, "1.1.0", ShaOf("other"), Len(NEW_CONTENT))), _
                      "1.1.0", folder & "\b", TOOL_FILE
    n = Err.Number
    On Error GoTo 0
    AssertEqual UPD_ERR_BASE + ueHash, n, "checksum error"
    AssertEqual False, FileExists(folder & "\b\" & TOOL_FILE), "nothing saved"
End Sub

Private Sub Test_Swap_InstallsAndKeepsPrevious()
    Dim base As String
    StartTest "Swap: new version in addins\, old one in previous\"
    base = PreparedBase("swap")
    Updater.UseTestDoubles New UpdFakeHttp, New Collection, base, Dict(), False
    AssertEqual irInstalled, Updater.SwapAddin(TOOL_FILE, "1.1.0"), "installed"
    AssertEqual NEW_CONTENT, ReadAscii(base & "\addins\" & TOOL_FILE), "new in place"
    AssertEqual OLD_CONTENT, ReadAscii(base & "\previous\" & TOOL_FILE), "old kept"
    AssertEqual False, FileExists(base & "\pending\" & TOOL_FILE), "pending emptied"
    Updater.ResetTestDoubles
End Sub

Private Sub Test_Swap_RestoresPreviousWhenNewDoesNotStart()
    Dim base As String
    StartTest "Swap: a new version that does not start is replaced by the old one and marked bad"
    base = PreparedBase("swap-fail")
    Updater.UseTestDoubles New UpdFakeHttp, New Collection, base, Dict(), True
    AssertEqual irFailedKeptOld, Updater.SwapAddin(TOOL_FILE, "1.1.0"), "failed"
    AssertEqual OLD_CONTENT, ReadAscii(base & "\addins\" & TOOL_FILE), "old back in place"
    AssertEqual "1.1.0", GetSetting("GCV Tools", "BadVersions", TOOL_FILE, ""), "marked bad"
    Updater.ResetTestDoubles
    CleanRegistry
End Sub

' ----------------------------------------------------------------------------
' Whole update (fake network, queued answers)
' ----------------------------------------------------------------------------

Private Sub Test_Update_OfflineSaysSoAndChangesNothing()
    Dim fake As UpdFakeHttp
    StartTest "Offline: one message, nothing changed"
    Set fake = New UpdFakeHttp
    fake.Offline = True
    Updater.UseTestDoubles fake, Answers(True), PreparedBase("offline"), Dict(TOOL_ID, "1.0.0"), False
    Updater.UpdateTools
    AssertEqual 1, Updater.TestMessages().Count, "one message"
    AssertEqual Updater.T_Offline(), Updater.TestMessages()(1), "offline"
    Updater.ResetTestDoubles
End Sub

Private Sub Test_Update_UpToDateListsVersions()
    StartTest "Up to date: message lists the installed versions"
    Updater.UseTestDoubles FakeRepo(ShaOf(NEW_CONTENT)), Answers(True), PreparedBase("uptodate"), _
                           Dict(TOOL_ID, "1.1.0", EXTRA_ID, "0.1.0"), False
    Updater.UpdateTools
    AssertEqual 1, Updater.TestMessages().Count, "one message"
    AssertEqual True, InStr(Updater.TestMessages()(1), "GCV Test Tool 1.1.0") > 0, "versions listed"
    Updater.ResetTestDoubles
End Sub

Private Sub Test_Update_InstallsUpdateAndOfferedNewTool()
    Dim base As String
    StartTest "Update: optional new tool offered, then both installed"
    base = PreparedBase("install")
    ' answers: install the optional new tool? yes; install the plan? yes
    Updater.UseTestDoubles FakeRepo(ShaOf(NEW_CONTENT)), Answers(True, True), base, Dict(TOOL_ID, "1.0.0"), False
    Updater.UpdateTools
    AssertEqual NEW_CONTENT, ReadAscii(base & "\addins\" & TOOL_FILE), "update installed"
    AssertEqual NEW_CONTENT, ReadAscii(base & "\addins\" & EXTRA_FILE), "new tool installed"
    AssertEqual 3, Updater.TestMessages().Count, "offer, plan, results"
    AssertEqual True, InStr(Updater.TestMessages()(2), "1.0.0") > 0, "plan shows the old version"
    AssertEqual True, InStr(Updater.TestMessages()(3), Updater.T_Result(irInstalled)) > 0, "results"
    Updater.ResetTestDoubles
    CleanRegistry
End Sub

Private Sub Test_Update_DeclinedChangesNothing()
    Dim base As String
    StartTest "Update declined: nothing downloaded or replaced"
    base = PreparedBase("declined")
    Updater.UseTestDoubles FakeRepo(ShaOf(NEW_CONTENT)), Answers(False, False), base, Dict(TOOL_ID, "1.0.0"), False
    Updater.UpdateTools
    AssertEqual OLD_CONTENT, ReadAscii(base & "\addins\" & TOOL_FILE), "old still in place"
    AssertEqual False, FileExists(base & "\addins\" & EXTRA_FILE), "new tool not installed"
    Updater.ResetTestDoubles
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

Private Function ToolJson(ByVal id As String, ByVal file As String, ByVal repo As String, ByVal required As Boolean) As String
    ' Backslashes must be escaped in JSON (also in the name built from the file).
    ToolJson = "{""id"": """ & id & """, ""name"": """ & Replace(Replace(file, ".xlam", ""), "\", "\\") & """, ""file"": """ & _
               Replace(file, "\", "\\") & """, ""repo"": """ & repo & """, ""required"": " & LCase$(CStr(required)) & _
               ", ""description"": ""Test tool""}"
End Function

Private Function CatalogJson(ByVal withExtra As Boolean) As String
    CatalogJson = "{""tools"": [" & ToolJson(TOOL_ID, TOOL_FILE, TOOL_REPO, True) & _
                  IIf(withExtra, "," & ToolJson(EXTRA_ID, EXTRA_FILE, EXTRA_REPO, False), "") & "]}"
End Function

Private Function ReleaseJson(ByVal tag As String, ByVal draft As Boolean, ByVal prerelease As Boolean, _
                             ByVal urlBase As String) As String
    ReleaseJson = "{""tag_name"": """ & tag & """, ""draft"": " & LCase$(CStr(draft)) & ", ""prerelease"": " & _
                  LCase$(CStr(prerelease)) & ", ""body"": ""- Test change"", ""assets"": [" & _
                  "{""name"": ""manifest.json"", ""browser_download_url"": ""https://fake/" & urlBase & "/manifest.json""}, " & _
                  "{""name"": ""GCV-Test.xlam"", ""browser_download_url"": ""https://fake/" & urlBase & "/GCV-Test.xlam""}]}"
End Function

Private Function ToolRelease() As Object
    Set ToolRelease = JsonParse(ReleaseJson("test-tool-v1.1.0", False, False, "tool"))
End Function

Private Function ManifestJson(ByVal id As String, ByVal version As String, ByVal sha As String, ByVal size As Long) As String
    ManifestJson = "{""id"": """ & id & """, ""version"": """ & version & """, ""minLoaderVersion"": ""2.0.0"", " & _
                   """file"": {""name"": ""GCV-Test.xlam"", ""sha256"": """ & sha & """, ""size"": " & size & "}}"
End Function

' Catalog with the test tool (1.1.0) and the optional extra tool (0.1.0).
Private Function FakeRepo(ByVal sha As String) As UpdFakeHttp
    Dim fake As UpdFakeHttp
    Set fake = New UpdFakeHttp
    fake.RespondText Updater.CATALOG_URL, CatalogJson(True)
    fake.RespondText ReleasesUrl(TOOL_REPO), "[" & ReleaseJson("test-tool-v1.1.0", False, False, "tool") & "]"
    fake.RespondText "https://fake/tool/manifest.json", ManifestJson(TOOL_ID, "1.1.0", sha, Len(NEW_CONTENT))
    fake.RespondText "https://fake/tool/GCV-Test.xlam", NEW_CONTENT
    fake.RespondText ReleasesUrl(EXTRA_REPO), "[" & ReleaseJson("test-extra-v0.1.0", False, False, "extra") & "]"
    fake.RespondText "https://fake/extra/manifest.json", ManifestJson(EXTRA_ID, "0.1.0", sha, Len(NEW_CONTENT))
    fake.RespondText "https://fake/extra/GCV-Test.xlam", NEW_CONTENT
    Set FakeRepo = fake
End Function

' Temporary ExcelTools folder: old test tool installed, new one pending.
Private Function PreparedBase(ByVal name As String) As String
    PreparedBase = TempFolder(name)
    EnsureFolder PreparedBase & "\addins"
    EnsureFolder PreparedBase & "\pending"
    WriteAscii PreparedBase & "\addins\" & TOOL_FILE, OLD_CONTENT
    WriteAscii PreparedBase & "\pending\" & TOOL_FILE, NEW_CONTENT
End Function

Private Function ManifestError(ByVal json As String) As Long
    On Error Resume Next
    ValidateToolsManifest JsonParse(json), TOOL_ID, "1.1.0"
    ManifestError = Err.Number
    On Error GoTo 0
End Function

Private Function CatalogError(ByVal json As String) As Long
    On Error Resume Next
    ParseCatalog json
    CatalogError = Err.Number
    On Error GoTo 0
End Function

Private Function PlanIds(ByVal plan As Collection) As String
    Dim item As Variant
    For Each item In plan
        If Len(PlanIds) > 0 Then PlanIds = PlanIds & ","
        PlanIds = PlanIds & item("id")
    Next item
End Function

Private Function Dict(ParamArray pairs() As Variant) As Object
    Dim result As Object
    Dim i As Long
    Set result = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(pairs) Step 2
        result(pairs(i)) = pairs(i + 1)
    Next i
    Set Dict = result
End Function

Private Function Answers(ParamArray values() As Variant) As Collection
    Dim result As Collection
    Dim i As Long
    Set result = New Collection
    For i = LBound(values) To UBound(values)
        result.Add CBool(values(i))
    Next i
    Set Answers = result
End Function

Private Sub CleanRegistry()
    Dim fileName As Variant
    On Error Resume Next
    For Each fileName In Array(TOOL_FILE, EXTRA_FILE)
        DeleteSetting "GCV Tools", "BadVersions", fileName
        DeleteSetting "GCV Tools", "PendingVersions", fileName
    Next fileName
    On Error GoTo 0
End Sub

Private Function ShaOf(ByVal text As String) As String
    Dim data() As Byte
    data = AsciiBytes(text)
    ShaOf = Sha256Hex(data)
End Function

Private Function TempFolder(ByVal name As String) As String
    TempFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\gcv-updater-tests\" & _
                 name & "-" & Format$(Now, "hhnnss") & Int(Rnd() * 10000)
    EnsureFolder TempFolder
End Function

Private Function FileExists(ByVal path As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

Private Function ReadAscii(ByVal path As String) As String
    Dim data() As Byte
    If Not FileExists(path) Then Exit Function
    data = ReadFileBytes(path)
    ReadAscii = StrConv(data, vbUnicode)
End Function

Private Sub WriteAscii(ByVal path As String, ByVal text As String)
    Dim data() As Byte
    data = AsciiBytes(text)
    WriteFileBytes path, data
End Sub

Private Sub StartTest(ByVal name As String)
    currentTest = name
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    Dim same As Boolean
    If VarType(expected) = vbBoolean Or VarType(actual) = vbBoolean Then
        same = (CStr(expected) = CStr(actual))
    ElseIf IsNumeric(expected) And IsNumeric(actual) And VarType(expected) <> vbString And VarType(actual) <> vbString Then
        same = (CDbl(expected) = CDbl(actual))
    Else
        same = (CStr(expected) = CStr(actual))
    End If
    If same Then
        passedCount = passedCount + 1
    Else
        failedCount = failedCount + 1
        Debug.Print "FAIL: " & currentTest & " / " & label & " - expected [" & CStr(expected) & "] got [" & CStr(actual) & "]"
    End If
End Sub

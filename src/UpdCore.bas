Attribute VB_Name = "UpdCore"
Option Explicit

' ============================================================================
' GCV Updater - core helpers (no user interface): versions, SHA-256, files,
' GitHub release selection, add-in manifests and verified downloads.
' ============================================================================

Public Const UPD_ERR_BASE As Long = vbObjectError + 2000
Public Const MANIFEST_NAME As String = "manifest.json"

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByVal pbHashObject As LongPtr, _
    ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, _
    ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As Long, _
    ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, _
    ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As LongPtr) As Long

Private Const SHA256_BYTES As Long = 32
Private Const AD_TYPE_BINARY As Long = 1
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_SAVE_OVERWRITE As Long = 2

Public Enum UpdError
    ueVersion = 1
    ueHash = 2
    ueManifest = 3
    ueDownload = 4
    ueInstall = 5
    ueLoader = 6
End Enum

Public Sub RaiseUpd(ByVal code As UpdError, ByVal message As String)
    Err.Raise UPD_ERR_BASE + code, "GCV updater", message
End Sub

' ----------------------------------------------------------------------------
' Logging through GCV Logger (GcvLogClient keeps the caller's Err).
' ----------------------------------------------------------------------------
Public Sub UpdLog(ByVal levelName As String, ByVal message As String)
    GcvLogClient.LogAt levelName, "updater", message
End Sub

' ----------------------------------------------------------------------------
' Versions (MAJOR.MINOR.PATCH, optional leading "v")
' ----------------------------------------------------------------------------

Public Function IsValidVersion(ByVal text As String) As Boolean
    Dim parts() As String
    Dim i As Long

    text = Trim$(text)
    If LCase$(Left$(text, 1)) = "v" Then text = Mid$(text, 2)
    parts = Split(text, ".")
    If UBound(parts) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(parts(i)) = 0 Or Len(parts(i)) > 6 Then Exit Function
        If Not parts(i) Like String$(Len(parts(i)), "#") Then Exit Function
    Next i
    IsValidVersion = True
End Function

Public Function CompareVersions(ByVal a As String, ByVal b As String) As Long
    Dim pa() As Long
    Dim pb() As Long
    Dim i As Long

    pa = VersionParts(a)
    pb = VersionParts(b)
    For i = 0 To 2
        If pa(i) < pb(i) Then
            CompareVersions = -1
            Exit Function
        End If
        If pa(i) > pb(i) Then
            CompareVersions = 1
            Exit Function
        End If
    Next i
End Function

Private Function VersionParts(ByVal text As String) As Long()
    Dim parts() As String
    Dim result() As Long
    Dim i As Long

    If Not IsValidVersion(text) Then RaiseUpd ueVersion, "Invalid version: " & text
    text = Trim$(text)
    If LCase$(Left$(text, 1)) = "v" Then text = Mid$(text, 2)
    parts = Split(text, ".")
    ReDim result(0 To 2)
    For i = 0 To 2
        result(i) = CLng(parts(i))
    Next i
    VersionParts = result
End Function

' ----------------------------------------------------------------------------
' SHA-256 (Windows CNG, no .NET needed)
' ----------------------------------------------------------------------------

Public Function Sha256Hex(ByRef data() As Byte) As String
    Dim hAlg As LongPtr
    Dim hHash As LongPtr
    Dim digest() As Byte
    Dim status As Long
    Dim size As Long

    ReDim digest(0 To SHA256_BYTES - 1)
    status = BCryptOpenAlgorithmProvider(hAlg, StrPtr("SHA256"), 0, 0)
    If status <> 0 Then RaiseUpd ueHash, "Cannot open SHA-256 provider (0x" & Hex$(status) & ")."

    On Error GoTo Cleanup
    status = BCryptCreateHash(hAlg, hHash, 0, 0, 0, 0, 0)
    If status <> 0 Then RaiseUpd ueHash, "Cannot create hash (0x" & Hex$(status) & ")."
    size = ByteCount(data)
    If size > 0 Then
        status = BCryptHashData(hHash, VarPtr(data(LBound(data))), size, 0)
        If status <> 0 Then RaiseUpd ueHash, "Cannot hash data (0x" & Hex$(status) & ")."
    End If
    status = BCryptFinishHash(hHash, VarPtr(digest(0)), SHA256_BYTES, 0)
    If status <> 0 Then RaiseUpd ueHash, "Cannot finish hash (0x" & Hex$(status) & ")."
    Sha256Hex = BytesToHex(digest)

Cleanup:
    If hHash <> 0 Then BCryptDestroyHash hHash
    If hAlg <> 0 Then BCryptCloseAlgorithmProvider hAlg, 0
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ByteCount(ByRef data() As Byte) As Long
    ' UBound raises error 9 on an unallocated array; that means 0 bytes.
    On Error Resume Next
    ByteCount = UBound(data) - LBound(data) + 1
    If Err.Number <> 0 Then ByteCount = 0
    On Error GoTo 0
End Function

Public Function BytesToHex(ByRef data() As Byte) As String
    Dim i As Long
    Dim parts() As String

    If ByteCount(data) = 0 Then Exit Function
    ReDim parts(LBound(data) To UBound(data))
    For i = LBound(data) To UBound(data)
        parts(i) = Right$("0" & LCase$(Hex$(data(i))), 2)
    Next i
    BytesToHex = Join(parts, vbNullString)
End Function

' ----------------------------------------------------------------------------
' Text and files
' ----------------------------------------------------------------------------

Public Function Utf8Decode(ByRef data() As Byte) As String
    Dim stream As Object

    If ByteCount(data) = 0 Then Exit Function
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = AD_TYPE_BINARY
    stream.Open
    stream.Write data
    stream.Position = 0
    stream.Type = AD_TYPE_TEXT
    stream.Charset = "utf-8"
    Utf8Decode = stream.ReadText
    stream.Close
End Function

Public Function AsciiBytes(ByVal text As String) As Byte()
    AsciiBytes = StrConv(text, vbFromUnicode)
End Function

Public Function ReadFileBytes(ByVal path As String) As Byte()
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = AD_TYPE_BINARY
    stream.Open
    stream.LoadFromFile path
    ReadFileBytes = stream.Read
    stream.Close
End Function

Public Sub WriteFileBytes(ByVal path As String, ByRef data() As Byte)
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = AD_TYPE_BINARY
    stream.Open
    stream.Write data
    stream.SaveToFile path, AD_SAVE_OVERWRITE
    stream.Close
End Sub

Public Function ReadFileText(ByVal path As String) As String
    Dim data() As Byte
    data = ReadFileBytes(path)
    ReadFileText = Utf8Decode(data)
End Function

' FileSystemObject handles Unicode paths (Arabic Windows user names); MkDir/Dir$ do not.
Public Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Dim parts() As String
    Dim current As String
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    parts = Split(path, "\")
    current = parts(0)
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            current = current & "\" & parts(i)
            If Not fso.FolderExists(current) Then fso.CreateFolder current
        End If
    Next i
End Sub

' ----------------------------------------------------------------------------
' GitHub releases
' ----------------------------------------------------------------------------

Public Function ReleasesUrl(ByVal repo As String) As String
    ReleasesUrl = "https://api.github.com/repos/" & repo & "/releases?per_page=100"
End Function

' Highest non-draft, non-prerelease release whose tag is prefix + version.
Public Function PickLatestRelease(ByVal releases As Collection, ByVal tagPrefix As String) As Object
    Dim rel As Variant
    Dim tag As String
    Dim ver As String
    Dim best As Object
    Dim bestVersion As String

    For Each rel In releases
        If IsObject(rel) Then
            If Not JsonBool(rel, "draft") And Not JsonBool(rel, "prerelease") Then
                tag = JsonStr(rel, "tag_name")
                If Left$(tag, Len(tagPrefix)) = tagPrefix Then
                    ver = Mid$(tag, Len(tagPrefix) + 1)
                    If IsValidVersion(ver) Then
                        If best Is Nothing Then
                            Set best = rel
                            bestVersion = ver
                        ElseIf CompareVersions(ver, bestVersion) > 0 Then
                            Set best = rel
                            bestVersion = ver
                        End If
                    End If
                End If
            End If
        End If
    Next rel
    Set PickLatestRelease = best
End Function

Public Function ReleaseVersion(ByVal release As Object, ByVal tagPrefix As String) As String
    ReleaseVersion = Mid$(JsonStr(release, "tag_name"), Len(tagPrefix) + 1)
    If LCase$(Left$(ReleaseVersion, 1)) = "v" Then ReleaseVersion = Mid$(ReleaseVersion, 2)
End Function

Public Function AssetUrl(ByVal release As Object, ByVal assetName As String) As String
    Dim assets As Object
    Dim asset As Variant

    Set assets = JsonObj(release, "assets")
    If assets Is Nothing Then Exit Function
    For Each asset In assets
        If IsObject(asset) Then
            If JsonStr(asset, "name") = assetName Then
                AssetUrl = JsonStr(asset, "browser_download_url")
                Exit Function
            End If
        End If
    Next asset
End Function

' Release notes for display: markdown marks removed, length limited.
Public Function PlainChangelog(ByVal text As String, ByVal maxLength As Long) As String
    Dim lines() As String
    Dim i As Long
    Dim ln As String

    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(text, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        Do While Left$(ln, 1) = "#"
            ln = Trim$(Mid$(ln, 2))
        Loop
        If Left$(ln, 2) = "- " Or Left$(ln, 2) = "* " Then ln = ChrW$(&H2022) & " " & Mid$(ln, 3)
        ln = Replace(Replace(ln, "**", vbNullString), "`", vbNullString)
        lines(i) = ln
    Next i
    text = Trim$(Join(lines, vbCrLf))
    Do While InStr(text, vbCrLf & vbCrLf & vbCrLf) > 0
        text = Replace(text, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop
    If Len(text) > maxLength Then text = Left$(text, maxLength) & " ..."
    PlainChangelog = text
End Function

Public Function IsSafeFileName(ByVal name As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(name) = 0 Or Len(name) > 64 Then Exit Function
    If Left$(name, 1) = "." Then Exit Function
    For i = 1 To Len(name)
        ch = Mid$(name, i, 1)
        If Not (ch Like "[A-Za-z]" Or ch Like "#" Or ch = "_" Or ch = "-" Or ch = ".") Then Exit Function
    Next i
    IsSafeFileName = (InStr(name, "..") = 0)
End Function

Public Function IsSha256Hex(ByVal text As String) As Boolean
    Dim i As Long
    If Len(text) <> 64 Then Exit Function
    For i = 1 To 64
        If Not Mid$(text, i, 1) Like "[0-9a-f]" Then Exit Function
    Next i
    IsSha256Hex = True
End Function

' ----------------------------------------------------------------------------
' Add-in manifest (manifest.json of a release <package>-v<version>)
' ----------------------------------------------------------------------------
' {
'   "id": "gcv/student-lookup", "version": "3.0.0", "minLoaderVersion": "2.0.0",
'   "file": {"name": "GCV-Student-Lookup.xlam", "sha256": "...", "size": 123456},
'   "changelog": "..."
' }

Public Function FetchToolsManifest(ByVal http As IUpdHttp, ByVal release As Object, _
                                   ByVal toolsId As String, ByVal version As String) As Object
    Dim url As String
    Dim result As UpdHttpResult
    Dim body() As Byte
    Dim parsed As Variant

    url = AssetUrl(release, MANIFEST_NAME)
    If Len(url) = 0 Then RaiseUpd ueDownload, "The release has no " & MANIFEST_NAME & "."
    Set result = http.GetUrl(url, 20000)
    If Not result.IsOk Then RaiseUpd ueDownload, "Cannot download " & MANIFEST_NAME & ": " & result.Describe
    body = result.Body
    JsonParseInto Utf8Decode(body), parsed
    If Not IsObject(parsed) Then RaiseUpd ueManifest, MANIFEST_NAME & " is not a JSON object."
    If TypeName(parsed) <> "Dictionary" Then RaiseUpd ueManifest, MANIFEST_NAME & " is not a JSON object."
    ValidateToolsManifest parsed, toolsId, version
    Set FetchToolsManifest = parsed
End Function

Public Sub ValidateToolsManifest(ByVal manifest As Object, ByVal toolsId As String, ByVal version As String)
    Dim file As Object

    If JsonStr(manifest, "id") <> toolsId Then _
        RaiseUpd ueManifest, "Manifest is for '" & JsonStr(manifest, "id") & "', expected '" & toolsId & "'."
    If Not IsValidVersion(JsonStr(manifest, "version")) Then RaiseUpd ueManifest, "Manifest version is invalid."
    If CompareVersions(JsonStr(manifest, "version"), version) <> 0 Then _
        RaiseUpd ueManifest, "Manifest version " & JsonStr(manifest, "version") & " does not match release " & version & "."
    If Len(JsonStr(manifest, "minLoaderVersion")) > 0 Then
        If Not IsValidVersion(JsonStr(manifest, "minLoaderVersion")) Then RaiseUpd ueManifest, "minLoaderVersion is invalid."
    End If
    Set file = JsonObj(manifest, "file")
    If file Is Nothing Then RaiseUpd ueManifest, "Manifest has no file."
    If TypeName(file) <> "Dictionary" Then RaiseUpd ueManifest, "Manifest file entry is not an object."
    If Not IsSafeFileName(JsonStr(file, "name")) Then RaiseUpd ueManifest, "Unsafe file name: " & JsonStr(file, "name")
    If LCase$(Right$(JsonStr(file, "name"), 5)) <> ".xlam" Then RaiseUpd ueManifest, "The file must be an .xlam add-in."
    If Not IsSha256Hex(JsonStr(file, "sha256")) Then RaiseUpd ueManifest, "Invalid sha256."
    If Len(JsonStr(file, "size")) > 0 Then
        If Not IsNumeric(JsonStr(file, "size")) Then RaiseUpd ueManifest, "Invalid size."
        If CDbl(JsonStr(file, "size")) <= 0 Then RaiseUpd ueManifest, "Invalid size."
    End If
End Sub

' True when this loader can start the release (no requirement, or no loader
' found - a DEV copy opened by hand).
Public Function LoaderIsEnough(ByVal manifest As Object, ByVal loaderVersion As String) As Boolean
    LoaderIsEnough = True
    If Len(JsonStr(manifest, "minLoaderVersion")) = 0 Or Len(loaderVersion) = 0 Then Exit Function
    LoaderIsEnough = (CompareVersions(loaderVersion, JsonStr(manifest, "minLoaderVersion")) >= 0)
End Function

' ----------------------------------------------------------------------------
' Download of an add-in into <folder>\<finalName> (normally pending\).
' Written as .part, checked again after writing, then renamed, so the loader
' never sees a half-written file. Returns the final path.
' ----------------------------------------------------------------------------

Public Function DownloadToolsFile(ByVal http As IUpdHttp, ByVal release As Object, ByVal manifest As Object, _
                                  ByVal version As String, ByVal folder As String, _
                                  ByVal finalName As String) As String
    Dim fso As Object
    Dim file As Object
    Dim url As String
    Dim result As UpdHttpResult
    Dim body() As Byte
    Dim written() As Byte
    Dim partPath As String
    Dim finalPath As String
    Dim errNumber As Long
    Dim errDescription As String

    Set file = JsonObj(manifest, "file")
    url = AssetUrl(release, JsonStr(file, "name"))
    If Len(url) = 0 Then RaiseUpd ueDownload, "The release has no file " & JsonStr(file, "name") & "."
    Set result = http.GetUrl(url, 120000)
    If Not result.IsOk Then RaiseUpd ueDownload, "Cannot download " & JsonStr(file, "name") & ": " & result.Describe
    body = result.Body
    If Len(JsonStr(file, "size")) > 0 Then
        If ByteCount(body) <> CDbl(JsonStr(file, "size")) Then _
            RaiseUpd ueHash, "Size does not match: " & ByteCount(body) & " bytes, expected " & JsonStr(file, "size") & "."
    End If
    If Sha256Hex(body) <> JsonStr(file, "sha256") Then _
        RaiseUpd ueHash, JsonStr(file, "name") & ": checksum does not match (file changed or damaged)."

    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder folder
    finalPath = folder & "\" & finalName
    partPath = finalPath & ".part"
    On Error GoTo Failed
    If fso.FileExists(partPath) Then fso.DeleteFile partPath, True
    WriteFileBytes partPath, body
    written = ReadFileBytes(partPath)
    If Sha256Hex(written) <> JsonStr(file, "sha256") Then RaiseUpd ueHash, "The saved file is damaged (disk full?)."
    If fso.FileExists(finalPath) Then fso.DeleteFile finalPath, True
    fso.MoveFile partPath, finalPath
    DownloadToolsFile = finalPath
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error GoTo -1                       ' leave the handler so the clean-up can trap its own errors
    On Error Resume Next
    If fso.FileExists(partPath) Then fso.DeleteFile partPath, True
    On Error GoTo 0
    Err.Raise errNumber, "GCV updater", errDescription
End Function

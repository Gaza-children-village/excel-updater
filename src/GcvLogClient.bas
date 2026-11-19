Attribute VB_Name = "GcvLogClient"
Option Explicit

' ============================================================================
' Logging for GCV add-ins other than GCV Logger itself. Lines go to the
' logger add-in (Application.Run, so this add-in works and compiles without
' it); when the logger is not open they go to the Immediate window.
' Never log personal data (names, IDs) - only what the add-in does.
' The caller's Err is left unchanged.
' ============================================================================

Private Const LOGGER_FILE As String = "GCV Logger.xlam"

Private mCapturing As Boolean            ' tests: keep lines here instead of sending them
Private mCaptured As String

Public Sub LogDebug(ByVal source As String, ByVal message As String)
    SendLine "DEBUG", source, message
End Sub

Public Sub LogInfo(ByVal source As String, ByVal message As String)
    SendLine "INFO", source, message
End Sub

Public Sub LogWarn(ByVal source As String, ByVal message As String)
    SendLine "WARN", source, message
End Sub

' Adds the current VBA error (if any) to the message.
Public Sub LogError(ByVal source As String, ByVal message As String)
    If Err.Number <> 0 Then message = message & " (error " & Err.Number & ": " & Err.Description & ")"
    SendLine "ERROR", source, message
End Sub

' Tests: True = every line (all levels) is kept in memory and not sent;
' read them with CapturedLog. False = normal logging again.
Public Sub CaptureLogForTests(ByVal enabled As Boolean)
    mCapturing = enabled
    mCaptured = vbNullString
End Sub

Public Function CapturedLog() As String
    CapturedLog = mCaptured
End Function

' For callers that choose the level at run time (e.g. "INFO").
Public Sub LogAt(ByVal levelName As String, ByVal source As String, ByVal message As String)
    SendLine UCase$(levelName), source, message
End Sub

Private Sub SendLine(ByVal levelName As String, ByVal source As String, ByVal message As String)
    Dim savedNumber As Long
    Dim savedDescription As String
    Dim savedSource As String
    Dim logger As Workbook
    Dim sent As Boolean

    savedNumber = Err.Number
    savedDescription = Err.Description
    savedSource = Err.Source
    If mCapturing Then
        mCaptured = mCaptured & "[" & source & "] - " & levelName & ": " & message & vbCrLf
        Exit Sub
    End If
    On Error Resume Next
    ' Look the logger up first: Application.Run on a closed file would try to open it.
    Set logger = Workbooks(LOGGER_FILE)
    If Not logger Is Nothing Then
        Err.Clear
        Application.Run "'" & LOGGER_FILE & "'!GcvLog.LogWrite", levelName, source, message
        sent = (Err.Number = 0)
    End If
    If Not sent Then Debug.Print "[" & source & "] - " & levelName & ": " & Format$(Now, "yyyy-mm-dd - hh:nn:ss") & " " & message
    Err.Clear
    On Error GoTo 0
    If savedNumber <> 0 Then
        Err.Number = savedNumber
        Err.Description = savedDescription
        Err.Source = savedSource
    End If
End Sub

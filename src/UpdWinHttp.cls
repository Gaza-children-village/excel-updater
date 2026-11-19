VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdWinHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUpdHttp

' HTTPS GET with WinHTTP. GitHub requires a User-Agent header; redirects
' (release downloads) are followed automatically.

Private Const USER_AGENT As String = "GCV-Excel-Updater"

Private Function IUpdHttp_GetUrl(ByVal url As String, ByVal timeoutMs As Long) As UpdHttpResult
    Dim result As UpdHttpResult
    Dim http As Object
    Dim data() As Byte

    Set result = New UpdHttpResult
    ' Network failures are an expected outcome (offline schools): they are
    ' returned as StatusCode 0 with the error text, and logged.
    On Error GoTo Offline
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", USER_AGENT
    If InStr(1, url, "://api.github.com/", vbTextCompare) > 0 Then
        http.SetRequestHeader "Accept", "application/vnd.github+json"
    Else
        http.SetRequestHeader "Accept", "*/*"
    End If
    http.Send
    On Error GoTo 0

    result.StatusCode = http.Status
    If result.IsOk Then
        data = http.ResponseBody
        result.SetBody data
    End If
    UpdLog "DEBUG", "GET " & url & " -> HTTP " & http.Status
    Set IUpdHttp_GetUrl = result
    Exit Function

Offline:
    result.StatusCode = 0
    result.ErrorText = Err.Description
    UpdLog "DEBUG", "GET " & url & " failed: " & Err.Description
    Set IUpdHttp_GetUrl = result
End Function

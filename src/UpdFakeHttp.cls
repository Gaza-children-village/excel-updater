VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdFakeHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUpdHttp

' Test double: answers registered URLs; unknown URLs give HTTP 404.

Public Offline As Boolean
Public Requests As Collection

Private mStatus As Object
Private mBodies As Object

Private Sub Class_Initialize()
    Set mStatus = CreateObject("Scripting.Dictionary")
    Set mBodies = CreateObject("Scripting.Dictionary")
    Set Requests = New Collection
End Sub

Public Sub RespondText(ByVal url As String, ByVal text As String, Optional ByVal statusCode As Long = 200)
    mStatus(url) = statusCode
    mBodies(url) = StrConv(text, vbFromUnicode)
End Sub

Private Function IUpdHttp_GetUrl(ByVal url As String, ByVal timeoutMs As Long) As UpdHttpResult
    Dim result As UpdHttpResult
    Dim data() As Byte

    Requests.Add url
    Set result = New UpdHttpResult
    If Offline Then
        result.StatusCode = 0
        result.ErrorText = "fake: offline"
    ElseIf mStatus.Exists(url) Then
        result.StatusCode = mStatus(url)
        data = mBodies(url)
        result.SetBody data
    Else
        result.StatusCode = 404
    End If
    Set IUpdHttp_GetUrl = result
End Function

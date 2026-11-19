VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdHttpResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Result of one HTTP GET. StatusCode is 0 when the server could not be reached.

Public StatusCode As Long
Public ErrorText As String

Private mBody() As Byte

Public Property Get Body() As Byte()
    Body = mBody
End Property

Public Sub SetBody(ByRef data() As Byte)
    mBody = data
End Sub

Public Property Get IsOk() As Boolean
    IsOk = (StatusCode >= 200 And StatusCode < 300)
End Property

' True when the network or server could not be reached at all.
Public Property Get IsOffline() As Boolean
    IsOffline = (StatusCode = 0)
End Property

Public Function Describe() As String
    If StatusCode = 0 Then
        Describe = "no connection (" & ErrorText & ")"
    Else
        Describe = "HTTP " & StatusCode
    End If
End Function

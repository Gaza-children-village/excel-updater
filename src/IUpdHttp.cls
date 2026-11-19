VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUpdHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Boundary to the network. UpdWinHttp is the real implementation;
' UpdFakeHttp is used by the tests.

Public Function GetUrl(ByVal url As String, ByVal timeoutMs As Long) As UpdHttpResult
End Function

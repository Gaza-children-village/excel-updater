VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoaderEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Waits for the first workbook of the Excel session (opened, created, or
' activated - e.g. after "Enable Editing"), then lets GCV Loader open the
' GCV add-ins. Protected View windows are ignored (see StartWhenReady).

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Public Sub Attach()
    Set App = Application
End Sub

Public Sub Detach()
    Set App = Nothing
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    If Wb.IsAddin Then Exit Sub                  ' other add-ins loading at start-up
    GcvLoader.OnFirstWorkbook
End Sub

Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    GcvLoader.OnFirstWorkbook
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    If Wb.IsAddin Then Exit Sub
    GcvLoader.OnFirstWorkbook
End Sub

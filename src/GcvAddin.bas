Attribute VB_Name = "GcvAddin"
Option Explicit

' ============================================================================
' Identity of this GCV add-in (generated per add-in by build_addins.py).
' Every GCV add-in has this module, so GCV Loader and GCV Updater can ask any
' of them the same questions with Application.Run "'<file>'!GcvAddinPing".
' ============================================================================

Public Const ADDIN_ID As String = "gcv/updater"
Public Const ADDIN_NAME As String = "GCV Updater"
Public Const ADDIN_FILE As String = "GCV Updater.xlam"
Public Const ADDIN_VERSION As String = "2.0.0"
Public Const ADDIN_REPO As String = "Gaza-children-village/excel-updater"

Public Const ID_PROPERTY As String = "GcvAddinId"
Public Const VERSION_PROPERTY As String = "GcvAddinVersion"
Public Const FILE_PROPERTY As String = "GcvAddinFile"

' Called right after the add-in is opened. AddinSelfCheck (in the add-in's
' main module) touches every module, so VBA compiles them now: a broken
' release fails here and the previous version is put back. Returns the version.
Public Function GcvAddinPing() As String
    AddinSelfCheck
    GcvAddinPing = ADDIN_VERSION
End Function

' "id|version|repo|name" - used by the updater and the About / report texts.
Public Function GcvAddinInfo() As String
    GcvAddinInfo = ADDIN_ID & "|" & ADDIN_VERSION & "|" & ADDIN_REPO & "|" & ADDIN_NAME
End Function

' Workbook_BeforeSave: records id, version and file name in the .xlam, where
' make_addin_release.py reads them.
Public Sub OnAddinSave()
    SetProperty ID_PROPERTY, ADDIN_ID
    SetProperty VERSION_PROPERTY, ADDIN_VERSION
    SetProperty FILE_PROPERTY, ADDIN_FILE
End Sub

Private Sub SetProperty(ByVal propertyName As String, ByVal value As String)
    On Error Resume Next                   ' raises when the property does not exist yet
    ThisWorkbook.CustomDocumentProperties(propertyName).Delete
    On Error GoTo 0
    ThisWorkbook.CustomDocumentProperties.Add Name:=propertyName, LinkToContent:=False, _
                                              Type:=4, Value:=value     ' 4 = msoPropertyTypeString
End Sub

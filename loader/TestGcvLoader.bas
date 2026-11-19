Attribute VB_Name = "TestGcvLoader"
Option Explicit

' Tests for GCV Loader's choices (no add-in is opened). Run RunLoaderTests from
' the VBA editor (F5) in a DEV copy of the loader. Do NOT ship this module.

Private passedCount As Long
Private failedCount As Long
Private currentTest As String

Public Sub RunLoaderTests()
    passedCount = 0
    failedCount = 0
    Debug.Print String$(60, "-")
    Debug.Print "Loader tests"
    On Error GoTo Crash

    Test_IsAddinFileName_AcceptsOnlyGcvAddins
    Test_OpeningOrder_LoggerFirstUpdaterLastOthersSorted
    Test_OpeningOrder_SkipsMissingFiles
    Test_Lists_AddAndFindWithoutCase

Finish:
    Debug.Print "Passed: " & passedCount & "   Failed: " & failedCount
    MsgBox "Passed: " & passedCount & vbCrLf & "Failed: " & failedCount & vbCrLf & _
           "Details in the Immediate window (Ctrl+G).", _
           IIf(failedCount = 0, vbInformation, vbCritical), "Loader tests"
    Exit Sub

Crash:
    failedCount = failedCount + 1
    Debug.Print "CRASH in " & currentTest & ": " & Err.Number & " " & Err.Description
    Resume Finish
End Sub

Private Sub Test_IsAddinFileName_AcceptsOnlyGcvAddins()
    StartTest "Only 'GCV *.xlam' files other than the loader are add-ins"
    AssertEqual True, IsAddinFileName("GCV Logger.xlam"), "logger"
    AssertEqual True, IsAddinFileName("gcv english names.XLAM"), "case"
    AssertEqual False, IsAddinFileName("GCV Loader.xlam"), "loader"
    AssertEqual False, IsAddinFileName("GCV Logger.xlam.part"), "download in progress"
    AssertEqual False, IsAddinFileName("Other.xlam"), "other add-in"
    AssertEqual False, IsAddinFileName("GCV Administration.xlsx"), "workbook"
End Sub

Private Sub Test_OpeningOrder_LoggerFirstUpdaterLastOthersSorted()
    StartTest "Logger first, updater last, unknown add-ins sorted in between"
    AssertEqual "GCV Logger.xlam,GCV Student Lookup.xlam,GCV English Names.xlam,GCV Attendance.xlam,GCV Zeta.xlam,GCV Updater.xlam", _
                Joined(OpeningOrder(FileList("GCV Updater.xlam", "GCV Zeta.xlam", "GCV English Names.xlam", _
                                             "GCV Attendance.xlam", "GCV Logger.xlam", "GCV Student Lookup.xlam"))), "order"
End Sub

Private Sub Test_OpeningOrder_SkipsMissingFiles()
    StartTest "Only installed files are opened"
    AssertEqual "GCV Logger.xlam,GCV Updater.xlam", Joined(OpeningOrder(FileList("GCV Updater.xlam", "GCV Logger.xlam"))), "two"
    AssertEqual "", Joined(OpeningOrder(New Collection)), "none"
End Sub

Private Sub Test_Lists_AddAndFindWithoutCase()
    StartTest "Lists"
    AssertEqual "1.2.0", AddToList("", "1.2.0"), "first"
    AssertEqual "1.2.0;1.3.0", AddToList("1.2.0", "1.3.0"), "second"
    AssertEqual "1.2.0", AddToList("1.2.0", "1.2.0"), "no duplicate"
    AssertEqual True, InList("GCV Logger.xlam;GCV Updater.xlam", "gcv updater.xlam"), "case"
    AssertEqual False, InList("1.2.0;1.3.0", "1.3"), "whole entries only"
End Sub

Private Function FileList(ParamArray items() As Variant) As Collection
    Dim result As Collection
    Dim i As Long
    Set result = New Collection
    For i = LBound(items) To UBound(items)
        result.Add CStr(items(i))
    Next i
    Set FileList = result
End Function

Private Function Joined(ByVal items As Collection) As String
    Dim item As Variant
    For Each item In items
        If Len(Joined) > 0 Then Joined = Joined & ","
        Joined = Joined & item
    Next item
End Function

Private Sub StartTest(ByVal name As String)
    currentTest = name
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    Dim same As Boolean
    If VarType(expected) = vbBoolean Or VarType(actual) = vbBoolean Then
        same = (CStr(expected) = CStr(actual))
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

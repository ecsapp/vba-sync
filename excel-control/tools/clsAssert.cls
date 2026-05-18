VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Minimal assertion library for excel-control run_tests.
'
' Usage in a Test_* Sub:
'   Public Sub Test_Sales_Total()
'       Assert.AreEqual 100, ComputeTotal(50, 50)
'       Assert.IsTrue ComputeTotal(0, 0) >= 0
'   End Sub
'
' On failure, raises a vbObjectError-based error so the session host's
' Application.Run catches it and emits a `test_result` with status=fail.

Private Const ASSERT_ERR As Long = vbObjectError + 9001

Public Sub AreEqual(ByVal expected As Variant, ByVal actual As Variant, _
                    Optional ByVal message As String = "")
    Dim ok As Boolean
    If IsObject(expected) Or IsObject(actual) Then
        ok = (ObjPtr(expected) = ObjPtr(actual))
    Else
        ok = (expected = actual)
    End If
    If Not ok Then
        Dim msg As String
        msg = "AreEqual failed: expected=" & CStr(expected) & " actual=" & CStr(actual)
        If Len(message) > 0 Then msg = msg & " (" & message & ")"
        Err.Raise ASSERT_ERR, "Assert.AreEqual", msg
    End If
End Sub

Public Sub IsTrue(ByVal expr As Boolean, Optional ByVal message As String = "")
    If Not expr Then
        Dim msg As String
        msg = "IsTrue failed"
        If Len(message) > 0 Then msg = msg & ": " & message
        Err.Raise ASSERT_ERR, "Assert.IsTrue", msg
    End If
End Sub

Public Sub IsFalse(ByVal expr As Boolean, Optional ByVal message As String = "")
    If expr Then
        Dim msg As String
        msg = "IsFalse failed"
        If Len(message) > 0 Then msg = msg & ": " & message
        Err.Raise ASSERT_ERR, "Assert.IsFalse", msg
    End If
End Sub

Public Sub IsNothing(ByVal obj As Object, Optional ByVal message As String = "")
    If Not obj Is Nothing Then
        Dim msg As String
        msg = "IsNothing failed: object is not Nothing"
        If Len(message) > 0 Then msg = msg & " (" & message & ")"
        Err.Raise ASSERT_ERR, "Assert.IsNothing", msg
    End If
End Sub

Public Sub Fail(Optional ByVal message As String = "")
    Dim msg As String
    msg = "Fail()"
    If Len(message) > 0 Then msg = msg & ": " & message
    Err.Raise ASSERT_ERR, "Assert.Fail", msg
End Sub

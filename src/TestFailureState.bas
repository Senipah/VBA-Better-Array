Attribute VB_Name = "TestFailureState"
Option Explicit

Private pHasFailed As Boolean
Private pFailureMessage As String

Public Sub ResetTestFailureState()
    pHasFailed = False
    pFailureMessage = vbNullString
End Sub

Public Sub RecordTestFailure(ByVal Message As String)
    If Not pHasFailed Then
        pHasFailed = True
        pFailureMessage = Message
    End If
End Sub

Public Function TestFailed() As Boolean
    TestFailed = pHasFailed
End Function

Public Function TestFailureMessage() As String
    TestFailureMessage = pFailureMessage
End Function

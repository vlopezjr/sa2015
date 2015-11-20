Attribute VB_Name = "StringBuilder"
Option Explicit

Private myString As String

Public Property Get ToString() As String
    ToString = myString
End Property

Public Sub Add(s As String)
    myString = myString & s
End Sub

Public Sub AddLine(Optional s As Variant)
    If IsMissing(s) Then
        myString = myString & vbCrLf
    Else
        myString = myString & s & vbCrLf
    End If
End Sub

Public Sub Clear()
    myString = ""
End Sub

Public Sub PrintString()
    Debug.Print myString
End Sub


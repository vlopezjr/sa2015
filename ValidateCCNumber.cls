VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidateCCNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const k_sClassName = "ValidateCCNumber"

'=======================================================
' Standard stuff you would inherit from IValidationRule
' if VB 6.0 supported inheritance

Implements IValidationRule

Private m_oBaseRule As ValidationRule


Private Sub Class_Initialize()
    Debug.Print "Class_Initialize: " & k_sClassName
    Set m_oBaseRule = New ValidationRule
End Sub


Private Sub Class_Terminate()
    Debug.Print "Class_Terminate: " & k_sClassName & " " & m_oBaseRule.ID
    Set m_oBaseRule = Nothing
End Sub


Public Property Get IValidationRule_CtlWrapper() As ControlWrapper
    Set IValidationRule_CtlWrapper = m_oBaseRule.CtlWrapper
End Property

Public Property Let IValidationRule_CtlWrapper(ByVal RHS As ControlWrapper)
    m_oBaseRule.CtlWrapper = RHS
End Property


Private Property Get IValidationRule_Enabled() As Boolean
    IValidationRule_Enabled = m_oBaseRule.Enabled
End Property

Private Property Let IValidationRule_Enabled(ByVal RHS As Boolean)
    m_oBaseRule.Enabled = RHS
End Property


Public Property Get IValidationRule_ClassID() As Long
    IValidationRule_ClassID = m_oBaseRule.ClassID
End Property

Public Property Let IValidationRule_ClassID(ByVal RHS As Long)
    m_oBaseRule.ClassID = RHS
End Property


Public Property Get IValidationRule_CustomMsg() As String
    IValidationRule_CustomMsg = m_oBaseRule.CustomMsg
End Property

Public Property Let IValidationRule_CustomMsg(ByVal RHS As String)
    m_oBaseRule.CustomMsg = RHS
End Property


Public Property Get IValidationRule_ID() As Long
    IValidationRule_ID = m_oBaseRule.ID
End Property


Private Function IValidationRule_Validate() As String
    If m_oBaseRule.Enabled Then
        IValidationRule_Validate = m_oBaseRule.Validate(Validate)
    End If
End Function


' End of standard IValidationRule base class logic
'=======================================================


' This function is assummed by IValidationRule_Validate but its
' implementation will vary with each class that implements the
' IValidationRule interface.

Private Function Validate() As String
    If (Not IsCCNoValid(m_oBaseRule.CtlWrapper.Text)) Then
        Validate = "Invalid credit card number. Please confirm and reenter."
    End If
End Function


Private Function IsCCNoValid(sCardNo As String) As Boolean
    Const MAX_DIGITS = 16
    Dim digits(1 To MAX_DIGITS) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim c As String
    Dim multiplier As Integer
    Dim checksum As Integer
    Dim Value As Integer
    
    If Len(Trim$(sCardNo)) < 1 Then Exit Function
    
    'strip out non-numeric charaters
    For i = 1 To Len(sCardNo)
      c = Mid$(sCardNo, i, 1)
      If IsNumeric(c) Then
         j = j + 1
         If j > MAX_DIGITS Then Exit Function
         digits(j) = Val(c)
      End If
    Next i
    
    checksum = digits(j)

    For i = j - 1 To 1 Step -1
        If multiplier = 2 Then
            multiplier = 1
        Else
            multiplier = 2
        End If
        Value = digits(i) * multiplier
        checksum = checksum + Value
        If Value > 9 Then checksum = checksum - 9
    Next i

    IsCCNoValid = ((checksum Mod 10) = 0)
End Function




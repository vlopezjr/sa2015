VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidateFutureDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const k_sClassName = "ValidateFutureDate"
Private m_vMinValue As Variant

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
    On Error GoTo ErrorHandler

    Dim dValue As String

    With m_oBaseRule.CtlWrapper
        'to see if a valid date is entered
        If CInt(Mid$(.Text, 1, 2)) > 12 Then
            GoTo ErrorHandler
        Else
            dValue = .Text
        End If
    End With

    '***465 SMR 04-20-2006 - changed code to use same logic from matrix in CCEditor Init.
        'dValue = 052004; m_vMinValue = now
        '(pass in m_vMinValue (now) value, we may not always compare to the current date.)
'    If VarType(m_vMinValue) <> vbEmpty Then
'        If CInt(Mid$(dValue, 3, 4)) <= CInt(Mid$(m_vMinValue, 3, 4)) Then
'            If CInt(Mid$(dValue, 1, 2)) < CInt(Mid$(m_vMinValue, 1, 2)) Then
'                GoTo ErrorHandler
'            End If
'        End If
'    End If
    If VarType(m_vMinValue) <> vbEmpty Then
        If CInt(Mid$(dValue, 3, 4)) < CInt(Mid$(m_vMinValue, 3, 4)) Then
            GoTo ErrorHandler
        ElseIf CInt(Mid$(dValue, 3, 4)) = CInt(Mid$(m_vMinValue, 3, 4)) Then
            If CInt(Mid$(dValue, 1, 2)) <= CInt(Mid$(m_vMinValue, 1, 2)) Then
                GoTo ErrorHandler
            End If
        End If
    End If

    Exit Function

ErrorHandler:
    Validate = "Expiration Date must be a valid date in the future."
    'Validate = m_oBaseRule.CtlWrapper.Caption & " must be a valid date after " & m_vMinValue & "."
End Function


'=======================================================
' Public methods defined for this particular subclass

Public Property Get MinValue() As Variant
    MinValue = m_vMinValue
End Property

Public Property Let MinValue(ByVal RHS As Variant)
    m_vMinValue = RHS
End Property



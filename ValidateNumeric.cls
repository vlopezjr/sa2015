VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidateNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const k_sClassName = "ValidateNumeric"

Private m_vMinValue As Variant
Private m_vMaxValue As Variant
Private m_bForceInteger As Boolean

'=======================================================
' Standard stuff you would inherit from IValidationRule
' if VB 6.0 supported inheritance

Implements IValidationRule

Private m_oBaseRule As ValidationRule


Private Sub Class_Initialize()
'    Debug.Print "Class_Initialize: " & k_sClassName
    Set m_oBaseRule = New ValidationRule
End Sub


Private Sub Class_Terminate()
'    Debug.Print "Class_Terminate: " & k_sClassName & " " & m_oBaseRule.ID
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

    Dim dValue As Double

    With m_oBaseRule.CtlWrapper
        If m_bForceInteger Then
            .Text = Fix(.Text)
        End If
        dValue = CDbl(.Text)
    End With

    If VarType(m_vMinValue) <> vbEmpty Then
        If dValue < m_vMinValue Then
            GoTo ErrorHandler
        End If
    End If

    If VarType(m_vMaxValue) <> vbEmpty Then
        If dValue > m_vMaxValue Then
            GoTo ErrorHandler
        End If
    End If
    Exit Function

ErrorHandler:
    Validate = m_oBaseRule.CtlWrapper.Caption & " must be"
    If m_bForceInteger Then
        Validate = Validate + " an integer"
    Else
        Validate = Validate + " a number"
    End If

    If VarType(m_vMinValue) = vbEmpty Then
        If VarType(m_vMaxValue) = vbEmpty Then
            'nothing to add to errmsg since no min or max constraints
        Else
            Validate = Validate + " not greater than " & m_vMaxValue
        End If
    Else
        If VarType(m_vMaxValue) = vbEmpty Then
            Validate = Validate + " not less than " & m_vMinValue
        Else
            Validate = Validate + " between " & m_vMinValue & " and " & m_vMaxValue
        End If
    End If
    Validate = Validate + "."
End Function


'=======================================================
' Public methods defined for this particular subclass

Public Property Get MinValue() As Variant
    MinValue = m_vMinValue
End Property

Public Property Let MinValue(ByVal RHS As Variant)
    m_vMinValue = RHS
    If VarType(RHS) = vbInteger Or VarType(RHS) = vbLong Then
        ForceInteger = (RHS = Fix(RHS))
    End If
End Property


Public Property Get MaxValue() As Variant
    MaxValue = m_vMaxValue
End Property

Public Property Let MaxValue(ByVal RHS As Variant)
    m_vMaxValue = RHS
    If VarType(RHS) = vbInteger Or VarType(RHS) = vbLong Then
        ForceInteger = (RHS = Fix(RHS))
    End If
End Property


Public Property Get ForceInteger() As Boolean
    ForceInteger = m_bForceInteger
End Property

Public Property Let ForceInteger(ByVal RHS As Boolean)
    m_bForceInteger = RHS
End Property


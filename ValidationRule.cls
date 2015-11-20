VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_oCtlWrapper As ControlWrapper
Private m_bEnabled As Boolean
Private m_lClassID As Long
Private m_sCustomMsg As String
Private m_lID As Long


Private Sub Class_Initialize()
    m_lID = NextValidationRuleID
'    Debug.Print "Class_Initialize: ValidationRule " & m_lID
End Sub


Private Sub Class_Terminate()
'    Debug.Print "Class_Terminate: ValidationRule " & m_lID
    Set m_oCtlWrapper = Nothing
End Sub


Public Property Get CtlWrapper() As ControlWrapper
    Set CtlWrapper = m_oCtlWrapper
End Property

Public Property Let CtlWrapper(ByVal RHS As ControlWrapper)
    Set m_oCtlWrapper = RHS
End Property


Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal RHS As Boolean)
    m_bEnabled = RHS
End Property


Public Property Get ClassID() As Long
    ClassID = m_lClassID
End Property

Public Property Let ClassID(ByVal RHS As Long)
    m_lClassID = RHS
End Property


Public Property Get CustomMsg() As String
    CustomMsg = m_sCustomMsg
End Property

Public Property Let CustomMsg(ByVal RHS As String)
    m_sCustomMsg = RHS
End Property


Public Property Get ID() As Long
    ID = m_lID
End Property


Public Function Validate(ByRef i_sErrorMsg As String) As String
    If m_bEnabled And Len(i_sErrorMsg) > 0 Then
        If Len(m_sCustomMsg) > 0 Then
            Validate = m_sCustomMsg
        Else
            Validate = i_sErrorMsg
        End If
    End If
End Function


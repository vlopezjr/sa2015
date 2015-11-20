VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidateRegEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValidationRule
Private sRegularExpression As String


Private m_oBaseRule As ValidationRule


Private Sub Class_Initialize()
    Set m_oBaseRule = New ValidationRule
End Sub


Private Sub Class_Terminate()
    Set m_oBaseRule = Nothing
End Sub

Public Property Get RegularExpression() As String
    RegularExpression = sRegularExpression
End Property

Public Property Let RegularExpression(i_sRegularExpression As String)
    sRegularExpression = i_sRegularExpression
End Property

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

Private Function Validate() As String
    Dim myRegExp As RegExp
        
    Set myRegExp = New RegExp
    
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    myRegExp.Pattern = sRegularExpression
        
    If (myRegExp.Test(m_oBaseRule.CtlWrapper.Text) = False) Then
        Validate = m_oBaseRule.CtlWrapper.Caption
    End If
    
'    Dim myRegExp As RegExp
'    Dim myMatches As MatchCollection
'    Dim myMatch As Match
'
'    Set myRegExp = New RegExp
'
'    myRegExp.IgnoreCase = True
'    myRegExp.Global = True
'    myRegExp.Pattern = sRegularExpression
'
'    Set myMatches = myRegExp.Execute(m_oBaseRule.CtlWrapper.Text)
'
'    If myMatches.Count = 0 Then
'        Validate = m_oBaseRule.CtlWrapper.Caption
'    End If
End Function

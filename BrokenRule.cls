VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BrokenRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_oRule As IValidationRule
Private m_sErrorMsg As String


Private Sub Class_Terminate()
    Set m_oRule = Nothing
End Sub


Public Property Get Rule() As IValidationRule
    Set Rule = m_oRule
End Property

Public Property Let Rule(RHS As IValidationRule)
    Set m_oRule = RHS
End Property


Public Property Get ErrorMsg() As String
    ErrorMsg = m_sErrorMsg
End Property

Public Property Let ErrorMsg(ByVal RHS As String)
    m_sErrorMsg = RHS
End Property


Public Sub Initialize(ByVal i_oRule As IValidationRule, ByVal i_sErrorMsg As String)
    Rule = i_oRule
    ErrorMsg = i_sErrorMsg
End Sub


Public Sub SetFocus()
    Dim oCtl As Control
    Dim oContainer As Object
    
    Set oCtl = m_oRule.CtlWrapper.Ctl
    Set oContainer = oCtl.Container

    While TypeOf oContainer Is Frame
        Set oContainer = oContainer.Container
    Wend
    
    'For controls within tab panels, ensure panel is selected first
    If TypeOf oContainer Is SSActiveTabPanel Then
        With oContainer.Container.TabFromControl(oCtl)
            If .Visible Then
                .Selected = True
            Else
                Exit Sub
            End If
        End With
    End If

    oCtl.SetFocus
End Sub

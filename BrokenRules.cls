VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BrokenRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colBrokenRules As Collection
Private m_frm As Form
Private m_colControls As Collection
Private m_lAddCount As Long
Private m_lRemoveCount As Long
Private m_bInValidate As Boolean

Public Event AddBrokenRule(ByVal i_oBrokenRule As BrokenRule)
Public Event RemoveBrokenRule(ByVal i_oBrokenRule As BrokenRule)
Public Event BeforeValidateUpdates()
Public Event AfterValidateUpdates()
Public Event HaveBrokenRules()
Public Event NoBrokenRules()
Public Event ForceFocus(ByVal i_oCtl As Control, ByRef i_sErrorMsg As String, ByRef o_bCancel As Boolean)


Public Sub Destroy()
    Dim oCtrlWrapper As ControlWrapper
    Dim oBrokenRule As BrokenRule
    
    Set m_frm = Nothing
    
    For Each oCtrlWrapper In m_colControls
        oCtrlWrapper.BrokenRules = Nothing
'*** added 11/11/05 LR
        oCtrlWrapper.Ctl = Nothing
    Next
    Set m_colControls = Nothing

    For Each oBrokenRule In m_colBrokenRules
        oBrokenRule.Rule = Nothing
    Next
    Set m_colBrokenRules = Nothing
End Sub


Private Sub Class_Initialize()
    Debug.Print "Class_Initialize: BrokenRules"
    Set m_colBrokenRules = New Collection
    Set m_colControls = New Collection
End Sub


Private Sub Class_Terminate()
    Debug.Print "Class_Terminate: BrokenRules"
    Set m_colBrokenRules = Nothing
    Set m_colControls = Nothing
End Sub


Public Property Get Form() As Form
    Set Form = m_frm
End Property

Public Property Let Form(ByRef RHS As Form)
    Set m_frm = RHS
End Property


Public Function AddControl( _
        ByVal i_oCtl As Control, _
        ByVal i_sCaption As String, _
        Optional ByVal i_bRequired As Boolean = False, _
        Optional ByVal i_bTrim As Boolean = True _
)
    Dim oCtlWrapper As ControlWrapper
    
    Set oCtlWrapper = New ControlWrapper
    oCtlWrapper.Initialize Me, i_oCtl, i_sCaption, i_bRequired, i_bTrim
    m_colControls.Add oCtlWrapper
    Set AddControl = oCtlWrapper
End Function


Public Sub RemoveControl()
    
End Sub


Public Sub AddBrokenRule(ByVal i_oRule As IValidationRule, ByVal i_sErrorMsg As String)
    Dim i As Long
    Dim lTabIndex As Long

    lTabIndex = i_oRule.CtlWrapper.Ctl.TabIndex
    For i = 1 To m_colBrokenRules.Count
        Select Case m_colBrokenRules(i).Rule.CtlWrapper.Ctl.TabIndex
        Case Is > lTabIndex 'insert before later rule in tab order
            ReallyAddRule i_oRule, i_sErrorMsg, i
            Exit Sub

        Case Is = lTabIndex 'update if necessary
            With m_colBrokenRules(i)
                'only update if something is different
                If Not (.Rule Is i_oRule And .ErrorMsg = i_sErrorMsg) Then
                    RemoveBrokenRule .Rule
                    ReallyAddRule i_oRule, i_sErrorMsg, i
                End If
                Exit Sub
            End With
        End Select
    Next

    'if we get here, add rule to the end of the collection
    ReallyAddRule i_oRule, i_sErrorMsg
End Sub


Private Sub ReallyAddRule( _
        ByVal i_oRule As IValidationRule, _
        ByVal i_sErrorMsg As String, _
        Optional ByVal i_lInsertBefore As Long = 0 _
)
    Dim oBrokenRule As BrokenRule

    'If this is the first change during Validate, raise event
    If m_bInValidate And m_lAddCount + m_lRemoveCount = 0 Then
        RaiseEvent BeforeValidateUpdates
    End If
    
    Set oBrokenRule = New BrokenRule
    oBrokenRule.Initialize i_oRule, i_sErrorMsg

    If i_lInsertBefore > 0 And i_lInsertBefore <= m_colBrokenRules.Count Then
        m_colBrokenRules.Add oBrokenRule, CStr(i_oRule.ID), i_lInsertBefore
    Else
        m_colBrokenRules.Add oBrokenRule, CStr(i_oRule.ID)
    End If

    RaiseEvent AddBrokenRule(oBrokenRule)
    
    oBrokenRule.Rule.CtlWrapper.AssignErrorMsg i_sErrorMsg

    m_lAddCount = m_lAddCount + 1
End Sub


Public Sub RemoveBrokenRule(ByVal i_oRule As IValidationRule)

    'We will get an error if the rule does not exist
    On Error GoTo ErrorHandler
    Dim oBrokenRule As BrokenRule
    Set oBrokenRule = m_colBrokenRules(CStr(i_oRule.ID))

    'If this is the first change during Validate, raise event
    If m_bInValidate And m_lAddCount + m_lRemoveCount = 0 Then
        RaiseEvent BeforeValidateUpdates
    End If

    'Finally, we can remove the broken rule
    RaiseEvent RemoveBrokenRule(oBrokenRule)
    
    oBrokenRule.Rule.CtlWrapper.ClearErrorMsg
    
    m_colBrokenRules.Remove CStr(i_oRule.ID)
    m_lRemoveCount = m_lRemoveCount + 1
    Exit Sub

ErrorHandler:
    'Just ignore this error
End Sub


Public Sub EnableClass( _
        ByVal i_lClassID As Long, _
        ByVal i_bEnabled As Boolean _
)
    Dim oCtlWrapper As ControlWrapper

    For Each oCtlWrapper In m_colControls
        oCtlWrapper.EnableClass i_lClassID, i_bEnabled
    Next
End Sub


Public Sub LockClass( _
        ByVal i_lClassID As Long, _
        ByVal i_bLocked As Boolean _
)
    Dim oCtlWrapper As ControlWrapper

    For Each oCtlWrapper In m_colControls
        oCtlWrapper.LockClass i_lClassID, i_bLocked
    Next
End Sub


Public Sub Validate(Optional ByRef i_oCtl As Control = Nothing)
    Dim oCtlWrapper As ControlWrapper
    Dim lPrevCount As Long

    lPrevCount = Count
    m_lAddCount = 0
    m_lRemoveCount = 0
    m_bInValidate = True

    For Each oCtlWrapper In m_colControls
        If i_oCtl Is Nothing Or i_oCtl Is oCtlWrapper.Ctl Then
            oCtlWrapper.Validate
        End If
    Next

    If m_lAddCount + m_lRemoveCount > 0 Then
        RaiseEvent AfterValidateUpdates
    End If

    m_bInValidate = False

    'Deal with HaveBrokenRules/NoBrokenRules events
    If lPrevCount = 0 Then
        If Count > 0 Then
            RaiseEvent HaveBrokenRules
        End If
    Else
        If Count = 0 Then
            RaiseEvent NoBrokenRules
        End If
    End If
End Sub


Public Function IsValid(ByRef i_oCtl As Control) As Boolean
    Dim oRule As BrokenRule

    For Each oRule In m_colBrokenRules
        If oRule.Rule.CtlWrapper.Ctl Is i_oCtl Then
            Exit Function
        End If
    Next

    IsValid = True
End Function


Public Function FindBrokenRuleIndex(ByRef i_oCtl As Control) As Long
    Dim oRule As BrokenRule
    Dim i As Long
    
    For Each oRule In m_colBrokenRules
        i = i + 1
        If oRule.Rule.CtlWrapper.Ctl Is i_oCtl Then
            FindBrokenRuleIndex = i
            Exit Function
        End If
    Next
End Function


Public Function ErrorListText() As String
    Dim oBrokenRule As BrokenRule
    Dim sText As String
    
    For Each oBrokenRule In m_colBrokenRules
        sText = sText & oBrokenRule.ErrorMsg & vbCrLf
    Next

    ErrorListText = sText
End Function

Public Sub ForceFocus(ByVal i_oCtl As Control, ByRef i_sErrorMsg As String)
    Dim bCancel As Boolean
    Dim oContainer As Object

    RaiseEvent ForceFocus(i_oCtl, i_sErrorMsg, bCancel)
    If Not bCancel Then
        Set oContainer = i_oCtl.Container
    
        While TypeOf oContainer Is Frame
            Set oContainer = oContainer.Container
        Wend
        
        'For controls within tab panels, ensure panel is selected first
        If TypeOf oContainer Is SSActiveTabPanel Then
            With oContainer.Container.TabFromControl(i_oCtl)
                If .Visible Then
                    .Selected = True
                Else
                    Exit Sub
                End If
            End With
        End If
    
        TryToSetFocus i_oCtl
    End If
End Sub


Public Sub SetFocusNext()
    Dim lTabIndex As Long
    Dim i As Long
    Dim o As BrokenRule
    
    lTabIndex = m_frm.ActiveControl.TabIndex
    For i = 1 To m_colBrokenRules.Count
        If lTabIndex < m_colBrokenRules(i).Rule.CtlWrapper.Ctl.TabIndex Then
            m_colBrokenRules(i).SetFocus
            Exit Sub
        End If
    Next

    If m_colBrokenRules.Count > 0 Then
        m_colBrokenRules(1).SetFocus
    End If
End Sub


Public Property Get Count() As Long
    Count = m_colBrokenRules.Count
End Property


'This function returns the number of broken rules only counting rules that
'have class IDs including the same bits as passed in the i_lClassMask parameter
Public Property Get MaskedCount(ByRef i_lClassMask As Long) As Long
    Dim oRule As BrokenRule
    Dim lBrokenRuleCount As Long
    
    For Each oRule In m_colBrokenRules
        With oRule.Rule
            If (.ClassID And i_lClassMask) = i_lClassMask Then
                If .Enabled Then
                    If Len(.Validate) > 0 Then
                        lBrokenRuleCount = lBrokenRuleCount + 1
                    End If
                End If
            End If
        End With
    Next

    MaskedCount = lBrokenRuleCount
End Property


Public Function Item(vIndex As Variant) As BrokenRule
Attribute Item.VB_UserMemId = 0
    Set Item = m_colBrokenRules.Item(vIndex)
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colBrokenRules.[_NewEnum]
End Function


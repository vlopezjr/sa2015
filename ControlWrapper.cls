VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_oBrokenRules As BrokenRules
Private m_oCtl As Control
Private m_sCaption As String
Private m_bRequired As Boolean
Private m_bTrimText As Boolean
Private m_oRules As Collection
Private m_sToolTipText As String


Private Sub Class_Initialize()
'    Debug.Print "Class_Initialize: ControlWrapper"
    Set m_oRules = New Collection
End Sub


Private Sub Class_Terminate()
'    Debug.Print "Class_Terminate: ControlWrapper"
    Set m_oRules = Nothing
    Set m_oCtl = Nothing
End Sub


Public Property Get BrokenRules() As BrokenRules
    Set BrokenRules = m_oBrokenRules
End Property

Public Property Let BrokenRules(ByVal RHS As BrokenRules)
    Set m_oBrokenRules = RHS
End Property


Public Property Get Ctl() As Control
    Set Ctl = m_oCtl
End Property

Public Property Let Ctl(ByVal RHS As Control)
    Set m_oCtl = RHS
End Property


Public Property Get Caption() As String
    Caption = m_sCaption
End Property

Public Property Let Caption(ByVal RHS As String)
    m_sCaption = RHS
End Property


Public Property Let ToolTipText(ByVal RHS As String)
    m_sToolTipText = RHS
End Property


Public Property Get Text() As String
    If TypeOf m_oCtl Is CheckBox Then
        Select Case m_oCtl.Value
        Case vbChecked
            Text = "True"
        Case vbUnchecked
            Text = "False"
        Case Else
            Text = ""
        End Select
    Else
        If m_bTrimText Then
            Text = Trim(m_oCtl.Text)
        Else
            Text = m_oCtl.Text
        End If
    End If
End Property

Public Property Let Text(ByVal RHS As String)
    On Error Resume Next 'ignore not being able to update read-only controls
    
    If TypeOf m_oCtl Is CheckBox Then
        Select Case RHS
        Case "True", "Yes"
            m_oCtl.Value = vbChecked
        Case Else
            m_oCtl.Value = vbUnchecked
        End Select
    Else
        If m_bTrimText Then
            RHS = Trim(RHS)
        End If
    
        If m_oCtl.Text <> RHS Then
            m_oCtl.Text = RHS
        End If
    End If
End Property


Public Property Get Required() As Boolean
    Required = m_bRequired
End Property

Public Property Let Required(ByVal RHS As Boolean)
    m_bRequired = RHS
End Property


Public Property Get TrimText() As Boolean
    TrimText = m_bTrimText
End Property

Public Property Let TrimText(ByVal RHS As Boolean)
    m_bTrimText = RHS
End Property


Public Sub AssignErrorMsg(ByRef i_sErrorMsg As String)
    If Len(m_sToolTipText) = 0 Then
        m_sToolTipText = m_oCtl.ToolTipText
    End If
    m_oCtl.ToolTipText = i_sErrorMsg

    m_oCtl.BackColor = k_lCtlMarkColor
End Sub


Public Sub ClearErrorMsg()
    m_oCtl.ToolTipText = m_sToolTipText
    m_sToolTipText = ""

    m_oCtl.BackColor = k_lCtlBackColor
End Sub


Public Function AddRuleManual( _
        Optional ByVal i_bValid As Boolean = True, _
        Optional ByVal i_sErrorMsg As String = "", _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
) As ValidateManual
    Dim oRule As ValidateManual
    
    Set oRule = New ValidateManual
    With oRule
        .Valid = i_bValid
        .ErrorMsg = i_sErrorMsg
    End With
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
    Set AddRuleManual = oRule
End Function


Public Sub AddRuleNumeric( _
        Optional ByVal i_vMinValue As Variant = Empty, _
        Optional ByVal i_vMaxValue As Variant = Empty, _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateNumeric
    
    Set oRule = New ValidateNumeric
    With oRule
        .MinValue = i_vMinValue
        .MaxValue = i_vMaxValue
    End With
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


Public Sub AddRuleFutureDate( _
        Optional ByVal i_vMinValue As Variant = Empty, _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateFutureDate
    
    Set oRule = New ValidateFutureDate
    With oRule
        .MinValue = i_vMinValue
    End With
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


Public Sub AddRuleLength( _
        Optional ByVal i_lMinLength As Long = 0, _
        Optional ByVal i_lMaxLength As Long = 0, _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateLength
    
    Set oRule = New ValidateLength
    With oRule
        .MinLength = i_lMinLength
        .MaxLength = i_lMaxLength
    End With
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


Public Sub AddRuleRequired( _
        Optional ByVal i_sDefaultValue As String = "", _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateRequired
    
    Set oRule = New ValidateRequired
    oRule.DefaultValue = i_sDefaultValue
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


'4/28/08 LR & DH edited this

Public Sub AddRuleDuplicateCC( _
        Optional ByVal i_colCC As Collection, _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateDuplicateCC
    Set oRule = New ValidateDuplicateCC

    oRule.colCC = i_colCC

    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


'4/28/08 LR & DH added this for 502

Public Sub AddRuleCCNumber( _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateCCNumber
    Set oRule = New ValidateCCNumber
    
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


Public Function AddRuleZipCode( _
        Optional ByVal i_sPostalCodeMask = "", _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
) As ValidateZipCode
    Dim oRule As ValidateZipCode
    
    Set oRule = New ValidateZipCode
    With oRule
        .PostalCodeMask = i_sPostalCodeMask
    End With
    
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
    Set AddRuleZipCode = oRule
End Function


'SMR Intl - 01/17/2006 - for FThisOrderOnlyAddress StateID combo box validation
Public Function AddRuleStateID( _
        Optional ByVal i_sCountryID = "", _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
) As ValidateStateID
    Dim oRule As ValidateStateID
    
    Set oRule = New ValidateStateID
    With oRule
        .CountryID = i_sCountryID
    End With
    
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
    Set AddRuleStateID = oRule
End Function


Public Sub AddRuleCustomer( _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateCustomer
    
    Set oRule = New ValidateCustomer
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


Public Sub AddRuleAmExCustType( _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateAmExCustType
    
    Set oRule = New ValidateAmExCustType
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


Public Sub AddRuleInchWorm( _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateInchWorm
    
    Set oRule = New ValidateInchWorm
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


'added 10/11/05 LR

Public Sub AddRuleAlpha( _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateAlpha
    
    Set oRule = New ValidateAlpha
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


Public Sub AddRuleEMail( _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateEMail
    
    Set oRule = New ValidateEMail
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


'added 8/17/05 LR

Public Sub AddRuleAMDelivery( _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateAMDelivery
    
    Set oRule = New ValidateAMDelivery
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
End Sub


'***DH 2/20/09 Added
Public Sub AddRuleRegEx( _
        Optional ByVal i_sRegularExpression As String = "", _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    Dim oRule As ValidateRegEx
    
    Set oRule = New ValidateRegEx
    
    oRule.RegularExpression = i_sRegularExpression
    AddRule oRule, i_lClassID, i_bEnabled, i_sCustomMsg
    
End Sub


Public Sub Initialize( _
        ByVal i_oBrokenRules As BrokenRules, _
        ByVal i_oCtl As Control, _
        ByVal i_sCaption As String, _
        Optional ByVal i_bRequired As Boolean = False, _
        Optional ByVal i_bTrimText As Boolean = True _
)
    BrokenRules = i_oBrokenRules
    Ctl = i_oCtl
    Caption = i_sCaption
    Required = i_bRequired
    TrimText = i_bTrimText
End Sub


Public Sub EnableClass( _
        ByVal i_lClassID As Long, _
        ByVal i_bEnabled As Boolean _
)
    Dim oRule As IValidationRule

    For Each oRule In m_oRules
        With oRule
            If .ClassID = i_lClassID Then
                .Enabled = i_bEnabled
            End If
        End With
    Next
End Sub


Public Sub LockClass( _
        ByVal i_lClassID As Long, _
        ByVal i_bLocked As Boolean _
)
    Dim oRule As IValidationRule

    For Each oRule In m_oRules
        With oRule
            If .ClassID = i_lClassID Then
                LockControl i_bLocked
            End If
        End With
    Next
End Sub


Public Function LockControl(ByVal i_bLocked As Boolean)
    With Ctl
        If TypeOf Ctl Is CheckBox Then
            If i_bLocked Then
                .Enabled = False
'                .ForeColor = k_lCtlLockColor
'                .BackColor = k_lCtlBackColor
            Else
                .Enabled = True
'                .ForeColor = k_lCtlForeColor
'                .BackColor = k_lCtlWhiteColor
            End If
        Else
            If i_bLocked Then
                .Locked = True
                .ForeColor = k_lCtlLockColor
            Else
                .Locked = False
                .ForeColor = k_lCtlForeColor
            End If
        End If
    End With
End Function


Public Sub Validate()
    Dim oRule As IValidationRule
    Dim sMsg As String
    Dim bBrokenRuleFound As Boolean

    Me.Text = Me.Text 'enforce trim, if required

    For Each oRule In m_oRules
        'don't enforce length restrictions on empty, non-required field
        If Len(Me.Text) = 0 And Not Me.Required Then
            sMsg = ""
        Else
            sMsg = oRule.Validate
        End If
        
        If Len(sMsg) > 0 And Not bBrokenRuleFound Then
            bBrokenRuleFound = True
            m_oBrokenRules.AddBrokenRule oRule, sMsg
            
            'At one time we thought it would be cool to support
            'field validation that would force you to enter a legal
            'value before advancing to the next field.  But feedback
            'from users was negative and it also caused some problems
            'such as strange tab order when doing a full validation on
            'a form with multiple invalid fields.
            'For these reasons, the code below is commented out.
            
            'If m_bRequired And Len(Me.Text) > 0 _
            'And Not (TypeOf Me.Ctl Is ComboBox And Me.Text = "<none>") Then
            '    m_oBrokenRules.ForceFocus Ctl, sMsg
            'End If
        Else
            m_oBrokenRules.RemoveBrokenRule oRule
        End If
    Next
End Sub


'We need to call ClearRules before releasing a ControlWrapper because
'otherwise, the CtlWrapper references back to this object from the
'underlying ValidationRule object would prevent the Class_Termination
'event from firing, consequently causing a memory leak.

Public Sub ClearRules()
    Dim oRule As IValidationRule
    Dim i As Long
        
    For i = m_oRules.Count To 1 Step -1
        Set oRule = m_oRules(i)
        oRule.CtlWrapper = Nothing
        m_oRules.Remove i
    Next
End Sub


Private Sub AddRule( _
        ByRef i_oRule As IValidationRule, _
        Optional ByVal i_lClassID As Long = 0, _
        Optional ByVal i_bEnabled As Boolean = False, _
        Optional ByVal i_sCustomMsg As String = "" _
)
    With i_oRule
        .CtlWrapper = Me
        .Enabled = i_bEnabled
        .ClassID = i_lClassID
        .CustomMsg = i_sCustomMsg
    End With

    m_oRules.Add i_oRule
End Sub





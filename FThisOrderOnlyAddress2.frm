VERSION 5.00
Begin VB.Form FThisOrderOnlyAddress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3420
      Width           =   4635
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   60
      TabIndex        =   13
      Top             =   0
      Width           =   4515
      Begin VB.CommandButton cmdValidateWithUPS 
         Caption         =   "&Auto-Complete"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   435
         Left            =   2820
         TabIndex        =   8
         Top             =   2160
         Width           =   1515
      End
      Begin VB.ComboBox cboCountry 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox cboState 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtExtZip 
         Height          =   315
         Left            =   3765
         TabIndex        =   6
         Top             =   1680
         Width           =   555
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox txtAddr1 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   3000
      End
      Begin VB.TextBox txtAddr2 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   3000
      End
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   3000
      End
      Begin VB.TextBox txtZip 
         Height          =   315
         Left            =   2820
         TabIndex        =   5
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Country"
         Height          =   195
         Left            =   465
         TabIndex        =   20
         Top             =   2100
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "State"
         Height          =   255
         Left            =   345
         TabIndex        =   19
         Top             =   1740
         Width           =   795
      End
      Begin VB.Line Line1 
         X1              =   3660
         X2              =   3720
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Company"
         Height          =   255
         Left            =   300
         TabIndex        =   18
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Address 1"
         Height          =   255
         Left            =   300
         TabIndex        =   17
         Top             =   660
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Address 2"
         Height          =   255
         Left            =   300
         TabIndex        =   16
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "City"
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   1380
         Width           =   840
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Zip"
         Height          =   195
         Left            =   2400
         TabIndex        =   14
         Top             =   1740
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3540
      TabIndex        =   12
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2400
      TabIndex        =   11
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Height          =   435
      Left            =   1200
      TabIndex        =   10
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   435
      Left            =   60
      TabIndex        =   9
      Top             =   2880
      Width           =   1035
   End
End
Attribute VB_Name = "FThisOrderOnlyAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lClassAlways = 1

Private WithEvents m_oBrokenRules As BrokenRules
Attribute m_oBrokenRules.VB_VarHelpID = -1
Private WithEvents m_oState As State
Attribute m_oState.VB_VarHelpID = -1

Private m_oAddr As Address

Private m_bLoading As Boolean

'This is distinct from Form BrokenRules validation and m_oState.IsValid
Private m_bUPSValidated As Boolean

Private m_RetVal As VbMsgBoxResult
Private m_oUPSXAV As MSSOAPLib30.SoapClient30


'*********************
' Form events
'*********************

Private Sub Form_Initialize()
    Set m_oState = New State
End Sub


Private Sub Form_Load()

    'Hard-code the Country to USA
    cboCountry.AddItem "USA"
    cboCountry.ListIndex = 0

    'Load the States by Country (this is a global reference)
    g_rstStates.Filter = "CountryID = '" & Trim(cboCountry.text) & "'"
    LoadCombo cboState, g_rstStates, "StateID", , , True
    SetComboByText cboState, "", True
    g_rstStates.Filter = adFilterNone
    
    Set m_oUPSXAV = CreateUPSProxy()
    
End Sub


'**********
'Properties
'**********

'This property is being used to track changes that do not trigger a state transition.
Private Property Let UPSValidated(ByVal i_bNewValue As Boolean)
    m_bUPSValidated = i_bNewValue
    SetButtons
End Property


'**************
'Control events
'**************

Private Sub txtName_Change()
    If m_bLoading Then Exit Sub
    
    m_oAddr.AddrName = Trim$(txtName.text)
    m_oBrokenRules.Validate txtName
    m_oState.IsDirty = True
    SetButtons
End Sub

Private Sub txtAddr1_Change()
    If m_bLoading Then Exit Sub
    
    m_oAddr.Addr1 = Trim$(txtAddr1.text)
    m_oBrokenRules.Validate txtAddr1
    m_oState.IsDirty = True
    UPSValidated = False
End Sub

Private Sub txtAddr2_Change()
    If m_bLoading Then Exit Sub
   
    m_oAddr.Addr2 = Trim$(txtAddr2.text)
    m_oBrokenRules.Validate txtAddr2
    m_oState.IsDirty = True
    UPSValidated = False
End Sub

Private Sub txtCity_Change()
    If m_bLoading Then Exit Sub

    m_oAddr.City = Trim$(txtCity.text)
    m_oBrokenRules.Validate txtCity
    m_oState.IsDirty = True
    UPSValidated = False
End Sub

Private Sub cboState_Click()
    If m_bLoading Then Exit Sub

    m_oAddr.State = Trim$(cboState.text)
    m_oBrokenRules.Validate cboState
    m_oState.IsDirty = True
    UPSValidated = False
End Sub

Private Sub txtZip_Change()
    If m_bLoading Then Exit Sub

    m_oAddr.Zip = Trim$(txtZip.text) + Trim$(txtExtZip.text)
    m_oBrokenRules.Validate txtZip
    m_oState.IsDirty = True
    UPSValidated = False
End Sub

Private Sub txtExtZip_Change()
    If m_bLoading Then Exit Sub

    m_oAddr.Zip = Trim$(txtZip.text) + Trim$(txtExtZip.text)
    m_oBrokenRules.Validate txtExtZip
    m_oState.IsDirty = True
    UPSValidated = False
End Sub


'Buttons
Private Sub cmdValidateWithUPS_Click()
    Dim Addr1 As String
    Dim Addr2 As String
    Dim City As String
    Dim State As String
    Dim Zip5 As String
    Dim Zip4 As String
    Dim Status As Integer
    Dim Class As Integer
    Dim Count As Integer
    Dim Response As Integer

    With m_oAddr
        Addr1 = .Addr1
        Addr2 = .Addr2
        City = .City
        State = .State
        Zip5 = Left(.Zip, 5) '5 digit
        If Len(.Zip) = 9 Then
            Zip4 = Right(.Zip, 4)  '+4
        Else
            Zip4 = ""
        End If
    End With
    
    On Error GoTo EH
    
    Screen.MousePointer = vbHourglass
    Response = m_oUPSXAV.ValidateClassifyAddress(Addr1, Addr2, City, State, Zip5, Zip4, Status, Class, Count)
    Screen.MousePointer = vbDefault
    
    ' "XAV Error"
    If Response = -1 Then
        'Default to Commercial on error
        m_oAddr.Residential = False
        UPSValidated = True
        txtStatus.text = "Unable to perform validation. The UPS service is unavailable."
        txtStatus.ForeColor = vbRed
    Else
    
        'ValidAddress or (AmbiguousAddress and Count = 1) = Valid
        If Status = 2 Or (Status = 1 And Count = 1) Then

'added 10/20/09 LR
            If cboState.text <> State Then
                If vbNo = MsgBox("UPS Address Validation just changed your state from " & cboState.text & " to " & State & vbCrLf & "Do you want to accept this?", _
                    vbExclamation + vbYesNo, "UPS Address Validation") Then
                
                    'Allow UPS validation override
                    If m_oState.IsValid Then
                        cmdOK.Enabled = True
                        cmdOK.SetFocus
                        UPSValidated = True
                        txtStatus.text = "Overrode UPS address validation"
                        txtStatus.ForeColor = vbRed
                    End If
                    Exit Sub
                End If
            End If
            
            txtAddr1.text = Addr1
            txtAddr2.text = Addr2
            
            'truncate to fit the tciAddress field
            txtCity.text = Left(City, 20)
    
            cboState.text = State
            txtZip.text = Zip5
            txtExtZip.text = Zip4
            UPSValidated = True
            
            'If Unknown, try to re-classify.
            If (Class = 0) Then
                Screen.MousePointer = vbHourglass
                Response = m_oUPSXAV.ClassifyAddress(Addr1, Addr2, City, State, Zip5, Zip4, Class)
                Screen.MousePointer = vbDefault
                If Response = -1 Then   ' "XAV Error"
                    'Default to Commercial on error
                    m_oAddr.Residential = False
                    txtStatus.text = "Unable to perform validation. The UPS service is unavailable."
                    txtStatus.ForeColor = vbRed
                Else
                    'Default to Commercial if still unknown
                    m_oAddr.Residential = IIf((Class = 2), True, False)
                End If
            
            'Class 2 = Residential, anything else is Commercial
            Else
                m_oAddr.Residential = IIf((Class = 2), True, False)
            End If
            
            m_oBrokenRules.Validate
            If cmdOK.Enabled Then
                cmdOK.SetFocus
            End If
            
        'Ambiguous
        ElseIf (Status = 1 And Count > 1) Then
            m_oAddr.Residential = False
            UPSValidated = False
            txtStatus.text = "Ambiguous address. Please check your data and try again."
            txtStatus.ForeColor = vbRed

            'Allow UPS validation override
            If m_oState.IsValid Then
                cmdOK.Enabled = True
                cmdOK.SetFocus
            End If
            
        'NoCandidates (Status = 0)
        Else
            m_oAddr.Residential = False
            UPSValidated = False
            txtStatus.text = "Invalid address. Please check your data and try again."
            txtStatus.ForeColor = vbRed

            'Allow UPS validation override
            If m_oState.IsValid Then
                cmdOK.Enabled = True
                cmdOK.SetFocus
            End If
            
        End If
    End If
    
    Exit Sub
EH:
    'SOAP connection timeout
    If Err.Number = 5415 Then
        Screen.MousePointer = vbDefault
        'Default to Commercial on error
        m_oAddr.Residential = False
        UPSValidated = True
        txtStatus.text = "Unable to perform validation. The UPS service is unavailable."
        txtStatus.ForeColor = vbRed

        'Log the error
        LogEvent "FThisOrderOnlyAddress.frm", "cmdValidateWithUPS_Click", Err.Description
    End If
End Sub


Private Sub cmdClear_Click()
    m_oAddr.Clear False 'Do not clear payment terms.
    m_oAddr.AddrType = TOO
    
    m_bLoading = True
    LoadDisplayControls
    m_bLoading = False
    
    m_oBrokenRules.Validate
    m_oState.IsDirty = True
    UPSValidated = False
    txtName.SetFocus
End Sub


Private Sub cmdUndo_Click()
    m_oAddr.Restore
    
    m_bLoading = True
    LoadDisplayControls
    m_bLoading = False
    
    m_oBrokenRules.Validate
    m_oState.IsDirty = False
    UPSValidated = (m_oState.IsNew = False)
    txtName.SetFocus
End Sub


Private Sub cmdCancel_Click()
    m_RetVal = VbMsgBoxResult.vbCancel
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    Dim Response As VbMsgBoxResult
    Dim msg As String

    'Handle validation override
    msg = "UPS says this address can not be validated. " + vbCrLf _
         & "If you want to go ahead and use it anyway, click OK. " + vbCrLf _
         & "Otherwise, click Cancel and continue editing."
    
    If m_bUPSValidated = False Then
        Response = MsgBox(msg, vbOKCancel + vbExclamation, "Validation Override")
        
        'message limit is 512 Chars
        If Response = VbMsgBoxResult.vbOK Then
            LogEvent "FThisOrderOnlyAddress", "cmdOK_Click", "Address Validation Override. " _
                    & "CSR: " & GetUserName & ", AddrName: " & m_oAddr.AddrName & ", Addr1: " _
                    & m_oAddr.Addr1 & ", Addr2: " & m_oAddr.Addr2 & ", City: " & m_oAddr.City _
                    & ", State: " & m_oAddr.State & ", Zip: " & m_oAddr.Zip
            
            m_RetVal = VbMsgBoxResult.vbOK
            Me.Hide
        Else
            Exit Sub
        End If
    Else
        m_RetVal = VbMsgBoxResult.vbOK
        Me.Hide
    End If

End Sub


'*******
'Methods
'*******
Public Function EditShipAddress(ByRef i_oAddr As Address) As VbMsgBoxResult
   
    LoadValidationRules
    
    m_bLoading = True
    
    If i_oAddr.AddrType <> TOO Then
        Me.caption = "Create Shipping Address"
        Set m_oAddr = New Address
        m_oAddr.AddrType = TOO
        m_oState.IsNew = True
        m_oState.IsValid = False
        m_oState.IsDirty = False
        m_bUPSValidated = False
    Else
        Me.caption = "Edit Shipping Address"
        Set m_oAddr = i_oAddr
        LoadDisplayControls
        m_oState.IsNew = False
        m_oState.IsValid = True
        m_oState.IsDirty = False
        m_bUPSValidated = True
    End If
    
    m_oAddr.Backup
    
    m_oBrokenRules.Validate
    
    SetButtons
    
    m_bLoading = False
    
    Me.Show vbModal

    If m_RetVal = VbMsgBoxResult.vbOK Then
        If i_oAddr.AddrType <> TOO Then
            Set i_oAddr = m_oAddr
        End If
    Else
        If i_oAddr.AddrType = TOO Then
            m_oAddr.Restore
        End If
    End If
    
    EditShipAddress = m_RetVal
    
    'Clean up
    Set m_oAddr = Nothing
    m_oBrokenRules.Destroy
    Set m_oBrokenRules = Nothing
    Unload Me
    
End Function


'AddRuleLength params must match Sage tciAddress field lengths

Private Sub LoadValidationRules()
    Dim oCtlWrapper As ControlWrapper
    
    Set m_oBrokenRules = New BrokenRules
    m_oBrokenRules.Form = Me
    
    'Name
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtName, "Address Name", True, False)
    oCtlWrapper.AddRuleLength , 40, k_lClassAlways
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtName, "Address Name", True, False)
    oCtlWrapper.AddRuleRequired "", k_lClassAlways
    
    'Address 1
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtAddr1, "Address Line 1", True, False)
    oCtlWrapper.AddRuleLength , 40, k_lClassAlways
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtAddr1, "Address Line 1 cannot contain a # sign.", True, False)
    oCtlWrapper.AddRuleRegEx "^[^#]*$", k_lClassAlways
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtAddr1, "Address Line 1", True, False)
    oCtlWrapper.AddRuleRequired "", k_lClassAlways
    
    'Address 2
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtAddr2, "Address Line 2", True, False)
    oCtlWrapper.AddRuleLength , 40, k_lClassAlways

    'City
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtCity, "Address City", True, False)
    oCtlWrapper.AddRuleLength , 20, k_lClassAlways
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtCity, "Address City", True, False)
    oCtlWrapper.AddRuleRequired "", k_lClassAlways
    
    'State
    Set oCtlWrapper = m_oBrokenRules.AddControl(cboState, "Address State", True, False)
    oCtlWrapper.AddRuleStateID "USA", k_lClassAlways
    
    'Zip
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtZip, "Invalid Primary Zip Code.", True, True)
    oCtlWrapper.AddRuleRegEx "^(\d{5})$", k_lClassAlways
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtZip, "Zip Code", True, True)
    oCtlWrapper.AddRuleRequired "", k_lClassAlways
    
    'Extended Zip
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtExtZip, "Invalid Extended Zip Code.", True, True)
    oCtlWrapper.AddRuleRegEx "^(\d{4})?$", k_lClassAlways
    
    m_oBrokenRules.EnableClass k_lClassAlways, True
End Sub


Private Sub LoadDisplayControls()
    With m_oAddr
        txtName.text = .AddrName
        txtAddr1.text = .Addr1
        txtAddr2.text = .Addr2
        txtCity.text = .City
        
        SetComboByText cboState, .State, True
        
        txtZip.text = Left(.Zip, 5) '5 digit
        If Len(.Zip) = 9 Then
            txtExtZip.text = Right(.Zip, 4) '+4
        Else
            txtExtZip = ""
        End If
        
        cboCountry.text = .CountryID
    End With
End Sub


Private Sub SetButtons()
    cmdClear.Enabled = (m_oState.IsNew = False)
    cmdUndo.Enabled = (m_oState.IsDirty = True)
    
    If (m_oState.IsValid And m_oState.IsDirty And (m_bUPSValidated = True)) Then
        cmdOK.Enabled = True
        cmdValidateWithUPS.Default = False
        cmdOK.Default = True
    Else
        cmdOK.Enabled = False
        cmdOK.Default = False
        cmdValidateWithUPS.Default = True
    End If
    
    'Ready to Validate
    If m_oBrokenRules.IsValid(txtAddr1) _
        And m_oBrokenRules.IsValid(txtCity) _
        And m_oBrokenRules.IsValid(cboState) _
        And m_oBrokenRules.IsValid(txtZip) Then
        
        If m_bUPSValidated = False Then
            cmdValidateWithUPS.Enabled = True
            cmdValidateWithUPS.caption = "&Validate"
            txtStatus.text = "Address has not been validated. Click Validate to continue."
            txtStatus.ForeColor = vbRed
        Else
            cmdValidateWithUPS.Enabled = False
            cmdValidateWithUPS.caption = "&Validate"
            txtStatus.text = "Address has been validated."
            txtStatus.ForeColor = vbBlack
        End If
        
    'Ready to Auto-Complete
    ElseIf (m_oBrokenRules.IsValid(txtAddr1) And m_oBrokenRules.IsValid(txtZip) And m_bUPSValidated = False) _
        Or (m_oBrokenRules.IsValid(txtAddr1) And m_oBrokenRules.IsValid(txtCity) And m_bUPSValidated = False) Then

        cmdValidateWithUPS.Enabled = True
        cmdValidateWithUPS.caption = "&Auto-Complete"
        txtStatus.text = "Address has not been validated."
        txtStatus.ForeColor = vbRed
    Else
        cmdValidateWithUPS.Enabled = False
        cmdValidateWithUPS.caption = "&Auto-Complete"
        txtStatus.text = "Address has not been validated."
        txtStatus.ForeColor = vbRed
    End If
End Sub


'**************
'Event Handlers
'**************

Private Sub m_oBrokenRules_HaveBrokenRules()
    m_oState.IsValid = False
End Sub

Private Sub m_oBrokenRules_NoBrokenRules()
    m_oState.IsValid = True
End Sub

Private Sub m_oState_StateChange()
    If m_bLoading Then Exit Sub
    
    SetButtons
End Sub


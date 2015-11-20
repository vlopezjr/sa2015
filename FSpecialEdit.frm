VERSION 5.00
Begin VB.Form FLiteEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lite Edit"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   675
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3300
      Width           =   2835
   End
   Begin VB.ComboBox cboPmtTerms 
      Height          =   312
      Left            =   1440
      TabIndex        =   15
      Text            =   "cboPmtTerms"
      Top             =   2505
      Width           =   1635
   End
   Begin VB.TextBox txtPO 
      Height          =   312
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   1635
   End
   Begin VB.ComboBox cboShipVia 
      Height          =   312
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   516
      Width           =   1635
   End
   Begin VB.CheckBox chkShipComplete 
      Caption         =   "Ship Complete"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   2955
      Width           =   1452
   End
   Begin VB.CheckBox chkBillRecipient 
      Caption         =   "Bill Recipient"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2100
      Width           =   1335
   End
   Begin VB.TextBox txtUPSAcct 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   312
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   8
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1704
      Width           =   1635
   End
   Begin VB.TextBox txtShipToName 
      Height          =   312
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   4
      Top             =   912
      Width           =   1635
   End
   Begin VB.TextBox txtShipToPhone 
      Height          =   312
      Left            =   1440
      MaxLength       =   17
      TabIndex        =   3
      Top             =   1308
      Width           =   1635
   End
   Begin VB.CommandButton cmdUPSUpdate 
      Caption         =   "&Change"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   1740
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1140
      TabIndex        =   1
      Top             =   4140
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Top             =   4140
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Notes"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   3300
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Payment Terms"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer PO#"
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   200
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship Via"
      Height          =   255
      Index           =   24
      Left            =   540
      TabIndex        =   13
      Top             =   680
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "UPS Acct"
      Height          =   255
      Index           =   8
      Left            =   420
      TabIndex        =   12
      Top             =   1785
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship To Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1050
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship To Phone #"
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   1400
      Width           =   1275
   End
End
Attribute VB_Name = "FLiteEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancel As Boolean
Private m_bLoading As Boolean
Private m_oOrder As Order

Private m_bReceiptWasCancelled As Boolean

'validation rule classid
Private Const ccShipToData = 13
Private WithEvents m_oBrokenRules As BrokenRules
Attribute m_oBrokenRules.VB_VarHelpID = -1


Private Sub Form_Unload(Cancel As Integer)
    m_oBrokenRules.Destroy
    Set m_oBrokenRules = Nothing
End Sub


Public Function Load(ByRef i_oOrder As Order) As Boolean
    Dim HasReceipt As Boolean
    
    m_bLoading = True
    
    Set m_oBrokenRules = New BrokenRules
    m_oBrokenRules.Form = Me
    LoadValidationRules
    
    Set m_oOrder = i_oOrder

    SetUpShipVia cboShipVia, m_oOrder.WhseKey, m_oOrder.ShipMethKey
        
    If m_oOrder.ShipComplete Then
        chkShipComplete.value = vbChecked
        chkShipComplete.Enabled = True
    Else
        chkShipComplete.value = vbUnchecked
        chkShipComplete.Enabled = False
    End If
    
    txtPO.text = i_oOrder.PurchOrd
    
    If Left(Trim(cboShipVia.text), 3) = "UPS" And m_oOrder.Customer.HasAccount Then
        If Not m_oOrder.PmtTerms.IsCOD Then
        
            chkBillRecipient.Enabled = True
            If m_oOrder.UPSAcct <> "" Then
                txtUPSAcct.text = m_oOrder.UPSAcct
                chkBillRecipient.value = vbChecked
                cmdUPSUpdate.Enabled = True
            End If
        End If
   End If
   
    txtShipToName.text = m_oOrder.ShipToName
    txtShipToPhone.text = m_oOrder.ShipToPhone

    m_oOrder.PmtTerms.LoadComboBox cboPmtTerms
    SetComboByText cboPmtTerms, m_oOrder.PmtTerms.ID
    
    cboPmtTerms.Enabled = True
            
    If OrderHasShipment(m_oOrder.OPKey) Then
        txtNotes.text = txtNotes.text & "Order has been packed" & vbCrLf
        cboPmtTerms.Enabled = False
    End If
    
    If m_oOrder.IsDropShip Then
        txtNotes.text = txtNotes.text & "Order is a drop ship" & vbCrLf
        cboPmtTerms.Enabled = False
    End If
    
    m_bLoading = False
    
    Me.caption = "Edit Committed Order [OP-" & m_oOrder.OPKey & " SO-" & m_oOrder.TranNo & "]"
    Show vbModal
    
    'check for changes and write to Order History if necessary
    
    'need to check for changes
    If (chkShipComplete.value = vbChecked) <> m_oOrder.ShipComplete Then
        With m_oOrder
            LogDB.LogActivity "SA", "Lite Edit: Ship Complete turned off", _
                .OPKey, .soKey, .TranNo, , , , .WhseKey
        End With
    End If
        
    If Not m_bCancel Then
        With m_oOrder

            .PurchOrd = Trim(txtPO.text)
            .ShipComplete = (chkShipComplete.value = vbChecked)
            .ShipMethKey = cboShipVia.ItemData(cboShipVia.ListIndex)
            .PmtTerms.Key = cboPmtTerms.ItemData(cboPmtTerms.ListIndex)
            .UPSAcct = Trim(txtUPSAcct.text)
            .ShipToName = Trim$(txtShipToName.text)
            .ShipToPhone = txtShipToPhone.text

            'this triggers a special save (on a read-only order)
            If Not .Save(i_bForcePending:=False, i_bSage:=(m_oOrder.soKey > 0), i_bCommitOrder:=False) Then
                msg "Error happened while saving lite edits"
            Else
                Load = True
            End If
        
        End With
    End If
    Unload Me
End Function


'set flag if the box started as checked and was unchecked
Private Sub chkHasReceipt_Click()
    m_bReceiptWasCancelled = True
End Sub


Private Sub cboShipVia_Click()
    
    EnableShipToContactCtrls

    If m_bLoading Then Exit Sub
    
    m_bLoading = True

    If Not m_oOrder.PmtTerms.IsCOD Then
    
        If Left(Trim(cboShipVia.text), 3) = "UPS" Then
                If m_oOrder.Customer.HasAccount Then
                    chkBillRecipient.Enabled = True
                Else
                    DisableBillRecipient
                End If
        Else
            'The cascading event of changing warehouse maybe cause Bill Recipient to be disabled
            DisableBillRecipient
        End If
    Else
        DisableBillRecipient
    End If
        
    m_bLoading = False
End Sub

Private Sub txtShipToName_Change()
    m_oBrokenRules.Validate txtShipToName
    SetSaveButton
End Sub

Private Sub txtShipToPhone_Change()
    m_oBrokenRules.Validate txtShipToPhone
    SetSaveButton
End Sub


Private Sub DisableBillRecipient()
    chkBillRecipient.Enabled = False
    chkBillRecipient.value = vbUnchecked
    txtUPSAcct.text = ""
    cmdUPSUpdate.Enabled = False
End Sub


Private Sub chkBillRecipient_Click()
    If m_bLoading Then Exit Sub
    
    Dim sTemp As String
    
    m_bLoading = True

    If chkBillRecipient.value = vbChecked Then
        If m_oOrder.UPSAcct <> "" Then
            sTemp = m_oOrder.UPSAcct
        Else
            sTemp = BillRecipientStat
        End If
        If Trim(sTemp) = "" Then
            msg "This customer has not been set-up to support 'Bill Recipient' freight terms", vbOKOnly + vbExclamation, "Bill Recipient"
            cmdUPSUpdate.Enabled = False
            txtUPSAcct.text = ""
            chkBillRecipient.value = vbUnchecked
        Else
            cmdUPSUpdate.Enabled = True
            txtUPSAcct.text = Trim(sTemp)
        End If
    Else
        cmdUPSUpdate.Enabled = False
        txtUPSAcct.text = ""
    End If
    m_bLoading = False
End Sub


Private Sub cmdCancel_Click()
    m_bCancel = True
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    m_bCancel = False
    Me.Hide
End Sub


Private Function BillRecipientStat() As String
    Dim rst As ADODB.Recordset
    Dim dfltRst As ADODB.Recordset
    
    If m_oOrder.Customer.ShipAddr.AddrKey > 0 Then
        BillRecipientStat = SetUPSAcct(m_oOrder.Customer.ShipAddr.AddrKey)
        If BillRecipientStat <> "" Then Exit Function
    End If
        
    Set dfltRst = LoadDiscRst("Select DfltShipToAddrKey from tarCustomer where CustKey = " & m_oOrder.Customer.Key)
    If Not dfltRst.EOF Then
        BillRecipientStat = SetUPSAcct(dfltRst.Fields("DfltShipToAddrKey"))
    End If
    Set dfltRst = Nothing
End Function


Private Function SetUPSAcct(lAddrKey As Long) As String
    Dim rst As ADODB.Recordset
    
    Set rst = LoadDiscRst("Select * from tcpUPSAcct where CustAddrKey = " & lAddrKey)
    If Not rst.EOF Then
        SetUPSAcct = Trim(rst.Fields("UPSAcct").value)
    End If
    Set rst = Nothing
End Function


Private Sub cmdUPSUpdate_Click()
    Dim sNewUPSAcct As String
    Dim oFrm As FChangeUPSAcct
    Dim lNewAddrKey As Long
    
    
    Set oFrm = New FChangeUPSAcct
    sNewUPSAcct = oFrm.SearchUPSAcct(m_oOrder.Customer.ID, m_oOrder.Customer.Key, Trim(txtUPSAcct.text))
    
    If Trim(sNewUPSAcct) = "" Then Exit Sub
    
    SetWaitCursor True
    If vbYes = msg("Are you sure that you want to change UPS Account for this order?" _
                , vbYesNo + vbExclamation, "Change UPS Account") Then
            txtUPSAcct.text = Trim(sNewUPSAcct)
    End If
    SetWaitCursor False
End Sub

Private Sub EnableShipToContactCtrls()

    If cboShipVia.text = "UPS STND" Or cboShipVia.text = "UPS Red AM" Then
        txtShipToName.Enabled = True
        txtShipToPhone.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        m_oBrokenRules.EnableClass ccShipToData, True
    Else
        txtShipToName.Enabled = False
        txtShipToPhone.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False
        m_oBrokenRules.EnableClass ccShipToData, False
    End If
    
    m_oBrokenRules.Validate txtShipToName
    m_oBrokenRules.Validate txtShipToPhone
    SetSaveButton
End Sub

Private Sub LoadValidationRules()
    Dim oCtlWrapper As ControlWrapper

    With m_oBrokenRules
        Set oCtlWrapper = .AddControl(txtShipToName, "Ship To Contact Name", True, True)
        oCtlWrapper.AddRuleRequired "", ccShipToData
        Set oCtlWrapper = .AddControl(txtShipToPhone, "Ship To Contact Phone Number", True, True)
        oCtlWrapper.AddRuleRequired "", ccShipToData
    End With
End Sub

Private Sub SetSaveButton()
    If m_oBrokenRules.Count = 0 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

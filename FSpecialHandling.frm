VERSION 5.00
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "mmremark.ocx"
Begin VB.Form FSpecialHandling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Special Handling"
   ClientHeight    =   3210
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkBillDifferentRate 
      Caption         =   "Bill Different Rate"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   585
      Width           =   1575
   End
   Begin VB.CheckBox chkFreeFreight 
      Caption         =   "Free Freight"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1515
      Width           =   1575
   End
   Begin VB.ComboBox cboShipVia 
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox chkDeposit 
      Caption         =   "Deposit"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1515
      Width           =   1575
   End
   Begin VB.CheckBox chkInboundFreight 
      Caption         =   "Inbound Freight"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1980
      Width           =   1575
   End
   Begin VB.CheckBox chkReducedFreight 
      Caption         =   "Reduced Freight"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   1050
      Width           =   1575
   End
   Begin VB.CheckBox chkPartsNoCharge 
      Caption         =   "Parts No Charge"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   1980
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3900
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2820
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin MMRemark.RemarkViewer rvOrderSH 
      Height          =   1035
      Left            =   3660
      TabIndex        =   5
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1826
      ContextID       =   "ViewOrderSH"
      Caption         =   "Special Handling Remarks"
   End
   Begin VB.Label lblShipMethod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1740
      TabIndex        =   11
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Shipping Method"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "FSpecialHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bLoad As Boolean
Private m_bCancel As Boolean

Private m_ShipMethod As String

Public Property Get ShipMethod() As String
    ShipMethod = m_ShipMethod
End Property

Public Property Let ShipMethod(sNewValue As String)
    m_ShipMethod = sNewValue
End Property


Public Sub Load(ByRef i_oOrder As Order)
                
    Me.caption = "Special Handling for OP " & i_oOrder.OPKey
    
    rvOrderSH.OwnerID = i_oOrder.OPKey
    
    m_bLoad = True
    
    'setup the checkbox state to match the Order object
    
    lblShipMethod.caption = m_ShipMethod
    
    'this loads the combobox
    
    SetUpShipVia cboShipVia, i_oOrder.WhseKey, i_oOrder.ShipMethKey
        
    If i_oOrder.BillDifferentRate Then
        chkBillDifferentRate.value = vbChecked
        cboShipVia.Enabled = True
        SetComboByKey cboShipVia, i_oOrder.BillMethKey
    Else
        chkBillDifferentRate.value = vbUnchecked
        cboShipVia.Enabled = False
    End If

    If i_oOrder.ReducedFreight Then
        chkReducedFreight.value = vbChecked
    Else
        chkReducedFreight.value = vbUnchecked
    End If
    
    If i_oOrder.HasFreeFreight Then
        chkFreeFreight.value = vbChecked
    Else
        chkFreeFreight.value = vbUnchecked
    End If

    If i_oOrder.HasPartsNoCharge Then
        chkPartsNoCharge.value = vbChecked
    Else
        chkPartsNoCharge.value = vbUnchecked
    End If
    
    If i_oOrder.HasInboundFreight Then
        chkInboundFreight.value = vbChecked
    Else
        chkInboundFreight.value = vbUnchecked
    End If
    
    If i_oOrder.HasDeposit Then
        chkDeposit.value = vbChecked
    Else
        chkDeposit.value = vbUnchecked
    End If
    
    m_bLoad = False
    
    'block here
    Me.Show vbModal
    
    'on the way out
    If Not m_bCancel Then
        m_bLoad = True
        
        With i_oOrder
            .BillDifferentRate = (chkBillDifferentRate.value = vbChecked)
            If .BillDifferentRate Then
                .BillMethKey = cboShipVia.ItemData(cboShipVia.ListIndex)
            End If
            .ReducedFreight = (chkReducedFreight.value = vbChecked)
            .HasFreeFreight = (chkFreeFreight.value = vbChecked)
            .HasInboundFreight = (chkInboundFreight.value = vbChecked)
            .HasPartsNoCharge = (chkPartsNoCharge.value = vbChecked)
            .HasDeposit = (chkDeposit.value = vbChecked)
        End With

        m_bLoad = False
    End If

    Unload Me
End Sub


Private Function GetRemarkCount(RemarkType As String) As Integer
    Dim oRemark As MemoMeister.remark
    Dim Count As Integer
    For Each oRemark In rvOrderSH.RemarkContext.RemarkList
        If oRemark.RemarkType.TypeID = RemarkType Then
            Count = Count + 1
        End If
    Next
    GetRemarkCount = Count
End Function


Private Sub chkDeposit_Click()
    If m_bLoad Then Exit Sub
    If chkDeposit.value = vbChecked Then 'And GetRemarkCount("Order.Deposit") = 0 Then
        rvOrderSH.RemarkContext.EditMemos True
    End If

    If chkDeposit.value = vbUnchecked And GetRemarkCount("Order.Deposit") > 0 Then
        MsgBox "There are Deposit remarks. Delete them before exiting.", vbExclamation, "Alert"
    End If
End Sub


Private Sub chkFreeFreight_Click()
    If m_bLoad Then Exit Sub

    If chkFreeFreight.value = vbChecked Then
        m_bLoad = True
        chkReducedFreight.value = vbUnchecked
        chkBillDifferentRate.value = vbUnchecked
        m_bLoad = False
        'If GetRemarkCount("Order.FreeFreight") = 0 Then
            rvOrderSH.RemarkContext.EditMemos True
        'End If
    End If
    
    If chkFreeFreight.value = vbUnchecked And GetRemarkCount("Order.FreeFreight") > 0 Then
        MsgBox "There are FreeFreight remarks. Delete them before exiting.", vbExclamation, "Alert"
    End If
End Sub


Private Sub chkInboundFreight_Click()
    If m_bLoad Then Exit Sub
    If chkInboundFreight.value = vbChecked Then 'And GetRemarkCount("Order.InboundFreight") = 0 Then
        rvOrderSH.RemarkContext.EditMemos True
    End If

    If chkInboundFreight.value = vbUnchecked And GetRemarkCount("Order.InboundFreight") > 0 Then
        MsgBox "There are InboundFreight remarks. Delete them before exiting.", vbExclamation, "Alert"
    End If
End Sub


Private Sub chkPartsNoCharge_Click()
    If m_bLoad Then Exit Sub

    If chkPartsNoCharge.value = vbChecked Then 'And GetRemarkCount("Order.PartsNoCharge") = 0 Then
        rvOrderSH.RemarkContext.EditMemos True
    End If

    If chkPartsNoCharge.value = vbUnchecked And GetRemarkCount("Order.PartsNoCharge") > 0 Then
        MsgBox "There are PartsNoCharge remarks. Delete them before exiting.", vbExclamation, "Alert"
    End If
    
End Sub


Private Sub chkReducedFreight_Click()
    If m_bLoad Then Exit Sub

    If chkReducedFreight.value = vbChecked Then
        m_bLoad = True
        chkFreeFreight.value = vbUnchecked
        chkBillDifferentRate.value = vbUnchecked
        m_bLoad = False
        'If GetRemarkCount("Order.ReducedFreight") = 0 Then
            rvOrderSH.RemarkContext.EditMemos True
        'End If
    End If
    
    If chkReducedFreight.value = vbUnchecked And GetRemarkCount("Order.ReducedFreight") > 0 Then
        MsgBox "There are ReducedFreight remarks. Delete them before exiting.", vbExclamation, "Alert"
    End If
    
End Sub


Private Sub chkBillDifferentRate_Click()
    If m_bLoad Then Exit Sub
    m_bLoad = True
    If chkBillDifferentRate.value = vbChecked Then
        chkFreeFreight.value = vbUnchecked
        chkReducedFreight.value = vbUnchecked
    End If
    cboShipVia.Enabled = (chkBillDifferentRate.value = vbChecked)
    m_bLoad = False
End Sub



Private Sub cmdCancel_Click()
    m_bCancel = True
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    Dim Count As Integer
    Dim countsummary As String
    Dim HasMissMatch As Boolean
    
    HasMissMatch = False
    
    'contrast the checkbox settings and MM remarks and gernerate any warnings
    
    Count = GetRemarkCount("Order.FreeFreight")
    If chkFreeFreight.value = vbChecked And Count = 0 Then
        countsummary = countsummary & "FreeFreight is checked but there are no Remarks" & vbCrLf
        HasMissMatch = True
    ElseIf chkFreeFreight.value = vbUnchecked And Count > 0 Then
        countsummary = countsummary & "FreeFreight is not checked but there are " & Count & " Remark(s)" & vbCrLf
        HasMissMatch = True
    End If
    
    Count = GetRemarkCount("Order.ReducedFreight")
    If chkReducedFreight.value = vbChecked And Count = 0 Then
        countsummary = countsummary & "ReducedFreight is checked but there are no Remarks" & vbCrLf
        HasMissMatch = True
    ElseIf chkReducedFreight.value = vbUnchecked And Count > 0 Then
        countsummary = countsummary & "ReducedFreight is not checked but there are " & Count & " Remark(s)" & vbCrLf
        HasMissMatch = True
    End If
    
    Count = GetRemarkCount("Order.PartsNoCharge")
    If chkPartsNoCharge.value = vbChecked And Count = 0 Then
        countsummary = countsummary & "PartsNoCharge is checked but there are no Remarks" & vbCrLf
        HasMissMatch = True
    ElseIf chkPartsNoCharge.value = vbUnchecked And Count > 0 Then
        countsummary = countsummary & "PartsNoCharge is not checked but there are " & Count & " Remark(s)" & vbCrLf
        HasMissMatch = True
    End If
    
    Count = GetRemarkCount("Order.InboundFreight")
    If chkInboundFreight.value = vbChecked And Count = 0 Then
        countsummary = countsummary & "InboundFreight is checked but there are no Remarks" & vbCrLf
        HasMissMatch = True
    ElseIf chkInboundFreight.value = vbUnchecked And Count > 0 Then
        countsummary = countsummary & "InboundFreight is not checked but there are " & Count & " Remark(s)" & vbCrLf
        HasMissMatch = True
    End If
    
    Count = GetRemarkCount("Order.Deposit")
    If chkDeposit.value = vbChecked And Count = 0 Then
        countsummary = countsummary & "Deposit is checked but there are no Remarks" & vbCrLf
        HasMissMatch = True
    ElseIf chkDeposit.value = vbUnchecked And Count > 0 Then
        countsummary = countsummary & "Deposit is not checked but there are " & Count & " Remark(s)" & vbCrLf
        HasMissMatch = True
    End If
    
    If HasMissMatch Then
        countsummary = countsummary & vbCrLf & "Are you sure you want to exit?"
        
        Dim oFrm As FSpecialHandlingWarnings
        Dim Result As VbMsgBoxResult
        
        Set oFrm = New FSpecialHandlingWarnings
        Result = oFrm.ShowMessage(countsummary)
    
        If Result = vbOK Then
            m_bCancel = False
            Me.Hide
        End If
    Else
        m_bCancel = False
        Me.Hide
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    m_bCancel = True
    Me.Hide
End Sub


VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "MMRemark.ocx"
Begin VB.Form FRMACreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authorize Items for Return"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   7515
   StartUpPosition =   1  'CenterOwner
   Begin MMRemark.RemarkViewer rvRMA 
      Height          =   795
      Left            =   4320
      TabIndex        =   4
      Top             =   3720
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1402
      ContextID       =   "ViewRMA"
      Caption         =   "RMA Remarks"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   312
      Left            =   6360
      TabIndex        =   3
      Top             =   4200
      Width           =   972
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&OK"
      Height          =   312
      Left            =   5280
      TabIndex        =   2
      Top             =   4200
      Width           =   972
   End
   Begin GridEX20.GridEX gdxRMA 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   5530
      Version         =   "2.0"
      ScrollToolTips  =   -1  'True
      ShowToolTips    =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      ColumnsCount    =   15
      Column(1)       =   "FRMACreate.frx":0000
      Column(2)       =   "FRMACreate.frx":016C
      Column(3)       =   "FRMACreate.frx":02C0
      Column(4)       =   "FRMACreate.frx":0400
      Column(5)       =   "FRMACreate.frx":0514
      Column(6)       =   "FRMACreate.frx":06A0
      Column(7)       =   "FRMACreate.frx":082C
      Column(8)       =   "FRMACreate.frx":0950
      Column(9)       =   "FRMACreate.frx":0AC8
      Column(10)      =   "FRMACreate.frx":0C40
      Column(11)      =   "FRMACreate.frx":0DA0
      Column(12)      =   "FRMACreate.frx":0F44
      Column(13)      =   "FRMACreate.frx":1084
      Column(14)      =   "FRMACreate.frx":1178
      Column(15)      =   "FRMACreate.frx":12CC
      FormatStylesCount=   6
      FormatStyle(1)  =   "FRMACreate.frx":1458
      FormatStyle(2)  =   "FRMACreate.frx":1590
      FormatStyle(3)  =   "FRMACreate.frx":1640
      FormatStyle(4)  =   "FRMACreate.frx":16F4
      FormatStyle(5)  =   "FRMACreate.frx":17CC
      FormatStyle(6)  =   "FRMACreate.frx":1884
      ImageCount      =   0
      PrinterProperties=   "FRMACreate.frx":1964
   End
   Begin VB.Label lblHeader 
      Caption         =   "Select the item and edit the quantity authorized for return."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5592
   End
End
Attribute VB_Name = "FRMACreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This screen is invoked both to create an RMA for an order and
'to add items to an existing RMA.

Private Const m_skSource = "Create RMA"

'this is the key for "Parts Wiz Error" in tcpRMAReason
Private Const PartsWizErrorReasonKey = 27

Private m_bAddItem As Boolean
Private m_oOrder As Order
Private m_OPKey As Long
Private m_RMAKey As Long

Private m_oRMAList As RMAList

Private m_bSuccess As Boolean

'Public Methods

Public Function CreateRMA(i_oOrder As Order) As Boolean
    On Error GoTo ErrorHandler

    'persist these
    m_bAddItem = False
    Set m_oOrder = i_oOrder
    m_OPKey = i_oOrder.OPKey
    
    rvRMA.OwnerID = ""
    rvRMA.OwnerID = i_oOrder.OPKey

    'Load order line items
    Set m_oRMAList = Billing.LoadNewRMA(i_oOrder.OPKey)
    
    If m_oRMAList.Count > 0 Then
        
        RefreshRMAGrid

        'This is modal. We're gonna stop here until the form closes (Update or Cancel)
        Me.Caption = "Creating RMA For OP " & i_oOrder.OPKey
        Me.Show vbModal
        
        CreateRMA = m_bSuccess
        
    Else
        CreateRMA = False
    End If

    Exit Function
    
ErrorHandler:
    CreateRMA = False
    msg Err.Description, vbOKOnly + vbCritical, Err.Source
End Function


Public Sub AddRMAItem(i_oOrder As Order, ByVal i_RMAKey As Long)
    'persist these
    m_bAddItem = True
    Set m_oOrder = i_oOrder
    m_OPKey = i_oOrder.OPKey
    m_RMAKey = i_RMAKey
    
    Set m_oRMAList = Billing.LoadNewRMA(i_oOrder.OPKey, m_bAddItem)
   
    
    'if there are RMAs in the list, refresh the grid
    If m_oRMAList.Count > 0 Then
        RefreshRMAGrid
        'This is modal. We're gonna stop here until the form closes (Update or Cancel)
        Me.Caption = "Add RMA Line Item for OP " & i_oOrder.OPKey
        Me.Show vbModal
    Else
        msg "Sorry. No remaining order line items are available. Check the Line tab.", vbOKOnly + vbExclamation, m_skSource
    End If
End Sub

'Form Events

Private Sub Form_Load()
    Dim colTemp As JSColumn
    Dim vl As JSValueList
    
    Set colTemp = gdxRMA.Columns("Reason")
    colTemp.HasValueList = True
    Set vl = colTemp.ValueList
    vl.Add 0, "- Select One -"
    g_rstRMAReason.MoveFirst
    Do While Not g_rstRMAReason.EOF
        vl.Add g_rstRMAReason.Fields("RMAReasonKey").value, g_rstRMAReason.Fields("RMAReasonID").value
        g_rstRMAReason.MoveNext
    Loop
    colTemp.EditType = jgexEditDropDown

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_oRMAList = Nothing
End Sub


'Buttons

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdUpdate_Click()
    Dim lIndex As Long
    Dim bFirstItem As Boolean
    Dim lAuthCount As Long
    
    
    On Error GoTo ErrorHandler

'NOTE:  isn't the form only displayed if the list exists with items?
    If m_oRMAList Is Nothing Then Exit Sub
    If m_oRMAList.Count < 1 Then Exit Sub
    
    SetWaitCursor True

    With m_oRMAList

        'Precheck:
        'scan all items in the list
        'if none are authorized, report this
        'if any are authorized without a reason, report this
        For lIndex = 1 To .Count
            If .Item(lIndex).Authorized Then
                lAuthCount = lAuthCount + 1
                If .Item(lIndex).Reason = 0 Then
                    SetWaitCursor False
                    msg "Specify a reason for the return.", vbExclamation + vbOKOnly, m_skSource
                    Exit Sub
                End If
            End If
        Next
        If lAuthCount = 0 Then
            SetWaitCursor False
            msg "First choose the items you want to authorize for return.", vbExclamation + vbOKOnly, m_skSource
            Exit Sub
        End If
        
        g_DB.Connection.BeginTrans
        
        bFirstItem = True
        'for each item in the list
        For lIndex = 1 To .Count
            'if the item has been authorized for return
            If .Item(lIndex).Authorized Then
                'if we're creating an RMA and this is the first authorized item
                If Not m_bAddItem And bFirstItem Then
                    'insert the RMA header in the database and if this fails, get out (needs better error handling)
                    If Not InsertRMA Then
                        'NOTE: better error reporting/handling?
                        cmdUpdate.Enabled = False
                        SetWaitCursor False
                        g_DB.Connection.RollbackTrans
                        Exit Sub
                    End If
                    bFirstItem = False
                End If
                
                'NOTE: what if this fails?
                'PRN#96
                InsertRMALine m_RMAKey, .Item(lIndex).OPLineKey, .Item(lIndex).QtyAuthorized, .Item(lIndex).Reason, .Item(lIndex).Restock, .Item(lIndex).CreditFreight
                
                If .Item(lIndex).Reason = PartsWizErrorReasonKey Then
                    If GetUserWhseID(m_oOrder.UserKey) = "STL" Then
                        Dim subject As String
                        Dim message As String
                        Dim sendto As String
                        subject = "RMA created for OP " & m_oOrder.OPKey & " due to PartsWiz Error"
                        message = ""
                        sendto = "joannar@caseparts.com, operations@caseparts.com, " & m_oOrder.UserId & "@caseparts.com"
                        
                        SendNotification subject, message, sendto
                    End If
                End If
    
            End If
        Next
        
        g_DB.Connection.CommitTrans

    End With

    RefreshRMAGrid  'NOTE: I don't think this is necessary
    
    SetWaitCursor False
    m_bSuccess = True
    Me.Hide
    
    DoEvents        'NOTE: I don't think this is necessary
    Exit Sub
    
ErrorHandler:
    m_bSuccess = False
    g_DB.Connection.RollbackTrans
    msg Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Err.Source
    ClearWaitCursor
    
End Sub

'Grid events

Private Sub RefreshRMAGrid()
    Dim i As Integer
    With gdxRMA
        .HoldFields
        .HoldSortSettings = True
        .ItemCount = m_oRMAList.Count
        .Refetch
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
End Sub


Private Sub gdxRMA_LostFocus()
    gdxRMA.Update
End Sub


Private Sub gdxRMA_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxRMA.Update
End Sub


Private Sub gdxRMA_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_oRMAList Is Nothing Then Exit Sub
    If RowIndex > m_oRMAList.Count Then Exit Sub
    
    With m_oRMAList.Item(RowIndex)
       Values(1) = .SOLineKey
       Values(3) = .Authorized
       Values(4) = .ItemID
       Values(5) = .Cost
       Values(6) = .Price
       Values(7) = .QtyAuthorized
       Values(11) = .ExtPrice
       Values(13) = .Reason
       Values(14) = .Restock * 100
       Values(15) = .CreditFreight
    End With
End Sub


Private Sub gdxRMA_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    '09/05/02  TeddyX
    'Add ErrorHandler and Numeric checking for Restock, QtyAuth field
    'This change is related with the type mismatch error in Error.Log
    'Error Report Time: 09-04-2002 13:01:24.    Module: FRMACreate.frm User: LupeT
    
    On Error GoTo ErrorHandler
    'persist Authorize, Reason and Restock to the underlying object
    
    m_oRMAList.Item(RowIndex).Authorized = Values(3)
    m_oRMAList.Item(RowIndex).Reason = Values(13)
    m_oRMAList.Item(RowIndex).CreditFreight = Values(15)
                      
    If Not IsNumeric(Values(14)) Then
        m_oRMAList.Item(RowIndex).Restock = 0
    ElseIf Values(14) >= 0 And Values(14) <= 100 Then
        m_oRMAList.Item(RowIndex).Restock = Values(14) / 100
    End If
    
'        Msg "The value in the Restock field is not valid.", vbOKOnly + vbExclamation, "RMA Creating"
'    ElseIf Values(14) < 0 Or Values(14) > 100 Then
'        Msg "The value in the Restock field should be between 0 and 100.", vbOKOnly + vbExclamation, "RMA Creating"
'    Else
'        If m_oRMAList.Item(RowIndex).Restock <> Values(14) Then
'            m_oRMAList.Item(RowIndex).Restock = Values(14) / 100
'        End If
'    End If
    
    If Not IsNumeric(Values(7)) Then
        msg "The value in the Qty Authorized field is not valid.", vbOKOnly + vbExclamation, "RMA Creating"
    Else
        If Values(7) > m_oRMAList.Item(RowIndex).lMaxQtyAuth Then
            msg "Qty Authorized must be less than or equal to " & m_oRMAList.Item(RowIndex).lMaxQtyAuth, vbExclamation, m_skSource
        ElseIf Values(7) = 0 Then
            msg "Qty Authorized cannot be 0.", vbExclamation, m_skSource
        Else
            m_oRMAList.Item(RowIndex).QtyAuthorized = Values(7)
            m_oRMAList.Item(RowIndex).ExtPrice = Values(6) * Values(7)
        End If
    End If
    Exit Sub
    
ErrorHandler:
    msg "The value in the Qty Authorized or Restock field is not valid. Please enter effective Qty Authorized or Restock.", vbOKOnly + vbExclamation, "RMA Creating"

End Sub


'Private functions

Private Function InsertRMA() As Boolean
    Dim cmd As ADODB.Command
    
    Set cmd = CreateCommandSP("spcpcRMAInsert")
    InsertRMA = True
    With cmd
        .Parameters("@_iOPKey").value = m_OPKey
        .Parameters("@_iUserID").value = GetUserName
        '.Parameters("@UPSCallTagNbr").Value = Trim(txtUPSTag.Text)
        .Execute
        m_RMAKey = .Parameters("@_oRMAKey").value
    End With
    
    If m_RMAKey < 0 Then
        msg "CreateRMA Failed. Contact the computer guys.", vbExclamation + vbOKOnly, m_skSource
        InsertRMA = False
    End If
End Function


Private Sub InsertRMALine(ByVal lRMAKey As Long, _
                            ByVal lOPLineKey As Long, _
                            ByVal lQtyAuth As Long, _
                            ByVal lReason As Long, _
                            ByVal dRestock As Double, _
                            ByVal bCreditFreight As Boolean)
    Dim cmd As ADODB.Command
    Dim lStatus As Long
        
    Set cmd = CreateCommandSP("spcpcRMAInsertItem")
    cmd.Parameters("@_iRMAKey").value = lRMAKey
    cmd.Parameters("@_iSOLineKey").value = lOPLineKey
    cmd.Parameters("@_iAuthorized").value = True
    'PRN#96
    cmd.Parameters("@_iQty").value = lQtyAuth
    cmd.Parameters("@_iUserID").value = GetUserName

    cmd.Parameters("@_iReasonCode").value = lReason 'cboReason.ItemData(cboReason.ListIndex)
    
'    If Trim(txtRestock.Text) <> "" Then
'        cmd.Parameters("@Restock").Value = CDbl(CDbl(txtRestock.Text) / 100)
'    Else
'        cmd.Parameters("@Restock").Value = Null
'    End If
    cmd.Parameters("@_iRestock").value = dRestock
    cmd.Parameters("@_iCreditFreight").value = bCreditFreight
    cmd.Execute
    
    lStatus = cmd.Parameters("@_oStatusCode").value
    If lStatus < 0 Then
        msg "Creating RMA Line of " & lOPLineKey & " for RMA " & lRMAKey & " failed.", _
            vbExclamation + vbOKOnly, m_skSource
    End If

    Set cmd = Nothing
End Sub


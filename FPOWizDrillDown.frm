VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FPOWizDrillDown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show Orders"
   ClientHeight    =   4470
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   312
      Left            =   7380
      TabIndex        =   1
      Top             =   4080
      Width           =   1032
   End
   Begin GridEX20.GridEX gdxOrders 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   6800
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   270
      ColumnsCount    =   8
      Column(1)       =   "FPOWizDrillDown.frx":0000
      Column(2)       =   "FPOWizDrillDown.frx":013C
      Column(3)       =   "FPOWizDrillDown.frx":03A8
      Column(4)       =   "FPOWizDrillDown.frx":04C8
      Column(5)       =   "FPOWizDrillDown.frx":05F4
      Column(6)       =   "FPOWizDrillDown.frx":0734
      Column(7)       =   "FPOWizDrillDown.frx":0854
      Column(8)       =   "FPOWizDrillDown.frx":0978
      FormatStylesCount=   5
      FormatStyle(1)  =   "FPOWizDrillDown.frx":0AA8
      FormatStyle(2)  =   "FPOWizDrillDown.frx":0BE0
      FormatStyle(3)  =   "FPOWizDrillDown.frx":0C90
      FormatStyle(4)  =   "FPOWizDrillDown.frx":0D44
      FormatStyle(5)  =   "FPOWizDrillDown.frx":0E1C
      ImageCount      =   0
      PrinterProperties=   "FPOWizDrillDown.frx":0ED4
   End
   Begin MSComctlLib.ImageList imglRemarks 
      Left            =   60
      Top             =   3900
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOWizDrillDown.frx":10AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOWizDrillDown.frx":14FE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FPOWizDrillDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_gwOrders As GridEXWrapper

Private m_iMode As Integer


Private Sub Form_Unload(Cancel As Integer)
    '4/7/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwOrders = Nothing
End Sub


'pop up a dialog with a list of customers and orders comprising the line item
'OrderDate, CustID, CustName, ShipVia, CSR, Qty, MM line item purchaing remark viewer

Public Sub ShowOrders(i_sItemID As String, i_lWhseKey As Long, Mode As Integer)
    Dim oRst As ADODB.Recordset
    Dim ssql As String
    Dim col As JSColumn

    SetWaitCursor True

    LoadImageList imglRemarks, gdxOrders
    
    m_iMode = Mode
    
    Set m_gwOrders = New GridEXWrapper
    m_gwOrders.Grid = gdxOrders

    'set common grid caption/datafield properties
        With gdxOrders.Columns
            Set col = .Item(1)
            col.Caption = "Trans #" '"SO#"
            col.DataField = "TranID"
            Set col = .Item(2)
            col.Caption = "Remarks"
            col.DataField = "Remarks"
            Set col = .Item(3)
            col.Caption = "Date"
            col.DataField = "TranDate"
            Set col = .Item(4)
            col.Caption = "Qty Ordered"
            col.DataField = "QtyOrd"
        End With

    If Mode = 0 Then   'Show SO Orders
        Set oRst = CallSP("spCPCpoStockOrderDetail", "@i_ItemID", i_sItemID, "@i_WhseKey", i_lWhseKey)
        With gdxOrders.Columns
            Set col = .Item(5)
            col.Caption = "Ship Method"
            col.DataField = "ShipMethDesc"
            Set col = .Item(6)
            col.Caption = "CSR"
            col.DataField = "CreateUserID"
            Set col = .Item(7)
            col.Caption = "Cust ID"
            col.DataField = "CustID"
            Set col = .Item(8)
            col.Caption = "Cust Name"
            col.DataField = "CustName"
        End With
        Me.Caption = "Show SO Orders"
    Else                'Show PO Orders
        'PRN #566 - Show Vendor Name, Whse and Cost
        'Set oRst = CallSP("spCPCPOPurchOrderDetail", "@i_ItemID", Trim$(i_sItemID), "@i_WhseKey", i_lWhseKey)
        Set oRst = CallSP("spCPCPOPurchOrderDetail", "@i_ItemID", Trim$(i_sItemID))
        With gdxOrders.Columns
            Set col = .Item(5)
            col.Caption = "Qty Received"
            col.DataField = "QtyRcvd"
            
            Set col = .Item(6)
            col.Caption = "Cost"
            col.SortType = jgexSortTypeNumeric
            col.Format = ".00"
            col.DataField = "UnitCost"
            
            Set col = .Item(7)
            col.Caption = "Vendor"
            col.DataField = "VendName"
            Set col = .Item(8)
            col.Caption = "Whse"
            col.DataField = "WhseID"
        End With
        Me.Caption = "Show PO Orders"
    End If
    
    AttachGrid gdxOrders, oRst
    
    SetWaitCursor False
    Me.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


'***BIG NOTE:
'***Though PO remarks are accessed by POKey in tciMemo, MM RemarkContext wants PO TranNo.

Private Sub gdxOrders_DblClick()
'2/25/05 LR
'Can't do this, both recordsets don't contain this field
'This code only worked because the grid wrapper was returning Empty (as a variant) rather than throwing an error.
'    If m_gwOrders.value("POKey") > 0 Then
    If m_iMode = 0 Then
        Call EditRemarks("ViewOrderLine", m_gwOrders.value("OPLineKey"))
    Else
'PRN 565
'        Call EditRemarks("ViewPO", m_gwOrders.value("POKey"))
'translate TranID to TranNo
        Call EditRemarks("ViewPO", CLng(Mid$(m_gwOrders.value("TranID"), 4)))
    End If
End Sub

Private Sub gdxOrders_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'EditRemarks m_gwOrders.value("OPLineKey")
'        If m_gwOrders.value("POKey") > 0 Then
        If m_iMode = 0 Then
            Call EditRemarks("ViewOrderLine", m_gwOrders.value("OPLineKey"))
        Else
'PRN 565
'            Call EditRemarks("ViewPO", m_gwOrders.value("POKey"))
'translate TranID to TranNo
        Call EditRemarks("ViewPO", CLng(Mid$(m_gwOrders.value("TranID"), 4)))
        End If
    End If
End Sub

Private Sub EditRemarks(Context As String, Key As Long)
    Dim oRC As RemarkContext
    'SO - "ViewOrderLine", OPLineKey
    'PO - "ViewPO", POKey
    Set oRC = New RemarkContext
    oRC.Edit Context, Key
End Sub


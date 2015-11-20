VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "MMRemark.ocx"
Begin VB.Form FPOWiz 
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   11790
   Begin VB.CommandButton cmdOTWizDev 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Open Transfer POs"
      Height          =   312
      Left            =   6600
      TabIndex        =   27
      Top             =   6600
      Width           =   1875
   End
   Begin VB.CommandButton cmdAddPartNbr 
      Caption         =   "Add Part"
      Height          =   312
      Left            =   3780
      TabIndex        =   22
      Top             =   6600
      Width           =   1092
   End
   Begin VB.CommandButton cmdRemoveUnusedParts 
      Caption         =   "Remove Unused Parts"
      Height          =   312
      Left            =   1860
      TabIndex        =   21
      Top             =   6600
      Width           =   1752
   End
   Begin VB.CommandButton cmdShowAllParts 
      Caption         =   "Show All Parts"
      Height          =   312
      Left            =   60
      TabIndex        =   20
      Top             =   6600
      Width           =   1752
   End
   Begin VB.TextBox txtPartNbr 
      Height          =   288
      Left            =   4920
      TabIndex        =   19
      Top             =   6600
      Width           =   1272
   End
   Begin VB.Frame frmDetail 
      Height          =   1032
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   9852
      Begin VB.Label lblPartNbr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   2112
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "QOH"
         Height          =   252
         Left            =   6360
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   432
      End
      Begin VB.Label lblQOH 
         Alignment       =   1  'Right Justify
         Caption         =   "QOH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   6240
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   552
      End
      Begin VB.Label lblShowOrders 
         Alignment       =   1  'Right Justify
         Caption         =   "QSO"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   432
      End
      Begin VB.Label lblQSO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Left            =   2400
         TabIndex        =   13
         Top             =   480
         Width           =   552
      End
      Begin VB.Label lblShowOrders 
         Alignment       =   1  'Right Justify
         Caption         =   "QPO"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   1
         Left            =   3120
         TabIndex        =   12
         Top             =   240
         Width           =   432
      End
      Begin VB.Label lblQPO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   552
      End
   End
   Begin VB.Frame frmDetail 
      Height          =   1032
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   9852
      Begin VB.Label lblCommitDate 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "aa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7200
         TabIndex        =   26
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "OP# / SO#"
         Height          =   252
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   912
      End
      Begin VB.Label lblOrderNbr 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2940
         TabIndex        =   8
         Top             =   480
         Width           =   1632
      End
      Begin VB.Label lblShipMethod 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   4680
         TabIndex        =   7
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label lblCSR 
         BackColor       =   &H80000016&
         Caption         =   "aa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   6240
         TabIndex        =   6
         Top             =   480
         Width           =   852
      End
      Begin VB.Label Label5 
         Caption         =   "Model Nbr"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   912
      End
      Begin VB.Label lblModelNbr 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1152
      End
      Begin VB.Label lblSerialNbr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "Serial Nbr"
         Height          =   252
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   792
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Height          =   312
      Left            =   0
      TabIndex        =   0
      Top             =   7044
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin GridEX20.GridEX gdxLines 
      Height          =   3195
      Left            =   60
      TabIndex        =   18
      Top             =   1140
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5636
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      ShowToolTips    =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      SelectionStyle  =   1
      Options         =   2
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   270
      ColumnsCount    =   20
      Column(1)       =   "FPOWizDev.frx":0000
      Column(2)       =   "FPOWizDev.frx":020C
      Column(3)       =   "FPOWizDev.frx":0390
      Column(4)       =   "FPOWizDev.frx":04D8
      Column(5)       =   "FPOWizDev.frx":061C
      Column(6)       =   "FPOWizDev.frx":07E4
      Column(7)       =   "FPOWizDev.frx":0940
      Column(8)       =   "FPOWizDev.frx":0B24
      Column(9)       =   "FPOWizDev.frx":0CE8
      Column(10)      =   "FPOWizDev.frx":0EBC
      Column(11)      =   "FPOWizDev.frx":103C
      Column(12)      =   "FPOWizDev.frx":1210
      Column(13)      =   "FPOWizDev.frx":13D8
      Column(14)      =   "FPOWizDev.frx":1544
      Column(15)      =   "FPOWizDev.frx":1708
      Column(16)      =   "FPOWizDev.frx":1888
      Column(17)      =   "FPOWizDev.frx":19F8
      Column(18)      =   "FPOWizDev.frx":1B68
      Column(19)      =   "FPOWizDev.frx":1CC0
      Column(20)      =   "FPOWizDev.frx":1E84
      SortKeysCount   =   2
      SortKey(1)      =   "FPOWizDev.frx":2060
      SortKey(2)      =   "FPOWizDev.frx":20C8
      FmtConditionsCount=   4
      FmtCondition(1) =   "FPOWizDev.frx":2130
      FmtCondition(2) =   "FPOWizDev.frx":2288
      FmtCondition(3) =   "FPOWizDev.frx":23C4
      FmtCondition(4) =   "FPOWizDev.frx":24F8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPOWizDev.frx":262C
      FormatStyle(2)  =   "FPOWizDev.frx":270C
      FormatStyle(3)  =   "FPOWizDev.frx":2844
      FormatStyle(4)  =   "FPOWizDev.frx":28F4
      FormatStyle(5)  =   "FPOWizDev.frx":29A8
      FormatStyle(6)  =   "FPOWizDev.frx":2A80
      ImageCount      =   0
      PrinterProperties=   "FPOWizDev.frx":2B38
   End
   Begin MMRemark.RemarkViewer rvOrderLine 
      Height          =   804
      Left            =   8160
      TabIndex        =   23
      Top             =   120
      Width           =   804
      _ExtentX        =   1429
      _ExtentY        =   1429
      ContextID       =   "ViewOrderLine"
      Caption         =   "OrderLine"
   End
   Begin MMRemark.RemarkViewer rvItem 
      Height          =   804
      Left            =   9000
      TabIndex        =   24
      Top             =   120
      Width           =   804
      _ExtentX        =   1429
      _ExtentY        =   1429
      ContextID       =   "PurchItem"
      Caption         =   "Item"
   End
   Begin MMRemark.RemarkViewer rvVendor 
      Height          =   804
      Left            =   9840
      TabIndex        =   25
      Top             =   120
      Width           =   804
      _ExtentX        =   1429
      _ExtentY        =   1429
      ContextID       =   "ViewVendor"
      Caption         =   "Vendor"
   End
   Begin VB.Menu mnugdxLinesPopup 
      Caption         =   "gdxLinesPopUp"
      Begin VB.Menu mnuSortByType 
         Caption         =   "Order by Color"
      End
      Begin VB.Menu mnugdxLinesFont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnugdxLinesAutofit 
         Caption         =   "Autofit"
      End
      Begin VB.Menu mnugdxLinesSave 
         Caption         =   "Save Layout"
      End
      Begin VB.Menu mnugdxLinesPrint 
         Caption         =   "Print"
      End
   End
End
Attribute VB_Name = "FPOWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lMinWidth = 12200
Private Const k_lMinHeight = 7400

'cache this to support up/down arrow movement
Private m_lastcol As Integer

Private WithEvents m_gwLines As GridEXWrapper
Attribute m_gwLines.VB_VarHelpID = -1

'increment this number when changes are made to gdxLinesPopUp
'Private Const k_iPOWizGridRev = 1

Private m_rstLines As ADODB.Recordset

Private m_lVendKey As Long
Private m_sUserID As String
Private m_lWhseKey As Long

Private m_bDirty As Boolean

Private m_dLineTotal As Double


'*****************************************************************
' Standard MDI child Properties & Methods
'*****************************************************************

Private m_lWindowID As Long

Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property

Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property

Public Sub SetCaption(ByRef i_sTitle As String)
    Me.Caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub

Public Sub DoShowHelp()
    ShowHelp "FPOWizard"
End Sub


'********************************************************************
' Form Events
'********************************************************************

Private Sub Form_Load()
    Set m_gwLines = New GridEXWrapper
    m_gwLines.Grid = gdxLines
    
    Call m_gwLines.InitGridLayout(GetUserKey, g_POWizGridRev)
    
    'add 3 more panels to the statusbar
    StatusBar.Panels.Add
    StatusBar.Panels.Add
    StatusBar.Panels.Add
    
    m_bDirty = False
    
    'enable button if more than one Open Trans (based on Vendor and Whse)
    Dim lbVendorHasBT As Boolean
    lbVendorHasBT = VendHasBT()
    'cmdOTWizDev.Enabled = VendHasBT()
    cmdOTWizDev.Enabled = lbVendorHasBT
    If lbVendorHasBT Then cmdOTWizDev.Font.Bold = True
    
    'PRN 534
    cmdShowAllParts.Enabled = True
   
End Sub


Private Sub Form_Activate()
    'update the MDI toolbars
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If m_bDirty Then
        'PRN #579
        'If vbYes = Msg("Would you like to save your changes before closing?", vbYesNo, "Create PO") Then
        If vbYes = msg("Do you want to save your changes to Order Qty and Cost?", vbYesNo, "Create PO") Then
            UpdateDBLines
        End If
    End If

    'update the MDI toolbars
    MDIMain.UnloadTool m_lWindowID
    
    '3/31/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwLines = Nothing

    m_rstLines.Close
    Set m_rstLines = Nothing
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub

    With Me
        '2/25/05 PRN 563
        If .WindowState <> vbMaximized Then
            If .width < k_lMinWidth Then .width = k_lMinWidth
            If .Height < k_lMinHeight Then .Height = k_lMinHeight
        End If
    End With

    frmDetail(0).width = Me.width - 200
    frmDetail(1).width = Me.width - 200

    gdxLines.width = Me.width - 200
    gdxLines.Height = Me.Height - gdxLines.Top - cmdShowAllParts.Height - 850

    cmdShowAllParts.Top = gdxLines.Top + gdxLines.Height + 60
    cmdRemoveUnusedParts.Top = cmdShowAllParts.Top
    
    cmdAddPartNbr.Top = cmdShowAllParts.Top
    cmdAddPartNbr.Left = Me.width - cmdAddPartNbr.width - 300
    txtPartNbr.Top = cmdAddPartNbr.Top
    txtPartNbr.Left = cmdAddPartNbr.Left - txtPartNbr.width - 100
    
    cmdOTWizDev.Top = cmdShowAllParts.Top
    cmdOTWizDev.Left = cmdRemoveUnusedParts.Left + cmdRemoveUnusedParts.width + 60


    'Note: the Align = vbAlignBottom property has a bug in VB runtime when maximizing & restoring
    'So we handle this ourselves
    StatusBar.Top = cmdShowAllParts.Top + cmdShowAllParts.Height + 60
    StatusBar.width = Me.width - 150
    StatusBar.Panels.Item(1).width = StatusBar.width - (1500 + 2500 + 2500)
    StatusBar.Panels.Item(2).width = 1500
    StatusBar.Panels.Item(3).width = 2500
    StatusBar.Panels.Item(4).width = 2500

    rvOrderLine.ZOrder 0
    rvOrderLine.Top = frmDetail(0).Top + 160
    rvOrderLine.Left = frmDetail(0).width - 4 * (60 + rvOrderLine.width) - 180
    rvVendor.ZOrder 0
    rvVendor.Top = rvOrderLine.Top
    rvVendor.Left = rvOrderLine.Left + 2 * (60 + rvOrderLine.width)
    rvItem.ZOrder 0
    rvItem.Top = rvOrderLine.Top
    rvItem.Left = rvOrderLine.Left + (60 + rvOrderLine.width)
    
End Sub


'***********************************************************************
' Public Methods
'***********************************************************************

Public Sub Init(ByVal i_lVendKey As Long, ByVal i_sUserID As String, ByVal i_lWhseKey As Long, _
        ByVal b_ByVendor As Boolean)
    
    'cache the parameters
    m_lVendKey = i_lVendKey
    m_sUserID = Trim(i_sUserID)
    m_lWhseKey = i_lWhseKey
   
    'Fill disconnected rst from tcpPOProformaLn table
    FillRecordset
    
    'Add white line items if loading PO Wiz by Vendor button
    If b_ByVendor Then Call AddParts(i_lVendKey, i_lWhseKey, i_sUserID)
    
    If m_rstLines.EOF Then
        'there are no items for this warehouse, we don't need the form
        If b_ByVendor Then msg "Your inventory contains no parts from this vendor.", vbInformation, "CreatePO"
        Unload Me
    Else
        MDIMain.AddNewWindow Me
        SetCaption "PO Wizard - " & m_rstLines.Fields("VendName").Value & " - " & m_rstLines.Fields("WhseID").Value
        
        If Not b_ByVendor Then DisplaySuggestedOrderTotal
        
        BindRstToGrid
                        
        'this fires off a RowColChange event which in turn calls RefreshDetail
        'this causes the MM popup remark viewer to interfere with grid refresh
        'the m_bInit flag is used to correct this
        'This is required to keep the following line from disrupting the firing of the Form_Activate event
        DoEvents
        rvVendor.OwnerID = VendKeyToID(i_lVendKey)
        
        'PRN #534 - All parts are shown when loading by Vendor
        If b_ByVendor Then cmdShowAllParts.Enabled = False
    End If
    
End Sub


'Create a POWiz for a vendor with no Red/Orange/Yellow items.
'Load all of the parts we buy for this vendor for this warehouse.

'Public Sub LoadVendor(ByVal i_lVendKey As Long, _
'        ByVal i_sUserID As String, _
'        ByVal i_lWhseKey As Long _
')
'    'cache the parameters
'    m_lVendKey = i_lVendKey
'    m_lWhseKey = i_lWhseKey
'    m_sUserID = i_sUserID
'
'    'Fill disconnected rst from tcpPOProformaLn table (recordcount will be zero)
'    FillRecordset
'
'    'Add to the ProformaLn table all inventory items that are not already there.
'''    Set m_rstLines = CallSP("spcpcPOAddLines", _
'''        "@i_VendKey", i_lVendKey, _
'''        "@i_WhseKey", i_lWhseKey, _
'''        "@i_UserID", i_sUserID)
'    Call AddParts(i_lVendKey, i_lWhseKey, i_sUserID)
'
'    If m_rstLines.EOF Then
'        'there are no items for this warehouse, we don't need the form
'        Msg "Your inventory contains no parts from this vendor.", vbInformation, "CreatePO"
'        Unload Me
'    Else
'        MDIMain.AddNewWindow Me
'        SetCaption "PO Wizard - " & m_rstLines.Fields("VendName").value & " - " & m_rstLines.Fields("WhseID").value
'
'        BindRstToGrid
'
'        'this fires off a RowColChange event which in turn calls RefreshDetail
'        'this causes the MM popup remark viewer to interfere with grid refresh
'        'the m_bInit flag is used to correct this
'        'This is required to keep the following line from disrupting the firing of the Form_Activate event
'        DoEvents
'        rvVendor.OwnerID = VendKeyToID(i_lVendKey)
'        cmdShowAllParts.Enabled = False
'    End If
'
'End Sub


'***********************************************************************
' Private Methods
'***********************************************************************

Private Function VendHasBT()

    Dim lsSql As String
    Dim lors As ADODB.Recordset
    
    VendHasBT = False
    
    lsSql = "SELECT COUNT(dbo.tpoPOLine.POLineKey) AS OpenPOLines FROM dbo.tpoPurchOrder INNER JOIN " & _
        "dbo.tpoPOLine ON dbo.tpoPurchOrder.POKey = dbo.tpoPOLine.POKey INNER JOIN " & _
        "dbo.tapVendor ON dbo.tpoPurchOrder.VendKey = dbo.tapVendor.VendKey INNER JOIN " & _
        "dbo.timWarehouse ON dbo.tapVendor.VendDBA = dbo.timWarehouse.WhseID INNER JOIN " & _
        "dbo.tcpWhseItemVend ON dbo.tpoPOLine.ItemKey = dbo.tcpWhseItemVend.ItemKey AND " & _
        "dbo.timWarehouse.WhseKey = dbo.tcpWhseItemVend.WhseKey INNER JOIN " & _
        "dbo.tpoPOLineDist ON dbo.tpoPOLine.POLineKey = dbo.tpoPOLineDist.POLineKey " & _
        "Where (dbo.tpoPurchOrder.Status = 1) And (dbo.tpoPOLine.Status = 1) " & _
        "And (dbo.timWarehouse.WhseKey = " & m_lWhseKey & ") " & _
        "And (dbo.tcpWhseItemVend.VendKey = " & m_lVendKey & ") " & _
        "AND (dbo.tpoPOLineDist.QtyOrd > dbo.tpoPOLineDist.QtyRcvd)"
    
    Set lors = New ADODB.Recordset
    lors.Source = lsSql
    Set lors.ActiveConnection = g_DB.Connection
    lors.Open
    
    If lors!OpenPOLines > 0 Then VendHasBT = True
    
    Set lors = Nothing
End Function


'**************************************************************************
' Button Handlers
'**************************************************************************

Private Sub cmdOTWizDev_Click()
    SetWaitCursor True
    
    ' New logic to call FViewer - uses: "Crystal Reports 8.5 ActiveX Designer Run Time Library"
    Dim oFrm As FViewer
    Set oFrm = New FViewer
    Call oFrm.ParamAdd(1, "ShippedFromLocation", Trim(m_rstLines.Fields("WhseID").Value))
    Call oFrm.ParamAdd(2, "VendorKey", m_lVendKey)
    Call oFrm.ViewReportByType("Transfer Status Param")
    Set oFrm = Nothing

    SetWaitCursor False
End Sub


Private Sub cmdShowAllParts_Click()
    'PRN 534
    cmdShowAllParts.Enabled = False
    AddParts m_lVendKey, m_lWhseKey, m_sUserID
End Sub


Private Sub cmdRemoveUnusedParts_Click()
    'PRN 534
    cmdShowAllParts.Enabled = True
    DiscardUnusedParts m_lVendKey
End Sub


Private Sub cmdAddPartNbr_Click()
    Dim oCmd As ADODB.Command
    Dim ItemID As String
    Dim ItemKey As Variant
    Dim ItemDesc As String
    
    ItemID = Trim(txtPartNbr.text)
    If Len(ItemID) > 0 Then

'should I also check to see if the item is in the inventory of the warehouse
'placing the order?

        'lookup the itemid in timItem
        Set oCmd = CreateCommandSP("spcpcPOGetItemInfo")
        With oCmd
            .Parameters("@_iItemID").Value = ItemID
            .Execute
            ItemKey = .Parameters("@_oItemKey").Value
            
            'if found
            If Not IsNull(ItemKey) Then
                ItemDesc = .Parameters("@_oItemDescr").Value
            '   insert into recordset
            '   refresh the grid
            
                'm_rstLines.AddNew Array("ItemKey", "ItemID", "Descr", "ItemQty", "QtyToOrder"), Array(ItemKey, ItemID, ItemDesc, 0, 0)
                
                m_rstLines.AddNew _
                    Array("CreateDate", _
                        "Descr", _
                        "LineStatus", _
                        "IsSPO", _
                        "ItemID", _
                        "ItemKey", _
                        "ItemQty", _
                        "QtyToOrder", _
                        "UnitCost", _
                        "ExtCost", _
                        "UserID", _
                        "VendKey", _
                        "WhseKey", _
                        "QOH", _
                        "QSO", _
                        "QPO", _
                        "MinStockQty", _
                        "MaxStockQty", "QtySold", "OrderCount"), _
                    Array(Date, ItemDesc, 5, 0, ItemID, ItemKey, _
                        0, 0, 0, 0, m_sUserID, m_lVendKey, m_lWhseKey, 0, 0, 0, 0, 0, 0, 0)

                m_rstLines.Update
                gdxLines.Refetch
                
                MarkGridAsDirty
            Else
                msg "Item not found.", vbExclamation, "PO Wizard"
            End If
        End With
    End If
End Sub


Private Sub lblQPO_Click()
    'PRN #525 - new PO Orders command button
    Dim oFrm As FPOWizDrillDown
    
    Set oFrm = New FPOWizDrillDown
    Call oFrm.ShowOrders(m_gwLines.Value("ItemID"), m_gwLines.Value("WhseKey"), 1)
    Set oFrm = Nothing
End Sub

Private Sub lblQSO_Click()
  'PRN #525 - new PO Orders command button
    Dim oFrm As FPOWizDrillDown
    
    Set oFrm = New FPOWizDrillDown
    Call oFrm.ShowOrders(m_gwLines.Value("ItemID"), m_gwLines.Value("WhseKey"), 0)
    Set oFrm = Nothing
End Sub

Private Sub FillRecordset()
    Dim sSQL As String
       
    sSQL = "SELECT * FROM tcpPOProformaLn " _
        & "WHERE tcpPOProformaLn.VendKey = " & m_lVendKey & " AND " _
        & "tcpPOProformaLn.UserID = '" & m_sUserID & "' AND " _
        & "WhseKey=" & m_lWhseKey _
        & " ORDER BY LineStatus ASC, IsSPO DESC, ItemQty DESC"
        
    Set m_rstLines = LoadDiscRst(sSQL, , adLockBatchOptimistic)
End Sub


Public Sub BindRstToGrid()
    With gdxLines
        .HoldSortSettings = True
        .HoldFields
        Set gdxLines.ADORecordset = m_rstLines
        gdxLines.Refresh
    End With
    
    RefreshDetail
    DisplayOrderTotal
    gdxLines.SetFocus
End Sub


'reconnect the recordset and update the SQL table
Private Sub UpdateDBLines()
    
    On Error GoTo EH
    
    SetWaitCursor True
    
    'PRN #539 - 2/4/2005
    'if the last edited cell was not tabbed out of, and the Save button was pressed
    'none of the 3 places we update the underlying disconnected recordset get called
    '1) rowchange (automatic)
    '2) colchange (aftercolupdate event)
    '3) lostfocus (not firing because we've clicked on the MDI form)
    'We need to write the last change to the recordset.
    gdxLines.Update
    
    DiscardUnusedParts m_lVendKey
    
    With m_rstLines
        'PRN #541 - if m_bDirty = true then set dirty flag in rst.
        .Fields("DirtyData") = 1
    
        'Save all changes to tcpPOProformaLn
        .ActiveConnection = g_DB.Connection
        .UpdateBatch
        .ActiveConnection = Nothing
    End With
    
    m_bDirty = False
    MDIMain.UpdateToolbarStatus
    
    SetWaitCursor False
    Exit Sub
EH:
    SetWaitCursor False
    msg Err.Number & " " & Err.Description, vbCritical, "UpdateDBLines"
End Sub


Private Sub RefreshDetail()
    
    If m_gwLines.Value("IsSPO") Then
        frmDetail(0).Visible = True
        frmDetail(1).Visible = False
        lblOrderNbr.Caption = m_gwLines.Value("OPKey") & " / " & StripLeadingZeros(m_gwLines.Value("SOID"))
        lblShipMethod.Caption = m_gwLines.Value("ShipMethID")
        lblCSR.Caption = IIf(IsNull(m_gwLines.Value("CSR")), "", m_gwLines.Value("CSR"))
        lblModelNbr.Caption = m_gwLines.Value("ModelNbr")
        lblSerialNbr.Caption = m_gwLines.Value("SerialNbr")
        If IsDate(m_gwLines.Value("CommitDate")) Then
            lblCommitDate.Caption = Format(m_gwLines.Value("CommitDate"), "mm/dd/yy")
        Else
            lblCommitDate.Caption = ""
        End If
    Else
        frmDetail(0).Visible = False
        frmDetail(1).Visible = True
        lblPartNbr.Caption = m_gwLines.Value("ItemID")
        lblQSO.Caption = IIf(IsNull(m_gwLines.Value("QSO")), vbNullString, m_gwLines.Value("QSO"))
        lblQPO.Caption = IIf(IsNull(m_gwLines.Value("QPO")), vbNullString, m_gwLines.Value("QPO"))
    End If
        
    If m_gwLines.Value("IsSPO") Then
        rvOrderLine.OwnerID = m_gwLines.Value("OPLineKey")
    Else
        'SPOs don't have Item remarks (they're non-inventory items)
        rvItem.OwnerID = m_gwLines.Value("ItemID")
        rvOrderLine.OwnerID = vbNullString
    End If

End Sub


'**************************************************************************
' Grid Event Handlers
'**************************************************************************

Private Sub gdxLines_Change()
    'PRN #541 -  2/23/05 - get key of current column
        'only mark dirty flag when Order Qty or Cost changes
    Dim colCurrent As JSColumn
    
    Set colCurrent = gdxLines.Columns.ItemByPosition(gdxLines.col)
    If colCurrent.Key = "QtyToOrder" Or colCurrent.Key = "UnitCost" Then
        MarkGridAsDirty
    End If
End Sub


Private Sub MarkGridAsDirty()
    m_bDirty = True
    MDIMain.UpdateToolbarStatus
End Sub


Private Sub gdxLines_DblClick()
    If Not IsEmpty(gdxLines.Value(9)) Then ' ItemKey
        Dim oFrm As Form
        Set oFrm = New FInventoryHistory
        If oFrm.ShowHistory(m_lWhseKey, gdxLines.Value(9)) Then
            MDIMain.AddNewWindow oFrm
            oFrm.SetCaption "Inventory History (" & WhseKeyToID(m_lWhseKey) & "): " & gdxLines.Value(3)  ' ItemID
        Else
            Unload oFrm
            Set oFrm = Nothing
        End If
    End If
End Sub

'Private Sub gdxLines_LostFocus()
'    'if any un-updated change exists, save it
'    gdxLines.Update
'End Sub


Private Sub gdxLines_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Dim colCurrent As JSColumn

    'If an up or down arrow was pressed, row index increments or decrements
    'but the Col property is set to 0 and the LastCol parameter is -1.
    'Restore the Col to the last cached value.
    If LastCol = -1 Then
        gdxLines.col = m_lastcol
    End If
    
    'if GridEX isn't in edit mode and
    'there is a current column (Col is the active column in the grid)
    If gdxLines.EditMode = jgexEditModeOff And gdxLines.col <> 0 Then

        'enter edit mode
        gdxLines.EditMode = jgexEditModeOn
        
        'Get the current column.
        'When a column changes its position the Index property for that column remains the same
        'and only its ColPosition property changes.
        'For that reason, regardless of how many times a column is moved, it always appears in
        'the same position of the collection. However, there are times when you need to access
        'a column by its positional index, rather than its collection index. In those cases you
        'need to use this method instead of the Item property.
        Set colCurrent = gdxLines.Columns.ItemByPosition(gdxLines.col)
        
        'select all the text in the cell
        gdxLines.SelStart = 0
        gdxLines.SelLength = Len(gdxLines.Value(colCurrent.Index))
    Else
        'Debug.Print "There is no current column."
    End If

    StatusBar.Panels(2).text = "Row " & gdxLines.Row & " of " & CStr(m_rstLines.RecordCount)

    RefreshDetail
    
    m_lastcol = gdxLines.col
End Sub


Private Sub gdxLines_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    gdxLines.PrinterProperties.FooterString(jgexHFRight) = "Page " & PageNumber & " of " & nPages
End Sub


'Here's where we wire in the PopUp menu for the grid

Private Sub gdxLines_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnugdxLinesPopup
    End If
End Sub


Private Sub gdxLines_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX20.JSRetBoolean)
    'Error handling for updating min/max
    If ColIndex = 16 Or ColIndex = 17 Then
        If Len(gdxLines.Value(ColIndex)) < 1 Then
            Cancel = True
            MsgBox "Please enter a numeric value.", vbInformation, "Min/Max Value"
            Exit Sub
        ElseIf Not IsNumeric(gdxLines.Value(ColIndex)) Then
            Cancel = True
            MsgBox "Please enter a numeric value.", vbInformation, "Min/Max Value"
            Exit Sub
        ElseIf CLng(gdxLines.Value(16)) > CLng(gdxLines.Value(17)) Then
            Cancel = True
            MsgBox "'Min' quantity must be less than 'Max' quantity.", vbInformation, "Min/Max Value"
            Exit Sub
        End If
    End If
End Sub


' This is fired by leaving the edited column (an arrow key or a tab)
Private Sub gdxLines_AfterColUpdate(ByVal ColIndex As Integer)
    
    If ColIndex = 16 Or ColIndex = 17 Then
        'update min/max in timInventory
        CallSP "spcpcPOUpdateMinMax", _
                "@_iWhseKey", gdxLines.Value(11), _
                "@_iItemKey", gdxLines.Value(9), _
                "@_iMinQty", CLng(gdxLines.Value(16)), _
                "@_iMaxQty", CLng(gdxLines.Value(17))
    End If
    

    'By default, the Janus grid does not update the recordset
    'until you move off the record.  This is smart in the
    'general case because it allows the user to edit multiple
    'fields quickly then update the database in one hit.
    'In our case, however, we want to have the total order line
    'immediately reflect the change we made.  Fortunately,
    'GridEX provides a method for this purpose.

    'Commits changes made to the current row, writes to the database
    'and re-positions the record if necessary.
    gdxLines.Update
        
    gdxLines.ADORecordset.Fields("ExtCost").Value = m_gwLines.Value("QtyToOrder") * m_gwLines.Value("UnitCost")
    
    gdxLines.Update

    DisplayOrderTotal   'Update the total

    'PRN #299 - This will refresh the OA buttons based on the current
        'state after the col has been updated.  This will turn on/off the
        'Commit button, based on the linetotal amount.
    MDIMain.UpdateToolbarStatus

    'Since we are using a disconnected recordset, nothing is really
    'committed to the database yet.
'***    m_bDirty = True

End Sub


Public Sub AddParts(i_lVendorKey As Long, i_lWhseKey As Long, i_sUserID As String)
    Dim sSQL As String
    Dim orst As ADODB.Recordset

    SetWaitCursor True
    
    Set orst = CallSP("spcpcPOAddLines", _
        "@i_VendKey", i_lVendorKey, _
        "@i_WhseKey", i_lWhseKey, _
        "@i_UserID", i_sUserID)
        
    With orst
        Do While Not .EOF
        'PRN 493 added isnull logic to descr field below
            m_rstLines.AddNew _
                Array("CreateDate", "Descr", "LineStatus", "IsSPO", "ItemID", "ItemKey", "ItemQty", "QtyToOrder", _
                    "UnitCost", "ExtCost", "UserID", "VendID", "VendKey", "VendName", "WhseID", "WhseKey", _
                    "VendItemID", "QOH", "QSO", "QPO", "MinStockQty", "MaxStockQty", "QtySold", "OrderCount"), _
                Array(.Fields("CreateDate").Value, _
                    IIf(IsNull(.Fields("Descr").Value), "", .Fields("Descr").Value), _
                    .Fields("LineStatus").Value, _
                    .Fields("IsSPO").Value, Trim(.Fields("ItemID").Value), .Fields("ItemKey").Value, _
                    .Fields("ItemQty").Value, .Fields("QtyToOrder").Value, .Fields("UnitCost").Value, .Fields("ExtCost").Value, _
                    .Fields("UserID").Value, .Fields("VendID").Value, .Fields("VendKey").Value, _
                    Trim(.Fields("VendName").Value), .Fields("WhseID").Value, .Fields("WhseKey").Value, _
                    .Fields("VendItemID").Value, .Fields("QOH").Value, .Fields("QSO").Value, _
                    .Fields("QPO").Value, .Fields("MinStockQty").Value, .Fields("MaxStockQty").Value, IIf(IsNull(.Fields("QtySold").Value), 0, .Fields("QtySold").Value), IIf(IsNull(.Fields("OrderCount").Value), 0, .Fields("OrderCount").Value))
            orst.MoveNext
        Loop
    End With
    With gdxLines
        .HoldSortSettings = True '= False
        .HoldFields
        .Refetch
    End With
    orst.Close
    Set orst = Nothing
    
    SetWaitCursor False
End Sub


Private Sub DiscardUnusedParts(ByVal i_lVendKey As Long)
    SetWaitCursor True
    With m_rstLines
        .MoveFirst
        Do While Not .EOF
            If .Fields("LineStatus").Value = kiWhite And .Fields("QtyToOrder").Value = 0 Then
                .Delete
            End If
            .MoveNext
        Loop
    End With
    gdxLines.Refetch
    
    SetWaitCursor False
End Sub


Private Sub DisplayOrderTotal()
    m_dLineTotal = 0
    With m_rstLines
        If Not .BOF Then .MoveFirst
        Do While Not .EOF
            m_dLineTotal = m_dLineTotal + .Fields("ExtCost").Value
            .MoveNext
        Loop
        .MoveFirst      'reset
    End With
    StatusBar.Panels(4).text = "Order Total:  $" & Format$(m_dLineTotal, g_MoneyMask)
End Sub

Private Sub DisplaySuggestedOrderTotal()
    Dim dLineTotal As Double
    dLineTotal = 0
    With m_rstLines
        If Not .BOF Then .MoveFirst
        Do While Not .EOF
            If Not IsNull(.Fields("ItemQty").Value) And Not IsNull(.Fields("UnitCost").Value) Then
                dLineTotal = dLineTotal + .Fields("ItemQty").Value * .Fields("UnitCost").Value
            End If
            .MoveNext
        Loop
        .MoveFirst      'reset
    End With
    StatusBar.Panels(3).text = "Sugg. Order Total:  $" & Format$(dLineTotal, g_MoneyMask)
End Sub


'******************************************************************************************
' MDI ToolBar command functions
'******************************************************************************************

Public Function CommitButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean
    Dim oFrm As FPOCommit
    
    If i_bDoIt Then
        'PRN392 SPO order check
        With m_rstLines
            .MoveFirst
            Do While Not .EOF
                If .Fields("IsSPO") = 1 Then 'SPO item
                    If .Fields("QtyToOrder") <> .Fields("ItemQty") Then 'Qty order is not equal to suggested
                        If vbCancel = MsgBox("One or more SPO quantities is not equal suggested quantity.", vbOKCancel) Then
                            Exit Function
                        Else
                            Exit Do
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        Set oFrm = New FPOCommit
        oFrm.Init m_rstLines, m_lVendKey, m_sUserID, m_lWhseKey
        If oFrm.Cancel Then
            Unload oFrm
            Set oFrm = Nothing
        Else
            Unload oFrm
            Set oFrm = Nothing
            
            'PRN #299
            'If commit is not cancelled, set m_bDirty to False
            'we are not in a dirty state anymore, since the order was committed.
            m_bDirty = False
            'Delect rec(s) from tcpPOProformaLn table
            CallSP "spcpcPODeleteProformaLn", _
                "@_iUserID", m_sUserID, _
                "@_iVendKey", m_lVendKey, _
                "@_iWhseKey", m_lWhseKey
            
            Unload Me
        End If
        Else
        'tell MDIMain to enable the commit button if there is anything on the order
        CommitButton = (m_dLineTotal > 0)
    End If
End Function


Public Function SaveButton(Optional ByVal i_bDoIt As Boolean = True, Optional ByVal i_bClose As Boolean = False) As Boolean
    If i_bDoIt Then
        UpdateDBLines
    Else
        'tell MDIMain to enable/disable Save based on this form's dirty status
        SaveButton = m_bDirty
    End If
End Function


Public Function CancelButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean
    Dim iResponse As Integer

    If i_bDoIt Then
        If m_bDirty Then
            iResponse = msg("Would you like to save your changes before closing?", vbYesNoCancel, "Create PO")
            Select Case iResponse
                Case vbYes:
                    UpdateDBLines
                    Unload Me
                Case vbNo:
                    m_bDirty = False    'this prevents the Form_Unload event handler from prompting a second time
                    Unload Me
                Case vbCancel:
                    Exit Function
            End Select
        Else
            Unload Me
        End If
    Else
        CancelButton = True
    End If
End Function


'******************************************************************************************
' Pop-Up Menu code
'******************************************************************************************

Private Sub mnugdxLinesFont_click()
    ChangeGridFont gdxLines
End Sub

    
Private Sub mnugdxLinesAutofit_Click()
    '12/2004 - smr
    Call m_gwLines.GridAutoFit
End Sub


Private Sub mnugdxLinesPrint_Click()
    gdxLines.PrinterProperties.Orientation = jgexPPLandscape
    gdxLines.PrintGrid True
End Sub


'restore the intial default sort order in the grid
Private Sub mnuSortByType_Click()
    gdxLines.HoldSortSettings = False
    gdxLines.HoldFields
    m_rstLines.Sort = "LineStatus ASC, IsSPO DESC, ItemQty DESC"
    gdxLines.Refetch
End Sub


Private Sub mnugdxLinesSave_Click()
    '12/2004 - smr
    SetWaitCursor True
    Call m_gwLines.GridSaveLayout(GetUserKey)
    SetWaitCursor False
End Sub



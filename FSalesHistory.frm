VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FSalesHistory 
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   10980
   Begin VB.ComboBox cboVend_SalesHist 
      Height          =   315
      Left            =   6180
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   3312
   End
   Begin VB.TextBox txtItemID 
      Height          =   315
      Left            =   660
      MaxLength       =   30
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cboWhse_SalesHist 
      Height          =   315
      Left            =   3540
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   312
      Left            =   4440
      TabIndex        =   1
      Top             =   6480
      Width           =   1152
   End
   Begin GridEX20.GridEX gdxLines 
      Height          =   3195
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5636
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      ShowToolTips    =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      SelectionStyle  =   1
      Options         =   2
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   9
      Column(1)       =   "FSalesHistory.frx":0000
      Column(2)       =   "FSalesHistory.frx":0148
      Column(3)       =   "FSalesHistory.frx":0294
      Column(4)       =   "FSalesHistory.frx":0474
      Column(5)       =   "FSalesHistory.frx":0624
      Column(6)       =   "FSalesHistory.frx":0774
      Column(7)       =   "FSalesHistory.frx":095C
      Column(8)       =   "FSalesHistory.frx":0ADC
      Column(9)       =   "FSalesHistory.frx":0C6C
      SortKeysCount   =   2
      SortKey(1)      =   "FSalesHistory.frx":0DFC
      SortKey(2)      =   "FSalesHistory.frx":0E64
      GroupConditionCountTitle=   ""
      FormatStylesCount=   6
      FormatStyle(1)  =   "FSalesHistory.frx":0ECC
      FormatStyle(2)  =   "FSalesHistory.frx":0FAC
      FormatStyle(3)  =   "FSalesHistory.frx":10E4
      FormatStyle(4)  =   "FSalesHistory.frx":1194
      FormatStyle(5)  =   "FSalesHistory.frx":1248
      FormatStyle(6)  =   "FSalesHistory.frx":1320
      ImageCount      =   0
      PrinterProperties=   "FSalesHistory.frx":13D8
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Vendor"
      Height          =   195
      Left            =   5400
      TabIndex        =   7
      Top             =   180
      Width           =   675
   End
   Begin VB.Label lblVendItem 
      Alignment       =   1  'Right Justify
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lblVendWhse 
      Alignment       =   1  'Right Justify
      Caption         =   "Warehouse"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   975
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
Attribute VB_Name = "FSalesHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lMinWidth = 12200
Private Const k_lMinHeight = 7400

Private WithEvents m_gwLines As GridEXWrapper
Attribute m_gwLines.VB_VarHelpID = -1

Private m_rstLines As ADODB.Recordset


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

Private Sub cmdGo_Click()
    If cboVend_SalesHist.ListIndex = 0 Then
        Set m_rstLines = CallSP("cpoaGetSalesByPartAndVendor", _
            "@WhseKey", cboWhse_SalesHist.ItemData(cboWhse_SalesHist.ListIndex), _
            "@ItemID", txtItemID.Text)
    Else
        Set m_rstLines = CallSP("cpoaGetSalesByPartAndVendor", _
            "@WhseKey", cboWhse_SalesHist.ItemData(cboWhse_SalesHist.ListIndex), _
            "@ItemID", txtItemID.Text, _
            "@VendKey", cboVend_SalesHist.ItemData(cboVend_SalesHist.ListIndex))
    End If
    
    If m_rstLines.EOF Then
        'there are no items for this warehouse, we don't need the form
        Msg "No item found.", vbInformation, "CreatePO"
    Else
        BindRstToGrid
        'gdxLines.Refetch
        'GetGridLayout GetUserKey, gdxLines
        
        'this fires off a RowColChange event which in turn calls RefreshDetail
        'this causes the MM popup remark viewer to interfere with grid refresh
        'the m_bInit flag is used to correct this
        'This is required to keep the following line from disrupting the firing of the Form_Activate event
        DoEvents
    End If

'    With gdxLines
'        .HoldSortSettings = True
'        .HoldFields
'        Set gdxLines.ADORecordset = m_rstLines
'        gdxLines.Refresh
'    End With


End Sub

'********************************************************************
' Form Events
'********************************************************************

Private Sub Form_Load()
    LoadCombo cboWhse_SalesHist, g_rstWhses, "WhseID", "WhseKey"
    'g_rstVendors.Sort = "VendID"      'change sort order
    LoadCombo cboVend_SalesHist, g_rstVendors, "VendName", "VendKey"
    cboVend_SalesHist.AddItem "<ALL>", 0
    cboVend_SalesHist.ListIndex = 0
    
    'g_rstVendors.Sort = "VendName"    'restore
    
    Set m_gwLines = New GridEXWrapper
    m_gwLines.Grid = gdxLines
End Sub


Private Sub Form_Activate()
    'update the MDI toolbars
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '3/31/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwLines = Nothing
    
    'update the MDI toolbars
    MDIMain.UnloadTool m_lWindowID
    
    m_rstLines.Close
    Set m_rstLines = Nothing
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub

    With Me
        If .Width < k_lMinWidth Then .Width = k_lMinWidth
        If .Height < k_lMinHeight Then .Height = k_lMinHeight
    End With


    gdxLines.Width = Me.Width - 200
    gdxLines.Height = Me.Height - gdxLines.Top - cmdGo.Height - 850

End Sub


'***********************************************************************
' Public Methods
'***********************************************************************

Public Sub Init(ByVal i_lVendKey As Long, ByVal i_sUserID As String, ByVal i_lWhseKey As String)
    If m_rstLines.EOF Then
        'there are no items for this warehouse, we don't need the form
        Unload Me
    Else
        MDIMain.AddNewWindow Me
        SetCaption "PO Wizard - " & m_rstLines.Fields("VendName").value & " - " & m_rstLines.Fields("WhseID").value
        
        BindRstToGrid
        GetGridLayout GetUserKey, gdxLines
        
        'this fires off a RowColChange event which in turn calls RefreshDetail
        'this causes the MM popup remark viewer to interfere with grid refresh
        'the m_bInit flag is used to correct this
        'This is required to keep the following line from disrupting the firing of the Form_Activate event
        DoEvents
    End If
    
End Sub


'**************************************************************************
' Button Handlers
'**************************************************************************

Public Sub BindRstToGrid()
    With gdxLines
        .HoldSortSettings = True
        .HoldFields
        Set gdxLines.ADORecordset = m_rstLines
        gdxLines.Refresh
    End With
    
    gdxLines.SetFocus
End Sub


'**************************************************************************
' Grid Event Handlers
'**************************************************************************

Private Sub gdxLines_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    gdxLines.PrinterProperties.FooterString(jgexHFRight) = "Page " & PageNumber & " of " & nPages
End Sub


'Here's where we wire in the PopUp menu for the grid

Private Sub gdxLines_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnugdxLinesPopup
    End If
End Sub


'******************************************************************************************
' MDI ToolBar command functions
'******************************************************************************************

'******************************************************************************************
' Pop-Up Menu code
'******************************************************************************************

Private Sub mnugdxLinesFont_click()
    ChangeGridFont gdxLines
End Sub

    
Private Sub mnugdxLinesAutofit_Click()
    Dim oCol As JSColumn
    
    For Each oCol In gdxLines.Columns
        oCol.AutoSize
    Next
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
    SetWaitCursor True
    CallSP "spcpcDeleteUserPrefs", "@_iUserKey", GetUserKey, "@_iGridName", "gdxLines"
    CallSP "spcpcSaveUserPrefs", "@_iUserKey", GetUserKey, "@_iGridName", "gdxLines", "@_iLayoutString", gdxLines.LayoutString(False)
    SetWaitCursor False
End Sub



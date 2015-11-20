VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "MMRemark.ocx"
Begin VB.Form FOnPurchOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Research Open Purchase Orders"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   9705
   Begin MMRemark.RemarkViewer rvPurchOrd 
      Height          =   804
      Left            =   8460
      TabIndex        =   5
      Top             =   60
      Width           =   804
      _ExtentX        =   1429
      _ExtentY        =   1429
      ContextID       =   "ViewPO"
      Caption         =   "PO Remarks"
   End
   Begin GridEX20.GridEX gdxPurchOrds 
      Height          =   2892
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   9492
      _ExtentX        =   16748
      _ExtentY        =   5106
      Version         =   "2.0"
      ShowToolTips    =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      Options         =   8
      AllowColumnDrag =   0   'False
      RecordsetType   =   1
      RecordSource    =   $"FOnPurchOrder.frx":0000
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      ColumnsCount    =   11
      Column(1)       =   "FOnPurchOrder.frx":0006
      Column(2)       =   "FOnPurchOrder.frx":012A
      Column(3)       =   "FOnPurchOrder.frx":025A
      Column(4)       =   "FOnPurchOrder.frx":0396
      Column(5)       =   "FOnPurchOrder.frx":0502
      Column(6)       =   "FOnPurchOrder.frx":066E
      Column(7)       =   "FOnPurchOrder.frx":07DA
      Column(8)       =   "FOnPurchOrder.frx":08FE
      Column(9)       =   "FOnPurchOrder.frx":0A86
      Column(10)      =   "FOnPurchOrder.frx":0C02
      Column(11)      =   "FOnPurchOrder.frx":0D7E
      FormatStylesCount=   6
      FormatStyle(1)  =   "FOnPurchOrder.frx":0EBA
      FormatStyle(2)  =   "FOnPurchOrder.frx":0F9A
      FormatStyle(3)  =   "FOnPurchOrder.frx":10D2
      FormatStyle(4)  =   "FOnPurchOrder.frx":1182
      FormatStyle(5)  =   "FOnPurchOrder.frx":1236
      FormatStyle(6)  =   "FOnPurchOrder.frx":130E
      ImageCount      =   0
      PrinterProperties=   "FOnPurchOrder.frx":13C6
   End
   Begin VB.TextBox txtPartNbr 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtDescr 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Part Number"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "FOnPurchOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWindowID As Long

Private WithEvents m_gwPurchOrders As GridEXWrapper
Attribute m_gwPurchOrders.VB_VarHelpID = -1


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


Private Sub Form_Load()
    Set m_gwPurchOrders = New GridEXWrapper
    m_gwPurchOrders.Grid = gdxPurchOrds
End Sub

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Reference: http://www.devx.com/vb2themax/Tip/18461
    Set m_gwPurchOrders = Nothing
    MDIMain.UnloadTool m_lWindowID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        DoShowHelp
    Else
        MDIMain.GlobalKeyDownProcessing KeyCode, Shift
    End If
End Sub


Public Sub DoShowHelp()
    ShowHelp "ResearchPO", True
End Sub


Public Sub ShowPurchaseOrders( _
        ByVal i_sPartNbr As String, _
        ByVal i_sPartDescr As String, _
        ByVal i_lItemKey As Long, _
        Optional i_lWhseKey As Long _
)
    Dim sSQL As String
    Dim lWhseKey As Long

    MDIMain.AddNewWindow Me
    SetCaption "Open Purchase Orders on " & i_sPartNbr
    
    txtPartNbr.text = i_sPartNbr
    txtDescr.text = i_sPartDescr
    
    If i_lWhseKey = 0 Then
        Set gdxPurchOrds.ADORecordset = CallSP("spcpcGetOpenPOsByKey", "@_iItemKey", i_lItemKey)
    Else
        Set gdxPurchOrds.ADORecordset = CallSP("spcpcGetOpenPOsByKey", "@_iItemKey", i_lItemKey, "@_iWhseKey", i_lWhseKey)
    End If
    
    Show
    
End Sub


Private Sub gdxPurchOrds_SelectionChange()
    rvPurchOrd.OwnerID = m_gwPurchOrders.value("PONumber")
End Sub

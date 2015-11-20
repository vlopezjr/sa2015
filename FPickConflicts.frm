VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FPickConflicts 
   Caption         =   "x orders in Pick Conflict"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   10185
   Begin VB.ComboBox cboWhse 
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   9060
      Width           =   1215
   End
   Begin VB.CommandButton cmdToggleView 
      Caption         =   "<>"
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   0
      Width           =   435
   End
   Begin VB.Timer tmrScheduler 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   4500
      Top             =   9060
   End
   Begin VB.ListBox lstEntries 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   9915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Quantities"
      Height          =   675
      Left            =   180
      TabIndex        =   7
      Top             =   6180
      Width           =   9915
      Begin VB.CommandButton cmdInvFinder 
         Caption         =   "InvFinder"
         Height          =   315
         Left            =   8640
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAvail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7440
         TabIndex        =   21
         ToolTipText     =   "QtyOnHand - QtyOnPicks - QtyOnPendingShips"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtOnPurchOrders 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4860
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtOnSalesOrders 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtItemId 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblOrderInfo 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   8040
         TabIndex        =   24
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lblOrderInfo 
         Caption         =   "On POs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   18
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lblOrderInfo 
         Caption         =   "On SOs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   2820
         TabIndex        =   12
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "On Hand"
         Height          =   255
         Left            =   5580
         TabIndex        =   11
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Avail"
         Height          =   255
         Left            =   7020
         TabIndex        =   10
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "ItemId"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdViewOrder 
      Caption         =   "Open in OrderPad"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   9060
      Width           =   1455
   End
   Begin VB.CommandButton cmdPickOrder 
      Caption         =   "Pick"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8700
      TabIndex        =   2
      Top             =   9060
      Width           =   1335
   End
   Begin GridEX20.GridEX gdxAllConflictOrders 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   2100
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   2672
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      AutomaticSort   =   -1  'True
      ShowToolTips    =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      RowHeight       =   19
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   10
      Column(1)       =   "FPickConflicts.frx":0000
      Column(2)       =   "FPickConflicts.frx":0134
      Column(3)       =   "FPickConflicts.frx":0268
      Column(4)       =   "FPickConflicts.frx":03F4
      Column(5)       =   "FPickConflicts.frx":0538
      Column(6)       =   "FPickConflicts.frx":0688
      Column(7)       =   "FPickConflicts.frx":07B4
      Column(8)       =   "FPickConflicts.frx":0940
      Column(9)       =   "FPickConflicts.frx":0A8C
      Column(10)      =   "FPickConflicts.frx":0BD8
      SortKeysCount   =   1
      SortKey(1)      =   "FPickConflicts.frx":0D14
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPickConflicts.frx":0D7C
      FormatStyle(2)  =   "FPickConflicts.frx":0EB4
      FormatStyle(3)  =   "FPickConflicts.frx":0F64
      FormatStyle(4)  =   "FPickConflicts.frx":1018
      FormatStyle(5)  =   "FPickConflicts.frx":10F0
      FormatStyle(6)  =   "FPickConflicts.frx":11A8
      ImageCount      =   0
      PrinterProperties=   "FPickConflicts.frx":1288
   End
   Begin GridEX20.GridEX gdxLineItems 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "FPickConflicts.frx":1460
      Column(2)       =   "FPickConflicts.frx":15C4
      Column(3)       =   "FPickConflicts.frx":16E8
      Column(4)       =   "FPickConflicts.frx":1800
      Column(5)       =   "FPickConflicts.frx":1934
      Column(6)       =   "FPickConflicts.frx":1AA0
      Column(7)       =   "FPickConflicts.frx":1C30
      FmtConditionsCount=   1
      FmtCondition(1) =   "FPickConflicts.frx":1D58
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPickConflicts.frx":1E6C
      FormatStyle(2)  =   "FPickConflicts.frx":1F4C
      FormatStyle(3)  =   "FPickConflicts.frx":2084
      FormatStyle(4)  =   "FPickConflicts.frx":2134
      FormatStyle(5)  =   "FPickConflicts.frx":21E8
      FormatStyle(6)  =   "FPickConflicts.frx":22C0
      ImageCount      =   0
      PrinterProperties=   "FPickConflicts.frx":2378
   End
   Begin GridEX20.GridEX gdxInConflictForItem 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   2990
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      ShowToolTips    =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      Options         =   2
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   10
      Column(1)       =   "FPickConflicts.frx":2550
      Column(2)       =   "FPickConflicts.frx":2684
      Column(3)       =   "FPickConflicts.frx":2848
      Column(4)       =   "FPickConflicts.frx":29D0
      Column(5)       =   "FPickConflicts.frx":2B38
      Column(6)       =   "FPickConflicts.frx":2CC0
      Column(7)       =   "FPickConflicts.frx":2E4C
      Column(8)       =   "FPickConflicts.frx":2FA8
      Column(9)       =   "FPickConflicts.frx":310C
      Column(10)      =   "FPickConflicts.frx":3298
      SortKeysCount   =   1
      SortKey(1)      =   "FPickConflicts.frx":340C
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPickConflicts.frx":3474
      FormatStyle(2)  =   "FPickConflicts.frx":3554
      FormatStyle(3)  =   "FPickConflicts.frx":368C
      FormatStyle(4)  =   "FPickConflicts.frx":373C
      FormatStyle(5)  =   "FPickConflicts.frx":37F0
      FormatStyle(6)  =   "FPickConflicts.frx":38C8
      ImageCount      =   0
      PrinterProperties=   "FPickConflicts.frx":3980
   End
   Begin VB.Label lblOrdersInConflictForItem 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6960
      Width           =   5115
   End
   Begin VB.Label Label6 
      Caption         =   "Today's Notifications"
      Height          =   315
      Left            =   180
      TabIndex        =   16
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "Orders in Conflict"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1860
      Width           =   3675
   End
   Begin VB.Label lblLineItems 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   5115
   End
End
Attribute VB_Name = "FPickConflicts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FormWidth = 10305
Private Const FormFullHeight = 9990
Private Const FormShortHeight = 2200

Private m_oOSItemList As OSItemList

Private m_gwAllConflictOrders As GridEXWrapper
Private m_gwInConflictForItem As GridEXWrapper

Private m_bIsLoading As Boolean

Private fullView As Boolean
Private lastNotificationCount As Integer
Private selectedItemKey As Long
Private selectedItemType As Long

'*******************************************************************
'Extended form property & method
'*******************************************************************

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


'*******************************************************************
'Std form events
'*******************************************************************

Private Sub Form_Load()
    Initialize
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.width < FormWidth Then Me.width = FormWidth
    If fullView Then
        If Me.Height < FormFullHeight Then Me.Height = FormFullHeight
    Else
        If Me.Height < FormShortHeight Then Me.Height = FormShortHeight
    End If
    
End Sub

Private Sub Initialize()
    
    fullView = False
    
    m_bIsLoading = True
    g_rstWhses.Filter = "transit = 0"
    LoadCombo cboWhse, g_rstWhses, "WhseID", "WhseKey", User.GetUserWhseKey(GetUserKey(GetUserName))
    g_rstWhses.Filter = adFilterNone
    m_bIsLoading = False
    
    Set m_gwAllConflictOrders = New GridEXWrapper
    m_gwAllConflictOrders.Grid = gdxAllConflictOrders
    
    Set m_gwInConflictForItem = New GridEXWrapper
    m_gwInConflictForItem.Grid = gdxInConflictForItem
    
    'do this to get a count for the caption (also used in RefreshGrids)
    'better replaced with a scalar sproc
    Dim orst As ADODB.Recordset
    Set orst = GetAllOrdersInConflict(cboWhse.ItemData(cboWhse.ListIndex))
    SetCaption orst.RecordCount & " order(s) in Pick Conflict"
    Set orst = Nothing
    
    LoadPickConflictOrdersLog
    tmrScheduler.Enabled = True
    
    Me.width = FormWidth
    Me.Height = FormShortHeight

    Exit Sub
    
Initialize_EH:
    ErrorUI.FatalError "PickMgr.Init", "Pick Conflicts initialization failed."
End Sub

Private Sub gdxInConflictForItem_Click()
    If (CBool(m_gwInConflictForItem.Value("Pickable"))) Then
        cmdPickOrder.Enabled = True
    Else
        cmdPickOrder.Enabled = False
    End If
    
    cmdViewOrder.Enabled = True
End Sub

Private Sub tmrScheduler_Timer()
    LoadPickConflictOrdersLog
End Sub


Private Sub LoadPickConflictOrdersLog()
    Dim orst As ADODB.Recordset
    Set orst = CallSP("spcpcGetConflictOrdersLog", "@WhseKey", User.GetUserWhseKey)
    
    Dim iOpKey As Long
    Dim iSoKey As Long
    Dim iItemKey As Long
    Dim sTranNo As String
    Dim sItemID As String
    Dim iWhseKey As Integer
    Dim sAction As String
    Dim dCreatedDate As Date
    Dim previousNotificationCount
    Dim sMessage As String
    
    previousNotificationCount = lstEntries.ListCount
    
    lstEntries.Clear
    
    With orst
        lastNotificationCount = .RecordCount
        Do While Not .EOF
            iOpKey = .Fields("OpKey").Value
            iSoKey = .Fields("SoKey").Value
            iItemKey = .Fields("SoKey").Value
            sTranNo = .Fields("TranNo").Value
            sItemID = .Fields("ItemId").Value
            iWhseKey = .Fields("WhseKey").Value
            sAction = .Fields("Description").Value
            dCreatedDate = .Fields("CreatedDate").Value

           Dim entry As String

            entry = PadRight(TimeValue(dCreatedDate), 11) & "  " & PadRight(sAction, 9) & " OP-" & iOpKey & ", SO-" & Trim(sTranNo) & ", Item-" & sItemID
            
            lstEntries.AddItem Trim(entry)
            
            .MoveNext
        Loop
    End With
    
    If orst.RecordCount > previousNotificationCount Then
        RefreshGrids
    End If
    
End Sub


Function PadRight(text As Variant, totalLength As Integer) As String
    Dim temp As String
    Dim tempLength As Integer
    Dim length As Integer
    
    temp = Trim(CStr(text))
    tempLength = Len(temp)
    length = totalLength - tempLength
    
    If (length > 0) Then
        PadRight = temp & String(length, " ")
    Else
        PadRight = temp
    End If
    
End Function


Private Sub cmdToggleView_Click()
    If fullView Then
        fullView = False
        Me.Height = FormShortHeight
    Else
        RefreshGrids
        fullView = True
        Me.Height = FormFullHeight
    End If
End Sub


Private Sub cmdRefresh_Click()
    RefreshGrids
End Sub


Private Sub cboWhse_Click()
    If Not m_bIsLoading Then RefreshGrids
End Sub


Private Sub RefreshGrids()
    SetWaitCursor True
    LoadConflictPickGrid
    UpdatePickLineItems
    
    If (m_oOSItemList.Count > 0) Then
        LoadConflictOrders m_oOSItemList.Item(1), m_gwAllConflictOrders.Value("WhseKey")
        'cmdViewOrder.Enabled = True
        'cmdPickOrder.Enabled = True
    Else
        LoadConflictOrders Nothing, m_gwAllConflictOrders.Value("WhseKey")
        'cmdViewOrder.Enabled = False
        'cmdPickOrder.Enabled = False
    End If
    
    SetWaitCursor False
End Sub


'Grid 1
Private Sub LoadConflictPickGrid()
    Dim i As Long
    Dim orst As ADODB.Recordset
    
    Set orst = GetAllOrdersInConflict(cboWhse.ItemData(cboWhse.ListIndex))
    SetCaption orst.RecordCount & " order(s) in Pick Conflict"
    
    With gdxAllConflictOrders
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
End Sub


'Grid 2
Private Sub UpdatePickLineItems()
    Set m_oOSItemList = New OSItemList

    'Load OSItem objects for all items on order
    m_oOSItemList.Load m_gwAllConflictOrders.Value("OPID"), m_gwAllConflictOrders.Value("DropShip")
    
    Dim i As Long
    With gdxLineItems
        .HoldFields
        .ItemCount = m_oOSItemList.Count
        .Refetch
         For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    
    If m_oOSItemList.Count = 0 Then
        lblLineItems.Caption = ""
    Else
        lblLineItems.Caption = "Line item(s) on OP " & m_gwAllConflictOrders.Value("OPID")
    End If

End Sub


Private Sub gdxAllConflictOrders_Click()
    SetWaitCursor True
    
    'RESET THE BUTTONS WHEN REFRESHED
    cmdViewOrder.Enabled = False
    cmdPickOrder.Enabled = False
    
    UpdatePickLineItems
    If (m_oOSItemList.Count > 0) Then
        LoadConflictOrders m_oOSItemList.Item(1), m_gwAllConflictOrders.Value("WhseKey")
    Else
        LoadConflictOrders Nothing, m_gwAllConflictOrders.Value("WhseKey")
    End If
    SetWaitCursor False
End Sub


Private Sub gdxAllConflictOrders_KeyUp(KeyCode As Integer, Shift As Integer)
    SetWaitCursor True
    UpdatePickLineItems
    If (m_oOSItemList.Count > 0) Then
        LoadConflictOrders m_oOSItemList.Item(1), m_gwAllConflictOrders.Value("WhseKey")
    Else
        LoadConflictOrders Nothing, m_gwAllConflictOrders.Value("WhseKey")
    End If
    SetWaitCursor False
End Sub


Private Sub gdxLineItems_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_oOSItemList Is Nothing Then Exit Sub
    If RowIndex > m_oOSItemList.Count Then Exit Sub
    
    With m_oOSItemList.Item(RowIndex)
        Values(1) = .itemkey
        Values(2) = .ItemID
        Values(3) = .Description
        Values(4) = .QtyOrdered
        Values(5) = .QtyOpenToShip
        Values(6) = .Conflict
        Values(7) = .TypeDesc
    End With
    
End Sub


Private Sub gdxLineItems_Click()
    SetWaitCursor True
    
    'RESET THE BUTTONS WHEN REFRESHED
    cmdViewOrder.Enabled = False
    cmdPickOrder.Enabled = False
    
    If (m_oOSItemList.Count > 0) Then
        LoadConflictOrders m_oOSItemList.Item(gdxLineItems.RowIndex(gdxLineItems.Row)), m_gwAllConflictOrders.Value("WhseKey")
    Else
        LoadConflictOrders Nothing, m_gwAllConflictOrders.Value("WhseKey")
    End If
    SetWaitCursor False
End Sub

Private Sub gdxLineItems_KeyUp(KeyCode As Integer, Shift As Integer)
    SetWaitCursor True
    If (m_oOSItemList.Count > 0) Then
        LoadConflictOrders m_oOSItemList.Item(gdxLineItems.RowIndex(gdxLineItems.Row)), m_gwAllConflictOrders.Value("WhseKey")
    Else
        LoadConflictOrders Nothing, m_gwAllConflictOrders.Value("WhseKey")
    End If
    SetWaitCursor False
End Sub

' call when a row is selected in Grid 2

Public Sub LoadConflictOrders(oOSItem As OSItem, ByVal lWhseKey As Long)
    Dim sWhseID As String
    Dim i As Integer

    If oOSItem Is Nothing Then
        txtItemId.text = ""
        txtOnSalesOrders.text = ""
        txtOnPurchOrders.text = ""
        txtOnHand.text = ""
        txtAvail.text = ""
        lblOrdersInConflictForItem.Caption = ""
        With gdxInConflictForItem
            .HoldFields
            .HoldSortSettings = True
            'call with itemkey=0 to get an empty recordset
            Set .ADORecordset = CallSP("spcpcGetPickConflictOrders", "@iItemKey", 0, "@iWhseKey", lWhseKey)
            For i = 1 To .Columns.Count
                'skip Summary column
                If (i <> 5) Then .Columns(i).AutoSize
            Next
        End With
    Else
        lblOrdersInConflictForItem.Caption = "Orders in conflict for " & oOSItem.ItemID
        txtItemId.text = oOSItem.ItemID
        txtOnSalesOrders.text = oOSItem.QtyOnSO
        txtOnPurchOrders.text = GetQtyOnPO(oOSItem.itemkey, lWhseKey)
        txtOnHand.text = oOSItem.QtyOnHand
        txtAvail.text = oOSItem.QtyAvail ' GetQtyAvail(oOSItem.itemkey, lWhseKey)
        selectedItemKey = oOSItem.itemkey
        selectedItemType = oOSItem.ItemType
        
        With gdxInConflictForItem
            .HoldFields
            .HoldSortSettings = True
            'Set .ADORecordset = CallSP("spcpcGetPickConflictOrders", "@iItemKey", oOSItem.ItemKey, "@iWhseKey", lWhseKey)
            Set .ADORecordset = GetOrdersInConflictForItem(oOSItem.itemkey, lWhseKey)
            For i = 1 To .Columns.Count
                'skip Summary column
                If (i <> 5) Then .Columns(i).AutoSize
            Next
        End With

    End If
End Sub


Private Sub cmdPickOrder_Click()
    Dim OPKey As Long
    Dim whsekey As Long
    Dim rdCurrent As JSRowData
    
    On Error GoTo EH
    
    Set rdCurrent = gdxInConflictForItem.GetRowData(gdxInConflictForItem.Row)
    OPKey = CLng(rdCurrent.Value(1))
    whsekey = cboWhse.ItemData(cboWhse.ListIndex)
    
    If vbYes = MsgBox("Are you sure you want to pick OP " & OPKey & "?", vbYesNo, "Pick and Print Order") Then
        SetWaitCursor True
        Dim proxy As MSSOAPLib30.SoapClient30
        Set proxy = New MSSOAPLib30.SoapClient30
        proxy.MSSoapInit g_AutoPickUrl & "?WSDL"
        proxy.PickOrder OPKey, 0, whsekey, User.GetUserName
        proxy.EvaluatePicks
        SetWaitCursor False
        
        RefreshGrids
        
    End If
    
    Exit Sub
EH:
    SetWaitCursor False
    MsgBox "Your data is stale. It's likely this order was aleady picked or is on backorder. Click the refresh button to update your display.", vbOKOnly + vbCritical, "Pick and Print Order"
End Sub


Private Sub cmdViewOrder_Click()
    Dim lOPKey As Long

    lOPKey = m_gwInConflictForItem.Value("opkey")
    
    If lOPKey = 0 Then Return
    
    LogEvent "FConflictOrders", "cmdViewOrder_Click", GetUserName & " instantiating FOrder from FConflictOrders for OP " & lOPKey
    
    SetWaitCursor True
    
    Dim oFrm As FOrder
    Set oFrm = New FOrder
    MDIMain.AddNewWindow oFrm
    With oFrm
        .Show
        .Order.Load lOPKey
        .lblCustName.Visible = True
        .lblCustType(0).Visible = True
        .txtCustName.Visible = False
        .TransitionTabs False
    End With
    
    SetWaitCursor False
End Sub

Private Sub lblOrderInfo_Click(Index As Integer)
    Dim oFrm As FPickConflictInfo
    Set oFrm = New FPickConflictInfo
    oFrm.ShowInfo Me, selectedItemKey, cboWhse.ItemData(cboWhse.ListIndex), Index
    Set oFrm = Nothing
End Sub


Private Sub cmdInvFinder_Click()
    If Len(txtItemId.text) > 0 Then
        Dim oFrm As FInvFinder
        Set oFrm = New FInvFinder
        MDIMain.AddNewWindow oFrm
        oFrm.LoadItemByKey txtItemId.text, selectedItemKey, selectedItemType, cboWhse.ItemData(cboWhse.ListIndex)
    End If
End Sub


'*******************************************************************************************
' Data Access
'*******************************************************************************************


Private Function GetAllOrdersInConflict(ByVal whsekey As Long) As ADODB.Recordset
    Set GetAllOrdersInConflict = CallSP("spcpcGetPickConflict", "@_iWhseKey", whsekey)
End Function


Private Function GetQtyAvail(ByVal itemkey As Long, ByVal whsekey As Long) As Integer
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spcpcGetItemQtyAvailable")
    cmd.Parameters("@ItemKey").Value = itemkey
    cmd.Parameters("@WhseKey").Value = whsekey
    cmd.Execute
    GetQtyAvail = cmd.Parameters("@RetVal").Value
End Function



Private Function GetQtyOnPO(ByVal itemkey As Long, ByVal whsekey As Long) As Integer
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spcpcGetItemQtyOnPO")
    cmd.Parameters("@ItemKey").Value = itemkey
    cmd.Parameters("@WhseKey").Value = whsekey
    cmd.Execute
    GetQtyOnPO = cmd.Parameters("@RetVal").Value
End Function



Private Function GetOrdersInConflictForItem(ByVal itemkey As Long, ByVal whsekey As Long) As ADODB.Recordset
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = g_DB.Connection
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spcpcGetPickConflictOrders"
    cmd.Parameters("@iItemKey").Value = itemkey
    cmd.Parameters("@iWhseKey").Value = whsekey

    Set rst = New ADODB.Recordset

    rst.Open cmd, , adOpenStatic, adLockReadOnly

    Set GetOrdersInConflictForItem = rst
End Function


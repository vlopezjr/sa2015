VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "mmremark.ocx"
Object = "{0FA91D91-3062-44DB-B896-91406D28F92A}#54.0#0"; "SOTACalendar.ocx"
Begin VB.Form FInvFinder 
   Caption         =   "Inventory Finder"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   8820
   Begin VB.Frame frmFind 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8595
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrintLabel 
         Caption         =   "Print Label"
         Height          =   372
         Left            =   4080
         TabIndex        =   19
         Top             =   600
         Width           =   1092
      End
      Begin VB.CheckBox chkDescr 
         Caption         =   "Is Description"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtPartNbr 
         Height          =   285
         Left            =   480
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox cboWhse 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin MMRemark.RemarkViewer rvItem 
         Height          =   810
         Left            =   7620
         TabIndex        =   21
         Top             =   180
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1429
         ContextID       =   "ManageItem"
         Caption         =   "Item Remarks"
      End
      Begin VB.Label Label1 
         Caption         =   "Part"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmTree 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   8595
      Begin VB.Frame frmDrillDown 
         Height          =   3135
         Left            =   180
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   8235
         Begin GridEX20.GridEX gdxDrillDown 
            Height          =   2595
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4577
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ScrollToolTipColumn=   ""
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   2
            Column(1)       =   "FInvFinder.frx":0000
            Column(2)       =   "FInvFinder.frx":00C8
            FormatStylesCount=   6
            FormatStyle(1)  =   "FInvFinder.frx":016C
            FormatStyle(2)  =   "FInvFinder.frx":02A4
            FormatStyle(3)  =   "FInvFinder.frx":0354
            FormatStyle(4)  =   "FInvFinder.frx":0408
            FormatStyle(5)  =   "FInvFinder.frx":04E0
            FormatStyle(6)  =   "FInvFinder.frx":0598
            ImageCount      =   0
            PrinterProperties=   "FInvFinder.frx":0678
         End
         Begin VB.Label lblClose 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7620
            TabIndex        =   17
            Top             =   120
            Width           =   195
         End
         Begin VB.Label lblSizing 
            Caption         =   "lblSizing"
            Height          =   1515
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   6615
         End
      End
      Begin ActiveTabs.SSActiveTabs stInventory 
         Height          =   3735
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6588
         _Version        =   262144
         TabCount        =   2
         Tabs            =   "FInvFinder.frx":0850
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
            Height          =   3345
            Left            =   -99969
            TabIndex        =   10
            Top             =   360
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   5900
            _Version        =   262144
            TabGuid         =   "FInvFinder.frx":08DF
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "Refresh"
               Height          =   315
               Left            =   2520
               TabIndex        =   14
               Top             =   2880
               Width           =   1095
            End
            Begin SOTACalendarControl.SOTACalendar calStartDate 
               Height          =   315
               Left            =   960
               TabIndex        =   12
               Top             =   2880
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskedText      =   "  /  /    "
               Text            =   "  /  /    "
            End
            Begin GridEX20.GridEX gdxInventoryTran 
               Height          =   2655
               Left            =   120
               TabIndex        =   11
               Top             =   120
               Width           =   8100
               _ExtentX        =   14288
               _ExtentY        =   4683
               Version         =   "2.0"
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               MethodHoldFields=   -1  'True
               AllowEdit       =   0   'False
               GroupByBoxVisible=   0   'False
               ColumnHeaderHeight=   285
               IntProp1        =   0
               IntProp2        =   0
               IntProp7        =   0
               ColumnsCount    =   6
               Column(1)       =   "FInvFinder.frx":0907
               Column(2)       =   "FInvFinder.frx":0A53
               Column(3)       =   "FInvFinder.frx":0B83
               Column(4)       =   "FInvFinder.frx":0CA7
               Column(5)       =   "FInvFinder.frx":0DCB
               Column(6)       =   "FInvFinder.frx":0EF3
               FormatStylesCount=   6
               FormatStyle(1)  =   "FInvFinder.frx":101F
               FormatStyle(2)  =   "FInvFinder.frx":10FF
               FormatStyle(3)  =   "FInvFinder.frx":1237
               FormatStyle(4)  =   "FInvFinder.frx":12E7
               FormatStyle(5)  =   "FInvFinder.frx":139B
               FormatStyle(6)  =   "FInvFinder.frx":1473
               ImageCount      =   0
               PrinterProperties=   "FInvFinder.frx":152B
            End
            Begin VB.Label lblStartDate 
               Caption         =   "Start Date"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   2880
               Width           =   800
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
            Height          =   3345
            Left            =   30
            TabIndex        =   8
            Top             =   360
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   5900
            _Version        =   262144
            TabGuid         =   "FInvFinder.frx":1703
            Begin MSComctlLib.TreeView tvwItem 
               Height          =   3135
               Left            =   120
               TabIndex        =   9
               Top             =   120
               Width           =   8175
               _ExtentX        =   14420
               _ExtentY        =   5530
               _Version        =   393217
               LabelEdit       =   1
               Style           =   7
               Appearance      =   1
            End
         End
      End
   End
   Begin VB.Frame frmGrid 
      BorderStyle     =   0  'None
      Height          =   4275
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   8595
      Begin GridEX20.GridEX gdxItems 
         Height          =   3855
         Left            =   0
         TabIndex        =   22
         Top             =   60
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6800
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         ShowToolTips    =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   16
         Column(1)       =   "FInvFinder.frx":172B
         Column(2)       =   "FInvFinder.frx":18C3
         Column(3)       =   "FInvFinder.frx":19E7
         Column(4)       =   "FInvFinder.frx":1B43
         Column(5)       =   "FInvFinder.frx":1C5B
         Column(6)       =   "FInvFinder.frx":1D7F
         Column(7)       =   "FInvFinder.frx":1EAF
         Column(8)       =   "FInvFinder.frx":1FFF
         Column(9)       =   "FInvFinder.frx":212F
         Column(10)      =   "FInvFinder.frx":2253
         Column(11)      =   "FInvFinder.frx":2377
         Column(12)      =   "FInvFinder.frx":249B
         Column(13)      =   "FInvFinder.frx":25DB
         Column(14)      =   "FInvFinder.frx":271F
         Column(15)      =   "FInvFinder.frx":2863
         Column(16)      =   "FInvFinder.frx":299B
         FormatStylesCount=   6
         FormatStyle(1)  =   "FInvFinder.frx":2AC7
         FormatStyle(2)  =   "FInvFinder.frx":2BA7
         FormatStyle(3)  =   "FInvFinder.frx":2CDF
         FormatStyle(4)  =   "FInvFinder.frx":2D8F
         FormatStyle(5)  =   "FInvFinder.frx":2E43
         FormatStyle(6)  =   "FInvFinder.frx":2F1B
         ImageCount      =   0
         PrinterProperties=   "FInvFinder.frx":2FD3
      End
   End
End
Attribute VB_Name = "FInvFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lMaxRecs = 50
Private Const k_lMinWidth = 9060
Private Const k_lMinHeight = 6300

Private m_bLoading As Boolean

Private m_oItem As IItem
Private m_lItemKey As Long

Private m_gwItems As GridEXWrapper
Private m_oRst As ADODB.Recordset

Private m_oItemInventories As Inventories
Private m_oItemAttributes As ADODB.Recordset

Private m_lp As LabelPrinter

Private m_lWindowID As Long


Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property


Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property


Public Sub SetCaption(ByRef i_sTitle As String)
    Me.caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub

Private Sub Form_Load()
    Dim lWhseKey
    m_bLoading = True

    frmGrid.ZOrder 0
    
    frmDrillDown.Visible = False
    
    'lblSizing.Visible = True 'for dev
    'lblSizing.ZOrder 0
    gdxDrillDown.Visible = True
    
    lWhseKey = GetWhseKeyFromUserKey(GetUserKey)

    g_rstWhses.Filter = "Transit=0"
    LoadCombo cboWhse, g_rstWhses, "WhseID", "WhseKey", lWhseKey
    g_rstWhses.Filter = adFilterNone

    Set m_gwItems = New GridEXWrapper
    m_gwItems.Grid = gdxItems
    
    cmdPrintLabel.Visible = False   'the grid is empty, don't show the button
    
    If HasLabelPrinter Then
        Set m_lp = New LabelPrinter
    End If

    calStartDate.value = Now
    
    With Me
        .Height = 5204
        .width = 7650
    End With
    
    m_bLoading = False
    
    txtPartNbr.Enabled = True
    SetCaption "Inventory Finder"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not m_oRst Is Nothing Then
        If m_oRst.State <> adStateClosed Then
            m_oRst.Close
        End If
        Set m_oRst = Nothing
    End If
    
    If Not m_oItem Is Nothing Then
        Set m_oItem = Nothing
    End If
    
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        DoShowHelp
    Else
        If KeyCode = Asc(vbCrLf) And Me.ActiveControl.Name = "txtPartNbr" Then
            cmdFind_Click
        Else
            MDIMain.GlobalKeyDownProcessing KeyCode, Shift
        End If
    End If
End Sub


Private Sub Form_Resize()
    Dim msg As String
    
    Dim i As Long
    Dim lBorder As Long
    lBorder = 240 '120

    If Me.WindowState = 1 Then Exit Sub   'Minimized
    If Me.width < k_lMinWidth Then Me.width = k_lMinWidth
    If Me.Height < k_lMinHeight Then Me.Height = k_lMinHeight

    With frmFind
        .width = Me.width - 600
        rvItem.Left = .width - rvItem.width - 200
    End With
    
    With frmGrid
        .width = Me.width - 600
        .Height = Me.Height - .Top - 650
        gdxItems.width = .width
        gdxItems.Height = .Height - 200
        For i = 1 To gdxItems.Columns.Count
            gdxItems.Columns(i).AutoSize
        Next
        
    End With
    
    With frmTree
        .width = Me.width - 600
        .Height = Me.Height - .Top - 650
        
        stInventory.width = .width
        stInventory.Height = .Height - 200
        
        tvwItem.width = stInventory.width - 400 '- (2 * lBorder)
        tvwItem.Height = stInventory.Height - 600 '- (3 * lBorder)
        
        gdxInventoryTran.width = tvwItem.width
        gdxInventoryTran.Height = tvwItem.Height - 600 '- (3 * lBorder)
        
        lblStartDate.Top = gdxInventoryTran.Top + gdxInventoryTran.Height + lBorder
        calStartDate.Top = lblStartDate.Top
        cmdRefresh.Top = lblStartDate.Top
    End With
    
    With frmDrillDown
        .Top = 360 + 120
        .Left = 30 + 120
        .width = tvwItem.width
        .Height = tvwItem.Height + 100
        lblClose.Left = .width - 240
        
        msg = "frmTree T " & frmTree.Top & ", L " & frmTree.Left
        msg = msg & vbCrLf & "stInventory T " & stInventory.Top & ", L " & stInventory.Left & ", W " & stInventory.width & ", H " & stInventory.Height
        msg = msg & vbCrLf & "SSActiveTabPanel1 T " & SSActiveTabPanel1.Top & ", L " & SSActiveTabPanel1.Left & ", W " & SSActiveTabPanel1.width & ", H " & SSActiveTabPanel1.Height
        msg = msg & vbCrLf & "tvwItem T " & tvwItem.Top & ", L " & tvwItem.Left & ", W " & tvwItem.width & ", H " & tvwItem.Height
        msg = msg & vbCrLf & "frmDrillDown T " & .Top & ", L " & .Left & ", W " & .width & ", H " & .Height
        lblSizing.caption = msg
    End With
    
    With gdxDrillDown
        .width = frmDrillDown.width - 250
        .Height = frmDrillDown.Height - 450
    End With

End Sub


Public Sub DoShowHelp()
    ShowHelp "FindInv", True
End Sub


Public Sub LookupPart(ByVal i_sPartNbr As String, ByVal i_lWhseIndex As Long)
    txtPartNbr.text = i_sPartNbr
    cboWhse.ListIndex = i_lWhseIndex
    cmdFind_Click
End Sub


Private Sub cboWhse_Click()
    If Not m_bLoading Then
        If Not tvwItem.Nodes.Count = 0 Then
            PopulateTree
        End If
    End If
End Sub


Private Sub cmdRefresh_Click()
    LoadInvTranHistory
End Sub


Private Sub cmdPrintLabel_Click()
    Dim cmd As ADODB.Command
    Dim BinID As String
    
    On Error GoTo EH
    
    Set cmd = CreateCommandSP("spcpcGetBinLoc")
    cmd.Parameters("@_iWhseKey").value = GetUserWhseKey
    cmd.Parameters("@_iItemKey").value = gdxItems.value(12)
    cmd.Execute
    
    BinID = IIf(IsNull(cmd.Parameters("@_oBinLoc").value), vbNullString, cmd.Parameters("@_oBinLoc").value)
    
    m_lp.Clear
    m_lp.AddLine gdxItems.value(2)
    m_lp.AddLine gdxItems.value(3)
    m_lp.AddLine BinID
    m_lp.NumLabels = 0
                
    m_lp.PrintLabel
    Exit Sub

EH:
    msg Err.Description, vbCritical, Err.Source
End Sub


Private Function HasLabelPrinter() As Boolean
    Dim p As Printer
    HasLabelPrinter = False
    For Each p In Printers
        If InStr(1, p.DeviceName, "Zebra") Then
            HasLabelPrinter = True
            Exit For
        End If
    Next
End Function


Private Sub cmdFind_Click()
    Dim cmd As ADODB.Command
    Dim InputText As String
    
    If txtPartNbr.text = "" Then
        msg "Please enter a part text first."
        Exit Sub
    End If

    InputText = Trim(txtPartNbr.text)
    
    'txtPartNbr.text = ""
    
    frmGrid.ZOrder 0
    
    SetCaption "Inventory Finder: " & UCase(InputText)
    SetWaitCursor True
    
    If chkDescr.value <> vbChecked Then
        Set m_oRst = CallSP("spCPOPInvFinderSearch", "@WhseKey", cboWhse.ItemData(cboWhse.ListIndex), "@SearchText", InputText, "@RowCount", k_lMaxRecs, "@SearchType", 1)
    Else
        Set m_oRst = CallSP("spCPOPInvFinderSearch", "@WhseKey", cboWhse.ItemData(cboWhse.ListIndex), "@SearchText", "%" & InputText & "%", "@RowCount", k_lMaxRecs, "@SearchType", 2)
    End If

    With gdxItems
        .Columns(4).Format = g_MoneyMask
        .Columns(5).Format = g_MoneyMask
        .Columns(6).Format = g_MoneyMask
        .HoldSortSettings = True
        .HoldFields
        Set .ADORecordset = m_oRst
    End With
    
    If gdxItems.ItemCount = 0 Then
        msg "Sorry, this search returned no records.", , "Search Failed"
    Else
        GroupFormat gdxItems
    End If

    'if the computer has a Zebra printer and the grid is not empty
    If HasLabelPrinter Then cmdPrintLabel.Visible = True    'And Not m_ILList Is Nothing

    SetWaitCursor False

    TryToSetFocus gdxItems
End Sub


Private Sub GroupFormat(oGdx As GridEX)
     Dim fmtcon As JSFmtCondition
    Dim col As JSColumn
    Dim group As JSGroup
    
    Set col = gdxItems.Columns("ItemType")
    Set fmtcon = gdxItems.FmtConditions.Add(col.Index, jgexEqual, 7)
    fmtcon.FormatStyle.BackColor = vbYellow
End Sub


Private Sub AutoSizeColumns(oGdx As GridEX)
    Dim i As Long
    
    For i = 1 To oGdx.Columns.Count
        oGdx.Columns(i).AutoSize
    Next
End Sub


Private Sub LoadInvTranHistory()
    Dim orst As ADODB.Recordset
    Dim i As Integer
    
    SetWaitCursor True
    m_bLoading = True
    
    Set orst = CallSP("spCPCGetInvHistTran", "@WhseKey", cboWhse.ItemData(cboWhse.ListIndex), "@ItemKey", m_lItemKey, "@StartDate", calStartDate.value)
    
    With gdxInventoryTran
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
        
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    
    Set orst = Nothing
    m_bLoading = False
    SetWaitCursor False
End Sub


Private Function LoadInventories(lItemType As Long, lItemKey As Long)
    Dim oitems As Items
    
    SetWaitCursor True
    
    Set oitems = New Items
    Select Case lItemType
        Case Is = 5
            Set m_oItem = oitems.CreateItem(itFinishedGood)
            Dim oFinGood As ItemFinGood
            Set oFinGood = m_oItem
            oFinGood.Load lItemKey, cboWhse.ItemData(cboWhse.ListIndex)
        Case Is = 7
            Set m_oItem = oitems.CreateItem(itBTOKit)
            Dim oBTOKit As ItemBTOKit
            Set oBTOKit = m_oItem
            oBTOKit.Load lItemKey, cboWhse.ItemData(cboWhse.ListIndex)
    End Select
    SetWaitCursor False
            
    cmdPrintLabel.Visible = False
        
    PopulateTree
    
    frmTree.ZOrder 0
        
End Function


Private Sub PopulateTree()
   
    tvwItem.Nodes.Clear
    
    If m_oItem.SageItemType = 5 Then
        AddFinGoodNodes m_oItem
    End If
    
    If m_oItem.SageItemType = 7 Then
        AddBTOKitNodes m_oItem
    End If
    
End Sub


Private Sub AddBTOKitNodes(i_oKit As ItemBTOKit)
    Dim oFinGood As ItemFinGood
    Dim oInv As Inventory
    
    With i_oKit
        
        tvwItem.Nodes.Add _
            , , _
            CStr(.IItem_ItemID) & "-Kit", _
            "(" & CStr(.Components.CalcQty(cboWhse.ItemData(cboWhse.ListIndex), "QtyAvail")) & ") " & .IItem_ItemID & " - " & .IItem_Descr

        'encode the itemkey in the node key as a parameter to attribute look up
        tvwItem.Nodes.Add _
            CStr(.IItem_ItemID) & "-Kit", tvwChild, _
            CStr(.IItem_ItemKey) & "PropertiesKit", _
            "Properties"

        tvwItem.Nodes.Add _
            CStr(.IItem_ItemKey) & "PropertiesKit", tvwChild, _
            , _
            "Catalog Page: " & CStr(.CatPage)

        'highlight CatPage (it's double-clickable)
        HiliteLinkNode tvwItem.Nodes
        
        'Components
        tvwItem.Nodes.Add _
            CStr(.IItem_ItemID) & "-Kit", tvwChild, _
            CStr(.IItem_ItemID) & "-Components", _
            "Components"
        
        For Each oFinGood In .Components
            AddFinGoodNodes oFinGood, .IItem_ItemID
        Next
        
        'Rolled Up inventory by Warehouse
'        tvwItem.Nodes.Add _
'            CStr(.IItem_ItemID) & "-Kit", tvwChild, _
'            CStr(.IItem_ItemID) & "-KitInventory", _
'            "Inventory"
        
        For Each oInv In .Inventories
            
                'CStr(.IItem_ItemID) & "-KitInventory", tvwChild,
                            
            tvwItem.Nodes.Add _
                CStr(.IItem_ItemID) & "-Kit", tvwChild, _
                oInv.WhseKey & CStr(.IItem_ItemID) & "-Kit", _
                oInv.whseid & " (" & CStr(.Components.CalcQty(oInv.WhseKey, "QtyAvail")) & ")"

            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "-Kit", tvwChild, _
                oInv.WhseKey & CStr(.IItem_ItemID) & "Vendor", _
                "Vendor" ' [" & oInv.VendName & "]"
            
            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "Vendor", tvwChild, _
                , _
                "Cost: " & Format$(.IItem_Cost, g_MoneyMask)
                
            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "Vendor", tvwChild, _
                , _
                "PartNbr: " & GetVendPartNbr(oInv.WhseKey, .IItem_ItemKey)
                
            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "-Kit", tvwChild, _
                , _
                "Qty Avail: " & CStr(.Components.CalcQty(oInv.WhseKey, "QtyAvail"))
            
'''         Expose '$' Kits on order when Kit Item Key is supplied. We want the user to drill down
'''         tvwItem.Nodes.Add "R" & sWhse & CStr(.IItem_ItemID), tvwChild, , "Qty On PO: " & CStr(oInv.QtyOnPO)
            
            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "-Kit", tvwChild, _
                oInv.WhseKey & .IItem_ItemID & "-OnPO", _
                "On PO (drill in for detail)"
                
            HiliteLinkNode tvwItem.Nodes
            
            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "-Kit", tvwChild, _
                oInv.WhseKey & .IItem_ItemID & "-OnSO", _
                "On SO: " & CStr(oInv.QtyOnSO)
                
            HiliteLinkNode tvwItem.Nodes
        Next
    End With
    
End Sub


Private Sub AddFinGoodNodes(i_oFinGood As ItemFinGood, Optional i_lKitID As Variant)
    Dim oInv As Inventory
    
    With i_oFinGood
    
        'cache a reference for node click processing
        Set m_oItemInventories = i_oFinGood.Inventories
        
        If IsMissing(i_lKitID) Then
          tvwItem.Nodes.Add _
            , , _
            CStr(.IItem_ItemID) & "-Item", _
            "(" & CStr(.IItem_QtyAvail(cboWhse.ItemData(cboWhse.ListIndex))) & ") " & .IItem_ItemID & " - " & .IItem_Descr
        Else
          tvwItem.Nodes.Add _
            i_lKitID & "-Components", tvwChild, _
            CStr(.IItem_ItemID) & "-Item", _
            "(" & CStr(.IItem_QtyAvail(cboWhse.ItemData(cboWhse.ListIndex))) & "/" & CStr(.IItem_Qty) & ") " & .IItem_ItemID & " - " & .IItem_Descr
        End If
        
        'encode the itemkey in the node key as a parameter to attribute look up
        tvwItem.Nodes.Add _
            CStr(.IItem_ItemID) & "-Item", tvwChild, _
            CStr(.IItem_ItemKey) & "Properties", _
            "Properties"
                    
        tvwItem.Nodes.Add _
            CStr(.IItem_ItemKey) & "Properties", tvwChild, _
            , _
            "Catalog Page: " & CStr(.CatPage)
    
        'highlight CatPage (it's double-clickable)
        HiliteLinkNode tvwItem.Nodes
                    
'        tvwItem.Nodes.Add _
'            CStr(.IItem_ItemID) & "-Item", tvwChild, _
'            .IItem_ItemID & "-Inventory", _
'            "Inventory"

        'for each warehouse
        For Each oInv In .Inventories
                  
                'CStr(.IItem_ItemID) & "-Inventory", tvwChild,
                
            tvwItem.Nodes.Add _
                CStr(.IItem_ItemID) & "-Item", tvwChild, _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", _
                oInv.whseid & " (" & CStr(oInv.QtyAvail) & ")"
            
            tvwItem.Nodes.Item(4).EnsureVisible
            
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                oInv.WhseKey & CStr(.IItem_ItemID) & "Vendor", _
                "Vendor [" & oInv.VendName & "]"

            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "Vendor", tvwChild, _
                , _
                "Cost: " & Format$(.IItem_Cost, g_MoneyMask)

            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "Vendor", tvwChild, _
                , _
                "PartNbr: " & GetVendPartNbr(oInv.WhseKey, .IItem_ItemKey)
            
            tvwItem.Nodes.Add _
               CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                oInv.WhseKey & CStr(.IItem_ItemID) & "Bins", _
                "Bins"
            
            tvwItem.Nodes.Add _
                oInv.WhseKey & CStr(.IItem_ItemID) & "Bins", tvwChild, _
                oInv.WhseKey & CStr(.IItem_ItemID) & "BinList", _
                "<bins>"
            
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                , _
                "In Conflict: " & IIf(oInv.InConflict, "Yes", "No")
                        
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                , _
                "Avail: " & CStr(oInv.QtyAvail)
            
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                oInv.WhseKey & CStr(.IItem_ItemID) & "-OnHand", _
                "On Hand: " & CStr(oInv.QtyOnHand)
            
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                oInv.WhseKey & .IItem_ItemID & "-OnPO", _
                "On PO: " & CStr(oInv.QtyOnPO)
            
            HiliteLinkNode tvwItem.Nodes
            
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                oInv.WhseKey & .IItem_ItemID & "-OnSO", _
                "On SO: " & CStr(oInv.QtyOnSO)

            HiliteLinkNode tvwItem.Nodes
            
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                oInv.WhseKey & .IItem_ItemID & "-Picked", _
                "Picked: " & CStr(oInv.Picked)
            
            HiliteLinkNode tvwItem.Nodes
            
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                oInv.WhseKey & .IItem_ItemID & "-Packed", _
                "Packed: " & CStr(oInv.Packed)
            
            HiliteLinkNode tvwItem.Nodes
                        
            tvwItem.Nodes.Add _
                CStr(oInv.WhseKey) & CStr(.IItem_ItemID) & "-Warehouse", tvwChild, _
                oInv.WhseKey & .IItem_ItemID & "-PendingAdj", _
                "Pending Adj: " & CStr(oInv.PendingAdjustments)
                
            HiliteLinkNode tvwItem.Nodes
                    
        Next oInv

    End With
End Sub


'operates on the last node in the collection

Private Sub HiliteLinkNode(i_nodes As Nodes)
    i_nodes.Item(i_nodes.Count).ForeColor = &HFF0000
    i_nodes.Item(i_nodes.Count).Bold = True
End Sub


'Used by FOrder

Public Sub LoadItem(i_oItem As IItem, l_iWhseKey As Long)

    Set m_oItem = i_oItem
    frmTree.ZOrder 0
    
    txtPartNbr.text = i_oItem.ItemID
    SetComboByKey cboWhse, l_iWhseKey
    
    cboWhse.Enabled = False
    txtPartNbr.Enabled = False
    cmdPrintLabel.Visible = False
    cmdFind.Visible = False

    PopulateTree
    
End Sub


'Used by FPickConflicts

Public Sub LoadItemByKey(ItemID As String, ByVal ItemKey As Long, ByVal ItemType As Long, ByVal WhseKey As Long)

    txtPartNbr.text = ItemID
    SetComboByKey cboWhse, WhseKey
    
    LoadInventories ItemType, ItemKey
    LoadInvTranHistory
    
End Sub


Private Sub gdxItems_DBlClick()
    LoadDetail
End Sub


Private Sub gdxItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        LoadDetail
    End If
End Sub


Private Sub LoadDetail()
    'This error happens when the user double clicks the empty grid. OP still tries to
    'call PopulateTree to add nodes to the TreeView and it cause this error.
    'Add guard condition in gdxItems_DBClick to exit when gdxItems has no items at all
    
    SetCaption "Inventory Finder: " & UCase(gdxItems.value(2))
    If Not IsEmpty(gdxItems.value(1)) Then
        m_lItemKey = gdxItems.value(12)
        LoadInventories lItemType:=gdxItems.value(1), lItemKey:=gdxItems.value(12)
        LoadInvTranHistory
    End If
End Sub


Private Sub gdxItems_SelectionChange()

    rvItem.OwnerID = m_gwItems.value("PartNbr")

End Sub


Private Sub lblClose_Click()
    frmDrillDown.Visible = False
    Set gdxDrillDown.ADORecordset = Nothing
End Sub


Private Sub tvwItem_DblClick()
    Dim SelectedNode As Node
    Dim lWhseKey As Long
    Dim oInv As Inventory
    
    SetWaitCursor True

    Set SelectedNode = tvwItem.SelectedItem
    
    If Left(Trim(SelectedNode.text), 5) = "On PO" Then
        
        lWhseKey = CLng(Mid(SelectedNode.Key, 1, 2))
        Set oInv = m_oItemInventories.WhseInventory(lWhseKey)
        
        ShowGridFrame Left(Trim(SelectedNode.text), 5)
        Set gdxDrillDown.ADORecordset = oInv.GetPurchaseOrders
        SizeGridColumns gdxDrillDown
        
    ElseIf Left(Trim(SelectedNode.text), 5) = "On SO" Then
        
        lWhseKey = CLng(Mid(SelectedNode.Key, 1, 2))
        Set oInv = m_oItemInventories.WhseInventory(lWhseKey)
            
        ShowGridFrame Left(Trim(SelectedNode.text), 5)
        Set gdxDrillDown.ADORecordset = oInv.GetOrders
        SizeGridColumns gdxDrillDown
    
    ElseIf Left(Trim(SelectedNode.text), 6) = "Picked" Then
        
        lWhseKey = CLng(Mid(SelectedNode.Key, 1, 2))
        Set oInv = m_oItemInventories.WhseInventory(lWhseKey)
        
        ShowGridFrame Left(Trim(SelectedNode.text), 6)
        Set gdxDrillDown.ADORecordset = oInv.GetPickedOrders
        SizeGridColumns gdxDrillDown
        
    ElseIf Left(Trim(SelectedNode.text), 6) = "Packed" Then
        
        lWhseKey = CLng(Mid(SelectedNode.Key, 1, 2))
        Set oInv = m_oItemInventories.WhseInventory(lWhseKey)
        
        ShowGridFrame Left(Trim(SelectedNode.text), 6)
        Set gdxDrillDown.ADORecordset = oInv.GetPackedOrders
        SizeGridColumns gdxDrillDown
        
    ElseIf Left(Trim(SelectedNode.text), 11) = "Pending Adj" Then
        
        lWhseKey = CLng(Mid(SelectedNode.Key, 1, 2))
        Set oInv = m_oItemInventories.WhseInventory(lWhseKey)
        
        ShowGridFrame Left(Trim(SelectedNode.text), 11)
        Set gdxDrillDown.ADORecordset = oInv.GetPendingInventoryAdjustments
        SizeGridColumns gdxDrillDown
        
    ElseIf Left(Trim(SelectedNode.text), 12) = "Catalog Page" Then
        
        Dim frmCatPage As FCatPage
    
        Set frmCatPage = New FCatPage
        MDIMain.AddNewWindow frmCatPage
        With frmCatPage
            .Show
            .PartNo = m_oItem.ItemID
            'Set default CustType to be end user
            .CustType = 1
            .ShowPage
        End With
    End If
    
    SetWaitCursor False
End Sub


Sub ShowGridFrame(i_caption As String)
    frmDrillDown.Visible = True
    frmDrillDown.ZOrder 0
    frmDrillDown.caption = i_caption
End Sub

Sub SizeGridColumns(i_grid As GridEX)
    Dim i As Integer
    For i = 1 To i_grid.Columns.Count
        i_grid.Columns(i).AutoSize
    Next
End Sub


Private Sub tvwItem_Expand(ByVal Node As MSComctlLib.Node)
    Dim lWhseKey As Long
    Dim lItemKey As Long
    Dim orst As ADODB.Recordset
    
    'LR 11/16/15 this is a kludge. the children count is to prevent the nodes from being added more than once
    'why node.children = 2 is uncertain
    If Node.text = "Properties" And Node.Children <= 2 Then

        lItemKey = CLng(Mid(Node.Key, 1, InStr(Node.Key, "Properties") - 1))
        
        Set m_oItemAttributes = GetItemAttributes(lItemKey)
        
        While Not m_oItemAttributes.EOF
            tvwItem.Nodes.Add _
                Node.Key, _
                tvwChild, _
                , _
                m_oItemAttributes.Fields("AttrName").value & " : " & m_oItemAttributes("AttrValue").value
            m_oItemAttributes.MoveNext
        Wend
    End If
    
    Select Case Node.Child.text
                                            
        Case Is = "<bins>"
        
            SetWaitCursor True
            
            lWhseKey = CLng(Mid(Node.Child.Key, 1, 2))

            If ConvertSageItemType(m_oItem.SageItemType) = itBTOKit Then
                'Parse the ItemId out of the node label
                'Assume that the post-fix character length is always 2
                Set orst = CallSP("spCPCGetBinByItemID", "@i_ItemID", Mid(Node.Key, 4, Len(Node.Key) - 5), "@i_WhseKey", lWhseKey)
            Else
                Set orst = CallSP("spCPCGetInvFinderLocs", "@i_ItemKey", m_oItem.ItemKey, "@i_WhseKey", lWhseKey)
            End If

            'If we've found bins, and them into the tree
            
            If orst.RecordCount > 0 Then
                'This solves the problem on not inserting a unique node key (?)
                'Remove node
                tvwItem.Nodes.Remove Node.Child.Index
                'add new nodes
                While Not orst.EOF
                    tvwItem.Nodes.Add _
                        Node.Key, tvwChild, _
                        orst.Fields("PrefNo") & " - " & orst.Fields("WhseBinID") & " - " & Node.Key, _
                        orst.Fields("PrefNo") & " - " & orst.Fields("WhseBinID")
                    orst.MoveNext
                Wend
            Else
                tvwItem.Nodes.Remove Node.Child.Index
                tvwItem.Nodes.Add _
                    Node.Key, tvwChild, _
                    , _
                    "<none>"
            End If
            
            CloseRst orst
            Set orst = Nothing
            
            SetWaitCursor False
            
        End Select
End Sub


Private Function GetItemAttributes(i_ItemKey As Long) As ADODB.Recordset
    Dim sql As String
    sql = "SELECT cmAttribute.AttrName, cmItemAttribute.AttrValue " & _
        "FROM catalog_staging.dbo.cmItemAttribute INNER JOIN " & _
        "catalog_staging.dbo.cmAttribute ON catalog_staging.dbo.cmItemAttribute.AttrKey = catalog_staging.dbo.cmAttribute.AttributeKey INNER JOIN " & _
        "catalog_staging.dbo.cmItem ON catalog_staging.dbo.cmItemAttribute.ItemKey = catalog_staging.dbo.cmItem.ItemKey " & _
        "where cmItem.timItemKey = " & i_ItemKey
    Set GetItemAttributes = LoadDiscRst(sql)
End Function


Private Function GetVendPartNbr(i_WhseKey As Long, i_ItemKey As Long) As String
    Dim orst As ADODB.Recordset
    Set orst = CallSP("spCPCGetVendItemID", "@i_ItemKey", i_ItemKey, "@i_WhseKey", i_WhseKey)
    If Not orst.EOF Then
        GetVendPartNbr = orst.Fields("VendPartNbr").value
    End If
    Set orst = Nothing
End Function


Private Sub txtPartNbr_GotFocus()
    txtPartNbr.SelStart = 0
    txtPartNbr.SelLength = Len(txtPartNbr.text)
End Sub

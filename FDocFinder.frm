VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FDocFinder 
   Caption         =   "Document Finder"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8887.865
   ScaleMode       =   0  'User
   ScaleWidth      =   9375
   Begin VB.Frame Frame1 
      Caption         =   "1. Select a  Document Number and Type"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9120
      Begin VB.TextBox txtDocNbr 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1440
      End
      Begin VB.CommandButton cmdFindDocs 
         Caption         =   "Find"
         Height          =   375
         Left            =   7680
         TabIndex        =   2
         Top             =   360
         Width           =   1200
      End
      Begin VB.ComboBox cboDocs 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2000
      End
      Begin GridEX20.GridEX gdxDocs 
         Height          =   5175
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9128
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "FDocFinder.frx":0000
         FormatStyle(2)  =   "FDocFinder.frx":0138
         FormatStyle(3)  =   "FDocFinder.frx":01E8
         FormatStyle(4)  =   "FDocFinder.frx":029C
         FormatStyle(5)  =   "FDocFinder.frx":0374
         FormatStyle(6)  =   "FDocFinder.frx":042C
         FormatStyle(7)  =   "FDocFinder.frx":050C
         ImageCount      =   0
         PrinterProperties=   "FDocFinder.frx":052C
      End
      Begin VB.Label Label1 
         Caption         =   "Document Number"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "2.Select a Related Document"
      Height          =   6135
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9120
      Begin VB.CommandButton cmdBackTo1 
         Caption         =   "Back"
         Height          =   312
         Left            =   7020
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwPT 
         Height          =   2235
         Left            =   4800
         TabIndex        =   8
         Top             =   3600
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   3942
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwShip 
         Height          =   2235
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwSO 
         Height          =   2235
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwInv 
         Height          =   2235
         Left            =   4800
         TabIndex        =   11
         Top             =   960
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Source Document"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblSrcDoc 
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   1740
         TabIndex        =   17
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "Sales Orders"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Invoices"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Shipments"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label lblCustName 
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   3360
         TabIndex        =   13
         Top             =   360
         Width           =   3312
      End
      Begin VB.Label Label6 
         Caption         =   "Pick Tickets"
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   3360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "3. View the Document Detail"
      Height          =   6135
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   9120
      Begin VB.CommandButton cmdBackTo2 
         Caption         =   "Back"
         Height          =   375
         Left            =   8040
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdEmailInvoice 
         Height          =   315
         Left            =   8040
         Picture         =   "FDocFinder.frx":0704
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Email this invoice to the customer"
         Top             =   960
         Visible         =   0   'False
         Width           =   315
      End
      Begin MSComctlLib.ListView lvwDocDetail 
         Height          =   5535
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   9763
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FDocFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWindowID As Long
Private WithEvents m_gw As GridEXWrapper
Attribute m_gw.VB_VarHelpID = -1

Private Const k_lDocMaxRecs = 50

'Private Const k_lMinWidth = 9120
'Private Const k_lMinHeight = 4455

Dim m_rstDocs As ADODB.Recordset
Dim m_rstRelDocs As ADODB.Recordset

'added 1/20/2010 LR to support email function
Private m_lInvcKey As Long

' Vax Viewer Manager
'Private m_oRstMetaData As ADODB.Recordset
'Private m_sSearchType As String
'Private m_gwVaxDataViewer As GridEXWrapper
'Private m_bLoading As Boolean
'Private Const k_sLookup = "Lookup"
'Private Const k_sFreeStanding = "FreeStanding"
'Private Const k_lPCHeight = 2415
'Private Const k_lNonPCHeight = 4815

Private Const LVM_SETCOLUMNWIDTH As Long = &H1000 + 30
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



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


Private Sub cmdEmailInvoice_Click()
    Dim frm As FEmailInvoice
    Set frm = New FEmailInvoice
    frm.InvcKey = m_lInvcKey
    frm.Init
'
'    If frm.LogEvent Then rvOrder.AddRemark "Order.Private", "Emailed quote to " & frm.EMailAddr
'
    Unload frm
    Set frm = Nothing
End Sub



Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        DoShowHelp
    Else
        MDIMain.GlobalKeyDownProcessing KeyCode, Shift
    End If
End Sub


Public Sub DoShowHelp()
    ShowHelp "FindDoc", True
End Sub


Private Sub cmdBackTo1_Click()
    UnloadFrame2
End Sub


Private Sub cmdBackTo2_Click()
    UnloadFrame3
End Sub

Private Sub cmdFindDocs_Click()
    Dim cmd As ADODB.Command

    SetWaitCursor True
    
    Set cmd = CreateCommandSP("spCPOPDocSearch")

    If m_rstDocs.State = adStateOpen Then
        m_rstDocs.Close
    End If
    
    With cmd
        .Parameters("@DocID").value = txtDocNbr.text
        .Parameters("@RowCount").value = k_lDocMaxRecs
        .Parameters("@DocType").value = cboDocs.text
        m_rstDocs.Open cmd
    End With
    
    Set gdxDocs.ADORecordset = m_rstDocs
    gdxDocs.Columns("SearchID").Visible = False
    'clean-up
    Set cmd = Nothing
    SetWaitCursor False
End Sub


Private Sub Form_Load()
    Set m_rstDocs = New ADODB.Recordset
    Frame1.ZOrder 0
    With Me
        .Height = 7500
        .width = 9915
    End With
    SetCaption "Document Finder"
    Set m_gw = New GridEXWrapper
    LoadDocCombo
    m_gw.Grid = gdxDocs
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '4/7/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gw = Nothing
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not m_rstDocs Is Nothing Then
        MDIMain.UnloadTool m_lWindowID
        Set m_rstDocs = Nothing
    End If
End Sub


Private Sub LoadFrame2(sTranID As String, sCustName As String, sDocType As String)
    lblSrcDoc.caption = sTranID
    lblCustName.caption = sCustName
    FindRelDocs sDocType
    Frame2.ZOrder 0
End Sub


Private Sub UnloadFrame2()
    Frame1.ZOrder 0
End Sub


Private Sub LoadFrame3(sDocType As String, lDocKey As Long)
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    SetWaitCursor True
    
    Set cmd = CreateCommandSP("spCPOPDocDetail")
    
    cmdEmailInvoice.Visible = False
    
    Select Case sDocType
        Case Is = "SalesOrder"
            cmd.Parameters("@SOKey").value = lDocKey
        Case Is = "Shipment"
            cmd.Parameters("@ShipKey").value = lDocKey
        Case Is = "Invoice"
            cmd.Parameters("@InvcKey").value = lDocKey
            cmdEmailInvoice.Visible = True
        Case Is = "PickTicket"
            cmd.Parameters("@PTKey").value = lDocKey
    End Select
    
    Set rst = New ADODB.Recordset
    rst.Open cmd
    
    If Not rst.EOF Then
        LoadDetailView rst
        Frame3.ZOrder 0
    Else
        msg "Sorry. No " & sDocType & " details available now."
    End If
    
    SetWaitCursor False
    Set cmd = Nothing
    rst.Close
    Set rst = Nothing
End Sub

Private Sub UnloadFrame3()
    Frame2.ZOrder 0
End Sub


Private Sub LoadListViews(rst As ADODB.Recordset)
    Dim oListItem As ListItem
    Dim olstItem As ListItem
    
    'Clear
    lvwSO.ListItems.Clear
    lvwInv.ListItems.Clear
    lvwShip.ListItems.Clear
    lvwPT.ListItems.Clear
    
    'Set Up Columns
    With lvwSO.ColumnHeaders
        .Clear
        .Add , "SalesOrder", "Sales Order"
        .Add , "SODate", "Date"
        .Add , "SOKey", "Key"
        lvwSO.ColumnHeaders("SOKey").width = 0
    End With
    
    With lvwInv.ColumnHeaders
        .Clear
        .Add , "Invoice", "Invoice"
        .Add , "InvDate", "Date"
        .Add , "InvcKey", "Key"
        lvwInv.ColumnHeaders("InvcKey").width = 0
    End With
    
    With lvwShip.ColumnHeaders
        .Clear
        .Add , "Shipment", "Shipment"
        .Add , "ShipDate", "Date"
        .Add , "ShipKey", "Key"
        lvwShip.ColumnHeaders("ShipKey").width = 0
    End With
    
     With lvwPT.ColumnHeaders
        .Clear
        .Add , "PickListNo", "Pick Ticket"
        .Add , "PickListDate", "Date"
        .Add , "PickListKey", "Key"
        lvwPT.ColumnHeaders("PickListKey").width = 0
    End With

    On Error Resume Next
    With rst
        Do While Not .EOF
            If Not IsNull(.Fields("SalesOrder").value) Then
                Set olstItem = lvwSO.ListItems.Add(, .Fields("SalesOrder").value, .Fields("SalesOrder").value)
                If Not olstItem Is Nothing Then
                    olstItem.SubItems(1) = .Fields("SODate").value
                    olstItem.SubItems(2) = .Fields("SOKey").value
                    Set olstItem = Nothing
                End If
            End If
        
            If Not IsNull(.Fields("Invoice").value) Then
                Set olstItem = lvwInv.ListItems.Add(, .Fields("Invoice").value, .Fields("Invoice").value)
                If Not olstItem Is Nothing Then
                    olstItem.SubItems(1) = .Fields("InvDate").value
                    olstItem.SubItems(2) = .Fields("InvcKey").value
                    Set olstItem = Nothing
                End If
            End If

            If Not IsNull(.Fields("Shipment").value) Then
                Set olstItem = lvwShip.ListItems.Add(, .Fields("Shipment").value, .Fields("Shipment").value)
                If Not olstItem Is Nothing Then
                    olstItem.SubItems(1) = .Fields("ShipDate").value
                    olstItem.SubItems(2) = .Fields("ShipKey").value
                    Set olstItem = Nothing
                End If
            End If
            
            'PRN391
            If Not IsNull(.Fields("PendShipment").value) Then
                Set olstItem = lvwShip.ListItems.Add(, .Fields("PendShipment").value, .Fields("PendShipment").value)
                If Not olstItem Is Nothing Then
                    olstItem.SubItems(1) = .Fields("PendShipDate").value
                    olstItem.SubItems(2) = .Fields("PendShipKey").value
                    Set olstItem = Nothing
                End If
            End If
            
            If Not IsNull(.Fields("PickList").value) Then
                Set olstItem = lvwPT.ListItems.Add(, .Fields("PickList").value & "-PL", .Fields("PickList").value)
                If Not olstItem Is Nothing Then
                    olstItem.SubItems(1) = .Fields("PickDate").value
                    olstItem.SubItems(2) = .Fields("PickListKey").value
                    Set olstItem = Nothing
                End If
            End If
            
            .MoveNext
        Loop
    End With
End Sub


Private Sub Frame1ControlResize()
    Dim lBorder As Long
    Dim ControlWidth As Long
    Dim newWidth As Long
    Dim i As Long
    
    lBorder = 240
    ControlWidth = cmdFindDocs.width + cboDocs.width + _
                    txtDocNbr.width + Label1.width
    newWidth = (Frame1.width - ControlWidth - 2 * lBorder) / 3

    cmdFindDocs.Left = Frame1.width - cmdFindDocs.width - 240
    cboDocs.Left = lBorder + txtDocNbr.width + Label1.width + newWidth
    
    With gdxDocs
        lBorder = 360
        .width = Me.width - 3 * lBorder
        .Height = Me.Height - (4 * lBorder + 120)
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    
End Sub


Private Sub Frame2ControlResize()
    Dim lBorder As Long
    Dim lWidth As Long
    Dim lHeight As Long
    Dim lTop As Long
    Dim lSecTop As Long
    Dim ControlWidth As Long
    Dim newWidth As Long
    Dim lListWidth As Long
    Dim lListHeight As Long
    
    lBorder = 240
    lWidth = 640
    lTop = 960
    lSecTop = 350
    ControlWidth = Label2.width + lblCustName.width + lblSrcDoc.width + cmdBackTo1.width
    newWidth = (Frame2.width - ControlWidth - 2 * lBorder) / 2
    
    cmdBackTo1.Left = Frame2.width - cmdBackTo1.width - lBorder
    lblCustName.Left = cmdBackTo1.Left - newWidth - lblCustName.width
    
    lListWidth = (Frame2.width - lWidth - 2 * lBorder) / 2
    lListHeight = (Frame2.Height - lTop - lSecTop - lBorder / 2) / 2
    
    With lvwInv
        .Height = lListHeight
        .width = lListWidth
        .Left = lListWidth + lBorder + lWidth
        Label4.Left = .Left
    End With
    
    With lvwSO
        .Height = lListHeight
        .width = lListWidth
    End With
    
    With lvwShip
        .Height = lListHeight
        .width = lListWidth
        .Top = Frame2.Height - lBorder / 2 - .Height
        Label5.Top = .Top - Label5.Height
    End With
    
    With lvwPT
        .Height = lListHeight
        .width = lListWidth
        .Left = lListWidth + lBorder + lWidth
        .Top = Frame2.Height - lBorder / 2 - .Height
        Label6.Left = .Left
        Label6.Top = .Top - Label6.Height
    End With
End Sub


Private Sub Frame3ControlResize()
    Dim lBorder As Long
    Dim lTop As Long
    
    lTop = 360
    lBorder = 2385
    With lvwDocDetail
        .width = Frame3.width - lBorder
        cmdBackTo2.Left = .width + .Left + 350
        .Height = Frame3.Height - 2 * lTop
    End With
End Sub

Private Sub lvwInv_DblClick()
    If lvwInv.SelectedItem Is Nothing Then Exit Sub
    
    'to support email function
    m_lInvcKey = lvwInv.SelectedItem.SubItems(2)
    
    LoadFrame3 "Invoice", lvwInv.SelectedItem.SubItems(2)
End Sub


Private Sub lvwPT_DblClick()
    If lvwPT.SelectedItem Is Nothing Then Exit Sub
    
    LoadFrame3 "PickTicket", lvwPT.SelectedItem.SubItems(2)
End Sub


Private Sub lvwShip_DblClick()
    If lvwShip.SelectedItem Is Nothing Then Exit Sub
    
    LoadFrame3 "Shipment", lvwShip.SelectedItem.SubItems(2)
End Sub


Private Sub lvwSO_DblClick()
    If lvwSO.SelectedItem Is Nothing Then Exit Sub
    
    LoadFrame3 "SalesOrder", lvwSO.SelectedItem.SubItems(2)
End Sub


Private Sub LoadDetailView(rst As ADODB.Recordset)
    Dim oFld As ADODB.field
    Dim oListItem As ListItem
    
    lvwDocDetail.ColumnHeaders.Clear
    lvwDocDetail.ListItems.Clear
    
    lvwDocDetail.ColumnHeaders.Add , "Field", "Field"
    lvwDocDetail.ColumnHeaders.Add , "Value", "Value"
        
    For Each oFld In rst.Fields
        Set oListItem = lvwDocDetail.ListItems.Add(, , oFld.Name)
        oListItem.SubItems(1) = oFld.value
    Next
End Sub


Private Sub FindRelDocs(sDocType As String)
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset

    SetWaitCursor True
    Set cmd = CreateCommandSP("spCPOPRelatedDocs")

    Set rst = New ADODB.Recordset
    
    With cmd
        Select Case sDocType
        Case "Sales Order":
            .Parameters("@SOTranID").value = Trim(lblSrcDoc.caption)
        Case "Shipment":
            .Parameters("@SHTranID").value = Trim(lblSrcDoc.caption)
        Case "Invoice":
            .Parameters("@INTranID").value = Trim(lblSrcDoc.caption)
        Case "Pick Ticket":
            .Parameters("@PTTranID").value = Trim(lblSrcDoc.caption)
        Case "OP":
            .Parameters("@OPID").value = Trim(lblSrcDoc.caption)
        End Select
        rst.Open cmd
    End With
    
    LoadListViews rst
    SetWaitCursor False

    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
End Sub


Private Sub m_gw_RowChosen()
    With m_gw
        LoadFrame2 .value("TranID"), .value("CustID") & " " & .value("CustName"), .value("DocType")
    End With
End Sub


Private Sub LoadDocCombo()
    With cboDocs
        .Clear
        .AddItem "OP"
        .AddItem "Sales Order"
        .AddItem "Pick Ticket"
        .AddItem "Shipment"
        .AddItem "Invoice"
        .text = "Sales Order"
    End With
    
End Sub


Private Sub txtDocNbr_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtDocNbr.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            cmdFindDocs_Click
        End If
    End If
End Sub


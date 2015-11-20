VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FPhoneFlagger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phone Flagger Notes"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   8550
   Begin MSComctlLib.ListView lvwResearch 
      Height          =   1695
      Left            =   120
      TabIndex        =   26
      Top             =   5880
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame frmLineDetail 
      Caption         =   "Order Line Detail"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   8295
      Begin VB.TextBox txtSerial 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtModel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtMake 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDescr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   660
         Width           =   4695
      End
      Begin VB.TextBox txtCustID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtOPID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial"
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Model"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Make"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer"
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "OP#"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdLoadOrder 
      Caption         =   "&Load Order"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin GridEX20.GridEX gdxOrder 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3836
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
      ColumnsCount    =   5
      Column(1)       =   "FPhoneFlagger.frx":0000
      Column(2)       =   "FPhoneFlagger.frx":0148
      Column(3)       =   "FPhoneFlagger.frx":025C
      Column(4)       =   "FPhoneFlagger.frx":0388
      Column(5)       =   "FPhoneFlagger.frx":04A0
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPhoneFlagger.frx":05EC
      FormatStyle(2)  =   "FPhoneFlagger.frx":0724
      FormatStyle(3)  =   "FPhoneFlagger.frx":07D4
      FormatStyle(4)  =   "FPhoneFlagger.frx":0888
      FormatStyle(5)  =   "FPhoneFlagger.frx":0960
      FormatStyle(6)  =   "FPhoneFlagger.frx":0A18
      ImageCount      =   0
      PrinterProperties=   "FPhoneFlagger.frx":0AF8
   End
   Begin VB.Frame Frame1 
      Caption         =   "Need-To-Research Orders"
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.TextBox lblCreateDay 
         Height          =   285
         Left            =   5040
         TabIndex        =   24
         Top             =   360
         Width           =   375
      End
      Begin MSComCtl2.UpDown UDDays 
         Height          =   300
         Left            =   5480
         TabIndex        =   7
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   7
         BuddyControl    =   "lblCreateDay"
         BuddyDispid     =   196628
         OrigLeft        =   4800
         OrigTop         =   360
         OrigRight       =   5040
         OrigBottom      =   615
         Max             =   60
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Day(s)"
         Height          =   255
         Left            =   5880
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Order Created in Last "
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblBranch 
         Caption         =   "Branch ID"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Recent order line items that contain matching Model and Serial number."
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5640
      Width           =   5175
   End
End
Attribute VB_Name = "FPhoneFlagger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWindowID As Long
Private WithEvents m_gwLines As GridEXWrapper
Attribute m_gwLines.VB_VarHelpID = -1


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
    ShowHelp "FPhoneFlagger", True
End Sub


'Form Event Handlers

Private Sub Form_Load()
   
    SetCaption "Phone Flagger"
    
    lblCreateDay.text = UDDays.value

    '***note: GetUserWhseKey sets a filter on g_rstWhses too, so need to do this separately
    
    g_rstWhses.Filter = "transit = 0"
    LoadCombo cboBranch, g_rstWhses, "WhseID", "WhseKey", GetUserWhseKey
    g_rstWhses.Filter = adFilterNone
    
    Set m_gwLines = New GridEXWrapper
    m_gwLines.Grid = gdxOrder
    
    lvwResearch.Gridlines = True
    lvwResearch.LabelEdit = lvwManual
    lvwResearch.MultiSelect = False
    lvwResearch.View = lvwReport
    lvwResearch.ColumnHeaders.Add , , "OP#", 900, lvwColumnLeft
    lvwResearch.ColumnHeaders.Add , , "Date", 1100, lvwColumnLeft
    lvwResearch.ColumnHeaders.Add , , "Customer", 1640, lvwColumnLeft
    lvwResearch.ColumnHeaders.Add , , "Description", 2360, lvwColumnLeft
    lvwResearch.ColumnHeaders.Add , , "Part#", 1400, lvwColumnLeft
    lvwResearch.ColumnHeaders.Add , , "Cost", 810, lvwColumnLeft
    
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_gwLines = Nothing
    MDIMain.UnloadTool m_lWindowID
End Sub


'Control Event Handlers

Private Sub cmdLoadOrder_Click()
    If IsEmpty(m_gwLines.value("OPKey")) Then Exit Sub
    
    Dim lOPKey As Long
    Dim oFrm As FOrder
    
    SetWaitCursor True
    lOPKey = m_gwLines.value("OPKey")

    LogEvent "FPhoneFlagger", "cmdLoadOrder_Click", GetUserName & " instantiating FOrder from FPhoneFlagger for OP " & lOPKey
    
    Set oFrm = New FOrder
    MDIMain.AddNewWindow oFrm
    With oFrm
        .Show
        .Order.Load lOPKey
        .lblCustName.Visible = True
        .lblCustType(0).Visible = True
        .txtCustName.Visible = False
        .cboCustType.Visible = False

        .TransitionTabs False
    End With
    SetWaitCursor False
End Sub


Private Sub cmdRefresh_Click()
    Dim orst As ADODB.Recordset
    
    SetWaitCursor True
    Set orst = CallSP("spcpcPhoneFlagger", "@iDays", CLng(lblCreateDay.text), "@iBranchID", cboBranch.List(cboBranch.ListIndex))
    
    With gdxOrder
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
    End With
    
    Set orst = Nothing
    SetWaitCursor False
End Sub


Private Sub gdxOrder_SelectionChange()
    Dim orst As ADODB.Recordset

    SetWaitCursor True
    lvwResearch.ListItems.Clear
        
    If IsEmpty(m_gwLines.value("OPLineKey")) Then Exit Sub
    
    Set orst = CallSP("spCPCPhoneFlaggerDetail", "@iOPLineKey", m_gwLines.value("OPLineKey"))
    
    If Not orst.EOF Then
        txtDate.text = orst.Fields("CreateDate").value
        txtOPID.text = m_gwLines.value("OPKey")
        txtCustID.text = Trim(orst.Fields("CustID")) & ": " & Trim(orst.Fields("CustName").value)
        txtCustID.ToolTipText = txtCustID.text
        txtDescr.text = Trim(orst.Fields("Descr").value)
        txtMake.text = Trim(orst.Fields("MakeText").value)
        txtModel.text = Trim(orst.Fields("CabModelNbr").value)
        txtSerial.text = Trim(orst.Fields("CabSerialNbr").value)
        
        GetResearch lblCreateDay.text, m_gwLines.value("OPLineKey")
    Else
        txtOPID.text = ""
        txtDate.text = ""
        txtCustID.text = ""
        txtDescr.text = ""
        txtMake.text = ""
        txtModel.text = ""
        txtSerial.text = ""
    End If
    
    Set orst = Nothing

    SetWaitCursor False
End Sub


Private Sub GetResearch(lDays As Long, lOPLineKey As Long)
    Dim orst As ADODB.Recordset
    Dim itmVar As ListItem
    
    Set orst = CallSP("spCPCPhoneFlaggerResearch", "@iDays", lDays, "@iOPLineKey", lOPLineKey)
    
    Do While Not orst.EOF
        Set itmVar = lvwResearch.ListItems.Add(, , orst.Fields("OPNo").value)
        itmVar.SubItems(1) = orst.Fields("UpdateDate").value
        itmVar.SubItems(2) = orst.Fields("CustName").value
        itmVar.SubItems(3) = orst.Fields("Descr").value
        itmVar.SubItems(4) = orst.Fields("PartNo").value
        itmVar.SubItems(5) = orst.Fields("Cost").value
        orst.MoveNext
    Loop
    
    Set orst = Nothing
End Sub


Private Sub lblCreateDay_LostFocus()
    On Error GoTo ErrorHandler
    
    If IsNumeric(lblCreateDay.text) Then
        If CLng(lblCreateDay.text) > 60 Then
            msg "The maximum create date is 60 days"
            lblCreateDay.text = 60
            TryToSetFocus lblCreateDay
        End If
        
        cmdRefresh.Enabled = True
        UDDays.Enabled = True
    ElseIf Trim(lblCreateDay.text) = "" Then
        lblCreateDay.text = 7
        cmdRefresh.Enabled = True
        UDDays.Enabled = True
    Else
        msg "Please enter valid Create Date"
        lblCreateDay.SelStart = 0
        lblCreateDay.SelLength = Len(lblCreateDay.text)
        TryToSetFocus lblCreateDay
        cmdRefresh.Enabled = False
        UDDays.Enabled = False
    End If
    
    Exit Sub
    
ErrorHandler:
    msg "There is an error for inputting Create Date. Please check and input Create Date again"
    TryToSetFocus lblCreateDay
    cmdRefresh.Enabled = False
End Sub


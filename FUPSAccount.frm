VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FUPSAcct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   7815
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   4215
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7435
      _Version        =   262144
      TabCount        =   2
      Tabs            =   "FUPSAccount.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   3825
         Left            =   -99969
         TabIndex        =   7
         Top             =   360
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   6747
         _Version        =   262144
         TabGuid         =   "FUPSAccount.frx":0093
         Begin VB.Frame Frame1 
            Height          =   1455
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   7335
            Begin VB.ComboBox cboCSWShipper 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtCSWCoName 
               Height          =   285
               Left            =   1080
               TabIndex        =   22
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtCSWUPSTrkNo 
               Height          =   285
               Left            =   5400
               TabIndex        =   25
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtCSWOPNo 
               Height          =   285
               Left            =   4320
               TabIndex        =   24
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtCSWPkgID 
               Height          =   285
               Left            =   3000
               TabIndex        =   23
               Top             =   960
               Width           =   1215
            End
            Begin VB.CommandButton cmdCSWFind 
               Caption         =   "Find"
               Height          =   288
               Left            =   6240
               TabIndex        =   26
               Top             =   240
               Width           =   975
            End
            Begin MSComCtl2.DTPicker dtpCSWBeginTime 
               Height          =   315
               Left            =   1440
               TabIndex        =   19
               Top             =   195
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   102825985
               CurrentDate     =   37312
            End
            Begin MSComCtl2.DTPicker dtpCSWEndTime 
               Height          =   315
               Left            =   3120
               TabIndex        =   20
               Top             =   195
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   102825985
               CurrentDate     =   37312
            End
            Begin VB.Label Label10 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2950
               TabIndex        =   16
               Top             =   130
               Width           =   135
            End
            Begin VB.Label Label9 
               Caption         =   "Ship Date Range:"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label8 
               Caption         =   "Company Name"
               Height          =   255
               Left            =   1080
               TabIndex        =   13
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "UPS Tracking Number"
               Height          =   255
               Left            =   5400
               TabIndex        =   12
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label6 
               Caption         =   "OP Number"
               Height          =   255
               Left            =   4320
               TabIndex        =   11
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Package ID"
               Height          =   255
               Left            =   3000
               TabIndex        =   10
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Warehouse"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   720
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Results"
            Height          =   2175
            Left            =   120
            TabIndex        =   14
            Top             =   1560
            Width           =   7335
            Begin VB.ListBox lstCSWResults 
               Height          =   1425
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   7095
            End
            Begin VB.CommandButton cmdCSWDisplay 
               Caption         =   "Display"
               Height          =   288
               Left            =   6240
               TabIndex        =   28
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lblCSWResults 
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   1800
               Width           =   1815
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3825
         Left            =   30
         TabIndex        =   0
         Top             =   360
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   6747
         _Version        =   262144
         TabGuid         =   "FUPSAccount.frx":00BB
         Begin VB.Frame frmUPSSearch 
            Caption         =   "Load Account"
            Height          =   735
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   7332
            Begin VB.TextBox txtCustID 
               Height          =   285
               Left            =   780
               TabIndex        =   5
               Top             =   300
               Width           =   1575
            End
            Begin VB.CommandButton cmdFind 
               Caption         =   "Find"
               Height          =   288
               Left            =   2580
               TabIndex        =   4
               Top             =   300
               Width           =   975
            End
            Begin VB.Label lblCustID 
               Caption         =   "CustID"
               Height          =   252
               Left            =   180
               TabIndex        =   6
               Top             =   360
               Width           =   492
            End
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Update"
            Height          =   320
            Left            =   6420
            TabIndex        =   1
            Top             =   3480
            Width           =   975
         End
         Begin GridEX20.GridEX gdxUPSAcct 
            Height          =   2412
            Left            =   60
            TabIndex        =   2
            Top             =   900
            Width           =   7332
            _ExtentX        =   12938
            _ExtentY        =   4260
            Version         =   "2.0"
            ScrollToolTips  =   -1  'True
            ShowToolTips    =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   10
            Column(1)       =   "FUPSAccount.frx":00E3
            Column(2)       =   "FUPSAccount.frx":0277
            Column(3)       =   "FUPSAccount.frx":03C7
            Column(4)       =   "FUPSAccount.frx":0527
            Column(5)       =   "FUPSAccount.frx":077B
            Column(6)       =   "FUPSAccount.frx":08CB
            Column(7)       =   "FUPSAccount.frx":0A1B
            Column(8)       =   "FUPSAccount.frx":0B6B
            Column(9)       =   "FUPSAccount.frx":0CA3
            Column(10)      =   "FUPSAccount.frx":0DE3
            FormatStylesCount=   6
            FormatStyle(1)  =   "FUPSAccount.frx":0F3F
            FormatStyle(2)  =   "FUPSAccount.frx":1077
            FormatStyle(3)  =   "FUPSAccount.frx":1127
            FormatStyle(4)  =   "FUPSAccount.frx":11DB
            FormatStyle(5)  =   "FUPSAccount.frx":12B3
            FormatStyle(6)  =   "FUPSAccount.frx":136B
            ImageCount      =   0
            PrinterProperties=   "FUPSAccount.frx":144B
         End
      End
   End
End
Attribute VB_Name = "FUPSAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***464 ***
'Private Const kUPSStmtFilePath = "\\cpfnp\shared\UPS"

'This form uses an object collection (of type VaxAcctParaList) with an unbound grid.
'VaxAcctParaList is also used by FVaxAcct.frm.

Private m_lWindowID As Long

Private m_oUPSAccts As UPSAcctList


'******************************************************************************
' Public Properties and Methods
'******************************************************************************

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
    ShowHelp "FUPSAcct", True
End Sub


'*******************************************************************************
' Form events
'*******************************************************************************

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Load()
    SetCaption "UPS Account Manager"
    With gdxUPSAcct
        .ItemCount = 0
        .Refetch
    End With
    
    'SMR (10-12-05) ConnectShip Find tab
    Dim lWhseKey As Long
    lWhseKey = GetUserWhseKey
    g_rstWhses.Filter = "transit = 0"
    LoadCombo cboCSWShipper, g_rstWhses, "WhseID", "WhseKey", lWhseKey
    g_rstWhses.Filter = adFilterNone
    
    dtpCSWBeginTime.value = Now
    dtpCSWEndTime.value = Now
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


'Private Sub Form_Resize()
'    If Me.WindowState = 1 Then Exit Sub
'
'    gdxUPSAcct.Width = Me.Width - 310
'    gdxUPSAcct.Height = Me.Height - 1980
'    cmdUpdate.Top = gdxUPSAcct.Top + gdxUPSAcct.Height + 120
'    cmdUpdate.Left = gdxUPSAcct.Left + gdxUPSAcct.Width - cmdUpdate.Width
'    gdxUPSAcct.Refresh
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.UnloadTool m_lWindowID
    Set m_oUPSAccts = Nothing
End Sub


'*******************************************************************************
' Control events
'*******************************************************************************

Private Sub cmdFind_Click()
    Dim oCustomer As Customer
    Dim lCustKey As Long
    Dim sCustID As String
    
    Set oCustomer = New Customer
    
    gdxUPSAcct.ItemCount = 0
    gdxUPSAcct.Refetch
    
    If Trim(txtCustID.text) = "" Then
        TryToSetFocus txtCustID
        Exit Sub
    End If
    
    If lCustKeyAcct(lCustKey, sCustID) > 1 Then
        lCustKey = Search.FindCustomer(Trim(txtCustID.text), 0, oCustomer)
        oCustomer.Load lCustKey
        sCustID = oCustomer.ID
    End If
    
    If lCustKey = 0 Then
        txtCustID.SelStart = 0
        txtCustID.SelLength = Len(txtCustID.text)
        TryToSetFocus txtCustID
        Exit Sub
    Else
        txtCustID.text = sCustID
        Set m_oUPSAccts = New UPSAcctList
        m_oUPSAccts.LoadUpsAcct lCustKey
        
        If m_oUPSAccts.Count > 0 Then
            With gdxUPSAcct
                .HoldFields
                .ItemCount = m_oUPSAccts.Count
                .Refetch
                .Row = 1
                TryToSetFocus gdxUPSAcct
            End With
        End If
    End If
End Sub


Private Sub cmdUpdate_Click()
    On Error GoTo ErrorHandler
    
    Dim cmd As ADODB.Command
    Dim lIndex As Long
    
    If (m_oUPSAccts Is Nothing) Then Exit Sub
    If m_oUPSAccts.Count = 0 Then Exit Sub

    If Not bUpdated Then
        msg "There are no address selected for UPS updating!", vbOKOnly + vbExclamation, "Updating UPS Acct"
    Else
        With m_oUPSAccts
            For lIndex = 1 To m_oUPSAccts.Count
                If .Item(lIndex).Selected Then
                    Set cmd = New ADODB.Command
                    Set cmd = CreateCommandSP("Delete tcpUPSAcct where CustAddrKey = " & .Item(lIndex).AddrKey, adCmdText)
                    cmd.Execute
                    If Trim(.Item(lIndex).UPSAcct) <> "" Then
                        Set cmd = New ADODB.Command
                        Set cmd = CreateCommandSP("Insert into tcpUPSAcct Values(" & .Item(lIndex).AddrKey _
                                        & ", '" & .Item(lIndex).UPSAcct & "')", adCmdText)
                        cmd.Execute
                    End If
                End If
            Next
            msg "UPS Acct updating succeeds", vbOKOnly + vbExclamation, "Updating Results"
            cmdFind_Click
            Set cmd = Nothing
        End With
    End If
    Exit Sub
    
ErrorHandler:
    msg Err.Description, vbOKOnly + vbCritical, Err.Source
    ClearWaitCursor
End Sub


Private Sub txtCustID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFind_Click
    End If
End Sub


Private Sub txtCustID_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub


Private Sub gdxUPSAcct_LostFocus()
    gdxUPSAcct.Update
End Sub


Private Sub gdxUPSAcct_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxUPSAcct.Update
End Sub


Private Sub gdxUPSAcct_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_oUPSAccts Is Nothing Then Exit Sub
    If RowIndex > m_oUPSAccts.Count Then Exit Sub
    
    With m_oUPSAccts.Item(RowIndex)
        Values(1) = .Selected
        Values(2) = .UPSAcct
        Values(3) = .AddrKey
        Values(4) = .AddrType
        Values(5) = .AddrName
        Values(6) = .AddrLine1
        Values(7) = .AddrLine2
        Values(8) = .City
        Values(9) = .StateID
        Values(10) = .PostalCode
    End With
End Sub


Private Sub gdxUPSAcct_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    m_oUPSAccts.Item(RowIndex).Selected = Values(1)
    m_oUPSAccts.Item(RowIndex).UPSAcct = Values(2)
End Sub


'*******************************************************************************
' Control Events - ConnectShip Find Tab
'    cmdCSWFind_Click
'    cmdCSWDisplay_Click
'    lstCSWResults_DblClick
'*******************************************************************************
Private Sub cmdCSWFind_Click()
    Dim lsSql As String
    Dim moConnect As ADODB.Connection
    Dim lors As ADODB.Recordset

    SetWaitCursor True
    Set moConnect = New ADODB.Connection
    Set lors = New ADODB.Recordset

    lstCSWResults.Clear

    moConnect.ConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Password=" & g_DB.Password & ";Initial Catalog=CSWReportDatabase;Data Source=" & g_DB.Server
    moConnect.CursorLocation = adUseClient
    moConnect.Open

    lsSql = "SELECT * From Packages INNER JOIN package_lists ON" & _
       " packages.packagelist_id = package_lists.packagelist_id" & _
       " WHERE shipper_symbol = '" & cboCSWShipper.text & "' AND packages.reference_2 NOT LIKE 'RMA%'" & _
       " AND shipdate >= '" & Format(dtpCSWBeginTime.value, "MM/DD/YYYY") & "'" & _
       " AND shipdate <= '" & Format(dtpCSWEndTime.value, "MM/DD/YYYY") & "'"
       If Len(Trim(txtCSWCoName)) > 0 Then lsSql = lsSql & " AND Consignee_Company = '" & Replace(Trim(txtCSWCoName), "'", "''") & "'"
       If Len(Trim(txtCSWPkgID)) > 0 And IsNumeric(Trim(txtCSWPkgID)) = True Then lsSql = lsSql & " AND Reference_1 = '" & Trim(txtCSWPkgID) & "'"
       If Len(Trim(txtCSWOPNo)) > 0 And IsNumeric(Trim(txtCSWOPNo)) = True Then lsSql = lsSql & " AND Reference_2 = '" & Trim(txtCSWOPNo) & "'"
       If Len(Trim(txtCSWUPSTrkNo)) > 0 Then lsSql = lsSql & " AND Tracking_Number = '" & Replace(Trim(txtCSWUPSTrkNo), "'", "''") & "'"
       lsSql = lsSql & " order by msn"
    
    '***464 - SMR - Removed void filter from above query; voided pkgs now show in red on the pkg history rpt
    'AND void_flag = 0
    
    lors.Source = lsSql
    Set lors.ActiveConnection = moConnect
    lors.Open
    If lors.EOF Then lblCSWResults = "": MsgBox "No results found.": GoTo ClearObjects
    
    Do Until lors.EOF
        lstCSWResults.AddItem "Tran# " & lors.Fields("reference_1") & "    OP-" & lors.Fields("reference_2") & "    " & lors.Fields("Consignee_Company") & "  -  MSN#" & lors.Fields("msn")
        lors.MoveNext
    Loop
    lblCSWResults.Caption = "Found " & lstCSWResults.ListCount & " matches."
    lstCSWResults.Selected(0) = True
    
ClearObjects:
    Set lors = Nothing
    moConnect.Close
    Set moConnect = Nothing
    SetWaitCursor False
End Sub

Private Sub cmdCSWDisplay_Click()
    If lstCSWResults.ListCount > 0 Then
        Dim llHolder As Long
        Dim oFrm As FViewer

        SetWaitCursor True
        Set oFrm = New FViewer
        llHolder = Mid(lstCSWResults.text, InStr(1, lstCSWResults.text, "MSN#") + 4, Len(lstCSWResults.text))
        Call oFrm.ParamAdd(1, "@MSN", llHolder)
        Call oFrm.ViewReportByType("CSWPkgHistory")
        Set oFrm = Nothing
        SetWaitCursor False
    End If
End Sub

Private Sub lstCSWResults_DblClick()
    Call cmdCSWDisplay_Click
End Sub


Private Function bUpdated() As Boolean
    Dim lIndex As Long
    
    For lIndex = 1 To m_oUPSAccts.Count
        If m_oUPSAccts.Item(lIndex).Selected Then
            bUpdated = True
        End If
    Next
End Function


' Returns lCustKey and sCustID by reference

Private Function lCustKeyAcct(ByRef lCustKey As Long, ByRef sCustID As String) As Long
    Dim rst As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT DISTINCT c.CustKey, c.CustID " _
         & "FROM tarCustomer c " _
         & "INNER JOIN tarCustAddr ca ON c.CustKey = ca.CustKey " _
         & "INNER JOIN tciAddress a ON ca.AddrKey = a.AddrKey " _
         & "WHERE ((c.CompanyID = 'CPC') AND (c.Status = 1)) AND " _
         & "c.CustID LIKE '" & Trim(txtCustID.text) & "%'"
    Set rst = LoadDiscRst(sSQL)
    
    lCustKeyAcct = rst.RecordCount
    If rst.RecordCount = 1 Then
        lCustKey = rst.Fields("CustKey").value
        sCustID = Trim(rst.Fields("CustID").value)
    End If
End Function


Private Sub GetCSWFields(ByVal sTrackNo As String, ByRef sRef1 As String, ByRef iTotalPkgCount As Integer)
    Dim lsSql As String
    Dim moConnect As ADODB.Connection
    Dim lors As ADODB.Recordset

    Set moConnect = New ADODB.Connection
    Set lors = New ADODB.Recordset

    moConnect.ConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Initial Catalog=CSWReportDatabase;Data Source=CPSQLPRO"
    moConnect.CursorLocation = adUseClient
    moConnect.Open

    lsSql = "SELECT reference_1, nofn_total  From Packages " & _
        "INNER JOIN package_lists ON packages.packagelist_id = package_lists.packagelist_id " & _
        "WHERE tracking_number = '" & sTrackNo & "'"
    
    lors.Source = lsSql
    Set lors.ActiveConnection = moConnect
    lors.Open
    If lors.EOF Then
        sRef1 = ""
        iTotalPkgCount = 0
    Else
        sRef1 = lors.Fields("reference_1")
        iTotalPkgCount = lors.Fields("nofn_total")
    End If
    
End Sub



'*** NO LONGER USED (5/26/04 LR) ***

Private Function ValidTrackingNumber(ByVal sTrackingNumber As String, ByRef sRemark As String) As Boolean
    If Len(sTrackingNumber) < 18 Then
        sRemark = " - tracking number is too short"
        ValidTrackingNumber = False
    ElseIf InStr(1, sTrackingNumber, "1Z815210") Then
        sRemark = vbNullString
        ValidTrackingNumber = True
    ElseIf InStr(1, sTrackingNumber, "1ZV65W05") Then
        sRemark = vbNullString
        ValidTrackingNumber = True
    ElseIf InStr(1, sTrackingNumber, "1Z665660") Then
        sRemark = vbNullString
        ValidTrackingNumber = True
    Else
        sRemark = " - not a CPC shipper #"
        ValidTrackingNumber = False
    End If
End Function


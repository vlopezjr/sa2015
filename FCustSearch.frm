VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FCustSearch 
   Caption         =   "Customer Search"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFilterByTerritory 
      Caption         =   "Only show addresses in my sales territory"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   4620
      Value           =   2  'Grayed
      Width           =   3735
   End
   Begin GridEX20.GridEX gdxCustSearch 
      Height          =   4095
      Left            =   0
      TabIndex        =   2
      Top             =   180
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7223
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      ScrollToolTips  =   -1  'True
      ShowToolTips    =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   "Addr2"
      CursorLocation  =   3
      ReadOnly        =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      CardCaptionPrefix=   "Customer Information"
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "FCustSearch.frx":0000
      DataMode        =   99
      ColumnHeaderHeight=   285
      ColumnsCount    =   16
      Column(1)       =   "FCustSearch.frx":0C52
      Column(2)       =   "FCustSearch.frx":0D76
      Column(3)       =   "FCustSearch.frx":0FAA
      Column(4)       =   "FCustSearch.frx":10E2
      Column(5)       =   "FCustSearch.frx":11FA
      Column(6)       =   "FCustSearch.frx":1306
      Column(7)       =   "FCustSearch.frx":1412
      Column(8)       =   "FCustSearch.frx":154A
      Column(9)       =   "FCustSearch.frx":1682
      Column(10)      =   "FCustSearch.frx":19F2
      Column(11)      =   "FCustSearch.frx":1BAE
      Column(12)      =   "FCustSearch.frx":1D06
      Column(13)      =   "FCustSearch.frx":1E46
      Column(14)      =   "FCustSearch.frx":1F5E
      Column(15)      =   "FCustSearch.frx":2046
      Column(16)      =   "FCustSearch.frx":2156
      SortKeysCount   =   2
      SortKey(1)      =   "FCustSearch.frx":2266
      SortKey(2)      =   "FCustSearch.frx":22CE
      FmtConditionsCount=   2
      FmtCondition(1) =   "FCustSearch.frx":2336
      FmtCondition(2) =   "FCustSearch.frx":241E
      FormatStylesCount=   7
      FormatStyle(1)  =   "FCustSearch.frx":2552
      FormatStyle(2)  =   "FCustSearch.frx":2632
      FormatStyle(3)  =   "FCustSearch.frx":276A
      FormatStyle(4)  =   "FCustSearch.frx":281A
      FormatStyle(5)  =   "FCustSearch.frx":28CE
      FormatStyle(6)  =   "FCustSearch.frx":29A6
      FormatStyle(7)  =   "FCustSearch.frx":2A5E
      ImageCount      =   1
      ImagePicture(1) =   "FCustSearch.frx":2B0E
      PrinterProperties=   "FCustSearch.frx":3760
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   7260
      TabIndex        =   1
      Top             =   4560
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   6120
      TabIndex        =   0
      Top             =   4560
      Width           =   1092
   End
   Begin VB.Label lblTooMany 
      Caption         =   "WARNING: Your search was too general. Only the first 250 matches are shown here."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   60
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   8292
   End
End
Attribute VB_Name = "FCustSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This dialog's function is influenced by the value of the global variable g_bFilterCustSearch
'which defaults to True on startup and is overridden/controlled by a Toolbar setting.

'Note: This custom dialog unloads itself (rather than hiding itself and leaving the unload
'the responsibility of its client). This is because this dialog communicates everything back
'to its client through reference parameters and function return values, not properties.

Private Const k_lMinWidth = 7800
Private Const k_lMinHeight = 3600

'*** this is for bound grids. needs to come out.
'Private WithEvents m_gw As GridEXWrapper

'This recordset is now simply a transport mechanism. it's no longer bound to the grid.
'Note: it's recordcount can now differ from the row count of the new array.
Private m_rst As ADODB.Recordset
Private m_arrayCustomers() As Variant

'cache the user's sales territory key
Private m_lTerritoryKey As Long

'cache partial SQL that created m_rst
Private m_sSearchSQL As String

Private m_bLoadTheAddress As Boolean
Private m_bInitializing As Boolean

'set to true by FindByCustKey()
'effects logic in chkFilterByTerritory_Click()
Private m_bSearchViaCustKey As Boolean


'**********************************************************************
' Form Events
'**********************************************************************

Private Sub Form_Activate()
    With gdxCustSearch
        TryToSetFocus gdxCustSearch
        If .RowCount >= 1 Then
            '9/4/03 LR changed from 2 to 1 to highlite the group header (at request of GT)
            .Row = 1
        End If
    End With
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        DoShowHelp
    Else
        MDIMain.GlobalKeyDownProcessing KeyCode, Shift
    End If
End Sub


Private Sub Form_Resize()
    Dim lBorder As Long

    If WindowState <> 0 Then Exit Sub
    
    If Me.width < k_lMinWidth Then Me.width = k_lMinWidth
    If Me.Height < k_lMinHeight Then Me.Height = k_lMinHeight

    With gdxCustSearch
        lBorder = 60
        .width = Me.width - (2 * lBorder + 120)
        .Height = Me.Height - 1450
    End With

    lblTooMany.Top = gdxCustSearch.Top + gdxCustSearch.Height + 120
    cmdOK.Top = lblTooMany.Top + 250
    cmdCancel.Top = cmdOK.Top
    cmdCancel.Left = Me.width - (cmdCancel.width + 120 + lBorder)
    cmdOK.Left = cmdCancel.Left - (cmdOK.width + lBorder)
    chkFilterByTerritory.Top = lblTooMany.Top + lblTooMany.Height + 120
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Debug.Print Me.Name & " Form_Unload"
    Set m_rst = Nothing

End Sub


'**********************************************************************
' Public Methods
'**********************************************************************

Public Sub DoShowHelp()
    ShowHelp "CustSearch", True
End Sub


'Called By
'    ValidateCustomer.Validate
'    Search.FindCustomer
'
'This method receives a Customer object BYREF
'It loads the Customer's Address objects
'Returns a CustKey or 0.

Public Function Find(i_sCaption As String, _
                        i_sInput As String, _
                        i_oCust As Customer, _
                        Optional bFindOnly As Boolean = False) As Long

    Dim sSQL As String
    
    SetWaitCursor True
    
    CacheUsersTerritory
    
    sSQL = BuildSQLSelect(i_sCaption, i_sInput)
    Set m_rst = LoadDiscRst(sSQL)

    SetWaitCursor False
    
    Select Case m_rst.RecordCount
        Case 0
            msg "No customer accounts found that satisfy this request."
            Find = 0
    
        Case 1
            If bFindOnly Then
                Find = ChooseFromGrid("CustKey")
            Else
                Find = m_rst("CustKey").value
            End If
            'if we have a CustKey
            If Find > 0 Then
                LoadAddresses m_rst, i_oCust, bFindOnly
            End If
            
        Case Else
            'if we have a CustKey
            Find = ChooseFromGrid("CustKey")
            If Find > 0 Then
                LoadAddresses m_rst, i_oCust, bFindOnly
            End If
    End Select

    Unload Me

End Function


'This method has only one client
'   FEditAddress.cmdMoreAddrs_Click()

Public Function FindByCustKey(i_lCustKey As Long, _
                                i_lCurShipAddrKey As Long) As Long
    Dim sSQL As String

Debug.Print Me.Name & " FindByCustKey"

    m_bSearchViaCustKey = True
    Me.Caption = "Customer Search by Headquarters"
    
    CacheUsersTerritory
    sSQL = CustKeySQL(i_lCustKey, i_lCurShipAddrKey)
    
    If g_bFilterCustSearch Then
        sSQL = sSQL & " and SalesTerritoryKey = " & m_lTerritoryKey & " SET ROWCOUNT 0"
'9/23/05 added this to fix a long standing bug
    Else
        sSQL = sSQL & " SET ROWCOUNT 0"
    End If
    
    Set m_rst = LoadDiscRst(sSQL)
    
    If m_rst.RecordCount = 0 Then
        msg "There are no other shipping addresses defined for this customer.", , "No Other Addresses Found"
        FindByCustKey = 0
    Else
        FindByCustKey = ChooseFromGrid("AddrKey")
    End If
    
    Unload Me

End Function


'**********************************************************************
' Button Events
'**********************************************************************

Private Sub cmdOK_Click()
    With gdxCustSearch
        If .RowIndex(.Row) <= 0 Then
            msg "Please select the desired address from the grid.", , "No Address Selected"
            Exit Sub
        End If
    End With
    m_bLoadTheAddress = True
    Me.Hide
End Sub


Private Sub cmdCancel_Click()
    m_bLoadTheAddress = False
    Me.Hide
End Sub


'**********************************************************************
' Control Events
'**********************************************************************

Private Sub chkFilterByTerritory_Click()
    Dim lRecordCount As Long
    Dim i As Integer
    
    If m_bInitializing Then Exit Sub

    SetWaitCursor True
    
    lRecordCount = m_rst.RecordCount
    If chkFilterByTerritory.value = vbChecked Then
        If m_bSearchViaCustKey Then
            Set m_rst = LoadDiscRst(m_sSearchSQL & " and SalesTerritoryKey = " & m_lTerritoryKey & " Order By c.CustID SET ROWCOUNT 0")
        Else
            Set m_rst = LoadDiscRst(m_sSearchSQL & " and ca.SalesTerritoryKey = " & m_lTerritoryKey & " Order By c.CustID SET ROWCOUNT 0")
        End If
        '.Filter = "SalesTerritoryKey = " & m_lTerritoryKey
        If lRecordCount > 0 And m_rst.RecordCount = 0 Then
            If vbYes = msg("There are hidden addresses.  Would you like to turn" & vbCrLf _
                         & "the filter off to see them?", vbYesNo, "Turn Off Filter?") Then
                chkFilterByTerritory.value = vbUnchecked
                Set m_rst = LoadDiscRst(m_sSearchSQL & " Order By c.CustID SET ROWCOUNT 0")
            End If
        End If
    Else
        Set m_rst = LoadDiscRst(m_sSearchSQL & " Order By c.CustID SET ROWCOUNT 0")
    End If
    
    ShowWarning m_rst.RecordCount
    
    LoadGridUsingArray

    SetWaitCursor False
    
End Sub


Private Sub gdxCustSearch_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    Dim i As Integer
    With gdxCustSearch
        For i = 1 To .FmtConditions.Count
            If .FmtConditions(i).Key = "Shade" Then
                gdxCustSearch.FmtConditions.Remove ("Shade")
                Exit Sub
            End If
        Next
    End With
End Sub



Private Sub gdxCustSearch_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim numCols As Integer
    Dim ColIndex As Integer
'10/3/05 LR Kludge!
On Error GoTo EH
    numCols = UBound(m_arrayCustomers, 1)
    For ColIndex = 1 To numCols
        Values(ColIndex) = m_arrayCustomers(ColIndex, RowIndex)
    Next
Exit Sub
EH:
'   the .exe is raising this event unexpectedly before the dynamic array has been dimensioned and intitalized.
'   this does not occur in the IDE
'   swallow it
End Sub


Private Sub gdxCustSearch_DblClick()
    cmdOK_Click
End Sub


'**********************************************************************
' Private Subroutines
'**********************************************************************

'Called By
'   Find

Private Function BuildSQLSelect(i_sCaption As String, i_sInput As String) As String
    Dim sInput As String
    Dim sWhere As String
    Dim sAddrType As String

    sInput = Trim$(PrepSQLText(i_sInput))

    Me.Caption = "Customer Search by " & FormatCaption(i_sCaption & " starting with: " & UCase(i_sInput))

    'build the WHERE clause
     Select Case i_sCaption
        Case k_sCustID
            If (Len(sInput) > 10) And (Right(sInput, 1) = "N") Then
                'No wildcard search
                sWhere = "tarNationalAcct.NationalAcctID = '" & sInput & "'"
            Else
                sWhere = "c.CustID LIKE '" & sInput & "%'"
            End If
        Case k_sCustName
            sWhere = "a.AddrName LIKE '" & sInput & "%'"
        Case k_sCustNameOrID
            sWhere = "a.AddrName LIKE '" & sInput & "%' OR c.CustID LIKE '" & sInput & "%'"
'        Case k_sCustVaxAcct
'            sWhere = "cav.VaxAcct = '" & sInput & "'"
        Case k_sCustID
            sWhere = "c.CustID LIKE '" & sInput & "%'"
        Case k_sCustPhone
            sWhere = "ct.Phone LIKE '" & sInput & "%'"
        Case k_sCustZip
            sWhere = "a.PostalCode LIKE '" & sInput & "%'"
        Case k_sNationalAccount
            sWhere = "tarNationalAcct.NationalAcctID Like '" & sInput & "%'"
    End Select

    sAddrType = "AcctType = CASE " _
        & "WHEN not tarNationalAcctMember_1.ParentCustKey is NULL AND tarNationalAcctMember.ChildCustKey is NULL THEN 'HQ ' " _
        & "WHEN tarNationalAcctMember_1.ParentCustKey is NULL AND not tarNationalAcctMember.ChildCustKey is NULL THEN 'BR ' " _
        & "Else '' END " _
        & "+ CASE WHEN a.AddrKey = c.DfltBillToAddrKey AND a.AddrKey = c.DfltShipToAddrKey THEN 'B&S' " _
        & "WHEN a.AddrKey = c.DfltBillToAddrKey AND a.AddrKey <> c.DfltShipToAddrKey THEN 'Bill' " _
        & "WHEN a.AddrKey <> c.DfltBillToAddrKey AND a.AddrKey = c.DfltShipToAddrKey THEN 'Ship' " _
        & "Else 'CSA' END, "
    
'!!! Need to fix query for phone #
'   LEFT OUTER JOIN  tciContact ct ON c.CustKey = ct.CntctOwnerKey
                 
    'To speed up the search indexing,
    'replaced tciAddress with aaaCustAddrZip.
    'aaaCustAddrZip is maintained by triggers
    
    BuildSQLSelect = "SET ROWCOUNT " & g_MaxCustRows & vbCrLf _
        & "SELECT distinct c.CustKey, c.CustID, ISNULL(hs.Hold, tcpHoldStatus_1.Hold) AS Hold, " _
        & "ISNULL(hs.HoldStatusDescr, tcpHoldStatus_1.HoldStatusDescr) AS HoldStatusDescr, " _
        & "a.AddrKey, ltrim(rtrim(ISNULL(a.AddrName, ''))) AS Name, ltrim(rtrim(ISNULL(a.City, ''))) AS City, " _
        & "ISNULL(a.StateID, '') AS St, " _
        & sAddrType _
        & "ltrim(rtrim(ISNULL(a.PostalCode, ''))) AS Zip, ltrim(rtrim(ISNULL(a.AddrLine1, ''))) AS Addr1, ltrim(rtrim(ISNULL(a.AddrLine2, ''))) AS Addr2, isnull(tarCustomer_1.DfltBillToAddrKey, " _
        & "c.DfltBillToAddrKey) as DfltBillToAddrKey, c.DfltShipToAddrKey, ca.SalesTerritoryKey, tarNationalAcct.NationalAcctID as NatAcct, ct.Phone " _
        & "FROM tarCustomer tarCustomer_1 with(nolock) INNER JOIN " _
        & "tarNationalAcctMember with(nolock) ON tarCustomer_1.CustKey = tarNationalAcctMember.ParentCustKey LEFT OUTER JOIN " _
        & "tcpHoldStatus tcpHoldStatus_1 with(nolock) INNER JOIN " _
        & "tcpCustHold tcpCustHold_1 with(nolock) ON tcpHoldStatus_1.HoldStatusKey = tcpCustHold_1.HoldStatusKey ON " _
        & "tarNationalAcctMember.ParentCustKey = tcpCustHold_1.CustKey RIGHT OUTER JOIN " _
        & "aaaCustAddrZip a with(nolock) INNER JOIN " _
        & "tarCustAddr ca with(nolock) ON a.AddrKey = ca.AddrKey INNER JOIN " _
        & "tarCustomer c with(nolock) ON ca.CustKey = c.CustKey LEFT OUTER JOIN " _
        & "tciContact ct with(nolock) ON c.PrimaryCntctKey = ct.CntctKey LEFT OUTER JOIN " _
        & "tarNationalAcct with(nolock) INNER JOIN " _
        & "tarNationalAcctLevel with(nolock) ON tarNationalAcct.NationalAcctKey = tarNationalAcctLevel.NationalAcctKey ON " _
        & "c.NationalAcctLevelKey = tarNationalAcctLevel.NationalAcctLevelKey LEFT OUTER JOIN " _
        & "tarNationalAcctMember tarNationalAcctMember_1 with(nolock) ON c.CustKey = tarNationalAcctMember_1.ParentCustKey ON " _
        & "tarNationalAcctMember.ChildCustKey = c.CustKey LEFT OUTER JOIN " _
        & "tcpCustHold ch with(nolock) ON c.CustKey = ch.CustKey " _
        & "LEFT OUTER JOIN tcpHoldStatus hs with(nolock) ON ch.HoldStatusKey = hs.HoldStatusKey " _
        & "WHERE ((c.CompanyID = 'CPC') AND (c.Status = 1) AND (c.CustID NOT LIKE '%-MISC%')) AND (ca.ShipDays < 90)"
        
        '& "tcpVaxAcct2 cav with(nolock) ON c.CustKey = cav.CustKey LEFT OUTER JOIN"
        
    BuildSQLSelect = BuildSQLSelect & " AND (" & sWhere & ")"

    'cache this for use by chkFilterByTerritory_Click()
    m_sSearchSQL = BuildSQLSelect

    If g_bFilterCustSearch Then
        BuildSQLSelect = BuildSQLSelect & " and ca.SalesTerritoryKey = " & m_lTerritoryKey
    End If
    BuildSQLSelect = BuildSQLSelect & " Order By c.CustID SET ROWCOUNT 0"

End Function


'get the user's territory
'called by
'   Find
'   FindByCustKey

Private Sub CacheUsersTerritory()
    Dim sWhseID As String
    
    With g_rstUsers
        .Filter = "UserKey=" & GetUserKey
        sWhseID = .Fields("BranchID").value
        .Filter = adFilterNone
    End With
    
    With g_rstWhses
        .Filter = "WhseID = '" & sWhseID & "'"
        chkFilterByTerritory.Caption = "Only show " & .Fields("BranchName").value & " customers"
        m_lTerritoryKey = .Fields("SalesTerritoryKey").value
        .Filter = adFilterNone
    End With
End Sub


Private Function CustKeySQL(i_lCustKey As Long, i_lCurShipAddrKey As Long) As String
    CustKeySQL = "SET ROWCOUNT " & g_MaxCustRows & " " _
         & "SELECT * FROM vwOPCustSearch " _
         & "WHERE CustKey = " & CStr(i_lCustKey) _
         & " AND AddrKey <> " & CStr(i_lCurShipAddrKey)
    m_sSearchSQL = CustKeySQL
End Function


Private Sub ShowWarning(lRecordCount As Long)
    lblTooMany.Visible = (lRecordCount >= g_MaxCustRows)
End Sub


'Called by
'   Find()
'   FindByCustKey()

Private Function ChooseFromGrid(i_sKeyField As String) As Long
    Dim i As Integer
    
Debug.Print Me.Name & " ChooseFromGrid"

    ShowWarning m_rst.RecordCount

    m_bInitializing = True
    'set the state of the checkbox control based on this global variable
    If g_bFilterCustSearch Then
        chkFilterByTerritory.value = vbChecked
    Else
        chkFilterByTerritory.value = vbUnchecked
    End If
    m_bInitializing = False
    
    LoadGridUsingArray

    Me.Show vbModal

    'the OK/Cancel button handlers set this flag
    If m_bLoadTheAddress Then
        'filter the recordset for the upcoming call to LoadAddresses()
        m_rst.Filter = "AddrKey = " & CLng(m_arrayCustomers(16, gdxCustSearch.RowIndex(gdxCustSearch.Row)))
        ChooseFromGrid = m_arrayCustomers(15, gdxCustSearch.RowIndex(gdxCustSearch.Row))
    End If

    MDIMain.DoRefresh

End Function


Private Sub LoadGridUsingArray()
    Dim i As Integer
    Dim NumRows As Integer
    Dim RowIndex As Integer
    
    Dim sAcctType As String
    Dim holdCustID As String
    Dim shadeRow As Boolean
    
    shadeRow = True
    holdCustID = ""
    
    NumRows = m_rst.RecordCount
    
    If NumRows > 0 Then
        ReDim m_arrayCustomers(1 To 16, 1 To NumRows) As Variant
        RowIndex = 1
        m_rst.MoveFirst
        
        Do Until m_rst.EOF
        
            'exclude distinct DfltBillToAddr keys for Branch accounts (not typical)
            sAcctType = m_rst.Fields("AcctType").value
            If sAcctType = "BR Bill" Then
                NumRows = NumRows - 1
                ReDim Preserve m_arrayCustomers(1 To 16, 1 To NumRows) As Variant
                m_rst.MoveNext
            Else
                'Map Rst to Array
                m_arrayCustomers(1, RowIndex) = m_rst.Fields("CustID").value
    
                'change the addr qualification from B&S to Ship for Branch accounts (more likely)
                If sAcctType = "BR B&S" Then sAcctType = "BR Ship"
                m_arrayCustomers(2, RowIndex) = sAcctType
                    
                m_arrayCustomers(3, RowIndex) = m_rst.Fields("Name").value
                m_arrayCustomers(4, RowIndex) = m_rst.Fields("City").value
                m_arrayCustomers(5, RowIndex) = m_rst.Fields("St").value
                m_arrayCustomers(6, RowIndex) = m_rst.Fields("Zip").value
                m_arrayCustomers(7, RowIndex) = m_rst.Fields("Addr1").value
                m_arrayCustomers(8, RowIndex) = m_rst.Fields("Addr2").value
                m_arrayCustomers(9, RowIndex) = m_rst.Fields("Hold").value
                m_arrayCustomers(10, RowIndex) = m_rst.Fields("SalesTerritoryKey").value
                'm_arrayCustomers(11, RowIndex) = m_rst.Fields("IDPassword").value
                m_arrayCustomers(12, RowIndex) = m_rst.Fields("NatAcct").value
                m_arrayCustomers(13, RowIndex) = m_rst.Fields("Phone").value
        
                If (holdCustID <> m_rst.Fields("CustID").value) Then
                    holdCustID = m_rst.Fields("CustID").value
                    shadeRow = Not shadeRow
                End If
                
                m_arrayCustomers(14, RowIndex) = shadeRow
                m_arrayCustomers(15, RowIndex) = m_rst.Fields("CustKey").value
                m_arrayCustomers(16, RowIndex) = m_rst.Fields("AddrKey").value
                RowIndex = RowIndex + 1
                m_rst.MoveNext
            End If
        Loop
        
    Else
        ReDim m_arrayCustomers(1 To 16, 1) As Variant
    End If
    
    With gdxCustSearch
        .HoldFields
        .HoldSortSettings = True
        .ItemCount = NumRows
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With

End Sub

'9/22/05 added
'The grid wrapper was providing this previously

Private Sub gdxCustSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And gdxCustSearch.RowIndex(gdxCustSearch.Row) <> 0 Then
        KeyCode = 0
        cmdOK_Click
    'map Tab to ctrl-Tab
    ElseIf KeyCode = 9 And Shift = 0 Then
        SendKeys "^{TAB}"
    End If
End Sub


'Called by Find()

Private Sub LoadAddresses(i_rst As ADODB.Recordset, i_oCust As Customer, Optional bFindOnly As Boolean = False)
    Dim sHoldStatusDescr As String

    If Not bFindOnly Then
        If Not IsNull(i_rst("HoldStatusDescr").value) Then
            sHoldStatusDescr = i_rst("HoldStatusDescr").value
        End If
    
        If Len(sHoldStatusDescr) > 0 Then
            Select Case i_rst("Hold").value
            Case 0
                msg sHoldStatusDescr, vbInformation + vbOKOnly, "Please Note..."
            Case 1
                msg sHoldStatusDescr, vbExclamation + vbOKOnly, "Customer On Hold"
            End Select
        End If
    End If

    With i_oCust
        If i_rst.Fields("AddrKey").value = i_rst.Fields("DfltBillToAddrKey").value Then
            .BillAddr.Load i_rst.Fields("AddrKey").value
            .ShipAddr.Load i_rst.Fields("DfltShipToAddrKey").value
        Else
            .BillAddr.Load i_rst.Fields("DfltBillToAddrKey").value
            .ShipAddr.Load i_rst.Fields("AddrKey").value
        End If
    End With
    
End Sub


'debug routine
Private Sub ShowFields(orst As ADODB.Recordset)
    Dim i As Integer
    For i = 0 To orst.Fields.Count - 1
        Debug.Print orst.Fields(i).Name & " : " & orst.Fields(i).value
    Next
End Sub


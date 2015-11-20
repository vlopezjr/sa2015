VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FChangeAddr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Shipping Address for "
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkOtherAddr 
      Caption         =   "Show deprecated addresses"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   7860
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   6540
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin GridEX20.GridEX gdxAddress 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   4895
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   7
      Column(1)       =   "FChangeAddr.frx":0000
      Column(2)       =   "FChangeAddr.frx":0144
      Column(3)       =   "FChangeAddr.frx":0264
      Column(4)       =   "FChangeAddr.frx":0384
      Column(5)       =   "FChangeAddr.frx":04A4
      Column(6)       =   "FChangeAddr.frx":05BC
      Column(7)       =   "FChangeAddr.frx":06D8
      FormatStylesCount=   5
      FormatStyle(1)  =   "FChangeAddr.frx":0814
      FormatStyle(2)  =   "FChangeAddr.frx":094C
      FormatStyle(3)  =   "FChangeAddr.frx":09FC
      FormatStyle(4)  =   "FChangeAddr.frx":0AB0
      FormatStyle(5)  =   "FChangeAddr.frx":0B88
      ImageCount      =   0
      PrinterProperties=   "FChangeAddr.frx":0C40
   End
   Begin GridEX20.GridEX gdxOtherAddr 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   4895
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   7
      Column(1)       =   "FChangeAddr.frx":0E18
      Column(2)       =   "FChangeAddr.frx":0F58
      Column(3)       =   "FChangeAddr.frx":1078
      Column(4)       =   "FChangeAddr.frx":1198
      Column(5)       =   "FChangeAddr.frx":12B8
      Column(6)       =   "FChangeAddr.frx":13D0
      Column(7)       =   "FChangeAddr.frx":14EC
      FormatStylesCount=   5
      FormatStyle(1)  =   "FChangeAddr.frx":1628
      FormatStyle(2)  =   "FChangeAddr.frx":1760
      FormatStyle(3)  =   "FChangeAddr.frx":1810
      FormatStyle(4)  =   "FChangeAddr.frx":18C4
      FormatStyle(5)  =   "FChangeAddr.frx":199C
      ImageCount      =   0
      PrinterProperties=   "FChangeAddr.frx":1A54
   End
   Begin VB.Label lblOldAddr 
      Caption         =   "Deprecated addresses"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3660
      Width           =   2895
   End
End
Attribute VB_Name = "FChangeAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lMinHeight = 4005
Private Const k_lMaxHeight = 7215

Private ActiveGrid As Integer

Private RetVal As VbMsgBoxResult

Private WithEvents m_ogwAddr As GridEXWrapper
Attribute m_ogwAddr.VB_VarHelpID = -1
Private WithEvents m_ogwOtherAddr As GridEXWrapper
Attribute m_ogwOtherAddr.VB_VarHelpID = -1

Private m_oCust As Customer
Private m_oRst As ADODB.Recordset
Private m_orstold As ADODB.Recordset


Private Sub Form_Activate()
    Set m_ogwAddr = New GridEXWrapper
    m_ogwAddr.Grid = gdxAddress
    Set m_ogwOtherAddr = New GridEXWrapper
    m_ogwOtherAddr.Grid = gdxOtherAddr
    ActiveGrid = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print Me.Name & " Form_Unload"
    Set m_oCust = Nothing
    Set m_ogwAddr = Nothing
    Set m_ogwOtherAddr = Nothing
    Set m_oRst = Nothing
End Sub


Public Function Load(oCust As Customer) As VbMsgBoxResult
    'load grid with all tarCustAddr ship address records for this customer, besides the one currently selected
    'oCust.ShipAddr.AddrKey
    Dim sSQL As String
    Dim i As Integer

    Me.Height = k_lMinHeight
    Me.Caption = Me.Caption & oCust.ID

    sSQL = "SELECT a.AddrKey, " _
            & "AddrType = CASE " _
            & "WHEN a.AddrKey = c.DfltBillToAddrKey AND a.AddrKey = c.DfltShipToAddrKey THEN 'B&S' " _
            & "WHEN a.AddrKey = c.DfltBillToAddrKey THEN 'Bill' " _
            & "WHEN a.AddrKey = c.DfltShipToAddrKey THEN 'Ship' " _
            & "Else 'CSA' END, " _
            & "LTRIM(RTRIM(ISNULL(a.AddrName, ''))) AS AddrName, " _
            & "LTRIM(RTRIM(ISNULL(a.AddrLine1, ''))) AS AddrLine1, " _
            & "LTRIM(RTRIM(ISNULL(a.AddrLine2, ''))) AS AddrLine2, " _
            & "LTRIM(RTRIM(ISNULL(a.City, ''))) AS City, " _
            & "ISNULL(a.StateID, '') AS StateID, " _
            & "LTRIM(RTRIM(ISNULL(a.PostalCode, ''))) AS PostalCode " _
            & "FROM tarCustAddr ca INNER JOIN " _
            & "tciAddress a ON ca.AddrKey = a.AddrKey INNER JOIN " _
            & "tarCustomer c ON ca.CustKey = c.CustKey " _
            & "WHERE (ca.CustKey=" & oCust.Key _
            & ") AND (ca.ShipDays < 90) AND " _
            & "(((a.AddrKey = c.DfltShipToAddrKey) AND (a.AddrKey = c.DfltBillToAddrKey)) OR " _
            & "((a.AddrKey <> c.DfltShipToAddrKey) AND (a.AddrKey <> c.DfltBillToAddrKey)) OR " _
            & "(a.AddrKey = c.DfltShipToAddrKey)) AND a.AddrKey <> " & oCust.ShipAddr.AddrKey _
            & " ORDER BY City"

    Set m_oRst = LoadDiscRst(sSQL)
    With gdxAddress
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = m_oRst
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With

    Set m_oCust = oCust
    
    Me.Show vbModal
    
    Load = RetVal
    
    Unload Me
End Function


Private Sub cmdCancel_Click()
    RetVal = VbMsgBoxResult.vbCancel
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    Dim sSQL As String
    Dim oCmd As ADODB.Command
    Dim errmsg As String
    
    On Error GoTo EH
    
    RetVal = VbMsgBoxResult.vbCancel
    
    'reload the ShipAddr with the selection
    If ActiveGrid = 0 Then
        If m_ogwAddr.Value("AddrKey") > 0 Then
            m_oCust.ShipAddr.Load m_ogwAddr.Value("AddrKey")
            RetVal = VbMsgBoxResult.vbOK
        End If
    Else
    'SMR - 10/13/2005 - If nothing is selected in the deprecated grid & ok is clicked
        If gdxOtherAddr.RowCount = 0 Then
            Me.Hide
            Exit Sub
        End If
        
        If vbYes = Msg("This will make this a permanent Common Shipping Address. Do you want to continue?", vbYesNo) Then
            If m_ogwOtherAddr.Value("AddrKey") > 0 Then
                m_oCust.ShipAddr.Load m_ogwOtherAddr.Value("AddrKey")
                RetVal = VbMsgBoxResult.vbOK
                
                Set oCmd = New ADODB.Command
                sSQL = "Update tarcustaddr set shipdays = 0 where addrkey =" & m_ogwOtherAddr.Value("AddrKey")
                Set oCmd = CreateCommandSP(sSQL, adCmdText)
                oCmd.Execute
                Set oCmd = Nothing
            End If
        Else
            Exit Sub
        End If
    End If
    
    Me.Hide
    Exit Sub
EH:
    errmsg = "Error changing shipping address." & vbCrLf
    errmsg = errmsg & IIf(ActiveGrid = 0, "Address grid", "Deprecated Address grid") & vbCrLf
    errmsg = Err.number & ": " & Err.Description
    Msg errmsg, vbCritical
End Sub


Private Sub gdxAddress_GotFocus()
    ActiveGrid = 0
End Sub

Private Sub gdxOtherAddr_GotFocus()
    ActiveGrid = 1
End Sub


Private Sub m_ogwAddr_RowChosen()
    cmdOK_Click
End Sub

Private Sub m_ogwOtherAddr_RowChosen()
    cmdOK_Click
End Sub


Private Sub chkOtherAddr_Click()
    Dim sSQL As String
    Dim i As Integer
    
    If chkOtherAddr.Value = vbChecked Then
    
        SetWaitCursor True
        
        sSQL = "SELECT tciaddress.addrkey, shipdays, whseid, addrname, addrline1, addrline2, city, stateid, postalcode " _
            & "FROM tarCustAddr INNER JOIN tciAddress ON tarCustAddr.AddrKey = dbo.tciAddress.AddrKey INNER JOIN " _
            & "tarCustomer ON tarCustAddr.CustKey = tarCustomer.CustKey INNER JOIN timwarehouse on tarcustaddr.whsekey = timwarehouse.whsekey " _
            & "WHERE (dbo.tarCustAddr.ShipDays > 89) AND (dbo.tarCustomer.CustID='" & m_oCust.ID & "')"

        Set m_orstold = LoadDiscRst(sSQL)
        With gdxOtherAddr
            .HoldFields
            .HoldSortSettings = True
            Set .ADORecordset = m_orstold
            For i = 1 To .Columns.Count
                .Columns(i).AutoSize
            Next
        End With

        SetWaitCursor False
        
        Me.Height = Me.Height + 3210
        lblOldAddr.Top = lblOldAddr.Top - 500
        gdxOtherAddr.Top = gdxOtherAddr.Top - 500
        chkOtherAddr.Top = chkOtherAddr.Top + 3240
        cmdOK.Top = cmdOK.Top + 3240
        cmdCancel.Top = cmdCancel.Top + 3240

    Else
        Me.Height = Me.Height - 3210
        lblOldAddr.Top = lblOldAddr.Top + 500
        gdxOtherAddr.Top = gdxOtherAddr.Top + 500
        chkOtherAddr.Top = chkOtherAddr.Top - 3240
        cmdOK.Top = cmdOK.Top - 3240
        cmdCancel.Top = cmdCancel.Top - 3240
    End If
    
    TryToSetFocus gdxAddress
End Sub



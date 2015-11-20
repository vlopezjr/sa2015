VERSION 5.00
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "mmremark.ocx"
Begin VB.Form FTest 
   Caption         =   "Plastic tubes and pots and pans, Bits and pieces, and Magic from the hand"
   ClientHeight    =   11955
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11955
   ScaleWidth      =   13860
   Begin VB.TextBox txtAddrKey 
      Height          =   315
      Left            =   9960
      TabIndex        =   50
      Top             =   1980
      Width           =   975
   End
   Begin VB.TextBox txtuserid 
      Height          =   315
      Left            =   9900
      TabIndex        =   49
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdgetshipkey 
      Caption         =   "shipkey"
      Height          =   375
      Left            =   8760
      TabIndex        =   48
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdTestGasketCheck 
      Caption         =   "Check for Gaskets"
      Height          =   495
      Left            =   9300
      TabIndex        =   46
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreatePSNoKey 
      Caption         =   "SHIPMENT    Create Stand-Alone"
      Height          =   855
      Left            =   11880
      TabIndex        =   45
      Top             =   6420
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE PS"
      Height          =   375
      Left            =   11700
      TabIndex        =   44
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreatePS 
      Caption         =   "SHIPMENT    Create As Dialog"
      Height          =   855
      Left            =   11880
      TabIndex        =   42
      Top             =   5340
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreateProvisionalShipment 
      Caption         =   "Create Provisional Shipment"
      Height          =   375
      Left            =   2520
      TabIndex        =   32
      Top             =   7320
      Width           =   2295
   End
   Begin VB.TextBox txtOpKey 
      Height          =   285
      Left            =   960
      TabIndex        =   31
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdTestExecuteScalar 
      Caption         =   "Cmd W Scalar"
      Height          =   615
      Left            =   11640
      TabIndex        =   30
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtmmcs 
      Height          =   435
      Left            =   1560
      TabIndex        =   29
      Top             =   120
      Width           =   8955
   End
   Begin VB.CommandButton cmdMMRC 
      Caption         =   "MM RC"
      Height          =   435
      Left            =   360
      TabIndex        =   28
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Len's Stuff"
      Height          =   1755
      Left            =   360
      TabIndex        =   22
      Top             =   780
      Width           =   8115
      Begin VB.CommandButton cmdFixPO 
         Caption         =   "Fix PO"
         Height          =   495
         Left            =   2100
         TabIndex        =   47
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdDumpToolbars 
         Caption         =   "Dump Toolbars"
         Height          =   375
         Left            =   300
         TabIndex        =   43
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdXmas 
         Caption         =   "Xmas"
         Height          =   375
         Left            =   6720
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdLoadContext 
         Caption         =   "Load Context"
         Height          =   375
         Left            =   1980
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdTestFailSafe 
         Caption         =   "Test FailSafe"
         Height          =   375
         Left            =   300
         TabIndex        =   25
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton cmdTestState 
         Caption         =   "Test State"
         Height          =   375
         Left            =   5160
         TabIndex        =   24
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton cmdTestLogging 
         Caption         =   "Test Logging"
         Height          =   375
         Left            =   3540
         TabIndex        =   23
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.OptionButton optAddrType 
      Caption         =   "TOO"
      Height          =   195
      Index           =   1
      Left            =   8700
      TabIndex        =   21
      Top             =   3420
      Width           =   1035
   End
   Begin VB.OptionButton optAddrType 
      Caption         =   "CSA"
      Height          =   195
      Index           =   0
      Left            =   8700
      TabIndex        =   20
      Top             =   3120
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.CommandButton cmdToo 
      Caption         =   "Edit ShipAddr"
      Height          =   495
      Left            =   7380
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtCCNum 
      Height          =   375
      Left            =   4500
      TabIndex        =   18
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdInitRV 
      Caption         =   "Init RV"
      Height          =   495
      Left            =   11640
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin MMRemark.RemarkViewer RemarkViewer1 
      Height          =   855
      Left            =   11640
      TabIndex        =   16
      Top             =   1980
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      ContextID       =   "ARCustLoad"
   End
   Begin VB.TextBox txtUserName 
      Height          =   495
      Left            =   7800
      TabIndex        =   15
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetUser 
      Caption         =   "GetUser"
      Height          =   495
      Left            =   6480
      TabIndex        =   14
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdInteropTest 
      Caption         =   "Interop Test"
      Height          =   372
      Left            =   4620
      TabIndex        =   13
      Top             =   3900
      Width           =   1152
   End
   Begin VB.TextBox txtText 
      Height          =   312
      Left            =   12120
      TabIndex        =   12
      Top             =   4500
      Width           =   1332
   End
   Begin VB.TextBox txtList 
      Height          =   312
      Left            =   12120
      TabIndex        =   11
      Top             =   4140
      Width           =   1332
   End
   Begin VB.TextBox txtPTKey 
      Height          =   312
      Left            =   12120
      TabIndex        =   10
      Top             =   3720
      Width           =   1332
   End
   Begin VB.TextBox txtCrCardTermsKey 
      Height          =   312
      Left            =   12120
      TabIndex        =   9
      Top             =   3420
      Width           =   1332
   End
   Begin VB.ComboBox cboTerms 
      Height          =   315
      Left            =   9840
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3960
      Width           =   2052
   End
   Begin VB.CheckBox chkImpliedDecimal 
      Caption         =   "Implied Decimal"
      Height          =   372
      Left            =   2220
      TabIndex        =   7
      Top             =   3900
      Width           =   1512
   End
   Begin NEWSOTALib.SOTANumber SOTANumber1 
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   3900
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2667
      _ExtentY        =   550
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.17
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
      text            =   "           0.00"
      sDecimalPlaces  =   2
   End
   Begin NEWSOTALib.SOTACurrency SOTACurrency1 
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   3540
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2667
      _ExtentY        =   550
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.17
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
      text            =   "           0.00"
      sDecimalPlaces  =   2
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   6180
      TabIndex        =   4
      Top             =   2940
      Width           =   855
   End
   Begin VB.TextBox txtCustType 
      Height          =   375
      Left            =   2220
      TabIndex        =   3
      Text            =   "1"
      Top             =   2940
      Width           =   3855
   End
   Begin VB.TextBox txtItemID 
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Text            =   "18510-G2"
      Top             =   2940
      Width           =   1935
   End
   Begin GridEX20.GridEX gdxProvisionalShipLines 
      Height          =   1455
      Left            =   2040
      TabIndex        =   33
      Top             =   9480
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   2566
      Version         =   "2.0"
      ShowToolTips    =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   10
      Column(1)       =   "FTest.frx":0000
      Column(2)       =   "FTest.frx":0148
      Column(3)       =   "FTest.frx":028C
      Column(4)       =   "FTest.frx":0400
      Column(5)       =   "FTest.frx":058C
      Column(6)       =   "FTest.frx":06D8
      Column(7)       =   "FTest.frx":080C
      Column(8)       =   "FTest.frx":094C
      Column(9)       =   "FTest.frx":0A98
      Column(10)      =   "FTest.frx":0BD0
      FormatStylesCount=   6
      FormatStyle(1)  =   "FTest.frx":0D1C
      FormatStyle(2)  =   "FTest.frx":0DFC
      FormatStyle(3)  =   "FTest.frx":0F34
      FormatStyle(4)  =   "FTest.frx":0FE4
      FormatStyle(5)  =   "FTest.frx":1098
      FormatStyle(6)  =   "FTest.frx":1170
      ImageCount      =   0
      PrinterProperties=   "FTest.frx":1228
   End
   Begin GridEX20.GridEX gdxProvisionalShipment 
      Height          =   1095
      Left            =   600
      TabIndex        =   34
      Top             =   8280
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   1931
      Version         =   "2.0"
      ShowToolTips    =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   9
      Column(1)       =   "FTest.frx":1400
      Column(2)       =   "FTest.frx":158C
      Column(3)       =   "FTest.frx":1710
      Column(4)       =   "FTest.frx":1868
      Column(5)       =   "FTest.frx":19C4
      Column(6)       =   "FTest.frx":1B24
      Column(7)       =   "FTest.frx":1C84
      Column(8)       =   "FTest.frx":1DB8
      Column(9)       =   "FTest.frx":1EEC
      FormatStylesCount=   6
      FormatStyle(1)  =   "FTest.frx":203C
      FormatStyle(2)  =   "FTest.frx":211C
      FormatStyle(3)  =   "FTest.frx":2254
      FormatStyle(4)  =   "FTest.frx":2304
      FormatStyle(5)  =   "FTest.frx":23B8
      FormatStyle(6)  =   "FTest.frx":2490
      ImageCount      =   0
      PrinterProperties=   "FTest.frx":2548
   End
   Begin VB.Frame Frame2 
      Caption         =   "Provisional Shipments"
      Height          =   3975
      Left            =   1800
      TabIndex        =   35
      Top             =   5940
      Width           =   9015
      Begin VB.Label Label4 
         Caption         =   "Items"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5295
      Left            =   300
      TabIndex        =   37
      Top             =   5160
      Width           =   9495
      Begin VB.TextBox txtOpKey2 
         Height          =   285
         Left            =   6000
         TabIndex        =   40
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewProvisionalShipment 
         Caption         =   "View Prov Shipment"
         Height          =   375
         Left            =   7440
         TabIndex        =   39
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "OpKey:"
         Height          =   375
         Left            =   5280
         TabIndex        =   41
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblOpKey 
         Caption         =   "Op Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cust Type { 1=End User  |  2=Dealer  |  3=Wholesale }:"
      Height          =   195
      Left            =   2220
      TabIndex        =   2
      Top             =   2700
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item ID:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   2700
      Width           =   555
   End
End
Attribute VB_Name = "FTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const KasonWestern = 309

Dim WithEvents cmdPriceHistory As VB.CommandButton
Attribute cmdPriceHistory.VB_VarHelpID = -1

Dim m_oCrypto As Crypto

Private m_rstTerms As ADODB.Recordset
Private m_rstProvisionalShipments As ADODB.Recordset
Private m_rstProvisionalShipmentLines As ADODB.Recordset

' Border handler class
Private m_bdr As CFormBorder

Public Property Get UserName() As String
    UserName = GetUserName
End Property


Private Sub cmdCreateProvisionalShipment_Click()
    'Create the Provisional Shipment Row
    Dim cmd As ADODB.Command
    Dim sql As String
    Dim psKey As Integer
    
    sql = "insert into tcpProvisionalShipment (OpKey, CreateDate, CreateUserId) Values(" & txtOpKey.text & ", '" & DateValue(Now) & "','VictorL')"
    g_DB.Connection.Execute (sql)
    
    'Return that row for updating
    sql = "select * from tcpProvisionalShipment where opkey=" & txtOpKey.text
    Set m_rstProvisionalShipments = LoadRst(sql)
    'gdxProvisionalShipment.HoldFields
    'Set gdxProvisionalShipment.ADORecordset = m_rstProvisionalShipments
    AttachGrid gdxProvisionalShipment, m_rstProvisionalShipments
    
    'Return ALL the items that have not been shipped yet
    '
    'sql = "select * from tcpProvisionalShipLine where pskey=" & psKey
    
    sql = "select l.ItemKey, l.ItemID, l.SOLineKey, ol.PoLineKey, l.Qty "
    sql = sql & "from tcpProvisionalShipment s "
    sql = sql & "join tcpSOLine l on s.OPKey=l.OPKey "
    sql = sql & "join tsoSoLine ol on l.solinekey=ol.solinekey"
    sql = sql & "left join tcpProvisionalShipLine sl on l.SOLineKey=sl.solinekey "
    sql = sql & "Where s.OpKey =" & txtOpKey.text & " "
    sql = sql & "and sl.pslinekey is null"
    Set m_rstProvisionalShipmentLines = LoadRst(sql)
    'gdxProvisionalShipment.HoldFields
    'Set gdxProvisionalShipLines.ADORecordset = m_rstProvisionalShipmentLines
    AttachGrid gdxProvisionalShipLines, m_rstProvisionalShipmentLines
    
End Sub

Private Sub cmdCreatePS_Click()
    Dim oFrm As FProvisionalShipment

    Set oFrm = New FProvisionalShipment
    oFrm.OPKey = 3142496
    oFrm.createShipment Me
End Sub

Private Sub cmdCreatePSNoKey_Click()
    Dim oFrm As FProvisionalShipment
    Set oFrm = New FProvisionalShipment
    oFrm.OPKey = 3142496
    oFrm.Show
End Sub

Private Sub cmdDumpToolbars_Click()
    MDIMain.DumpToolbars
End Sub







Private Sub cmdFixPO_Click()
    Dim frm As FFixPO
    Set frm = New FFixPO
    frm.FixPO "0000196804"
End Sub

Private Sub cmdgetshipkey_Click()
     txtAddrKey.text = User.GetUserWhseShipAddrKey(txtuserid.text)
End Sub

Private Sub cmdViewProvisionalShipment_Click()
    Dim sql As String
    
    sql = "select * from tcpProvisionalShipment where opkey=" & txtOpKey2.text
    Set m_rstProvisionalShipments = LoadRst(sql)
    AttachGrid gdxProvisionalShipment, m_rstProvisionalShipments
    
    
    sql = "select sl.ItemKey, sl.ItemID, sl.SOLineKey, l.Qty, sl.QtyBackordered, sl.QtyShipped "
    sql = sql & "from tcpProvisionalShipment s "
    sql = sql & "join tcpProvisionalShipLine sl on s.pskey=sl.pskey "
    sql = sql & "join tcpSoLine l on sl.SoLineKey=l.SoLineKey "
    sql = sql & "Where s.OpKey =" & txtOpKey2.text & " "
   
    Set m_rstProvisionalShipmentLines = LoadRst(sql)
    AttachGrid gdxProvisionalShipLines, m_rstProvisionalShipmentLines
End Sub


Private Sub cmdGetUser_Click()
    Dim sc As ScriptControl
    Dim sqlField As String
    
    sqlField = "FTest.UserName"
    Set sc = New ScriptControl
    sc.Language = "VBScript"
    sc.AddObject "FTest", Me
    sc.AddCode "sub Main() FTest.txtUserName.text = " & sqlField & " end sub"
    sc.Run "Main"
End Sub


'I need to update the Best Edit Controls (SOTA) on my computer.

Private Sub chkImpliedDecimal_Click()
    If chkImpliedDecimal.value = vbChecked Then
        SOTANumber1.ImpliedDecimal = True
    Else
        SOTANumber1.ImpliedDecimal = False
    End If
    Debug.Print SOTANumber1.ImpliedDecimal
End Sub



Private Sub cmdInteropTest_Click()
    EMail.SendToList "1", GetUserName & "@caseparts.com", "Interop Test", "Mail Interop Test"
    SendNotification "SendNotification Test", "This is a test", Array("lennyr@caseparts.com")
End Sub


'created 4/14/09 LR
'trying to track down the cause of a FS error thrown by MMRemark in AR ROOH
Private Sub cmdLoadContext_Click()
    Dim oRemarkContext As RemarkContext
    Dim sContextID As String
    Dim sOwnerID As String
    Dim sUserID As String
    
    sContextID = "ARCustLoad"
    sOwnerID = "SOURC92801"
    sUserID = "EreniaM"
    
    Set oRemarkContext = New RemarkContext
    oRemarkContext.Load sContextID, sOwnerID, sUserID
    oRemarkContext.PopupMemos
End Sub


Private Sub cmdMMRC_Click()
    Dim mm As MemoMeister.RemarkContext
    Set mm = New MemoMeister.RemarkContext
    mm.Load "ViewOrder"
    txtmmcs.text = mm.ConnectionString
End Sub

Private Sub cmdPriceHistory_Click()
    MsgBox "click"
End Sub

Private Sub cmdGo_Click()
    FPriceHistory.ShowPriceHistory txtItemID.text, txtCustType.text
End Sub


Private Sub cmdTestExecuteScalar_Click()
    Dim returnValue As Boolean
    
    returnValue = Billing.IsItemInInventory("01B36-020A", 23)
    
    MsgBox "Value: " & returnValue, vbInformation
    
End Sub

Private Sub cmdTestFailSafe_Click()
    TestFailSafe
End Sub

Private Sub TestFailSafe()
    Dim RC As MemoMeister.RemarkContext
    Set RC = New MemoMeister.RemarkContext
    'rc.Load "ARCustLoad", "TURBO90746"
    RC.Edit "ARCusterLoad", "TURBO90746"

'10          Dim z As Double
'20          Dim s(3) As String
'30          Dim i As Integer
'
'40      On Error GoTo EH:
'
'50          z = calculate(1, 0)
'
'60          For i = 0 To UBound(s) + 1
'70              Debug.Print s(i)
'80          Next
        
    Exit Sub

EH:
    LogError "FTest", "TestFailSafe", "", Err.Source, Err.Number, Err.Description
    
End Sub

Private Function calculate(X As Double, Y As Double) As Double
    calculate = X / Y
End Function


Private Sub cmdTestLogging_Click()
'    LogEvent "FTemp.frm", "cmdTestLogging_Click", "Testing the LogEvent function"
'
'    On Error Resume Next
'    Kill "temp.txt"
'    LogError "FTemp.frm", "cmdTestLogging_Click", "Testing the LogError function", Err.Source, Err.number, Err.Description
'    On Error GoTo 0
    
    Dim s As String
    Dim fso As FileSystemObject
    Dim xmlfile As File
    Dim ts As TextStream
    
    Set fso = New FileSystemObject
    Set xmlfile = fso.GetFile(App.path & "\domaingroup.txt")
    Set ts = xmlfile.OpenAsTextStream(ForReading)
    s = ts.ReadAll
    
    LogEventExt "FTest.frm", "cmdTestLogging_Click", "Testing LogEventExt function", s
    
End Sub


'4/6/09 LR: test Shay's observation about the State cbo behavior

Private Sub cmdTestState_Click()
    Dim cust As New Customer
    Dim frm As New FThisOrderOnlyAddress
    cust.Load 19440
    frm.EditShipAddress cust.ShipAddr
    frm.Show
    Unload frm
End Sub



Private Sub cmdInitRV_Click()
    RemarkViewer1.ContextID = "ARCustLoad"
    RemarkViewer1.OwnerID = "CASEP91754-2"

    Dim oRemark As MemoMeister.remark
    For Each oRemark In RemarkViewer1.RemarkContext.RemarkList
        Debug.Print oRemark.RemarkType.TypeID & " - " & oRemark.RemarkType.Caption
    Next
    Debug.Print GetRemarkCount("Cust.AR.Coll") & " Collection History"
End Sub

Private Function GetRemarkCount(RemarkType As String) As Integer
    Dim oRemark As MemoMeister.remark
    Dim Count As Integer
    For Each oRemark In RemarkViewer1.RemarkContext.RemarkList
        If oRemark.RemarkType.TypeID = RemarkType Then
            Count = Count + 1
        End If
    Next
    GetRemarkCount = Count
End Function

Private Sub cmdToo_Click()
'*************************************
'This is to mimic m_oCustomer.ShipAddr
    Dim moCustomer_ShipAddr As Address
    Set moCustomer_ShipAddr = New Address
    
'CSA Test
    If optAddrType(0).value Then
        'The default Addr Type is TOO not Undefined as expected.
        moCustomer_ShipAddr.AddrType = CSA
        moCustomer_ShipAddr.AddrName = "Dan"
        moCustomer_ShipAddr.Addr1 = "5333 Lorelei Ave"
        moCustomer_ShipAddr.Addr2 = ""
        moCustomer_ShipAddr.City = "Lakewood"
        moCustomer_ShipAddr.State = "CA"
        moCustomer_ShipAddr.Zip = "90712"
        moCustomer_ShipAddr.CountryID = "USA"
        moCustomer_ShipAddr.Residential = True
        moCustomer_ShipAddr.IsDirty = False
    Else
'Existing TOO Test
        moCustomer_ShipAddr.AddrType = TOO
        moCustomer_ShipAddr.AddrName = "Invalid address test"
        moCustomer_ShipAddr.Addr1 = "Couch 340015"
        moCustomer_ShipAddr.Addr2 = ""
        moCustomer_ShipAddr.City = "Prudhoe Bay"
        moCustomer_ShipAddr.State = "AK"
        moCustomer_ShipAddr.Zip = "99734"
        moCustomer_ShipAddr.CountryID = "USA"
        moCustomer_ShipAddr.Residential = False
        moCustomer_ShipAddr.IsDirty = False
    End If
    
'For testing
    Dim sResults As String
    sResults = moCustomer_ShipAddr.Export.ExportString & vbCrLf
'************************************

    Dim FTOO As FThisOrderOnlyAddress
    Set FTOO = New FThisOrderOnlyAddress
    
    'Pass in existing ship address object.
    If FTOO.EditShipAddress(moCustomer_ShipAddr) = VbMsgBoxResult.vbOK Then
'        UpdateShipAddrInfo
'
'        'Load SalesTax
'        m_oOrder.SalesTax.Init m_oCustomer
'        If m_oOrder.isWillCall Then
'            m_oOrder.SalesTax.WillCallTaxOverride m_oOrder.WhseID
'        End If
'
'        TestShipComplete
'        ConfirmPricePackList
'        chkPricePackList.Value = vbUnchecked
'
'        If m_oCust.IsTemp Then
'            m_oCust.Name = m_oCust.BillAddr.AddrName
'        End If
    End If
    Set FTOO = Nothing
    
'******************************
'For Testing Only
    sResults = sResults & moCustomer_ShipAddr.Export.ExportString
    Debug.Print sResults

    Set moCustomer_ShipAddr = Nothing
'******************************
    
End Sub




Private Sub cmdXmas_Click()
    Dim oForm As FXmasGift
    
    Set oForm = New FXmasGift
    oForm.Init 111111, 111111, Me
    Set oForm = Nothing
End Sub




Private Sub cboTerms_Click()
    txtList.text = cboTerms.List(cboTerms.ListIndex)
    txtPTKey.text = cboTerms.ItemData(cboTerms.ListIndex)
    txtText.text = cboTerms.text
End Sub


Private Function CrCardTermsKey() As Long
    With m_rstTerms
        .Filter = "pmtTermsID='CrCard'"
        CrCardTermsKey = .Fields("pmtTermsKey").value
        .Filter = adFilterNone
    End With
End Function


Private Sub Label6_Click()

End Sub


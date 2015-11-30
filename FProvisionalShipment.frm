VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form FProvisionalShipment 
   Caption         =   "Provisional Shipment"
   ClientHeight    =   10260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10829.17
   ScaleMode       =   0  'User
   ScaleWidth      =   10605
   Begin Threed.SSPanel pnlShipments 
      Height          =   10335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   18230
      _Version        =   262144
      Caption         =   "SSPanel1"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Frame Frame2 
         Caption         =   "ITEMS OPEN TO SHIP"
         Height          =   2415
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   10215
         Begin VB.CommandButton cmdCreateShipment 
            Caption         =   "Create Shipment"
            Height          =   375
            Left            =   8040
            TabIndex        =   14
            Top             =   1920
            Width           =   1935
         End
         Begin GridEX20.GridEX gdxOrderLines 
            Height          =   1095
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   1931
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   7
            Column(1)       =   "FProvisionalShipment.frx":0000
            Column(2)       =   "FProvisionalShipment.frx":01B8
            Column(3)       =   "FProvisionalShipment.frx":0390
            Column(4)       =   "FProvisionalShipment.frx":05BC
            Column(5)       =   "FProvisionalShipment.frx":0794
            Column(6)       =   "FProvisionalShipment.frx":0930
            Column(7)       =   "FProvisionalShipment.frx":0AE0
            FormatStylesCount=   6
            FormatStyle(1)  =   "FProvisionalShipment.frx":0C90
            FormatStyle(2)  =   "FProvisionalShipment.frx":0DC8
            FormatStyle(3)  =   "FProvisionalShipment.frx":0E78
            FormatStyle(4)  =   "FProvisionalShipment.frx":0F2C
            FormatStyle(5)  =   "FProvisionalShipment.frx":1004
            FormatStyle(6)  =   "FProvisionalShipment.frx":10BC
            ImageCount      =   0
            PrinterProperties=   "FProvisionalShipment.frx":119C
         End
         Begin VB.Label lblCustVendor 
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
            Left            =   5520
            TabIndex        =   22
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label lblOpKey 
            Caption         =   "OP#:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblSoNumber 
            Caption         =   "SO#:"
            Height          =   255
            Left            =   1560
            TabIndex        =   20
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblPoNumber 
            Caption         =   "PO#:"
            Height          =   255
            Left            =   3000
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblOpKeyValue 
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
            Left            =   600
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblSoNumberValue 
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
            Left            =   2040
            TabIndex        =   17
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblPoNumberValue 
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
            Left            =   3480
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CLOSED SHIPMENTS"
         Height          =   3495
         Left            =   120
         TabIndex        =   10
         Top             =   6600
         Width           =   10215
         Begin GridEX20.GridEX gdxOSShipments 
            Height          =   1335
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2355
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            ColumnsCount    =   5
            Column(1)       =   "FProvisionalShipment.frx":1374
            Column(2)       =   "FProvisionalShipment.frx":14C4
            Column(3)       =   "FProvisionalShipment.frx":1624
            Column(4)       =   "FProvisionalShipment.frx":17BC
            Column(5)       =   "FProvisionalShipment.frx":18E8
            FormatStylesCount=   6
            FormatStyle(1)  =   "FProvisionalShipment.frx":1A28
            FormatStyle(2)  =   "FProvisionalShipment.frx":1B08
            FormatStyle(3)  =   "FProvisionalShipment.frx":1C40
            FormatStyle(4)  =   "FProvisionalShipment.frx":1CF0
            FormatStyle(5)  =   "FProvisionalShipment.frx":1DA4
            FormatStyle(6)  =   "FProvisionalShipment.frx":1E7C
            ImageCount      =   0
            PrinterProperties=   "FProvisionalShipment.frx":1F34
         End
         Begin GridEX20.GridEX gdxOSShipItems 
            Height          =   1335
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2355
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            ColumnsCount    =   8
            Column(1)       =   "FProvisionalShipment.frx":210C
            Column(2)       =   "FProvisionalShipment.frx":2254
            Column(3)       =   "FProvisionalShipment.frx":2390
            Column(4)       =   "FProvisionalShipment.frx":24B4
            Column(5)       =   "FProvisionalShipment.frx":25EC
            Column(6)       =   "FProvisionalShipment.frx":2770
            Column(7)       =   "FProvisionalShipment.frx":28EC
            Column(8)       =   "FProvisionalShipment.frx":2A14
            FormatStylesCount=   6
            FormatStyle(1)  =   "FProvisionalShipment.frx":2B4C
            FormatStyle(2)  =   "FProvisionalShipment.frx":2C2C
            FormatStyle(3)  =   "FProvisionalShipment.frx":2D64
            FormatStyle(4)  =   "FProvisionalShipment.frx":2E14
            FormatStyle(5)  =   "FProvisionalShipment.frx":2EC8
            FormatStyle(6)  =   "FProvisionalShipment.frx":2FA0
            ImageCount      =   0
            PrinterProperties=   "FProvisionalShipment.frx":3058
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "OPEN SHIPMENTS"
         Height          =   3615
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   10215
         Begin VB.CommandButton cmdOk 
            Caption         =   "Ok"
            Height          =   375
            Left            =   7200
            TabIndex        =   23
            Top             =   3120
            Width           =   975
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   8280
            TabIndex        =   7
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   9240
            TabIndex        =   6
            Top             =   3120
            Width           =   855
         End
         Begin GridEX20.GridEX gdxShipLines 
            Height          =   1095
            Left            =   120
            TabIndex        =   8
            Top             =   1920
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   1931
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   5
            Column(1)       =   "FProvisionalShipment.frx":3230
            Column(2)       =   "FProvisionalShipment.frx":33D0
            Column(3)       =   "FProvisionalShipment.frx":35B0
            Column(4)       =   "FProvisionalShipment.frx":374C
            Column(5)       =   "FProvisionalShipment.frx":38A4
            FormatStylesCount=   6
            FormatStyle(1)  =   "FProvisionalShipment.frx":39F0
            FormatStyle(2)  =   "FProvisionalShipment.frx":3B28
            FormatStyle(3)  =   "FProvisionalShipment.frx":3BD8
            FormatStyle(4)  =   "FProvisionalShipment.frx":3C8C
            FormatStyle(5)  =   "FProvisionalShipment.frx":3D64
            FormatStyle(6)  =   "FProvisionalShipment.frx":3E1C
            ImageCount      =   0
            PrinterProperties=   "FProvisionalShipment.frx":3EFC
         End
         Begin GridEX20.GridEX gdxShipments 
            Height          =   1335
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2355
            Version         =   "2.0"
            ShowToolTips    =   -1  'True
            DefaultGroupMode=   1
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   9
            Column(1)       =   "FProvisionalShipment.frx":40D4
            Column(2)       =   "FProvisionalShipment.frx":42AC
            Column(3)       =   "FProvisionalShipment.frx":4464
            Column(4)       =   "FProvisionalShipment.frx":4624
            Column(5)       =   "FProvisionalShipment.frx":47D0
            Column(6)       =   "FProvisionalShipment.frx":4954
            Column(7)       =   "FProvisionalShipment.frx":4ADC
            Column(8)       =   "FProvisionalShipment.frx":4C5C
            Column(9)       =   "FProvisionalShipment.frx":4D74
            FormatStylesCount=   6
            FormatStyle(1)  =   "FProvisionalShipment.frx":4EF4
            FormatStyle(2)  =   "FProvisionalShipment.frx":502C
            FormatStyle(3)  =   "FProvisionalShipment.frx":50DC
            FormatStyle(4)  =   "FProvisionalShipment.frx":5190
            FormatStyle(5)  =   "FProvisionalShipment.frx":5268
            FormatStyle(6)  =   "FProvisionalShipment.frx":5320
            ImageCount      =   0
            PrinterProperties=   "FProvisionalShipment.frx":5400
         End
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1085
      _Version        =   262144
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboDocType 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtIdToFind 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Find"
         Height          =   315
         Left            =   3660
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FProvisionalShipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FormWidth = 10725
Private Const FormHeight = 10905

Private m_lWindowID As Long


Private m_sDocId As String

Private m_sOPID As String
Private m_sSOId As String
Private m_sPOId As String

Private m_lOPKey As Long
Private m_lPOKey As Long
Private m_lSOKey As Long

Private m_lPSKey As Long

Private WithEvents m_gwShipments As GridEXWrapper
Attribute m_gwShipments.VB_VarHelpID = -1
Private WithEvents m_gwShipItems As GridEXWrapper
Attribute m_gwShipItems.VB_VarHelpID = -1

Private o_rstShipments As ADODB.Recordset
Private o_rstShipmentLines As ADODB.Recordset
Private o_rstSoLines As ADODB.Recordset

Private m_arrayShipments As Variant
Private m_arrayShipLines As Variant
Private m_arrayOrderLines As Variant

Private m_iShipmentCount As Integer
Private m_iShipLinesCount As Integer
Private m_iOrderLinesCount As Integer

Private m_bAllowShipmentCreation As Boolean

Private m_bDialogMode As Boolean

Private m_fCaller As Form

' Event raised when the form is closed
Public Event OnClose()

' CFormBorder provides runtime access to "read-only" properties such as ControlBox, MaxButton, MinButton,
' Moveable, and ShowInTaskbar. Adds new properties such as AutoDrag, Sizeable, Titlebar, and Topmost.
' http://vb.mvps.org/samples/FormBdr/
Private m_bdr As CFormBorder


'****** MDI Interface **************************

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


'***********************************************

Public Property Get OPKey() As Long
    OPKey = m_lOPKey
End Property

Public Property Let OPKey(ByVal lNewValue As Long)
    m_lOPKey = lNewValue
End Property

Public Property Get POKey() As Long
    POKey = m_lPOKey
End Property

Public Property Let POKey(ByVal lNewValue As Long)
    m_lPOKey = lNewValue
End Property


'********************************************************

' this is the entry point when this form is invoked as a dialog box

Public Sub CreateShipment(Caller As Form)
        
    Caller.Enabled = False
    Set m_fCaller = Caller

    Me.Show
    Me.ZOrder
End Sub


Private Sub cmdOk_Click()
    MsgBox ("OK")
End Sub

'********** Form Events ********************************

Private Sub Form_Load()

    m_bDialogMode = Not m_fCaller Is Nothing

    m_bAllowShipmentCreation = False
    
    Set m_gwShipments = New GridEXWrapper
    m_gwShipments.Grid = gdxOSShipments
    
    Set m_gwShipItems = New GridEXWrapper
    m_gwShipItems.Grid = gdxOSShipItems
    
    If m_bDialogMode Then
        
        'HIDE THE SEARCH PANEL IN DIALOG MODE
        pnlSearch.Visible = False
        
        'MOVE THE SHIPMENT PANEL UP
        pnlShipments.Top = 0
        pnlShipments.Left = 0
        
        ' ** NOTE - the call to SetCaption or any other communication with the MDI form
        ' invalidates any changes made by the CFormBorder class
        If m_lOPKey > 0 Then
            LoadOrderHeader
            LoadOrderLines
            LoadShipments
            LoadClosedShipments
            SetCaption "Provisional Shipments for OP " & m_lOPKey
            
        ElseIf m_lPOKey > 0 Then
            LoadOrderHeader
            LoadOrderLines
            LoadShipments
            LoadClosedShipments
            SetCaption "Provisional Shipments for PO " & m_sPOId
        End If
        
        ' Remove/disable the minimize and maximize button on my mdi form
        Set m_bdr = New CFormBorder
        Set m_bdr.Client = Me
        m_bdr.ControlBox = False
        
        ' Hide the buttons in Dialog Mode
        txtIdToFind.Enabled = False
        cmdSearch.Enabled = False
        cboDocType.Enabled = False
        
    Else
        SetCaption "Provisional Shipments"
        With cboDocType
            .AddItem "OP#"
            .AddItem "SO#"
            .AddItem "PO#"
            .text = "OP#"
        End With
    
    End If
    
    cmdCreateShipment.Enabled = m_bAllowShipmentCreation
    
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
'    If m_bDialogMode = True Then
'        Me.width = FormWidth
'        Me.Height = FormHeight - 735
'    Else
        Me.width = FormWidth
        Me.Height = FormHeight
'    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not m_fCaller Is Nothing Then
        m_fCaller.Enabled = True
        m_fCaller.ZOrder
    End If
    Set m_fCaller = Nothing
    
    MDIMain.UnloadTool m_lWindowID
    
End Sub


'********************************************************

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdSearch_Click()
    Cursor.SetWaitCursor True

    m_lPSKey = 0
    
    m_bAllowShipmentCreation = False
    
    If IsNumeric(Trim(txtIdToFind.text)) Then
        m_sDocId = Trim(txtIdToFind.text)
        
         LoadOrderHeader
         LoadOrderLines
         LoadShipments
        
         cmdCreateShipment.Enabled = m_bAllowShipmentCreation
    End If
    
    Cursor.ClearWaitCursor
End Sub


'This query needs to be designed to look up the relationships by any three of the Ids and return the other two.

Private Sub LoadOrderHeader()
    Dim sel As String
    Dim from As String
    Dim where As String
    Dim sql As String
    
    sel = "select DISTINCT " _
    & "tcpSO.OPKey, " _
    & "SUBSTRING(tsoSalesOrder.TranNo, PATINDEX('%[^0]%', tsoSalesOrder.Tranno), LEN(tsoSalesOrder.Tranno)) AS SOID, " _
    & "tsoSalesOrder.SOKey, " _
    & "tsoSalesOrder.TranDate AS SODate, " _
    & "LTRIM(RTRIM(tarCustomer.CustID)) AS CustID, " _
    & "tarCustomer.CustName, " _
    & "tapVendor.VendID, " _
    & "LTRIM(RTRIM(tapVendor.VendName)) AS VendName, " _
    & "SUBSTRING(tpoPurchOrder.TranNo, PATINDEX('%[^0]%', tpoPurchOrder.TranNo), LEN(tpoPurchOrder.TranNo)) AS POID, " _
    & "tpoPurchOrder.POKey, " _
    & "tpoPurchOrder.TranDate AS PODate, " _
    & "tciContact.Phone, " _
    & "tciContact.PhoneExt "

    from = from & "From " _
    & "dbo.tcpSO INNER JOIN " _
    & "dbo.tsoSalesOrder ON dbo.tcpSO.SOKey = dbo.tsoSalesOrder.SOKey INNER JOIN " _
    & "dbo.tsoSOLine ON dbo.tsoSalesOrder.SOKey = dbo.tsoSOLine.SOKey INNER JOIN " _
    & "dbo.tsoSOLineDist ON dbo.tsoSOLine.SOLineKey = dbo.tsoSOLineDist.SOLineKey INNER JOIN " _
    & "dbo.tpoPOLine ON dbo.tsoSOLine.POLineKey = dbo.tpoPOLine.POLineKey INNER JOIN " _
    & "dbo.tarCustomer ON dbo.tsoSalesOrder.CustKey = dbo.tarCustomer.CustKey INNER JOIN " _
    & "dbo.tpoPurchOrder ON dbo.tpoPOLine.POKey = dbo.tpoPurchOrder.POKey  INNER JOIN " _
    & "dbo.tciContact INNER JOIN " _
    & "dbo.tapVendor ON dbo.tciContact.CntctKey = dbo.tapVendor.PrimaryCntctKey ON dbo.tpoPurchOrder.VendKey = dbo.tapVendor.VendKey "

    If m_bDialogMode Then
        If OPKey > 0 Then
            where = "Where " & "tcpSo.OPKey=" & m_lOPKey
        ElseIf POKey > 0 Then
            where = "Where " & "tpoPurchOrder.POKey=" & m_lPOKey
        End If
    Else
        Select Case cboDocType.text
            Case "OP#":
                where = "Where " & "tcpSo.OPKey=" & m_sDocId
            Case "SO#":
                where = "Where " & "tsoSalesOrder.TranNo='" & Format(m_sDocId, String(10, "0")) & "'"
            Case "PO#":
                where = "Where " & "tpoPurchOrder.TranNo='" & Format(m_sDocId, String(10, "0")) & "'"
        End Select
    End If
    
    sql = sel & from & where
    'MsgBox (sql)

    Dim rst As ADODB.Recordset
    Set rst = LoadDiscRst(sql)

    If rst.RecordCount > 0 Then
        With rst
            m_sOPID = .Fields("OPKey")
            m_sSOId = .Fields("SOId")
            m_sPOId = .Fields("POId")
            m_lOPKey = .Fields("OPKey")
            m_lSOKey = .Fields("SOKey")
            m_lPOKey = .Fields("POKey")

            lblOpKeyValue.caption = m_sOPID
            lblSoNumberValue.caption = m_sSOId
            lblPoNumberValue.caption = m_sPOId
            
            lblCustVendor.caption = .Fields("CustId") & " | " & .Fields("VendName")
        End With
    End If

    Set rst = Nothing
End Sub


Private Sub LoadOrderLines()
    Dim sql As String
    
'    sql = "SELECT l.ItemKey, l.POLineKey, " _
'            & "case when i.itemid like '%-mpk%' or i.itemid like '%-stl' or i.itemid like '%-sea%' then LTRIM(RTRIM(ql.ItemID)) else LTRIM(RTRIM(i.ItemID)) end AS ItemId, " _
'            & "ld.QtyOrd, " _
'            & "case when ps.status = 2 then ld.QtyOpenToShip else (ld.QtyOpenToShip - ISNULL(sum(psl.QtyShipped), 0)) end as QtyOpenToShip, " _
'            & " 0.00000000 as QtyToShip, l.SOLineKey " _
'            & "from " _
'            & "tsoSalesOrder o with (nolock) " _
'            & "join tsoSOLine l with (nolock) on o.sokey=l.sokey " _
'            & "join tsoSOLineDist ld  on l.SOLineKey=ld.SOLineKey " _
'            & "join tcpSoline ql with (nolock) on l.SOLineKey=ql.SOLineKey " _
'            & "left join tcpProvisionalShipLine psl on l.solinekey=psl.solinekey " _
'            & "left join tcpProvisionalShipment ps on psl.pskey=ps.pskey " _
'            & "join timItem i on l.ItemKey=i.ItemKey " _
'            & "where o.sokey =" & m_lSOKey & " " _
'            & "and (ps.Status = 0 or ps.Status is null) " _
'            & "group by l.ItemKey, ql.ItemID, l.POLineKey, i.ItemID, ld.QtyOpenToShip, ld.QtyOrd, l.SOLineKey, ps.status"
            
    Dim rst As ADODB.Recordset
'    Set rst = LoadDiscRst(sql)
    Set rst = CallSP("spcpcGetProvisionalShipmentLineCandidates", "@sokey", m_lSOKey)
    
    If rst.RecordCount > 0 Then
        m_arrayOrderLines = rst.GetRows
        m_iOrderLinesCount = UBound(m_arrayOrderLines, 2) + 1
        
        rst.MoveFirst

        While Not rst.EOF And m_bAllowShipmentCreation = False

            If CInt(rst.Fields("QtyOpenToShip")) > 0 Then
                m_bAllowShipmentCreation = True
            End If

            rst.MoveNext
        Wend
    Else
        m_arrayOrderLines = Empty
        m_iOrderLinesCount = 0
        gdxOrderLines.Rebind
        
        'THERE ARE NO SHIPMENTS ENABLE THE CREATE BUTTON
        m_bAllowShipmentCreation = True
    End If
    
    Dim i As Integer
    With gdxOrderLines
        .HoldFields
        .HoldSortSettings = True
        .ItemCount = m_iOrderLinesCount
        .Refetch
        .Row = 1
    End With
End Sub


Private Sub LoadShipments()
    Dim sql As String
    
    If m_lOPKey = 0 Then Exit Sub
    
    sql = "select *, 0 as Del from tcpProvisionalShipment where status = 0 and opkey=" & m_lOPKey
    
    Dim rst As ADODB.Recordset
    Set rst = LoadDiscRst(sql)
    
    If rst.RecordCount > 0 Then
        m_arrayShipments = rst.GetRows
        m_iShipmentCount = UBound(m_arrayShipments, 2) + 1
    Else
        m_arrayShipments = Empty
        m_iShipmentCount = 0
        'LoadShipmentLines
    End If
    
    Dim i As Integer
    With gdxShipments
        .HoldFields
        .HoldSortSettings = True
        .ItemCount = m_iShipmentCount
        .Refetch
        .Row = 1
    End With
    
End Sub


Private Sub LoadShipmentLines()
    Dim sql As String
    
    sql = "select ps.*, d.QtyOrd from tcpProvisionalShipLine ps join tsoSoLineDist d on ps.solinekey=d.solinekey where pskey=" & m_lPSKey
    
    Dim rst As ADODB.Recordset
    Set rst = LoadDiscRst(sql)
    
    If rst.RecordCount > 0 Then
        m_arrayShipLines = rst.GetRows
        m_iShipLinesCount = UBound(m_arrayShipLines, 2) + 1
    Else
        m_arrayShipLines = Empty
        m_iShipLinesCount = 0
    End If
    
   
    With gdxShipLines
        .HoldFields
        .HoldSortSettings = True
        .ItemCount = m_iShipLinesCount
        .Refetch
        .Row = 1
    End With
End Sub


Private Sub gdxOrderLines_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_iOrderLinesCount = 0 Then Exit Sub
    If RowIndex > m_iOrderLinesCount Then Exit Sub
    
    'REMINDER - the values variable is 1 based, the array is 0 based
    Values(1) = m_arrayOrderLines(2, RowIndex - 1)    'ItemID
    Values(2) = m_arrayOrderLines(3, RowIndex - 1)    'QtyOrd
    Values(3) = m_arrayOrderLines(4, RowIndex - 1)    'QtyOpenToShip
    Values(4) = m_arrayOrderLines(5, RowIndex - 1)    'QtyToShip
    Values(5) = m_arrayOrderLines(0, RowIndex - 1)    'ItemKey
    Values(6) = m_arrayOrderLines(1, RowIndex - 1)    'POLineKey
    Values(7) = m_arrayOrderLines(6, RowIndex - 1)    'SOLineKey
End Sub


Private Sub gdxOrderLines_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If IsNumeric(Values(4)) Then
        m_arrayOrderLines(5, RowIndex - 1) = Values(4)
    End If
End Sub




Private Sub gdxOSShipments_Click()
    UpdateClosedShipmentItem
End Sub

Private Sub gdxOSShipments_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        UpdateClosedShipmentItem
    End If
End Sub

Private Sub gdxShipments_LostFocus()
    gdxShipments.Update
End Sub



Private Sub gdxShipments_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxShipments.Update
End Sub


Private Sub gdxShipments_SelectionChange()
    Dim RowIndex As Long

    With gdxShipments
        RowIndex = .Row
    End With
    
    m_lPSKey = m_arrayShipments(0, RowIndex - 1)
    
    LoadShipmentLines
End Sub


Private Sub gdxShipments_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If m_iShipmentCount = 0 Then Exit Sub
    If RowIndex > m_iShipmentCount Then Exit Sub
           
    Values(1) = m_arrayShipments(3, RowIndex - 1)   'Freight
    Values(2) = m_arrayShipments(4, RowIndex - 1)   'Tracking
    Values(3) = m_arrayShipments(5, RowIndex - 1)   'Handling
    Values(4) = m_arrayShipments(6, RowIndex - 1)   'Packing
    Values(5) = m_arrayShipments(7, RowIndex - 1)   'Tax
    Values(7) = m_arrayShipments(21, RowIndex - 1)  'VendorPays
    Values(8) = m_arrayShipments(22, RowIndex - 1)  'Note
    Values(9) = m_arrayShipments(23, RowIndex - 1)  'Delete
End Sub


Private Sub gdxShipments_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    m_arrayShipments(3, RowIndex - 1) = Values(1)   'Freight
    m_arrayShipments(4, RowIndex - 1) = Values(2)   'Tracking
    m_arrayShipments(5, RowIndex - 1) = Values(3)   'Handling
    m_arrayShipments(6, RowIndex - 1) = Values(4)   'Packing
    m_arrayShipments(7, RowIndex - 1) = Values(5)   'Tax
    m_arrayShipments(21, RowIndex - 1) = Values(7)  'VendorPays
    m_arrayShipments(22, RowIndex - 1) = Values(8)  'Note
    m_arrayShipments(23, RowIndex - 1) = Values(9)  'Delete flag
End Sub

Private Sub gdxShipLines_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If m_iShipLinesCount = 0 Then Exit Sub
    If RowIndex > m_iShipLinesCount Then Exit Sub
    
    Values(1) = m_arrayShipLines(5, RowIndex - 1)    'ItemId
    Values(2) = m_arrayShipLines(7, RowIndex - 1)    'QtyShipped
    Values(3) = m_arrayShipLines(4, RowIndex - 1)    'ItemKey
    Values(4) = m_arrayShipLines(8, RowIndex - 1)    'QtyOrdered
    Values(5) = m_arrayShipLines(0, RowIndex - 1)    'PSLineKey
End Sub

Private Sub gdxShipLines_LostFocus()
    gdxShipLines.Update
End Sub


Private Sub gdxShipLines_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxShipLines.Update
End Sub

Private Sub gdxShipLines_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim iOrigQtyShipped As Integer
    Dim iNewQtyShipped As Integer
    Dim iQtyOrdered As Integer
    
    If IsNumeric(Values(2)) Then
    
        'PULL VALUES FROM ARRAY TO LOCAL VARIABLES
        iQtyOrdered = Values(4)
        iNewQtyShipped = Values(2)
               
        If (iNewQtyShipped > iQtyOrdered) Then
            MsgBox "You can't ship more than what was ordered!"
            Exit Sub
        ElseIf (iNewQtyShipped = 0) Then
            MsgBox "You entered 0 for the Qty, you probably want to delete the shipment instead"
            Exit Sub
        End If
        
        'WRITING VALUE BACK TO ARRAY, UPDATE WILL HAPPEN WHEN CLICKING SAVE BUTTON
        m_arrayShipLines(7, RowIndex - 1) = iNewQtyShipped
    End If
End Sub

Private Sub cmdCreateShipment_Click()
    'VL to get around GridEX limitation of updating last row
    'This way any changes are immediately applied to the current row
    'http://stackoverflow.com/questions/726965/janus-gridex-problem
    If gdxOrderLines.EditMode = jgexEditModeOn Then gdxOrderLines.Update
    
    'TODO: Validation - ONLY PROCEED IF AT LEAST ONE ITEM HAS QTYTOSHIP > 0
    'when qtytoship = 0
    'when qtytoship is not a number
    'when qtytoship > qtyopentoship
    
    If Not IsShipmentValid Then
        MsgBox ("The shipment is invalid. Make sure the Qty To Ship is a number greater than 0 and not greater than qty open to ship")
        Exit Sub
    End If
    
    
    Dim cmd As ADODB.Command
    Dim sql As String
    Dim psKey As Integer
    
    'INSERT INTO THE PROV SHIPMENT TABLE
    sql = "insert into tcpProvisionalShipment (OpKey, CreateDate, CreateUserId) " _
     & "Values(" & m_lOPKey & ", '" & DateValue(Now) & "','" & GetUserName & "')"
    
    'MsgBox (sql)
    
    g_DB.Connection.Execute (sql)
    
    'GRAB THE KEY AND ASSIGN IT TO A LOCAL VARIABLE FOR FUTURE USE
    sql = "select max(pskey) as pskey from tcpProvisionalShipment where opkey=" & m_lOPKey
    
    Dim rst As ADODB.Recordset
    Set rst = LoadDiscRst(sql)
    
    If rst.RecordCount > 0 Then
        With rst
            'Store the PSKey
            m_lPSKey = .Fields("pskey")
        End With
    End If

    Set rst = Nothing
    
    'INSERT INTO THE PROV SHIPMENT LINES TABLE
    Dim polinekey As Long
    Dim SOLineKey As Long
    Dim ItemKey As Long
    Dim ItemID As String
    Dim qtytoship As Integer
    Dim i As Integer
    
    For i = 0 To m_iOrderLinesCount - 1
       
       If IsNumeric(m_arrayOrderLines(5, i)) Then
            qtytoship = CInt(m_arrayOrderLines(5, i))
            
            'ONLY INSERT THE LINE ITEMS WITH QTY > 0
            If qtytoship > 0 Then
            
                ItemKey = CLng(m_arrayOrderLines(0, i))
                polinekey = CLng(m_arrayOrderLines(1, i))
                ItemID = m_arrayOrderLines(2, i)
                qtytoship = CInt(m_arrayOrderLines(5, i))
                SOLineKey = CLng(m_arrayOrderLines(6, i))
            
                sql = "insert into tcpProvisionalShipLine " _
                    & "select " & m_lPSKey & "," & polinekey & "," & SOLineKey & "," & ItemKey & ",'" & ItemID & "',0," & qtytoship
                    
                g_DB.Connection.Execute (sql)
            
                'FIND A WAY TO REDUCE THE QTYOPENTOSHIP VALUE IN ORDER LINES
                
            End If
        End If
        
    Next
    
    'RELOAD THE SHIPMENTS
    LoadOrderHeader
    LoadOrderLines
    LoadShipments

End Sub


Private Sub cmdSave_Click()
    'VL to get around GridEX limitation of updating last row
    'This way any changes are immediately applied to the current row
    'http://stackoverflow.com/questions/726965/janus-gridex-problem
    If gdxShipments.EditMode = jgexEditModeOn Then gdxShipments.Update
   
    Dim sql As String
    Dim i As Integer
    
    For i = 0 To m_iShipmentCount - 1
    
        'CHECK FOR THE DELETE CHECKBOX
        If m_arrayShipments(23, i) Then
            'MsgBox "Delete Shipment"
            
            g_DB.Connection.BeginTrans
            
            sql = "DELETE FROM tcpProvisionalShipLine where pskey=" & m_arrayShipments(0, i)
            g_DB.Connection.Execute (sql)
            
            sql = "DELETE FROM tcpProvisionalShipment where pskey=" & m_arrayShipments(0, i)
            g_DB.Connection.Execute (sql)
            
            g_DB.Connection.CommitTrans
            
           
        Else
            'UPDATE THE PROV SHIPMENT TABLE
            sql = "update tcpProvisionalShipment " _
            & "set FreightAmt=" & m_arrayShipments(3, i) _
            & " ,ShipTrackNo='" & m_arrayShipments(4, i) & "'" _
            & " ,Handling=" & m_arrayShipments(5, i) _
            & " ,Packing=" & m_arrayShipments(6, i) _
            & " ,Tax=" & m_arrayShipments(7, i) _
            & " ,VendorPays=" & IIf(IsNull(m_arrayShipments(21, i)), "null", Abs(m_arrayShipments(21, i))) _
            & " ,Note=" & IIf(IsNull(m_arrayShipments(22, i)), "null", "'" & Trim(m_arrayShipments(22, i)) & "'") _
            & " Where pskey=" & m_arrayShipments(0, i)
            
            'MsgBox (sql)
            g_DB.Connection.Execute (sql)
            
            Dim j As Integer
            For j = 0 To m_iShipLinesCount - 1
    
                'UPDATE THE PROV SHIPMENT LINE TABLE
                If m_arrayShipLines(6, j) <> m_arrayShipLines(7, j) Then
                
                    sql = "update tcpProvisionalShipLine " _
                    & "set QtyShipped=" & m_arrayShipLines(7, j) _
                    & " Where pslinekey=" & m_arrayShipLines(0, j)
                    
                    'MsgBox (sql)
                    g_DB.Connection.Execute (sql)
                    
                End If
                
            Next
    
        End If
    Next
    
    'Raise event so that calling form can reload its data
    RaiseEvent OnClose
    
    Unload Me
End Sub


Private Function IsShipmentValid() As Boolean
    Dim i As Integer
    Dim qtytoship As Integer
    Dim qtyopen As Integer
    
    IsShipmentValid = False
    
    
    For i = 0 To m_iOrderLinesCount - 1
       
       If IsNumeric(m_arrayOrderLines(5, i)) Then
            qtytoship = CInt(m_arrayOrderLines(5, i))
            qtyopen = CInt(m_arrayOrderLines(4, i))
            
            If qtytoship > 0 And qtytoship <= qtyopen Then
                IsShipmentValid = True
                Exit For
            End If
        End If
    Next
    
End Function


Private Function GetNote(ByVal noteValue As String) As String
    If IsNull(noteValue) Then
        GetNote = Null
    Else
        GetNote = "'" & Trim(noteValue) & "'"
    End If
End Function

Private Sub LoadClosedShipments()
    Dim rstShipments As ADODB.Recordset
    Dim i As Integer
    
    Set rstShipments = CallSP("spOPOrdStatGetShipment", "@_iOPKey", m_lOPKey, "@_bExcludeProvisional", True)
    
    If Not rstShipments.EOF Then
       
        With gdxOSShipments
            .HoldFields
            Set .ADORecordset = rstShipments

            'After Shipments grid is loaded, autosize the grid columns
            'to show the column contents to user clearly, especially ShipTrackNumber
            'column.
            For i = 1 To .Columns.Count
                .Columns(i).AutoSize
            Next
        End With
        
        UpdateClosedShipmentItem
    End If
End Sub

Private Sub UpdateClosedShipmentItem()
    Dim lShipKey As Long
    Dim rst As ADODB.Recordset
    
    If IsNull(m_gwShipments.value("Shipkey")) Then Exit Sub
    
    SetWaitCursor True
    lShipKey = m_gwShipments.value("ShipKey")
    
   
   Set rst = CallSP("spOPOrdStatShipDtl1", "@i_ShipKey", lShipKey)
  

    gdxOSShipItems.HoldFields
    
    'lblOSLGridCaption(1).Caption = "Item(s) contained on Shipment " & m_gwShipments.Value("TranNo") & " dated " & Format(m_gwShipments.Value("ShipDate"), "MM/DD/YY") & ":"
    Set gdxOSShipItems.ADORecordset = rst
    Set rst = Nothing
    
    SetWaitCursor False
End Sub


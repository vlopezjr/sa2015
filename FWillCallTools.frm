VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FWillCallTool 
   Caption         =   "Will Call Manager"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   13245
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   12060
      TabIndex        =   5
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrintLabel 
      Caption         =   "Print Label"
      Default         =   -1  'True
      Height          =   315
      Left            =   2460
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox txtOPNo 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   60
      Width           =   1575
   End
   Begin GridEX20.GridEX gdxOnHold 
      Height          =   1395
      Left            =   60
      TabIndex        =   1
      Top             =   6780
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   2461
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
      Column(1)       =   "FWillCallTools.frx":0000
      Column(2)       =   "FWillCallTools.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FWillCallTools.frx":016C
      FormatStyle(2)  =   "FWillCallTools.frx":024C
      FormatStyle(3)  =   "FWillCallTools.frx":0384
      FormatStyle(4)  =   "FWillCallTools.frx":0434
      FormatStyle(5)  =   "FWillCallTools.frx":04E8
      FormatStyle(6)  =   "FWillCallTools.frx":05C0
      ImageCount      =   0
      PrinterProperties=   "FWillCallTools.frx":0678
   End
   Begin GridEX20.GridEX gdxWillCall 
      Height          =   6255
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11033
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      ShowToolTips    =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   14
      Column(1)       =   "FWillCallTools.frx":0850
      Column(2)       =   "FWillCallTools.frx":09D8
      Column(3)       =   "FWillCallTools.frx":0AF0
      Column(4)       =   "FWillCallTools.frx":0C0C
      Column(5)       =   "FWillCallTools.frx":0D2C
      Column(6)       =   "FWillCallTools.frx":0E44
      Column(7)       =   "FWillCallTools.frx":0F78
      Column(8)       =   "FWillCallTools.frx":10A4
      Column(9)       =   "FWillCallTools.frx":11C8
      Column(10)      =   "FWillCallTools.frx":12F8
      Column(11)      =   "FWillCallTools.frx":1420
      Column(12)      =   "FWillCallTools.frx":1540
      Column(13)      =   "FWillCallTools.frx":1658
      Column(14)      =   "FWillCallTools.frx":1890
      FmtConditionsCount=   1
      FmtCondition(1) =   "FWillCallTools.frx":1B28
      FormatStylesCount=   6
      FormatStyle(1)  =   "FWillCallTools.frx":1C74
      FormatStyle(2)  =   "FWillCallTools.frx":1D54
      FormatStyle(3)  =   "FWillCallTools.frx":1E8C
      FormatStyle(4)  =   "FWillCallTools.frx":1F3C
      FormatStyle(5)  =   "FWillCallTools.frx":1FF0
      FormatStyle(6)  =   "FWillCallTools.frx":20C8
      ImageCount      =   0
      PrinterProperties=   "FWillCallTools.frx":2180
   End
   Begin MSComctlLib.ImageList imglRemarks 
      Left            =   8580
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FWillCallTools.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FWillCallTools.frx":27AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblOPNo 
      Alignment       =   1  'Right Justify
      Caption         =   "OP #"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FWillCallTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FormWidth = 13365
Private Const FormHeight = 8655

Private m_lp As LabelPrinter

Private m_lUserWhseKey As Long

Private WithEvents m_gwWillCall As GridEXWrapper
Attribute m_gwWillCall.VB_VarHelpID = -1

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
    Me.caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub


Private Sub cmdPrintLabel_Click()
    Dim CustID As String
    Dim OPKey As Long
    Dim CommitDate As Date
    Dim sSQL As String
    Dim orst As ADODB.Recordset
    
    If Len(txtOPNo.text) = 0 Then Exit Sub
    
    'load op info from opkey
    
    OPKey = CLng(txtOPNo.text)
    
    sSQL = "SELECT tcpSO.OPKey, tarCustomer.CustID, tsoSalesOrder.CreateDate " _
        & "FROM tcpSO INNER JOIN tarCustomer ON tcpSO.CustKey = tarCustomer.CustKey INNER JOIN " _
        & "tsoSalesOrder ON tcpSO.SOKey = tsoSalesOrder.SOKey WHERE tcpSO.OPKey=" & OPKey

    Set orst = LoadDiscRst(sSQL)
    
    If Not orst.EOF Then
        m_lp.LabelHeight = 4#
        m_lp.LabelWidth = 4#
        m_lp.LabelTop = 0.5
        m_lp.LabelLeft = 0.25
        m_lp.FontSize = 36
        m_lp.FontBold = True
        m_lp.NumLabels = 1
        m_lp.Clear
        m_lp.AddLine orst.Fields("CustID")
        m_lp.AddLine ""
        m_lp.AddLine orst.Fields("OPKey")
        m_lp.AddLine ""
        m_lp.AddLine Format$(orst.Fields("CreateDate"), "m/d/yy")
        m_lp.PrintLabel
    End If
    
End Sub



'*******************************************************************
'Std form events
'*******************************************************************

Private Sub Form_Load()
    Dim oCol As JSColumn
    'Dim orstHold As ADODB.Recordset

    SetCaption "Will Call Management Tool"

    m_lUserWhseKey = GetUserWhseKey(GetUserKey(GetUserName))

    Set m_gwWillCall = New GridEXWrapper
    m_gwWillCall.Grid = gdxWillCall
    
    LoadImageList imglRemarks, gdxWillCall

    Me.width = FormWidth
    Me.Height = FormHeight
    
    LoadOpenOrderGrid

    For Each oCol In gdxWillCall.Columns
        oCol.AutoSize
    Next
    
    With gdxOnHold
        Set .ADORecordset = GetOrdersOnHold
    End With

    For Each oCol In gdxOnHold.Columns
        oCol.AutoSize
    Next
    
    Set m_lp = New LabelPrinter
End Sub


Private Function GetOrdersOnHold() As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tcpSO.OPKey, tcpSO.CreateDate as Created, tcpSO.UserID AS CSR, RTRIM(tarCustomer.UserFld1) AS Collector, tarCustomer.CustID," _
        & "tcpSO.CustName, RTRIM(tciPaymentTerms.PmtTermsID) AS PmtTerms " _
        & "FROM tcpCustHold WITH (NOLOCK) INNER JOIN " _
        & "tciPaymentTerms WITH (NOLOCK) INNER JOIN " _
        & "tcpSO WITH (NOLOCK) INNER JOIN " _
        & "tarCustomer WITH (NOLOCK) ON tcpSO.CustKey = tarCustomer.CustKey ON " _
        & "tciPaymentTerms.PmtTermsKey = tcpSO.PmtTermsKey INNER JOIN " _
        & "tciShipMethod WITH (NOLOCK) ON tcpSO.ShipMethKey = tciShipMethod.ShipMethKey ON " _
        & "tcpCustHold.CustKey = tarCustomer.CustKey " _
        & "WHERE (tcpSO.StatusCode = 5) AND (tciShipMethod.ShipMethID = 'MPK-Will Call')"
        
    Set GetOrdersOnHold = LoadDiscRst(sSQL)
End Function


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
    
'    If Me.width < FormWidth Then Me.width = FormWidth
'    If Me.Height < FormHeight Then Me.Height = FormHeight
    gdxWillCall.Height = Me.Height - 1115 - 1395
    gdxWillCall.width = Me.width - 330
    
    gdxOnHold.Top = gdxWillCall.Top + gdxWillCall.Height + 50
    gdxOnHold.width = Me.width - 330
End Sub


Private Sub m_gwWillCall_ColumnChosen(columnName As String)
    Select Case columnName
        Case "OP #"
            LoadOrderPad m_gwWillCall.value("OPKey")
        Case "Remarks"
            EditRemarks m_gwWillCall.value("OPKey")
            LoadOpenOrderGrid
    End Select
End Sub

Private Sub cmdRefresh_Click()
    LoadOpenOrderGrid
    
    With gdxOnHold
        Set .ADORecordset = GetOrdersOnHold
    End With
End Sub

Private Sub LoadOpenOrderGrid()
    Dim orstOpen As ADODB.Recordset
    Set orstOpen = CallSP("spcpcGetOpenWillCallInfo")
    With gdxWillCall
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orstOpen
    End With
End Sub


Private Sub LoadOrderPad(OPKey As Long)
    Dim oFrm As FOrder
    
    SetWaitCursor True

    LogEvent "FWillCallTool", "LoadOrderPad", GetUserName & " instantiating FOrder from FWillCallTool for OP " & OPKey
    
    Set oFrm = New FOrder
    MDIMain.AddNewWindow oFrm
    With oFrm
        .Show
        .Order.Load OPKey
        .lblCustName.Visible = True
        .lblCustType(0).Visible = True
        .txtCustName.Visible = False
        .cboCustType.Visible = False

        .TransitionTabs False
    End With
    SetWaitCursor False
End Sub


Private Sub EditRemarks(ByVal OPKey As Long)
    Dim oRC As RemarkContext

    Set oRC = New RemarkContext
    oRC.Edit "ViewWillCall", OPKey
End Sub


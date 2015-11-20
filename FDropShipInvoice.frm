VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FDropShipInvoice 
   Caption         =   "Drop Ship Invoice Tool"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGridRows 
      Appearance      =   0  'Flat
      Height          =   312
      Left            =   8580
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2760
      Width           =   672
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter a Remark"
      Height          =   852
      Left            =   60
      TabIndex        =   12
      Top             =   5580
      Width           =   9192
      Begin VB.CheckBox chkShipComp 
         Caption         =   "Ship Complete"
         Height          =   312
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1512
      End
      Begin VB.TextBox txtFreight 
         Height          =   312
         Left            =   2400
         TabIndex        =   5
         Top             =   300
         Width           =   972
      End
      Begin VB.TextBox txtComment 
         Height          =   312
         Left            =   4380
         TabIndex        =   6
         Top             =   300
         Width           =   3012
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   312
         Left            =   7800
         TabIndex        =   7
         Top             =   300
         Width           =   1032
      End
      Begin VB.Label Label1 
         Caption         =   "Freight"
         Height          =   252
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   552
      End
      Begin VB.Label Label2 
         Caption         =   "Comment"
         Height          =   252
         Left            =   3600
         TabIndex        =   13
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   312
      Left            =   7140
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   312
      Left            =   8280
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin GridEX20.GridEX gdxPO 
      Height          =   1032
      Left            =   60
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4440
      Width           =   9192
      _ExtentX        =   16219
      _ExtentY        =   1826
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   4
      Column(1)       =   "FDropShipInvoice.frx":0000
      Column(2)       =   "FDropShipInvoice.frx":0148
      Column(3)       =   "FDropShipInvoice.frx":026C
      Column(4)       =   "FDropShipInvoice.frx":03A8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FDropShipInvoice.frx":052C
      FormatStyle(2)  =   "FDropShipInvoice.frx":060C
      FormatStyle(3)  =   "FDropShipInvoice.frx":0744
      FormatStyle(4)  =   "FDropShipInvoice.frx":07F4
      FormatStyle(5)  =   "FDropShipInvoice.frx":08A8
      FormatStyle(6)  =   "FDropShipInvoice.frx":0980
      ImageCount      =   0
      PrinterProperties=   "FDropShipInvoice.frx":0A38
   End
   Begin GridEX20.GridEX gdxSO 
      Height          =   1032
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3120
      Width           =   9192
      _ExtentX        =   16219
      _ExtentY        =   1826
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   4
      Column(1)       =   "FDropShipInvoice.frx":0C10
      Column(2)       =   "FDropShipInvoice.frx":0D58
      Column(3)       =   "FDropShipInvoice.frx":0E7C
      Column(4)       =   "FDropShipInvoice.frx":0FB8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FDropShipInvoice.frx":1140
      FormatStyle(2)  =   "FDropShipInvoice.frx":1220
      FormatStyle(3)  =   "FDropShipInvoice.frx":1358
      FormatStyle(4)  =   "FDropShipInvoice.frx":1408
      FormatStyle(5)  =   "FDropShipInvoice.frx":14BC
      FormatStyle(6)  =   "FDropShipInvoice.frx":1594
      ImageCount      =   0
      PrinterProperties=   "FDropShipInvoice.frx":164C
   End
   Begin GridEX20.GridEX gdxDSOrders 
      Height          =   2232
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   9192
      _ExtentX        =   16219
      _ExtentY        =   3942
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "FDropShipInvoice.frx":1824
      Column(2)       =   "FDropShipInvoice.frx":1948
      Column(3)       =   "FDropShipInvoice.frx":1A78
      Column(4)       =   "FDropShipInvoice.frx":1B9C
      Column(5)       =   "FDropShipInvoice.frx":1CCC
      Column(6)       =   "FDropShipInvoice.frx":1DF8
      Column(7)       =   "FDropShipInvoice.frx":1F18
      Column(8)       =   "FDropShipInvoice.frx":203C
      Column(9)       =   "FDropShipInvoice.frx":22A8
      Column(10)      =   "FDropShipInvoice.frx":23BC
      FormatStylesCount=   6
      FormatStyle(1)  =   "FDropShipInvoice.frx":24E8
      FormatStyle(2)  =   "FDropShipInvoice.frx":25C8
      FormatStyle(3)  =   "FDropShipInvoice.frx":2700
      FormatStyle(4)  =   "FDropShipInvoice.frx":27B0
      FormatStyle(5)  =   "FDropShipInvoice.frx":2864
      FormatStyle(6)  =   "FDropShipInvoice.frx":293C
      ImageCount      =   0
      PrinterProperties=   "FDropShipInvoice.frx":29F4
   End
   Begin MSComctlLib.ImageList imglRemarks 
      Left            =   2400
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FDropShipInvoice.frx":2BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FDropShipInvoice.frx":304F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Order"
      Height          =   192
      Left            =   8040
      TabIndex        =   16
      Top             =   2820
      Width           =   432
   End
   Begin VB.Label Label5 
      Caption         =   "Drop Ship Orders"
      Height          =   192
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   2112
   End
   Begin VB.Label Label4 
      Caption         =   "Purchase Order Line Items"
      Height          =   192
      Left            =   60
      TabIndex        =   10
      Top             =   4200
      Width           =   2112
   End
   Begin VB.Label Label3 
      Caption         =   "Sales Order Line Items"
      Height          =   192
      Left            =   60
      TabIndex        =   9
      Top             =   2880
      Width           =   2112
   End
End
Attribute VB_Name = "FDropShipInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oRstDSOrders As ADODB.Recordset
Private m_oRstSO As ADODB.Recordset
Private m_oRstPO As ADODB.Recordset

Private m_gwDSOrders As GridEXWrapper

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

Private Sub Form_Load()
    SetCaption "Drop Ship Invoice Tool"

    Set m_gwDSOrders = New GridEXWrapper
    m_gwDSOrders.Grid = gdxDSOrders

    LoadImageList imglRemarks, gdxDSOrders
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '3/31/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwDSOrders = Nothing

    MDIMain.FormUnregister Me
End Sub

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub

Private Sub cmdRefresh_Click()
        SetWaitCursor True
    
        Set m_oRstDSOrders = CallSP("spOPGetDropShipPO")
        AttachGrid gdxDSOrders, m_oRstDSOrders
    
        TryToSetFocus gdxDSOrders

        If Not m_oRstDSOrders.EOF Then cmdPrint.Enabled = True
        
        SetWaitCursor False
End Sub


Private Sub gdxDSOrders_DblClick()
    EditRemarks m_gwDSOrders.value("OPKey")
End Sub

Private Sub gdxDSOrders_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        EditRemarks m_gwDSOrders.value("OPKey")
    End If
End Sub

Private Sub EditRemarks(OPKey As Long)
    Dim oRC As RemarkContext
    
    Set oRC = New RemarkContext
    oRC.Edit "DropShipUpdate", OPKey
End Sub

Private Sub gdxDSOrders_SelectionChange()
    Set m_oRstSO = LoadDiscRst("SELECT ItemID, Description, QtyOrd, UnitPrice FROM tsoSOLine " _
                    & "inner join tsoSOLineDist on tsoSOLine.SOLineKey = tsoSOLineDist.SOLineKey " _
                    & "inner join timitem on tsoSOLine.ItemKey = timitem.itemkey " _
                    & "WHERE SOKey=" & m_gwDSOrders.value("SOKey"))
    Set m_oRstPO = LoadDiscRst("SELECT ItemID, Description, QtyOrd, UnitCost FROM tpoPOLine " _
                    & "inner join timitem on tpoPOLine.ItemKey = timitem.itemkey " _
                    & "inner join tpoPOLineDist on tpoPOLine.POLineKey = tpoPOLineDist.POLineKey " _
                    & "WHERE POKey=" & m_gwDSOrders.value("POKey"))
    AttachGrid gdxSO, m_oRstSO
    AttachGrid gdxPO, m_oRstPO
   
    txtGridRows.text = gdxDSOrders.Row & "/" & CStr(m_oRstDSOrders.RecordCount)

    Set m_oRstSO = Nothing
    Set m_oRstPO = Nothing
    
    ClearRemark
End Sub

Private Sub chkShipComp_Click()
    cmdSave.Enabled = True
End Sub

Private Sub txtComment_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtFreight_Change()
    If Len(txtFreight.text) > 0 Then
        If IsNumeric(txtFreight.text) Then
            cmdSave.Enabled = True
        Else
            msg "This must be a number.", vbCritical
        End If
    End If
End Sub

Private Sub txtFreight_LostFocus()
    If Len(txtFreight.text) > 0 Then
        txtFreight.text = Format(txtFreight.text, "Currency")
    End If
End Sub

Private Sub ClearRemark()
    chkShipComp.value = vbUnchecked
    txtFreight = vbNullString
    txtComment = vbNullString
    cmdSave.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    Dim remarks As String
    Dim s As String

    SetWaitCursor True
    
    Printer.Font = "Arial"
    Printer.FontSize = 14
    Printer.Print "Drop Ship Order Report" & "     " & Date
    Printer.Print
    Printer.FontSize = 10
    With m_oRstDSOrders
        .MoveFirst
        Do While Not .EOF
            remarks = FetchRemarks(.Fields("OPKey"))
            If Len(remarks) > 0 Then
                s = .Fields("SOID") & " " & .Fields("SODate") & "     "
                s = s & .Fields("POID") & " " & .Fields("PODate") & vbCrLf
                s = s & .Fields("CustID") & " " & .Fields("CustName") & vbCrLf
                s = s & .Fields("VendID") & "" & .Fields("VendName") & " " & FormatPhoneNumber(.Fields("Phone"), vbNullString) & vbCrLf
                s = s & remarks
                Printer.Print s
                Printer.Print
            End If
            .MoveNext
        Loop
    End With
    Printer.NewPage
    
    SetWaitCursor False
    msg "Report printed on " & Printer.DeviceName
End Sub

'fetch all dropship remarks for an order
'if there are none, a null string is returned
Private Function FetchRemarks(OPKey As Long) As String
    Dim orst As ADODB.Recordset
    Dim s As String

    Set orst = New ADODB.Recordset
    With orst
        .Open "SELECT EffectiveDate, Sender, MemoText FROM tciMemo WHERE tciMemo.MemoOwnerKey=" & OPKey _
                & " AND (tciMemo.Addressee = 'order.dropship')", g_DB.Connection
        Do While Not .EOF
            s = s & .Fields("EffectiveDate") & " " & .Fields("Sender") & " " & .Fields("MemoText") & vbCrLf
            .MoveNext
        Loop
    End With
    FetchRemarks = s
End Function

Private Sub cmdSave_Click()
    Dim oRC As RemarkContext
    Dim s As String

    'build the remark string
    If chkShipComp.value = vbChecked Then
        s = s & "Shipped Complete"
    End If
    If Len(txtFreight) > 0 Then
        If Len(s) > 0 Then s = s & "; "
        s = s & "Freight = " & txtFreight
    End If
    If Len(txtComment) > 0 Then
        If Len(s) > 0 Then s = s & "; "
        s = s & txtComment
    End If
    
    Set oRC = New RemarkContext
    oRC.Load "DropShipUpdate", m_gwDSOrders.value("OPKey")
    oRC.AddRemark "Order.DropShip", s
    oRC.Save True
    Set oRC = Nothing
    ClearRemark
End Sub



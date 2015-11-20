VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FPickConflictInfo 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQtyAvail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FPickConflictInfo.frx":0000
      Top             =   180
      Visible         =   0   'False
      Width           =   4155
   End
   Begin GridEX20.GridEX gdxPO 
      Height          =   3195
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5636
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   6
      Column(1)       =   "FPickConflictInfo.frx":00CE
      Column(2)       =   "FPickConflictInfo.frx":0206
      Column(3)       =   "FPickConflictInfo.frx":032E
      Column(4)       =   "FPickConflictInfo.frx":0452
      Column(5)       =   "FPickConflictInfo.frx":056E
      Column(6)       =   "FPickConflictInfo.frx":0682
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPickConflictInfo.frx":07BA
      FormatStyle(2)  =   "FPickConflictInfo.frx":089A
      FormatStyle(3)  =   "FPickConflictInfo.frx":09D2
      FormatStyle(4)  =   "FPickConflictInfo.frx":0A82
      FormatStyle(5)  =   "FPickConflictInfo.frx":0B36
      FormatStyle(6)  =   "FPickConflictInfo.frx":0C0E
      ImageCount      =   0
      PrinterProperties=   "FPickConflictInfo.frx":0CC6
   End
   Begin GridEX20.GridEX gdxSO 
      Height          =   3195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5636
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   6
      Column(1)       =   "FPickConflictInfo.frx":0E9E
      Column(2)       =   "FPickConflictInfo.frx":0FD2
      Column(3)       =   "FPickConflictInfo.frx":10F6
      Column(4)       =   "FPickConflictInfo.frx":121A
      Column(5)       =   "FPickConflictInfo.frx":132E
      Column(6)       =   "FPickConflictInfo.frx":1442
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPickConflictInfo.frx":1576
      FormatStyle(2)  =   "FPickConflictInfo.frx":1656
      FormatStyle(3)  =   "FPickConflictInfo.frx":178E
      FormatStyle(4)  =   "FPickConflictInfo.frx":183E
      FormatStyle(5)  =   "FPickConflictInfo.frx":18F2
      FormatStyle(6)  =   "FPickConflictInfo.frx":19CA
      ImageCount      =   0
      PrinterProperties=   "FPickConflictInfo.frx":1A82
   End
End
Attribute VB_Name = "FPickConflictInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowInfo(parent As Form, itemkey As Long, WhseKey As Long, mode As Integer)
    Dim col As JSColumn
    Dim orst As ADODB.Recordset
        
    Me.Top = parent.Top + parent.Height / 2
    Me.Left = parent.Left + parent.width / 2
        
    Select Case mode
        Case 0:
            Me.caption = "Open Sales Orders"
            Me.Height = 3600
            Me.width = 5910
            Me.Left = Me.Left - Me.width / 2
            Set orst = CallSP("spCPCGetOrdersforSO", "@i_ItemKey", itemkey, "@i_WhseKey", WhseKey)
            gdxSO.Visible = True
            AttachGrid gdxSO, orst
        Case 1:
            Me.caption = "Open Purchase Orders"
            Me.Height = 3600
            Me.width = 5910
            Me.Left = Me.Left - Me.width / 2
            Set orst = CallSP("spcpcGetPoForItem", "@i_ItemKey", itemkey, "@i_WhseKey", WhseKey)
            gdxPO.Visible = True
            AttachGrid gdxPO, orst
        Case 2:
            Me.caption = "What does Avail mean?"
            Me.Height = 2700
            Me.width = 4665
            Me.Left = Me.Left - Me.width / 2
            txtQtyAvail.Visible = True
    End Select
    
    Me.Show vbModal
End Sub

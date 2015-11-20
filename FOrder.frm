VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CCDE1390-FFE8-11D4-8122-AA0004000604}#17.0#0"; "InchWorm.ocx"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "mmremark.ocx"
Object = "{DAE11CD2-4384-11D7-9DBD-000102499D33}#1.0#0"; "currcontrol.ocx"
Begin VB.Form FOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   9405
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   9405
   Begin ActiveTabs.SSActiveTabs tabMain 
      Height          =   5955
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   10504
      _Version        =   262144
      TabCount        =   8
      TagVariant      =   ""
      Tabs            =   "FOrder.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   5565
         Left            =   30
         TabIndex        =   301
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder.frx":01BE
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print History"
            Height          =   375
            Left            =   7920
            TabIndex        =   303
            Top             =   5040
            Width           =   1095
         End
         Begin GridEX20.GridEX gdxOrderEvent 
            Height          =   4815
            Left            =   120
            TabIndex        =   302
            Top             =   120
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   8493
            Version         =   "2.0"
            ScrollToolTips  =   -1  'True
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            ColumnsCount    =   6
            Column(1)       =   "FOrder.frx":01E6
            Column(2)       =   "FOrder.frx":032E
            Column(3)       =   "FOrder.frx":0452
            Column(4)       =   "FOrder.frx":05CE
            Column(5)       =   "FOrder.frx":0CA6
            Column(6)       =   "FOrder.frx":0DF2
            SortKeysCount   =   1
            SortKey(1)      =   "FOrder.frx":0F5A
            FormatStylesCount=   6
            FormatStyle(1)  =   "FOrder.frx":0FC2
            FormatStyle(2)  =   "FOrder.frx":10FA
            FormatStyle(3)  =   "FOrder.frx":11AA
            FormatStyle(4)  =   "FOrder.frx":125E
            FormatStyle(5)  =   "FOrder.frx":1336
            FormatStyle(6)  =   "FOrder.frx":13EE
            ImageCount      =   0
            PrinterProperties=   "FOrder.frx":14CE
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpStatus 
         Height          =   5565
         Left            =   30
         TabIndex        =   8
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder.frx":16A6
         Begin VB.CommandButton cmdOSRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   7920
            TabIndex        =   291
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Frame Frame1 
            Height          =   1392
            Index           =   15
            Left            =   120
            TabIndex        =   9
            Top             =   60
            Width           =   8895
            Begin VB.Label lblShipComplete 
               BorderStyle     =   1  'Fixed Single
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
               Left            =   4800
               TabIndex        =   26
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   2
               Left            =   960
               TabIndex        =   25
               Top             =   960
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "PurchOrd"
               Height          =   255
               Index           =   54
               Left            =   120
               TabIndex        =   24
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   23
               Top             =   600
               Width           =   2775
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   6
               Left            =   7440
               TabIndex        =   22
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   4
               Left            =   4800
               TabIndex        =   21
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Index           =   3
               Left            =   4800
               TabIndex        =   20
               Top             =   240
               Width           =   1452
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   5
               Left            =   7440
               TabIndex        =   19
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   7
               Left            =   7440
               TabIndex        =   18
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   17
               Top             =   240
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ordered By"
               Height          =   252
               Index           =   49
               Left            =   6360
               TabIndex        =   16
               Top             =   600
               Width           =   972
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Order Date"
               Height          =   252
               Index           =   67
               Left            =   6360
               TabIndex        =   15
               Top             =   960
               Width           =   972
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "CSR"
               Height          =   252
               Index           =   29
               Left            =   4140
               TabIndex        =   14
               Top             =   600
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "SO#"
               Height          =   252
               Index           =   26
               Left            =   6360
               TabIndex        =   13
               Top             =   240
               Width           =   972
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "OP#"
               Height          =   252
               Index           =   27
               Left            =   3960
               TabIndex        =   12
               Top             =   240
               Width           =   672
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Name"
               Height          =   255
               Index           =   28
               Left            =   360
               TabIndex        =   11
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "CustID"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   10
               Top             =   240
               Width           =   615
            End
         End
         Begin ActiveTabs.SSActiveTabs SSAOrderDetails 
            Height          =   3732
            Left            =   120
            TabIndex        =   27
            Top             =   1680
            Width           =   8892
            _ExtentX        =   15690
            _ExtentY        =   6588
            _Version        =   262144
            TabCount        =   3
            TagVariant      =   ""
            Tabs            =   "FOrder.frx":16CE
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
               Height          =   3345
               Left            =   30
               TabIndex        =   28
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   5900
               _Version        =   262144
               TabGuid         =   "FOrder.frx":1782
               Begin GridEX20.GridEX gdxOSInvoice 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   33
                  Top             =   120
                  Width           =   8595
                  _ExtentX        =   15161
                  _ExtentY        =   2355
                  Version         =   "2.0"
                  ShowToolTips    =   -1  'True
                  BoundColumnIndex=   ""
                  ReplaceColumnIndex=   ""
                  MethodHoldFields=   -1  'True
                  AllowEdit       =   0   'False
                  GroupByBoxVisible=   0   'False
                  ColumnHeaderHeight=   285
                  IntProp1        =   0
                  ColumnsCount    =   7
                  Column(1)       =   "FOrder.frx":17AA
                  Column(2)       =   "FOrder.frx":18FE
                  Column(3)       =   "FOrder.frx":1A26
                  Column(4)       =   "FOrder.frx":1BB2
                  Column(5)       =   "FOrder.frx":1D32
                  Column(6)       =   "FOrder.frx":1EB2
                  Column(7)       =   "FOrder.frx":2036
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder.frx":2176
                  FormatStyle(2)  =   "FOrder.frx":2256
                  FormatStyle(3)  =   "FOrder.frx":238E
                  FormatStyle(4)  =   "FOrder.frx":243E
                  FormatStyle(5)  =   "FOrder.frx":24F2
                  FormatStyle(6)  =   "FOrder.frx":25CA
                  ImageCount      =   0
                  PrinterProperties=   "FOrder.frx":2682
               End
               Begin GridEX20.GridEX gdxOSInvoiceItem 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   34
                  Top             =   1920
                  Width           =   8595
                  _ExtentX        =   15161
                  _ExtentY        =   2355
                  Version         =   "2.0"
                  ShowToolTips    =   -1  'True
                  BoundColumnIndex=   ""
                  ReplaceColumnIndex=   ""
                  MethodHoldFields=   -1  'True
                  AllowEdit       =   0   'False
                  GroupByBoxVisible=   0   'False
                  ColumnHeaderHeight=   285
                  IntProp1        =   0
                  ColumnsCount    =   6
                  Column(1)       =   "FOrder.frx":285A
                  Column(2)       =   "FOrder.frx":29A2
                  Column(3)       =   "FOrder.frx":2ADE
                  Column(4)       =   "FOrder.frx":2C02
                  Column(5)       =   "FOrder.frx":2D3A
                  Column(6)       =   "FOrder.frx":2EC2
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder.frx":304A
                  FormatStyle(2)  =   "FOrder.frx":312A
                  FormatStyle(3)  =   "FOrder.frx":3262
                  FormatStyle(4)  =   "FOrder.frx":3312
                  FormatStyle(5)  =   "FOrder.frx":33C6
                  FormatStyle(6)  =   "FOrder.frx":349E
                  ImageCount      =   0
                  PrinterProperties=   "FOrder.frx":3556
               End
               Begin VB.Label lblOSLGridCaption 
                  Caption         =   "Ordered By"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   282
                  Top             =   1560
                  Width           =   7695
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
               Height          =   3345
               Left            =   30
               TabIndex        =   35
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   5900
               _Version        =   262144
               TabGuid         =   "FOrder.frx":372E
               Begin GridEX20.GridEX gdxOSShipments 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   31
                  Top             =   120
                  Width           =   8595
                  _ExtentX        =   15161
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
                  Column(1)       =   "FOrder.frx":3756
                  Column(2)       =   "FOrder.frx":38A6
                  Column(3)       =   "FOrder.frx":3A06
                  Column(4)       =   "FOrder.frx":3B9E
                  Column(5)       =   "FOrder.frx":3CCA
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder.frx":3E0A
                  FormatStyle(2)  =   "FOrder.frx":3EEA
                  FormatStyle(3)  =   "FOrder.frx":4022
                  FormatStyle(4)  =   "FOrder.frx":40D2
                  FormatStyle(5)  =   "FOrder.frx":4186
                  FormatStyle(6)  =   "FOrder.frx":425E
                  ImageCount      =   0
                  PrinterProperties=   "FOrder.frx":4316
               End
               Begin GridEX20.GridEX gdxOSShipItems 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   32
                  Top             =   1920
                  Width           =   8595
                  _ExtentX        =   15161
                  _ExtentY        =   2355
                  Version         =   "2.0"
                  BoundColumnIndex=   ""
                  ReplaceColumnIndex=   ""
                  MethodHoldFields=   -1  'True
                  AllowEdit       =   0   'False
                  GroupByBoxVisible=   0   'False
                  ColumnHeaderHeight=   285
                  IntProp1        =   0
                  ColumnsCount    =   9
                  Column(1)       =   "FOrder.frx":44EE
                  Column(2)       =   "FOrder.frx":4636
                  Column(3)       =   "FOrder.frx":4772
                  Column(4)       =   "FOrder.frx":4896
                  Column(5)       =   "FOrder.frx":49CE
                  Column(6)       =   "FOrder.frx":4B52
                  Column(7)       =   "FOrder.frx":4CCE
                  Column(8)       =   "FOrder.frx":4DF6
                  Column(9)       =   "FOrder.frx":4F2E
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder.frx":5056
                  FormatStyle(2)  =   "FOrder.frx":5136
                  FormatStyle(3)  =   "FOrder.frx":526E
                  FormatStyle(4)  =   "FOrder.frx":531E
                  FormatStyle(5)  =   "FOrder.frx":53D2
                  FormatStyle(6)  =   "FOrder.frx":54AA
                  ImageCount      =   0
                  PrinterProperties=   "FOrder.frx":5562
               End
               Begin VB.Label lblOSLGridCaption 
                  Caption         =   "Ordered By"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   281
                  Top             =   1560
                  Width           =   7095
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
               Height          =   3345
               Left            =   30
               TabIndex        =   36
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   5900
               _Version        =   262144
               TabGuid         =   "FOrder.frx":573A
               Begin GridEX20.GridEX gdxOSLine 
                  Height          =   1455
                  Left            =   120
                  TabIndex        =   29
                  Top             =   120
                  Width           =   8595
                  _ExtentX        =   15161
                  _ExtentY        =   2566
                  Version         =   "2.0"
                  ShowToolTips    =   -1  'True
                  BoundColumnIndex=   ""
                  ReplaceColumnIndex=   ""
                  MethodHoldFields=   -1  'True
                  AllowEdit       =   0   'False
                  GroupByBoxVisible=   0   'False
                  DataMode        =   99
                  ColumnHeaderHeight=   285
                  IntProp1        =   0
                  IntProp2        =   0
                  IntProp7        =   0
                  ColumnsCount    =   9
                  Column(1)       =   "FOrder.frx":5762
                  Column(2)       =   "FOrder.frx":58AA
                  Column(3)       =   "FOrder.frx":59CE
                  Column(4)       =   "FOrder.frx":5B0A
                  Column(5)       =   "FOrder.frx":5C46
                  Column(6)       =   "FOrder.frx":5DCE
                  Column(7)       =   "FOrder.frx":5F0E
                  Column(8)       =   "FOrder.frx":605A
                  Column(9)       =   "FOrder.frx":61A2
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder.frx":62D2
                  FormatStyle(2)  =   "FOrder.frx":63B2
                  FormatStyle(3)  =   "FOrder.frx":64EA
                  FormatStyle(4)  =   "FOrder.frx":659A
                  FormatStyle(5)  =   "FOrder.frx":664E
                  FormatStyle(6)  =   "FOrder.frx":6726
                  ImageCount      =   0
                  PrinterProperties=   "FOrder.frx":67DE
               End
               Begin GridEX20.GridEX gdxOSLineItems 
                  Height          =   1215
                  Left            =   120
                  TabIndex        =   30
                  Top             =   2040
                  Width           =   8595
                  _ExtentX        =   15161
                  _ExtentY        =   2143
                  Version         =   "2.0"
                  ShowToolTips    =   -1  'True
                  BoundColumnIndex=   ""
                  ReplaceColumnIndex=   ""
                  MethodHoldFields=   -1  'True
                  Options         =   8
                  RecordsetType   =   1
                  AllowEdit       =   0   'False
                  GroupByBoxVisible=   0   'False
                  DataMode        =   1
                  ColumnHeaderHeight=   285
                  ColumnsCount    =   10
                  Column(1)       =   "FOrder.frx":69B6
                  Column(2)       =   "FOrder.frx":6AFE
                  Column(3)       =   "FOrder.frx":6E56
                  Column(4)       =   "FOrder.frx":6F8A
                  Column(5)       =   "FOrder.frx":70CA
                  Column(6)       =   "FOrder.frx":71EE
                  Column(7)       =   "FOrder.frx":731A
                  Column(8)       =   "FOrder.frx":7446
                  Column(9)       =   "FOrder.frx":7566
                  Column(10)      =   "FOrder.frx":77D6
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder.frx":78F6
                  FormatStyle(2)  =   "FOrder.frx":79D6
                  FormatStyle(3)  =   "FOrder.frx":7B0E
                  FormatStyle(4)  =   "FOrder.frx":7BBE
                  FormatStyle(5)  =   "FOrder.frx":7C72
                  FormatStyle(6)  =   "FOrder.frx":7D4A
                  ImageCount      =   0
                  PrinterProperties=   "FOrder.frx":7E02
               End
               Begin VB.Label lblOSLGridCaption 
                  Caption         =   "Ordered By"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   280
                  Top             =   1680
                  Width           =   8175
               End
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpRMA 
         Height          =   5565
         Left            =   30
         TabIndex        =   37
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder.frx":7FDA
         Begin ActiveTabs.SSActiveTabs tabRMADetail 
            Height          =   5295
            Left            =   120
            TabIndex        =   346
            Top             =   120
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   9340
            _Version        =   262144
            TabCount        =   1
            TagVariant      =   ""
            Tabs            =   "FOrder.frx":8002
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
               Height          =   4905
               Left            =   30
               TabIndex        =   347
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   8652
               _Version        =   262144
               TabGuid         =   "FOrder.frx":804E
               Begin VB.CommandButton cmdRMARefresh 
                  Caption         =   "Refresh"
                  Height          =   325
                  Left            =   3000
                  TabIndex        =   351
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdUpdateRMALine 
                  Caption         =   "Save Changes"
                  Enabled         =   0   'False
                  Height          =   325
                  Left            =   1560
                  TabIndex        =   350
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdAddMoreItem 
                  Caption         =   "Add Items"
                  Height          =   325
                  Left            =   120
                  TabIndex        =   349
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdRMAVendor 
                  Caption         =   "Show Vendor"
                  Height          =   325
                  Left            =   4440
                  TabIndex        =   348
                  Top             =   240
                  Width           =   1215
               End
               Begin GridEX20.GridEX gdxRMALine 
                  Height          =   1635
                  Left            =   120
                  TabIndex        =   352
                  Top             =   1080
                  Width           =   8580
                  _ExtentX        =   15134
                  _ExtentY        =   2884
                  Version         =   "2.0"
                  ScrollToolTips  =   -1  'True
                  ShowToolTips    =   -1  'True
                  BoundColumnIndex=   ""
                  ReplaceColumnIndex=   ""
                  MethodHoldFields=   -1  'True
                  Options         =   2
                  RecordsetType   =   1
                  GroupByBoxVisible=   0   'False
                  DataMode        =   99
                  ColumnHeaderHeight=   285
                  ColumnsCount    =   19
                  Column(1)       =   "FOrder.frx":8076
                  Column(2)       =   "FOrder.frx":81E2
                  Column(3)       =   "FOrder.frx":835A
                  Column(4)       =   "FOrder.frx":849E
                  Column(5)       =   "FOrder.frx":862A
                  Column(6)       =   "FOrder.frx":87B6
                  Column(7)       =   "FOrder.frx":88DA
                  Column(8)       =   "FOrder.frx":8A1E
                  Column(9)       =   "FOrder.frx":8B9E
                  Column(10)      =   "FOrder.frx":8D0E
                  Column(11)      =   "FOrder.frx":8E7E
                  Column(12)      =   "FOrder.frx":8FEE
                  Column(13)      =   "FOrder.frx":9172
                  Column(14)      =   "FOrder.frx":92DE
                  Column(15)      =   "FOrder.frx":942A
                  Column(16)      =   "FOrder.frx":95B6
                  Column(17)      =   "FOrder.frx":974E
                  Column(18)      =   "FOrder.frx":98A2
                  Column(19)      =   "FOrder.frx":99DE
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder.frx":9B1E
                  FormatStyle(2)  =   "FOrder.frx":9C56
                  FormatStyle(3)  =   "FOrder.frx":9D06
                  FormatStyle(4)  =   "FOrder.frx":9DBA
                  FormatStyle(5)  =   "FOrder.frx":9E92
                  FormatStyle(6)  =   "FOrder.frx":9F4A
                  ImageCount      =   0
                  PrinterProperties=   "FOrder.frx":A02A
               End
               Begin GridEX20.GridEX gdxRMALineStatus 
                  Height          =   1635
                  Left            =   120
                  TabIndex        =   353
                  Top             =   3120
                  Width           =   8580
                  _ExtentX        =   15134
                  _ExtentY        =   2884
                  Version         =   "2.0"
                  ScrollToolTips  =   -1  'True
                  ShowToolTips    =   -1  'True
                  BoundColumnIndex=   ""
                  ReplaceColumnIndex=   ""
                  MethodHoldFields=   -1  'True
                  AllowEdit       =   0   'False
                  GroupByBoxVisible=   0   'False
                  ColumnHeaderHeight=   285
                  ColumnsCount    =   19
                  Column(1)       =   "FOrder.frx":A202
                  Column(2)       =   "FOrder.frx":A362
                  Column(3)       =   "FOrder.frx":A4CE
                  Column(4)       =   "FOrder.frx":A622
                  Column(5)       =   "FOrder.frx":A782
                  Column(6)       =   "FOrder.frx":A8A6
                  Column(7)       =   "FOrder.frx":AA06
                  Column(8)       =   "FOrder.frx":AB96
                  Column(9)       =   "FOrder.frx":ACF6
                  Column(10)      =   "FOrder.frx":AE1A
                  Column(11)      =   "FOrder.frx":AF3E
                  Column(12)      =   "FOrder.frx":B09E
                  Column(13)      =   "FOrder.frx":B1E2
                  Column(14)      =   "FOrder.frx":B346
                  Column(15)      =   "FOrder.frx":B486
                  Column(16)      =   "FOrder.frx":B5DE
                  Column(17)      =   "FOrder.frx":B71A
                  Column(18)      =   "FOrder.frx":B866
                  Column(19)      =   "FOrder.frx":B9D2
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder.frx":BADE
                  FormatStyle(2)  =   "FOrder.frx":BC16
                  FormatStyle(3)  =   "FOrder.frx":BCC6
                  FormatStyle(4)  =   "FOrder.frx":BD7A
                  FormatStyle(5)  =   "FOrder.frx":BE52
                  FormatStyle(6)  =   "FOrder.frx":BF0A
                  ImageCount      =   0
                  PrinterProperties=   "FOrder.frx":BFEA
               End
               Begin MMRemark.RemarkViewer rvRMA 
                  Height          =   810
                  Left            =   7920
                  TabIndex        =   354
                  Top             =   120
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   1429
                  ContextID       =   "ViewRMA"
                  Caption         =   "RMA Remarks"
               End
               Begin VB.Label lblRMALineStatus 
                  Caption         =   "RMA Line Status"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   355
                  Top             =   2880
                  Width           =   3075
               End
            End
         End
         Begin VB.Label Label1 
            Caption         =   "UPS Call Tag"
            Height          =   255
            Index           =   71
            Left            =   1800
            TabIndex        =   38
            Top             =   960
            Width           =   1095
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpFindOrder 
         Height          =   5565
         Left            =   30
         TabIndex        =   39
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder.frx":C1C2
         Begin VB.CommandButton cmdLoadOrder 
            Caption         =   "Load Ord&er"
            Height          =   312
            Index           =   1
            Left            =   120
            TabIndex        =   55
            Top             =   5220
            Width           =   1215
         End
         Begin VB.Frame Frame1 
            Caption         =   "Find by OP/SO/RMA"
            ClipControls    =   0   'False
            Height          =   1212
            Index           =   9
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   1800
            Begin VB.CommandButton cmdFindOrders 
               Caption         =   "Fin&d"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   42
               Top             =   720
               Width           =   900
            End
            Begin VB.TextBox txtFindOrder 
               Height          =   288
               Left            =   120
               TabIndex        =   41
               Top             =   300
               Width           =   1140
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Find Order via Related Information"
            ClipControls    =   0   'False
            Height          =   1212
            Index           =   7
            Left            =   1920
            TabIndex        =   43
            Top             =   0
            Width           =   7260
            Begin VB.ComboBox cboTimeInterval 
               Height          =   315
               ItemData        =   "FOrder.frx":C1EA
               Left            =   5760
               List            =   "FOrder.frx":C208
               Style           =   2  'Dropdown List
               TabIndex        =   334
               Top             =   300
               Width           =   1215
            End
            Begin VB.TextBox txtFindCust 
               Height          =   288
               Left            =   960
               TabIndex        =   45
               Top             =   300
               Width           =   1332
            End
            Begin VB.TextBox txtFindText 
               Height          =   288
               Left            =   960
               TabIndex        =   47
               Top             =   750
               Width           =   1332
            End
            Begin VB.CommandButton cmdFindOrders 
               Caption         =   "Fi&nd"
               Height          =   375
               Index           =   1
               Left            =   4740
               TabIndex        =   52
               Top             =   300
               Width           =   900
            End
            Begin VB.ComboBox cboFindCSR 
               Height          =   315
               ItemData        =   "FOrder.frx":C244
               Left            =   3000
               List            =   "FOrder.frx":C246
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   300
               Width           =   1632
            End
            Begin VB.ComboBox cboFindStatus 
               Height          =   315
               ItemData        =   "FOrder.frx":C248
               Left            =   3000
               List            =   "FOrder.frx":C282
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   750
               Width           =   1632
            End
            Begin VB.CommandButton cmdClearOrder 
               Caption         =   "R&eset"
               Height          =   375
               Left            =   4740
               TabIndex        =   53
               Top             =   720
               Width           =   900
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer"
               Height          =   252
               Index           =   4
               Left            =   180
               TabIndex        =   44
               Top             =   300
               Width           =   732
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Keywords"
               Height          =   252
               Index           =   77
               Left            =   180
               TabIndex        =   46
               Top             =   750
               Width           =   732
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "CSR"
               Height          =   252
               Index           =   78
               Left            =   2400
               TabIndex        =   48
               Top             =   300
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status"
               Height          =   252
               Index           =   79
               Left            =   2460
               TabIndex        =   50
               Top             =   750
               Width           =   492
            End
         End
         Begin GridEX20.GridEX gdxOrders 
            Height          =   3855
            Left            =   120
            TabIndex        =   54
            Top             =   1320
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   6800
            Version         =   "2.0"
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            LockType        =   1
            Options         =   1
            RecordsetType   =   3
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ItemCount       =   0
            DataMode        =   99
            ColumnHeaderHeight=   285
            ColumnsCount    =   24
            Column(1)       =   "FOrder.frx":C35B
            Column(2)       =   "FOrder.frx":C56B
            Column(3)       =   "FOrder.frx":C753
            Column(4)       =   "FOrder.frx":C89B
            Column(5)       =   "FOrder.frx":D2AB
            Column(6)       =   "FOrder.frx":D473
            Column(7)       =   "FOrder.frx":D5FB
            Column(8)       =   "FOrder.frx":D72B
            Column(9)       =   "FOrder.frx":D9EB
            Column(10)      =   "FOrder.frx":DB97
            Column(11)      =   "FOrder.frx":DCEF
            Column(12)      =   "FOrder.frx":DE37
            Column(13)      =   "FOrder.frx":DF8B
            Column(14)      =   "FOrder.frx":E0DF
            Column(15)      =   "FOrder.frx":E2B3
            Column(16)      =   "FOrder.frx":E43B
            Column(17)      =   "FOrder.frx":E5A3
            Column(18)      =   "FOrder.frx":E887
            Column(19)      =   "FOrder.frx":E9FB
            Column(20)      =   "FOrder.frx":EB2F
            Column(21)      =   "FOrder.frx":ECAF
            Column(22)      =   "FOrder.frx":EE2B
            Column(23)      =   "FOrder.frx":EFCF
            Column(24)      =   "FOrder.frx":F0E7
            SortKeysCount   =   1
            SortKey(1)      =   "FOrder.frx":F227
            FormatStylesCount=   6
            FormatStyle(1)  =   "FOrder.frx":F28F
            FormatStyle(2)  =   "FOrder.frx":F36F
            FormatStyle(3)  =   "FOrder.frx":F4A7
            FormatStyle(4)  =   "FOrder.frx":F557
            FormatStyle(5)  =   "FOrder.frx":F60B
            FormatStyle(6)  =   "FOrder.frx":F6E3
            ImageCount      =   0
            PrinterProperties=   "FOrder.frx":F79B
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpItems 
         Height          =   5565
         Left            =   30
         TabIndex        =   56
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder.frx":F973
         Begin VB.Frame frmPricing 
            Caption         =   "Pricing"
            ClipControls    =   0   'False
            Height          =   792
            Left            =   120
            TabIndex        =   78
            Top             =   1680
            Visible         =   0   'False
            Width           =   8952
            Begin VB.CommandButton cmdPriceHistory 
               Caption         =   "?"
               Height          =   315
               Left            =   6720
               TabIndex        =   227
               Top             =   300
               Width           =   255
            End
            Begin CurrControl.CurrencyInput txtPrice 
               Height          =   315
               Left            =   5760
               TabIndex        =   226
               Top             =   300
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
            End
            Begin NEWSOTALib.SOTACurrency txtPrice1 
               Height          =   315
               Left            =   5760
               TabIndex        =   229
               Top             =   300
               Visible         =   0   'False
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
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
            Begin NEWSOTALib.SOTANumber txtQtyOrdered 
               Height          =   315
               Left            =   960
               TabIndex        =   230
               Top             =   300
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               mask            =   "<ILH>##|,##<ILp0>#"
               text            =   "    0"
               sIntegralPlaces =   5
               sDecimalPlaces  =   0
            End
            Begin NEWSOTALib.SOTACurrency txtExtPrice 
               Height          =   315
               Left            =   7800
               TabIndex        =   79
               Top             =   300
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               bLocked         =   -1  'True
               mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               sDecimalPlaces  =   2
            End
            Begin CurrControl.CurrencyInput txtCost 
               Height          =   315
               Left            =   4080
               TabIndex        =   225
               Top             =   300
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
            End
            Begin NEWSOTALib.SOTACurrency txtCost1 
               Height          =   315
               Left            =   4080
               TabIndex        =   228
               Top             =   300
               Visible         =   0   'False
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
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
            Begin NEWSOTALib.SOTACurrency txtListPrice 
               Height          =   315
               Left            =   2520
               TabIndex        =   300
               Top             =   300
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               bLocked         =   -1  'True
               mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               sDecimalPlaces  =   2
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "List Price"
               Height          =   255
               Index           =   96
               Left            =   1740
               TabIndex        =   299
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ext. Price"
               Height          =   255
               Index           =   39
               Left            =   7020
               TabIndex        =   83
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Cost"
               Height          =   255
               Index           =   34
               Left            =   3540
               TabIndex        =   82
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Price"
               Height          =   255
               Index           =   32
               Left            =   5220
               TabIndex        =   81
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Quantity"
               Height          =   252
               Index           =   31
               Left            =   240
               TabIndex        =   80
               Top             =   300
               Width           =   612
            End
         End
         Begin VB.Frame frmInventory 
            Caption         =   "Inventory"
            ClipControls    =   0   'False
            Height          =   972
            Left            =   120
            TabIndex        =   134
            Top             =   2520
            Visible         =   0   'False
            Width           =   8952
            Begin VB.CommandButton cmdInvFinder 
               Caption         =   "iFinder..."
               Height          =   375
               Left            =   7620
               TabIndex        =   363
               Top             =   360
               Width           =   1095
            End
            Begin VB.ComboBox cboWarehouse 
               Height          =   315
               Index           =   2
               Left            =   6480
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   362
               Top             =   420
               Width           =   852
            End
            Begin VB.Label Label1 
               Caption         =   "On PO"
               Height          =   255
               Index           =   38
               Left            =   1980
               TabIndex        =   364
               Top             =   480
               Width           =   615
            End
            Begin VB.Label lblQtyPO 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2700
               TabIndex        =   361
               Top             =   360
               Width           =   735
            End
            Begin VB.Label lblQtyAvail 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1140
               TabIndex        =   360
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Available"
               Height          =   255
               Index           =   35
               Left            =   300
               TabIndex        =   359
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "On BackOrd"
               Height          =   255
               Index           =   68
               Left            =   5460
               TabIndex        =   141
               Top             =   270
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Warehouse"
               Height          =   255
               Index           =   33
               Left            =   5580
               TabIndex        =   140
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "On Hand"
               Height          =   255
               Index           =   36
               Left            =   3900
               TabIndex        =   139
               Top             =   270
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "On SalesOrd"
               Height          =   255
               Index           =   37
               Left            =   4680
               TabIndex        =   138
               Top             =   270
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lblQtyOH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3900
               TabIndex        =   137
               Top             =   510
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblQtySO 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   4680
               TabIndex        =   136
               Top             =   510
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblQtyBO 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5460
               TabIndex        =   135
               Top             =   510
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdItemOK 
            Caption         =   "&OK"
            Height          =   312
            Left            =   7965
            TabIndex        =   254
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdNextGasket 
            Caption         =   "&Next Gasket"
            Height          =   312
            Left            =   7965
            TabIndex        =   257
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdItemDelete 
            Caption         =   "&Delete"
            Height          =   312
            Left            =   7965
            TabIndex        =   256
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdItemCancel 
            Caption         =   "&Cancel"
            Height          =   312
            Left            =   7965
            TabIndex        =   255
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Frame frmSpecifyPart 
            Caption         =   "Specify Item"
            ClipControls    =   0   'False
            Height          =   1395
            Left            =   4320
            TabIndex        =   70
            Top             =   120
            Width           =   4752
            Begin VB.CommandButton cmdSpecifyItem 
               Caption         =   "SP&O"
               Height          =   855
               Index           =   3
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Add special order item to order"
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton cmdSpecifyItem 
               Caption         =   "&Gasket"
               Height          =   855
               Index           =   0
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   71
               ToolTipText     =   "Add Gasket to order"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   855
            End
            Begin VB.CommandButton cmdSpecifyItem 
               Caption         =   "Sh&elf"
               Height          =   855
               Index           =   1
               Left            =   1400
               Style           =   1  'Graphical
               TabIndex        =   72
               ToolTipText     =   "Add Wire Shelf to order"
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton cmdSpecifyItem 
               Caption         =   "W&ire"
               Height          =   855
               Index           =   2
               Left            =   2560
               Style           =   1  'Graphical
               TabIndex        =   73
               ToolTipText     =   "Add Warmer Wire to order"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame frmFindPart 
            Caption         =   "Find Item"
            ClipControls    =   0   'False
            Height          =   1395
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   4095
            Begin VB.CheckBox chkSearchItemDescr 
               Caption         =   "Search Item Description"
               Height          =   375
               Left            =   1600
               TabIndex        =   67
               Top             =   960
               Width           =   2055
            End
            Begin VB.PictureBox picItem 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   768
               Left            =   360
               ScaleHeight     =   765
               ScaleWidth      =   765
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   360
               Width           =   768
            End
            Begin VB.CommandButton cmdSearch 
               Caption         =   "Fi&nd"
               Height          =   360
               Left            =   3420
               TabIndex        =   68
               Top             =   600
               Width           =   555
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtItemSearch 
               Height          =   315
               Left            =   1620
               TabIndex        =   66
               Top             =   600
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               Caption         =   "Search for:"
               Height          =   195
               Index           =   30
               Left            =   1680
               TabIndex        =   69
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame frmBasicInfo 
            Caption         =   "General"
            ClipControls    =   0   'False
            Height          =   1395
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Visible         =   0   'False
            Width           =   7572
            Begin VB.CheckBox chkCGMPN 
               Caption         =   "Customer Gave Me Part Number"
               Height          =   255
               Left            =   1800
               TabIndex        =   60
               Top             =   1080
               Width           =   2775
            End
            Begin MSComctlLib.ImageCombo icbItemStatus 
               Height          =   330
               Left            =   3900
               TabIndex        =   253
               Top             =   240
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   582
               _Version        =   393216
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtItemPartNbr 
               Height          =   315
               Left            =   1800
               TabIndex        =   58
               Top             =   240
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               lMaxLength      =   30
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtItemDescr 
               Height          =   315
               Left            =   1800
               TabIndex        =   59
               Top             =   720
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7641
               _ExtentY        =   550
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               lMaxLength      =   50
            End
            Begin MMRemark.RemarkViewer rvOrderLine 
               Height          =   810
               Index           =   0
               Left            =   6480
               TabIndex        =   287
               Top             =   240
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewOrderLine"
               Caption         =   "Line Remarks"
            End
            Begin VB.Image imgType 
               Height          =   1050
               Left            =   120
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Descr"
               Height          =   210
               Index           =   41
               Left            =   1200
               TabIndex        =   63
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status"
               Height          =   216
               Index           =   42
               Left            =   3300
               TabIndex        =   62
               Top             =   288
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Part #"
               Height          =   204
               Index           =   40
               Left            =   1200
               TabIndex        =   61
               Top             =   280
               Width           =   492
            End
         End
         Begin VB.Frame frmGasket 
            BorderStyle     =   0  'None
            Caption         =   "Specify Gasket"
            ClipControls    =   0   'False
            Height          =   2892
            Left            =   120
            TabIndex        =   148
            Top             =   2640
            Visible         =   0   'False
            Width           =   8952
            Begin VB.Frame Frame1 
               Caption         =   "Gasket Specs"
               ClipControls    =   0   'False
               Height          =   2775
               Index           =   3
               Left            =   0
               TabIndex        =   149
               Top             =   0
               Width           =   5052
               Begin VB.Frame Frame1 
                  ClipControls    =   0   'False
                  Height          =   735
                  Index           =   2
                  Left            =   180
                  TabIndex        =   341
                  Top             =   240
                  Width           =   1695
                  Begin VB.OptionButton optGasketType 
                     Caption         =   "Magnetic"
                     Height          =   195
                     Index           =   0
                     Left            =   120
                     TabIndex        =   153
                     Top             =   180
                     Value           =   -1  'True
                     Width           =   1095
                  End
                  Begin VB.OptionButton optGasketType 
                     Caption         =   "Compression"
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   154
                     Top             =   420
                     Width           =   1335
                  End
               End
               Begin VB.CheckBox chkGasketOptions 
                  Caption         =   "Inverted"
                  Height          =   255
                  Index           =   1
                  Left            =   3000
                  TabIndex        =   162
                  Tag             =   "1"
                  Top             =   1620
                  Width           =   1215
               End
               Begin VB.CheckBox chkGasketOptions 
                  Caption         =   "Dart-to-Dart"
                  Height          =   255
                  Index           =   0
                  Left            =   3000
                  TabIndex        =   161
                  Tag             =   "8"
                  Top             =   1320
                  Width           =   1455
               End
               Begin VB.CheckBox chkGasketOptions 
                  Caption         =   "No Magnet LHS"
                  Height          =   255
                  Index           =   2
                  Left            =   3000
                  TabIndex        =   163
                  Tag             =   "4"
                  Top             =   1920
                  Width           =   1932
               End
               Begin VB.CheckBox chkGasketOptions 
                  Caption         =   "No Magnet RHS"
                  Height          =   255
                  Index           =   3
                  Left            =   3000
                  TabIndex        =   164
                  Tag             =   "2"
                  Top             =   2220
                  Width           =   1695
               End
               Begin VB.Frame Frame1 
                  ClipControls    =   0   'False
                  Height          =   735
                  Index           =   1
                  Left            =   180
                  TabIndex        =   340
                  Top             =   960
                  Width           =   1695
                  Begin VB.OptionButton optGasketSides 
                     Caption         =   "3-Sided"
                     Height          =   255
                     Index           =   1
                     Left            =   180
                     TabIndex        =   158
                     Top             =   420
                     Width           =   1215
                  End
                  Begin VB.OptionButton optGasketSides 
                     Caption         =   "4-Sided"
                     Height          =   255
                     Index           =   0
                     Left            =   180
                     TabIndex        =   157
                     Top             =   180
                     Value           =   -1  'True
                     Width           =   1215
                  End
               End
               Begin VB.ComboBox cboGasket 
                  Height          =   315
                  Left            =   3000
                  Style           =   2  'Dropdown List
                  TabIndex        =   156
                  Top             =   360
                  Width           =   1452
               End
               Begin InchWorm.LengthEntry lenGasket 
                  Height          =   285
                  Index           =   1
                  Left            =   780
                  TabIndex        =   159
                  Top             =   1860
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Caption         =   "Gasket Door Width"
                  InchesOnly      =   -1  'True
                  MinValue        =   5
                  MaxValue        =   96
               End
               Begin InchWorm.LengthEntry lenGasket 
                  Height          =   285
                  Index           =   0
                  Left            =   780
                  TabIndex        =   160
                  Top             =   2280
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Caption         =   "Gasket Door Height"
                  InchesOnly      =   -1  'True
                  MinValue        =   5
                  MaxValue        =   96
               End
               Begin VB.Label Label1 
                  Caption         =   "Height"
                  Height          =   255
                  Index           =   56
                  Left            =   180
                  TabIndex        =   339
                  Top             =   2340
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Caption         =   "Width"
                  Height          =   255
                  Index           =   55
                  Left            =   180
                  TabIndex        =   338
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.Label Label1 
                  Caption         =   "Molded in"
                  Height          =   255
                  Index           =   92
                  Left            =   2040
                  TabIndex        =   152
                  Top             =   780
                  Width           =   735
               End
               Begin VB.Label lblGasketMatlUsed 
                  Caption         =   "Los Angeles or St. Louis"
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   151
                  Top             =   780
                  Width           =   1935
               End
               Begin VB.Label Label1 
                  Caption         =   "Material"
                  Height          =   255
                  Index           =   57
                  Left            =   2040
                  TabIndex        =   150
                  Top             =   360
                  Width           =   735
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Line Remarks"
               Height          =   2895
               Index           =   17
               Left            =   5160
               TabIndex        =   277
               Top             =   0
               Width           =   3735
               Begin VB.ComboBox cboMMType 
                  Height          =   315
                  Index           =   1
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   249
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.TextBox txtMMRemark 
                  Height          =   2175
                  Index           =   1
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   250
                  Top             =   600
                  Width           =   2655
               End
               Begin MMRemark.RemarkViewer rvOrderLine 
                  Height          =   810
                  Index           =   2
                  Left            =   2830
                  TabIndex        =   289
                  ToolTipText     =   "Line Remarks"
                  Top             =   600
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   1429
                  ContextID       =   "ViewOrderLine"
                  Caption         =   ""
               End
            End
         End
         Begin VB.Frame frmAssembly 
            Caption         =   "Assembly Information"
            ClipControls    =   0   'False
            Height          =   1932
            Left            =   120
            TabIndex        =   131
            Top             =   3600
            Visible         =   0   'False
            Width           =   8952
            Begin VB.ComboBox cboMake 
               Height          =   315
               Index           =   1
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   231
               Top             =   360
               Width           =   3015
            End
            Begin VB.CommandButton cmdViewCat 
               Caption         =   "View Catalog Page..."
               Height          =   312
               Index           =   1
               Left            =   4320
               TabIndex        =   236
               Top             =   720
               Width           =   2052
            End
            Begin VB.CommandButton cmdCopyModel 
               Caption         =   "Copy From Previous Item"
               Height          =   312
               Index           =   1
               Left            =   4320
               TabIndex        =   235
               Top             =   360
               Width           =   2052
            End
            Begin VB.CommandButton cmdVendorDetails 
               Caption         =   "Vendor Details..."
               Height          =   312
               Index           =   1
               Left            =   4320
               TabIndex        =   238
               Top             =   1440
               Width           =   2052
            End
            Begin VB.CommandButton cmdResearchPO 
               Caption         =   "Purchase Orders..."
               Height          =   312
               Index           =   0
               Left            =   4320
               TabIndex        =   237
               Top             =   1080
               Width           =   2052
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtSerial 
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   233
               Top             =   1080
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtModel 
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   232
               Top             =   720
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MMRemark.RemarkViewer rvAssembly 
               Height          =   810
               Left            =   6600
               TabIndex        =   283
               Top             =   360
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewItem"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Make"
               Height          =   255
               Index           =   94
               Left            =   360
               TabIndex        =   297
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Model"
               Height          =   255
               Index           =   93
               Left            =   360
               TabIndex        =   295
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Serial #"
               Height          =   252
               Index           =   60
               Left            =   360
               TabIndex        =   133
               Top             =   1080
               Width           =   612
            End
            Begin VB.Label Label1 
               Caption         =   "Vendor"
               Height          =   255
               Index           =   65
               Left            =   480
               TabIndex        =   132
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label lblVendor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   234
               Top             =   1440
               Width           =   3015
            End
         End
         Begin VB.Frame frmItemList 
            BorderStyle     =   0  'None
            Caption         =   "Frame13"
            ClipControls    =   0   'False
            Height          =   4095
            Left            =   120
            TabIndex        =   142
            Top             =   1560
            Width           =   8952
            Begin VB.CommandButton cmdUnAuthorize 
               Caption         =   "UnAuthori&ze"
               Height          =   312
               Left            =   1140
               TabIndex        =   342
               Top             =   3600
               Width           =   1092
            End
            Begin VB.CommandButton cmdAuthorizeAll 
               Caption         =   "&Authorize All"
               Height          =   312
               Left            =   0
               TabIndex        =   76
               Top             =   3600
               Width           =   1092
            End
            Begin VB.ComboBox cboWarehouse 
               Height          =   315
               Index           =   1
               Left            =   7800
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   3600
               Width           =   972
            End
            Begin NEWSOTALib.SOTACurrency txtTotalPrice 
               Height          =   315
               Left            =   5640
               TabIndex        =   143
               Top             =   3600
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1926
               _ExtentY        =   550
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               sDecimalPlaces  =   2
            End
            Begin NEWSOTALib.SOTACurrency txtTotalTax 
               Height          =   315
               Left            =   3360
               TabIndex        =   144
               Top             =   3600
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1926
               _ExtentY        =   550
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               sDecimalPlaces  =   2
            End
            Begin GridEX20.GridEX gdxItems 
               Height          =   3375
               Left            =   0
               TabIndex        =   75
               Top             =   120
               Width           =   8952
               _ExtentX        =   15796
               _ExtentY        =   5953
               Version         =   "2.0"
               AutomaticSort   =   -1  'True
               ShowToolTips    =   -1  'True
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               OLEDropMode     =   1
               ColumnAutoResize=   -1  'True
               DetectRowDrag   =   -1  'True
               HideSelection   =   2
               MethodHoldFields=   -1  'True
               ContScroll      =   -1  'True
               AllowColumnDrag =   0   'False
               AllowEdit       =   0   'False
               GroupByBoxVisible=   0   'False
               ImageWidth      =   32
               ImageHeight     =   32
               DataMode        =   99
               ColumnHeaderHeight=   285
               FrozenColumns   =   3
               IntProp1        =   0
               ColumnsCount    =   12
               Column(1)       =   "FOrder.frx":F99B
               Column(2)       =   "FOrder.frx":FB57
               Column(3)       =   "FOrder.frx":FD37
               Column(4)       =   "FOrder.frx":FF33
               Column(5)       =   "FOrder.frx":1011F
               Column(6)       =   "FOrder.frx":1036B
               Column(7)       =   "FOrder.frx":105DB
               Column(8)       =   "FOrder.frx":10857
               Column(9)       =   "FOrder.frx":10ACF
               Column(10)      =   "FOrder.frx":10FEB
               Column(11)      =   "FOrder.frx":117FF
               Column(12)      =   "FOrder.frx":11BE7
               FmtConditionsCount=   1
               FmtCondition(1) =   "FOrder.frx":11D93
               FormatStylesCount=   7
               FormatStyle(1)  =   "FOrder.frx":11EDF
               FormatStyle(2)  =   "FOrder.frx":11FBF
               FormatStyle(3)  =   "FOrder.frx":120F7
               FormatStyle(4)  =   "FOrder.frx":121A7
               FormatStyle(5)  =   "FOrder.frx":1225B
               FormatStyle(6)  =   "FOrder.frx":12333
               FormatStyle(7)  =   "FOrder.frx":123EB
               ImageCount      =   0
               PrinterProperties=   "FOrder.frx":1249F
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Sales Amount"
               Height          =   210
               Index           =   62
               Left            =   4440
               TabIndex        =   147
               Top             =   3630
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Warehouse"
               Height          =   252
               Index           =   2
               Left            =   6840
               TabIndex        =   146
               Top             =   3630
               Width           =   852
            End
            Begin VB.Label Label1 
               Caption         =   "Sales Tax"
               Height          =   255
               Index           =   69
               Left            =   2520
               TabIndex        =   145
               Top             =   3630
               Width           =   735
            End
         End
         Begin VB.Frame frmStock 
            Caption         =   "Finished Good Information"
            ClipControls    =   0   'False
            Height          =   1932
            Left            =   120
            TabIndex        =   128
            Top             =   3600
            Visible         =   0   'False
            Width           =   8952
            Begin VB.ComboBox cboMake 
               Height          =   315
               Index           =   2
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   239
               Top             =   360
               Width           =   3015
            End
            Begin VB.CommandButton cmdViewCat 
               Caption         =   "View Catalog (pg 26)..."
               Height          =   312
               Index           =   0
               Left            =   4320
               TabIndex        =   244
               Top             =   720
               Width           =   2052
            End
            Begin VB.CommandButton cmdCopyModel 
               Caption         =   "Copy From Previous Item"
               Height          =   312
               Index           =   2
               Left            =   4320
               TabIndex        =   243
               Top             =   360
               Width           =   2052
            End
            Begin VB.CommandButton cmdVendorDetails 
               Caption         =   "Vendor Details..."
               Height          =   312
               Index           =   2
               Left            =   4320
               TabIndex        =   246
               Top             =   1440
               Width           =   2052
            End
            Begin VB.CommandButton cmdResearchPO 
               Caption         =   "Purchase Orders..."
               Height          =   312
               Index           =   1
               Left            =   4320
               TabIndex        =   245
               Top             =   1080
               Width           =   2052
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtSerial 
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   241
               Top             =   1080
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtModel 
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   240
               Top             =   720
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MMRemark.RemarkViewer rvFinGood 
               Height          =   810
               Left            =   6720
               TabIndex        =   285
               Top             =   360
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewItem"
               Caption         =   "Item Remarks"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Make"
               Height          =   255
               Index           =   95
               Left            =   480
               TabIndex        =   298
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Model"
               Height          =   255
               Index           =   89
               Left            =   480
               TabIndex        =   293
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Serial #"
               Height          =   252
               Index           =   61
               Left            =   360
               TabIndex        =   130
               Top             =   1080
               Width           =   612
            End
            Begin VB.Label Label1 
               Caption         =   "Vendor"
               Height          =   255
               Index           =   66
               Left            =   480
               TabIndex        =   129
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label lblVendor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   242
               Top             =   1440
               Width           =   3015
            End
         End
         Begin VB.Frame frmWire 
            BorderStyle     =   0  'None
            Caption         =   "Warmer Wire"
            ClipControls    =   0   'False
            Height          =   2892
            Left            =   120
            TabIndex        =   103
            Top             =   2640
            Visible         =   0   'False
            Width           =   8952
            Begin VB.Frame Frame1 
               Caption         =   "Wire Length"
               ClipControls    =   0   'False
               Height          =   2652
               Index           =   11
               Left            =   0
               TabIndex        =   104
               Top             =   120
               Width           =   4812
               Begin VB.OptionButton optLengthAlgorithm 
                  Caption         =   "Specify Overall Length"
                  Height          =   372
                  Index           =   0
                  Left            =   120
                  TabIndex        =   106
                  Top             =   360
                  Width           =   2052
               End
               Begin VB.OptionButton optLengthAlgorithm 
                  Caption         =   "Specify Door Dimensions"
                  Height          =   252
                  Index           =   1
                  Left            =   120
                  TabIndex        =   110
                  TabStop         =   0   'False
                  Top             =   1320
                  Value           =   -1  'True
                  Width           =   2532
               End
               Begin VB.Frame frmDoorStyle 
                  BorderStyle     =   0  'None
                  Caption         =   "optDoorStyle"
                  ClipControls    =   0   'False
                  Height          =   1212
                  Left            =   2880
                  TabIndex        =   115
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   1692
                  Begin VB.OptionButton optDoorStyle 
                     Caption         =   "4-Sided (Single Pass)"
                     Height          =   372
                     Index           =   0
                     Left            =   240
                     TabIndex        =   113
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1332
                  End
                  Begin VB.OptionButton optDoorStyle 
                     Caption         =   "3-Sided (Double Pass)"
                     Height          =   492
                     Index           =   1
                     Left            =   240
                     TabIndex        =   114
                     TabStop         =   0   'False
                     Top             =   720
                     Width           =   1332
                  End
               End
               Begin VB.Frame frmWirePasses 
                  BorderStyle     =   0  'None
                  Caption         =   "optWirePasses"
                  ClipControls    =   0   'False
                  Height          =   972
                  Left            =   2880
                  TabIndex        =   105
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1692
                  Begin VB.OptionButton optWirePasses 
                     Caption         =   "Single Pass"
                     Height          =   252
                     Index           =   0
                     Left            =   240
                     TabIndex        =   108
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1332
                  End
                  Begin VB.OptionButton optWirePasses 
                     Caption         =   "Double Pass"
                     Height          =   252
                     Index           =   1
                     Left            =   240
                     TabIndex        =   109
                     TabStop         =   0   'False
                     Top             =   600
                     Width           =   1332
                  End
               End
               Begin InchWorm.LengthEntry lenWireLength 
                  Height          =   288
                  Left            =   1200
                  TabIndex        =   107
                  Top             =   720
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Caption         =   "Wire Length"
                  InchesOnly      =   -1  'True
                  MinValue        =   12
                  MaxValue        =   1200
               End
               Begin InchWorm.LengthEntry lenDoorHeight 
                  Height          =   288
                  Left            =   1200
                  TabIndex        =   111
                  Top             =   1680
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Caption         =   "Door Height"
                  InchesOnly      =   -1  'True
                  MinValue        =   6
                  MaxValue        =   110
               End
               Begin InchWorm.LengthEntry lenDoorWidth 
                  Height          =   288
                  Left            =   1200
                  TabIndex        =   112
                  Top             =   2040
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Caption         =   "Door Width"
                  InchesOnly      =   -1  'True
                  MinValue        =   6
                  MaxValue        =   96
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Height"
                  Height          =   252
                  Index           =   43
                  Left            =   480
                  TabIndex        =   118
                  Top             =   1680
                  Width           =   612
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Width"
                  Height          =   252
                  Index           =   44
                  Left            =   480
                  TabIndex        =   117
                  Top             =   2040
                  Width           =   612
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Length"
                  Height          =   252
                  Index           =   45
                  Left            =   480
                  TabIndex        =   116
                  Top             =   720
                  Width           =   612
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Electrical Properties"
               ClipControls    =   0   'False
               Height          =   2652
               Index           =   13
               Left            =   5040
               TabIndex        =   119
               Top             =   120
               Width           =   3732
               Begin VB.ComboBox cboVoltage 
                  Height          =   315
                  ItemData        =   "FOrder.frx":12677
                  Left            =   480
                  List            =   "FOrder.frx":12684
                  TabIndex        =   120
                  Text            =   "cboVoltage"
                  Top             =   600
                  Width           =   852
               End
               Begin VB.ComboBox cboWires 
                  Height          =   315
                  Left            =   2160
                  Style           =   2  'Dropdown List
                  TabIndex        =   121
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  Caption         =   "Watts Per Foot"
                  Height          =   252
                  Index           =   46
                  Left            =   480
                  TabIndex        =   127
                  Top             =   1080
                  Width           =   1212
               End
               Begin VB.Label Label1 
                  Caption         =   "Amperage"
                  Height          =   252
                  Index           =   48
                  Left            =   480
                  TabIndex        =   126
                  Top             =   1800
                  Width           =   852
               End
               Begin VB.Label lblWattsPerFoot 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   252
                  Left            =   480
                  TabIndex        =   125
                  Top             =   1320
                  Width           =   852
               End
               Begin VB.Label lblAmperage 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   252
                  Left            =   480
                  TabIndex        =   124
                  Top             =   2040
                  Width           =   852
               End
               Begin VB.Label Label1 
                  Caption         =   "Available Wires"
                  Height          =   252
                  Index           =   47
                  Left            =   2160
                  TabIndex        =   123
                  Top             =   360
                  Width           =   1332
               End
               Begin VB.Label Label1 
                  Caption         =   "Voltage"
                  Height          =   252
                  Index           =   50
                  Left            =   480
                  TabIndex        =   122
                  Top             =   360
                  Width           =   612
               End
            End
         End
         Begin VB.Frame frmShelf 
            BorderStyle     =   0  'None
            Caption         =   "Specify Wire Shelf"
            ClipControls    =   0   'False
            Height          =   2892
            Left            =   120
            TabIndex        =   87
            Top             =   2640
            Visible         =   0   'False
            Width           =   8952
            Begin VB.Frame Frame1 
               Caption         =   "Line Remarks"
               Height          =   2775
               Index           =   16
               Left            =   4800
               TabIndex        =   276
               Top             =   0
               Width           =   4095
               Begin VB.TextBox txtMMRemark 
                  Height          =   2055
                  Index           =   0
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   252
                  Top             =   600
                  Width           =   2895
               End
               Begin VB.ComboBox cboMMType 
                  Height          =   315
                  Index           =   0
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   251
                  Top             =   240
                  Width           =   2055
               End
               Begin MMRemark.RemarkViewer rvOrderLine 
                  Height          =   810
                  Index           =   1
                  Left            =   3120
                  TabIndex        =   288
                  ToolTipText     =   "Line Remarks"
                  Top             =   600
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   1429
                  ContextID       =   "ViewOrderLine"
                  Caption         =   ""
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Dimensions"
               ClipControls    =   0   'False
               Height          =   2772
               Index           =   14
               Left            =   0
               TabIndex        =   98
               Top             =   0
               Width           =   2772
               Begin VB.ComboBox cboFinish 
                  Height          =   315
                  Left            =   1080
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   90
                  Top             =   960
                  Width           =   1452
               End
               Begin VB.ComboBox cboFrame 
                  Height          =   315
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   89
                  Top             =   360
                  Width           =   1455
               End
               Begin InchWorm.LengthEntry lenShelfWidth 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   92
                  Top             =   2040
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Caption         =   "Shelf Width"
                  InchesOnly      =   -1  'True
                  MinValue        =   6
                  MaxValue        =   96
               End
               Begin InchWorm.LengthEntry lenShelfDepth 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   91
                  Top             =   1560
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Caption         =   "Shelf Depth"
                  InchesOnly      =   -1  'True
                  MinValue        =   4
                  MaxValue        =   96
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Finish"
                  Height          =   252
                  Index           =   59
                  Left            =   360
                  TabIndex        =   102
                  Top             =   960
                  Width           =   612
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Frame Diameter"
                  Height          =   492
                  Index           =   51
                  Left            =   240
                  TabIndex        =   101
                  Top             =   360
                  Width           =   732
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Depth"
                  Height          =   252
                  Index           =   53
                  Left            =   240
                  TabIndex        =   100
                  Top             =   1560
                  Width           =   732
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Width"
                  Height          =   252
                  Index           =   52
                  Left            =   480
                  TabIndex        =   99
                  Top             =   2040
                  Width           =   492
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Options"
               ClipControls    =   0   'False
               Height          =   2772
               Index           =   12
               Left            =   3000
               TabIndex        =   88
               Top             =   0
               Width           =   1695
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Cut-Out"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   93
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Straight Leg"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   94
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Bent Leg"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   95
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Product Stop"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   96
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Support"
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   97
                  Top             =   1320
                  Width           =   1335
               End
            End
         End
         Begin VB.Frame frmSpecialOrder 
            Caption         =   "Specify Special Order"
            ClipControls    =   0   'False
            Height          =   2892
            Left            =   120
            TabIndex        =   84
            Top             =   2640
            Visible         =   0   'False
            Width           =   8952
            Begin VB.ComboBox cboMake 
               Height          =   315
               Index           =   3
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   218
               Top             =   360
               Width           =   2655
            End
            Begin VB.ComboBox cboVendor 
               Height          =   315
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   221
               Top             =   1440
               Width           =   2655
            End
            Begin VB.CommandButton cmdPartsWiz 
               Caption         =   "Parts Wiz..."
               Height          =   312
               Left            =   3840
               TabIndex        =   223
               Top             =   720
               Width           =   1935
            End
            Begin VB.CommandButton cmdCopyModel 
               Caption         =   "Copy From Previous Item"
               Height          =   312
               Index           =   3
               Left            =   3840
               TabIndex        =   222
               Top             =   360
               Width           =   1935
            End
            Begin VB.CommandButton cmdVendorDetails 
               Caption         =   "Vendor Details..."
               Height          =   312
               Index           =   3
               Left            =   3840
               TabIndex        =   224
               Top             =   1440
               Width           =   1935
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtSerial 
               Height          =   315
               Index           =   3
               Left            =   1080
               TabIndex        =   220
               Top             =   1080
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtModel 
               Height          =   315
               Index           =   3
               Left            =   1080
               TabIndex        =   219
               Top             =   720
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MMRemark.RemarkViewer rvVendor 
               Height          =   810
               Left            =   1080
               TabIndex        =   286
               Top             =   1920
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   1429
               ContextID       =   "ViewVendor"
               Caption         =   "Vendor Remarks"
            End
            Begin VB.Frame Frame1 
               Caption         =   "Line Remarks"
               Height          =   2655
               Index           =   18
               Left            =   5880
               TabIndex        =   278
               Top             =   120
               Width           =   2960
               Begin VB.ComboBox cboMMType 
                  Height          =   315
                  Index           =   2
                  ItemData        =   "FOrder.frx":12696
                  Left            =   120
                  List            =   "FOrder.frx":12698
                  Style           =   2  'Dropdown List
                  TabIndex        =   247
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.TextBox txtMMRemark 
                  Height          =   1935
                  Index           =   2
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   248
                  Top             =   600
                  Width           =   1905
               End
               Begin MMRemark.RemarkViewer rvOrderLine 
                  Height          =   810
                  Index           =   3
                  Left            =   2080
                  TabIndex        =   290
                  ToolTipText     =   "Line Remarks"
                  Top             =   600
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   1429
                  ContextID       =   "ViewOrderLine"
                  Caption         =   ""
               End
            End
            Begin VB.Label Label1 
               Caption         =   "Make"
               Height          =   255
               Index           =   91
               Left            =   360
               TabIndex        =   296
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Model"
               Height          =   255
               Index           =   90
               Left            =   357
               TabIndex        =   294
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Serial"
               Height          =   252
               Index           =   63
               Left            =   360
               TabIndex        =   86
               Top             =   1080
               Width           =   492
            End
            Begin VB.Label Label1 
               Caption         =   "Vendor"
               Height          =   252
               Index           =   64
               Left            =   360
               TabIndex        =   85
               Top             =   1440
               Width           =   612
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Total Price"
            Height          =   252
            Index           =   58
            Left            =   3720
            TabIndex        =   155
            Top             =   4920
            Width           =   972
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpFind 
         Height          =   5565
         Left            =   30
         TabIndex        =   165
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder.frx":1269A
         Begin VB.CommandButton cmdFilterByPart 
            Caption         =   "Filter"
            Height          =   315
            Left            =   7020
            TabIndex        =   345
            Top             =   5220
            Width           =   615
         End
         Begin VB.TextBox txtFilterByPart 
            Height          =   315
            Left            =   5520
            TabIndex        =   344
            Top             =   5220
            Width           =   1455
         End
         Begin VB.CommandButton cmdSelectCustomer 
            Caption         =   "Fin&d Account"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   420
            TabIndex        =   304
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton cmdContactMgr 
            Caption         =   "Contacts"
            Height          =   315
            Index           =   0
            Left            =   7860
            TabIndex        =   328
            Top             =   4860
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.ComboBox cboSearchType 
            Height          =   315
            ItemData        =   "FOrder.frx":126C2
            Left            =   4920
            List            =   "FOrder.frx":126D8
            Style           =   2  'Dropdown List
            TabIndex        =   306
            Top             =   540
            Width           =   1935
         End
         Begin VB.ComboBox cboOrderStatus 
            Height          =   315
            ItemData        =   "FOrder.frx":1271D
            Left            =   2460
            List            =   "FOrder.frx":12757
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   5220
            Width           =   1635
         End
         Begin VB.CommandButton cmdNewSearch 
            Caption         =   "New Se&arch"
            Height          =   312
            Left            =   7860
            TabIndex        =   178
            Top             =   5220
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton cmdSelectCustomer 
            Caption         =   "&New Customer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   420
            TabIndex        =   307
            Top             =   1380
            Width           =   1815
         End
         Begin VB.CommandButton cmdSelectCustomer 
            Caption         =   "Wal&k-up Order"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   2
            Left            =   420
            TabIndex        =   308
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CommandButton cmdNewOrder 
            Caption         =   "&New Order"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   175
            Top             =   5220
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton cmdLoadOrder 
            Caption         =   "Load Ord&er"
            Enabled         =   0   'False
            Height          =   312
            Index           =   0
            Left            =   120
            TabIndex        =   176
            Top             =   4860
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton cmdSelectCustomer 
            Caption         =   "Mis&c Customer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   3
            Left            =   420
            TabIndex        =   309
            Top             =   3180
            Width           =   1815
         End
         Begin NEWSOTALib.SOTAMaskedEdit txtCustSearch 
            Height          =   312
            Left            =   2580
            TabIndex        =   305
            ToolTipText     =   "Enter the first few characters of the customer name"
            Top             =   540
            Width           =   1980
            _Version        =   65536
            _ExtentX        =   3492
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin GridEX20.GridEX gdxCustOrders 
            Height          =   3315
            Left            =   120
            TabIndex        =   167
            Top             =   1440
            Visible         =   0   'False
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   5847
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            LockType        =   1
            Options         =   1
            RecordsetType   =   3
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   1
            ColumnHeaderHeight=   285
            ColumnsCount    =   18
            Column(1)       =   "FOrder.frx":12839
            Column(2)       =   "FOrder.frx":12A11
            Column(3)       =   "FOrder.frx":133ED
            Column(4)       =   "FOrder.frx":135B1
            Column(5)       =   "FOrder.frx":13711
            Column(6)       =   "FOrder.frx":139D1
            Column(7)       =   "FOrder.frx":13B7D
            Column(8)       =   "FOrder.frx":13CD5
            Column(9)       =   "FOrder.frx":13E6D
            Column(10)      =   "FOrder.frx":13FF5
            Column(11)      =   "FOrder.frx":1415D
            Column(12)      =   "FOrder.frx":142F5
            Column(13)      =   "FOrder.frx":1444D
            Column(14)      =   "FOrder.frx":145BD
            Column(15)      =   "FOrder.frx":14731
            Column(16)      =   "FOrder.frx":14865
            Column(17)      =   "FOrder.frx":149AD
            Column(18)      =   "FOrder.frx":14AC5
            SortKeysCount   =   1
            SortKey(1)      =   "FOrder.frx":14C05
            FmtConditionsCount=   1
            FmtCondition(1) =   "FOrder.frx":14C6D
            FormatStylesCount=   6
            FormatStyle(1)  =   "FOrder.frx":14DC5
            FormatStyle(2)  =   "FOrder.frx":14EA5
            FormatStyle(3)  =   "FOrder.frx":14FDD
            FormatStyle(4)  =   "FOrder.frx":1508D
            FormatStyle(5)  =   "FOrder.frx":15141
            FormatStyle(6)  =   "FOrder.frx":15219
            ImageCount      =   0
            PrinterProperties=   "FOrder.frx":152D1
         End
         Begin VB.Frame frmCustInfo 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1380
            Left            =   120
            TabIndex        =   168
            Top             =   120
            Visible         =   0   'False
            Width           =   9075
            Begin VB.CheckBox chkShowOrdersForShipAddr 
               Caption         =   "Show orders for all shipping addresses"
               Height          =   192
               Left            =   3300
               TabIndex        =   325
               Top             =   1020
               Width           =   3168
            End
            Begin VB.Label lblOrderCount 
               Caption         =   "lblOrderCount"
               Height          =   315
               Left            =   0
               TabIndex        =   326
               Top             =   1020
               Width           =   1935
            End
            Begin VB.Label lblCustAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblCustAddress"
               ForeColor       =   &H80000008&
               Height          =   960
               Left            =   3300
               TabIndex        =   174
               Top             =   0
               Width           =   2190
            End
            Begin VB.Label lblCustInfo 
               Caption         =   "Customer Ship Address"
               Height          =   495
               Index           =   0
               Left            =   2220
               TabIndex        =   173
               Top             =   0
               UseMnemonic     =   0   'False
               Width           =   975
            End
            Begin VB.Label lblOrderAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblOrderAddress"
               ForeColor       =   &H80000008&
               Height          =   960
               Left            =   6720
               TabIndex        =   172
               Top             =   0
               Visible         =   0   'False
               Width           =   2190
            End
            Begin VB.Label lblCustID 
               Appearance      =   0  'Flat
               Caption         =   "lblCustID(1)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   0
               TabIndex        =   171
               Top             =   0
               Width           =   2115
            End
            Begin VB.Label lblCustType 
               Appearance      =   0  'Flat
               Caption         =   "lblCustType(1)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   170
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label lblCustInfo 
               Caption         =   "Order         Ship Address"
               Height          =   375
               Index           =   1
               Left            =   5640
               TabIndex        =   169
               Top             =   0
               UseMnemonic     =   0   'False
               Visible         =   0   'False
               Width           =   1035
            End
         End
         Begin VB.Frame frmCustSearch 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   1095
            Left            =   60
            TabIndex        =   166
            Top             =   120
            Width           =   9015
         End
         Begin VB.Label lblFilterByPart 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter by Keywords"
            Height          =   195
            Left            =   4200
            TabIndex        =   343
            Top             =   5280
            Width           =   1275
         End
         Begin VB.Label lblExplain 
            Caption         =   "Use this to create a quote for a customer you intend to setup with an account before committing your order."
            Height          =   612
            Index           =   0
            Left            =   2580
            TabIndex        =   312
            Top             =   1380
            Width           =   4152
         End
         Begin VB.Label lblExplain 
            Caption         =   $"FOrder.frx":154A9
            Height          =   612
            Index           =   1
            Left            =   2580
            TabIndex        =   311
            Top             =   2280
            Width           =   4152
         End
         Begin VB.Label lblExplain 
            Caption         =   $"FOrder.frx":15555
            Height          =   612
            Index           =   2
            Left            =   2580
            TabIndex        =   310
            Top             =   3180
            Width           =   4152
         End
         Begin VB.Label lblOrderStatus 
            Caption         =   "Order Status"
            Height          =   315
            Left            =   1500
            TabIndex        =   292
            Top             =   5280
            Width           =   975
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpOrder 
         Height          =   5565
         Left            =   30
         TabIndex        =   179
         Top             =   360
         Visible         =   0   'False
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder.frx":15600
         Begin VB.Frame Frame2 
            Caption         =   "Ordered By"
            Height          =   1095
            Left            =   120
            TabIndex        =   335
            Top             =   60
            Width           =   4575
            Begin VB.CommandButton cmdEmailQuote 
               Height          =   315
               Left            =   3720
               Picture         =   "FOrder.frx":15628
               Style           =   1  'Graphical
               TabIndex        =   183
               ToolTipText     =   "Email a quote to customer"
               Top             =   660
               Width           =   315
            End
            Begin VB.TextBox txtInfo 
               Height          =   315
               Left            =   1140
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   180
               Text            =   "FOrder.frx":15703
               Top             =   240
               Width           =   2355
            End
            Begin VB.CommandButton cmdEditContact 
               Caption         =   "Edit"
               Height          =   315
               Left            =   3120
               TabIndex        =   182
               Top             =   660
               Width           =   495
            End
            Begin VB.ComboBox cboContact 
               Height          =   315
               Left            =   1140
               TabIndex        =   181
               Text            =   "cboContact"
               Top             =   660
               Width           =   1935
            End
            Begin VB.Label lblNote 
               Alignment       =   1  'Right Justify
               Caption         =   "Info"
               Height          =   255
               Left            =   120
               TabIndex        =   337
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Contact"
               Height          =   315
               Index           =   0
               Left            =   180
               TabIndex        =   336
               Top             =   720
               Width           =   795
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Shipping"
            ClipControls    =   0   'False
            Height          =   2715
            Index           =   4
            Left            =   120
            TabIndex        =   205
            Top             =   2760
            Width           =   4575
            Begin VB.TextBox txtShipToContact 
               Height          =   285
               Index           =   1
               Left            =   1080
               MaxLength       =   17
               TabIndex        =   195
               Top             =   1680
               Width           =   2655
            End
            Begin VB.TextBox txtShipToContact 
               Height          =   285
               Index           =   0
               Left            =   1080
               MaxLength       =   40
               TabIndex        =   194
               Top             =   1320
               Width           =   2655
            End
            Begin VB.CheckBox chkDropShip 
               Caption         =   "Drop Ship"
               Height          =   195
               Left            =   2880
               TabIndex        =   190
               Top             =   180
               Width           =   1215
            End
            Begin VB.ComboBox cboShipVia 
               Height          =   315
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   192
               Top             =   870
               Width           =   1680
            End
            Begin VB.ComboBox cboWarehouse 
               Height          =   315
               Index           =   0
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   191
               Top             =   480
               Width           =   1095
            End
            Begin VB.CheckBox chkShipComplete 
               Caption         =   "Ship Complete"
               Height          =   195
               Left            =   1080
               TabIndex        =   189
               Top             =   180
               Width           =   1425
            End
            Begin VB.CheckBox chkDefaultShipMeth 
               Caption         =   "Default Shipping Method"
               Enabled         =   0   'False
               Height          =   615
               Left            =   2880
               TabIndex        =   193
               Top             =   720
               Width           =   1575
            End
            Begin VB.CommandButton cmdUPSUpdate 
               Caption         =   "&Change"
               Height          =   300
               Left            =   2760
               TabIndex        =   197
               Top             =   2055
               Width           =   995
            End
            Begin VB.CheckBox chkBillRecipient 
               Caption         =   "Bill Recipient"
               Height          =   255
               Left            =   1080
               TabIndex        =   198
               Top             =   2355
               Width           =   1350
            End
            Begin VB.TextBox txtUPSAcct 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   8
               MultiLine       =   -1  'True
               TabIndex        =   196
               Top             =   2055
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ship To"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   330
               Top             =   1365
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Phone #"
               Height          =   255
               Index           =   13
               Left            =   240
               TabIndex        =   329
               Top             =   1725
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ship Via"
               Height          =   255
               Index           =   24
               Left            =   240
               TabIndex        =   206
               Top             =   945
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ship From"
               Height          =   255
               Index           =   25
               Left            =   240
               TabIndex        =   207
               Top             =   555
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "UPS Acct"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   208
               Top             =   2100
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdLiteEdit 
            Caption         =   "Lite Edit..."
            Height          =   375
            Left            =   7680
            TabIndex        =   201
            ToolTipText     =   "Edit a Committed Order"
            Top             =   5160
            Width           =   1380
         End
         Begin VB.CommandButton cmdCreateRMA 
            Caption         =   "Create RMA..."
            Height          =   375
            Left            =   4800
            TabIndex        =   199
            Top             =   5160
            Width           =   1380
         End
         Begin VB.Frame frmCreditCard 
            Caption         =   "Credit Card Information"
            ClipControls    =   0   'False
            Height          =   2592
            Left            =   4800
            TabIndex        =   209
            Top             =   2160
            Width           =   4215
            Begin VB.CommandButton cmdCCEdit 
               Caption         =   "Edit..."
               Height          =   375
               Left            =   3000
               TabIndex        =   357
               Top             =   2040
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label lblCCStatus 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   321
               Top             =   2100
               Width           =   1212
            End
            Begin VB.Label Label3 
               Caption         =   "Status"
               Height          =   255
               Left            =   180
               TabIndex        =   320
               Top             =   2160
               Width           =   795
            End
            Begin VB.Label lblCCZipCode 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   318
               Top             =   1740
               Width           =   1212
            End
            Begin VB.Label lblCCStreet 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   317
               Top             =   1380
               Width           =   2172
            End
            Begin VB.Label lblHolderName 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   316
               Top             =   1020
               Width           =   2232
            End
            Begin VB.Label lblCCNo 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   315
               Top             =   300
               Width           =   2292
            End
            Begin VB.Label lblCCType 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   180
               TabIndex        =   314
               Top             =   300
               Width           =   972
            End
            Begin VB.Label lblExpireDate 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   313
               Top             =   660
               Width           =   1332
            End
            Begin VB.Label Label1 
               Caption         =   "Street"
               Height          =   255
               Index           =   81
               Left            =   180
               TabIndex        =   279
               Top             =   1440
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "ZipCode"
               Height          =   255
               Index           =   82
               Left            =   180
               TabIndex        =   212
               Top             =   1800
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "Expires"
               Height          =   255
               Index           =   74
               Left            =   180
               TabIndex        =   210
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Name"
               Height          =   255
               Index           =   80
               Left            =   180
               TabIndex        =   211
               Top             =   1080
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Purchasing"
            ClipControls    =   0   'False
            Height          =   1935
            Index           =   6
            Left            =   4800
            TabIndex        =   202
            Top             =   60
            Width           =   4215
            Begin VB.CommandButton cmdManageDropShips 
               Caption         =   "Manage Drop Ships"
               Height          =   375
               Left            =   1560
               TabIndex        =   358
               Top             =   1440
               Width           =   1860
            End
            Begin VB.TextBox txtPO 
               Height          =   315
               Left            =   1560
               TabIndex        =   184
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox chkReqPO 
               Caption         =   "PO Required"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1560
               TabIndex        =   185
               TabStop         =   0   'False
               Top             =   600
               Width           =   1335
            End
            Begin VB.ComboBox cboTerms 
               Height          =   315
               ItemData        =   "FOrder.frx":1570D
               Left            =   1560
               List            =   "FOrder.frx":1570F
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   186
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Purchase Order"
               Height          =   252
               Index           =   20
               Left            =   180
               TabIndex        =   203
               Top             =   300
               Width           =   1212
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Terms"
               Height          =   252
               Index           =   21
               Left            =   900
               TabIndex        =   204
               Top             =   1020
               Width           =   492
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "General"
            ClipControls    =   0   'False
            Height          =   1515
            Index           =   5
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   4575
            Begin VB.ComboBox cboCSR 
               Height          =   315
               Left            =   1140
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   240
               Width           =   1935
            End
            Begin MMRemark.RemarkViewer rvOrder 
               Height          =   810
               Left            =   3360
               TabIndex        =   188
               Top             =   180
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   1429
               ContextID       =   "ViewOrder"
               Caption         =   "Order Remarks"
            End
            Begin VB.Label lblLastUpdate 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1140
               TabIndex        =   0
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Last Updated"
               Height          =   255
               Index           =   17
               Left            =   60
               TabIndex        =   1
               Top             =   1140
               Width           =   1020
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "CSR"
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   4
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblDate 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1140
               TabIndex        =   2
               Top             =   660
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Created"
               Height          =   255
               Index           =   22
               Left            =   120
               TabIndex        =   3
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdSpecialHandling 
            Caption         =   "Special Handling"
            Height          =   375
            Left            =   6240
            TabIndex        =   200
            ToolTipText     =   "Add special handling remarks"
            Top             =   5160
            Width           =   1380
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpCustomer 
         Height          =   5565
         Left            =   30
         TabIndex        =   213
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder.frx":15711
         Begin VB.Frame Frame1 
            Caption         =   "Customer Information"
            ClipControls    =   0   'False
            Height          =   1815
            Index           =   10
            Left            =   120
            TabIndex        =   264
            Top             =   120
            Width           =   8895
            Begin VB.PictureBox picCustHold 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   300
               Picture         =   "FOrder.frx":15739
               ScaleHeight     =   480
               ScaleWidth      =   480
               TabIndex        =   284
               Top             =   1260
               Visible         =   0   'False
               Width           =   480
            End
            Begin MMRemark.RemarkViewer rvCustomer 
               Height          =   810
               Left            =   7440
               TabIndex        =   267
               Top             =   360
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewCustomer"
               Caption         =   "Customer Remarks"
            End
            Begin VB.ComboBox cboCustType 
               Height          =   315
               ItemData        =   "FOrder.frx":1584C
               Left            =   840
               List            =   "FOrder.frx":15859
               Style           =   2  'Dropdown List
               TabIndex        =   275
               Top             =   780
               Width           =   1215
            End
            Begin VB.TextBox txtCustID 
               Height          =   312
               Left            =   5400
               TabIndex        =   266
               Top             =   360
               Visible         =   0   'False
               Width           =   1212
            End
            Begin VB.TextBox txtCustName 
               Height          =   312
               Left            =   840
               TabIndex        =   265
               Top             =   360
               Width           =   3432
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "ID"
               Height          =   255
               Index           =   6
               Left            =   4800
               TabIndex        =   356
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblCustType 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Index           =   0
               Left            =   840
               TabIndex        =   271
               Top             =   780
               Width           =   1092
            End
            Begin VB.Label lblCustID 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   0
               Left            =   5400
               TabIndex        =   274
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Type"
               Height          =   252
               Index           =   23
               Left            =   240
               TabIndex        =   272
               Top             =   780
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Name"
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   270
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblCustName 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   840
               TabIndex        =   269
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label lblCustHold 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   492
               Left            =   840
               TabIndex        =   268
               Top             =   1200
               Width           =   3492
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Bill To"
            ClipControls    =   0   'False
            Height          =   1755
            Index           =   8
            Left            =   120
            TabIndex        =   262
            Top             =   2040
            Width           =   2775
            Begin VB.Label lblBillAddr 
               BorderStyle     =   1  'Fixed Single
               Height          =   975
               Left            =   120
               TabIndex        =   263
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Ship To"
            ClipControls    =   0   'False
            Height          =   3360
            Index           =   0
            Left            =   3000
            TabIndex        =   214
            Top             =   2055
            Width           =   6015
            Begin VB.CheckBox chkPricePackList 
               Caption         =   "Price Pack Slip"
               Height          =   255
               Left            =   240
               TabIndex        =   323
               Top             =   2160
               Width           =   1455
            End
            Begin VB.TextBox txtShipToNote 
               Height          =   312
               Left            =   240
               MaxLength       =   50
               TabIndex        =   273
               Top             =   1680
               Width           =   2655
            End
            Begin VB.CommandButton cmdContactMgr 
               Caption         =   "Contacts"
               Height          =   315
               Index           =   1
               Left            =   4440
               TabIndex        =   333
               Top             =   2760
               Width           =   1335
            End
            Begin VB.CommandButton cmdEditAddr 
               Caption         =   "Change Address"
               Height          =   315
               Index           =   1
               Left            =   240
               TabIndex        =   327
               Top             =   2760
               Width           =   1335
            End
            Begin VB.CommandButton cmdEditAddr 
               Caption         =   "This Order Only"
               Height          =   315
               Index           =   0
               Left            =   1800
               TabIndex        =   331
               Top             =   2760
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Note"
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   332
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Cell"
               Height          =   255
               Left            =   2880
               TabIndex        =   324
               Top             =   1500
               Width           =   735
            End
            Begin VB.Label lblCellPhone 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3720
               TabIndex        =   322
               Top             =   1440
               Width           =   2175
            End
            Begin VB.Label lblShipPhone 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3720
               TabIndex        =   261
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label lblShipAddr 
               BorderStyle     =   1  'Fixed Single
               Height          =   975
               Left            =   240
               TabIndex        =   260
               Top             =   360
               Width           =   2655
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Phone"
               Height          =   255
               Index           =   9
               Left            =   2880
               TabIndex        =   259
               Top             =   780
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Fax"
               Height          =   255
               Index           =   10
               Left            =   2880
               TabIndex        =   258
               Top             =   1140
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Contact"
               Height          =   255
               Index           =   11
               Left            =   2880
               TabIndex        =   217
               Top             =   420
               Width           =   735
            End
            Begin VB.Label lblShipFax 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3720
               TabIndex        =   216
               Top             =   1080
               Width           =   2175
            End
            Begin VB.Label lblShipContact 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3720
               TabIndex        =   215
               Top             =   360
               Width           =   2175
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar sbOrderStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   6225
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSComctlLib.ImageList imglRemarks 
      Left            =   8160
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
            Picture         =   "FOrder.frx":15879
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":15CCB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglStatus16 
      Left            =   7560
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1611D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":16503
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":168CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":16CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":17077
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":17442
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":17794
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":17AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":18738
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":18A8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglTypeAndStatus32 
      Left            =   8160
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":18DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":19370
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1990C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":19EF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1A314
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1A96A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1AE9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1B40E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1B9D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1BED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1C326
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1C8C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1D519
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1E16B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1E5EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1EA84
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":1F6D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglType64 
      Left            =   8640
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":20328
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":20D61
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":21A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":222CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":227F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder.frx":2354D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   432
      Left            =   4200
      TabIndex        =   319
      Top             =   3000
      Width           =   972
   End
   Begin VB.Menu mnugdxCustOrder 
      Caption         =   "gdxCustOrdersMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCustOrderDetail 
         Caption         =   "Detail"
      End
      Begin VB.Menu mnuCustOrderRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuCustOrderPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustOrderGroup 
         Caption         =   "Grouping"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCustOrderExpand 
         Caption         =   "Expand All"
      End
      Begin VB.Menu mnuCustOrderCollapse 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu mnuCustOrderFont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuCustOrderAutofit 
         Caption         =   "Autofit"
      End
      Begin VB.Menu mnuCustOrderSaveLayout 
         Caption         =   "Save Layout"
      End
      Begin VB.Menu mnuCustOrderChangeColumns 
         Caption         =   "Select Columns..."
      End
   End
   Begin VB.Menu mnugdxOrder 
      Caption         =   "gdxOrderMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuOrderRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuOrderPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrderGroup 
         Caption         =   "Grouping"
      End
      Begin VB.Menu mnuOrderExpand 
         Caption         =   "Expand All"
      End
      Begin VB.Menu mnuOrderCollapse 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu mnuOrderFont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuOrdersAutofit 
         Caption         =   "Autofit"
      End
      Begin VB.Menu mnuOrderSaveLayout 
         Caption         =   "Save Layout"
      End
      Begin VB.Menu mnuOrderRestore 
         Caption         =   "Restore Layout"
      End
      Begin VB.Menu mnuOrderChangeColumns 
         Caption         =   "Select Columns..."
      End
   End
   Begin VB.Menu mnugdxOSShipItems 
      Caption         =   "gdxOSShipItems"
      Visible         =   0   'False
      Begin VB.Menu mnuTrackUPS 
         Caption         =   "Track UPS Shipment"
      End
   End
End
Attribute VB_Name = "FOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const m_skSource = "FOrder"

Private m_lWindowID As Long

'This is used to support UPS tracking
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

'This is used by the custom Contact control
'http://www.devx.com/vb2themax/Tip/18336
'Win32 API called used to send a message to a control
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Const CB_SHOWDROPDOWN = &H14F
Const CB_GETDROPPEDSTATE = &H157

Private Const k_lItemControlMask = 1024

'used in DragDrop
Private Const k_sTextDelimiter = "|<---->|"

Private Const k_sAmpMask = "####.00"
Private Const k_sWPFMask = "####.00"

'indices for the button control array cmdSpecifyItem()
Private Const btnGasket = 0
Private Const btnShelf = 1
Private Const btnWire = 2
Private Const btnSPO = 3

'parameter values for BrokenRules & CtlWrapper EnableClass()

Public Enum ControlClass
    ccCustomer = 1
    ccItem = 2
    ccWWireType = 3 + k_lItemControlMask
    ccWWireLength = 4 + k_lItemControlMask
    ccWWireDoorDim = 5 + k_lItemControlMask
    ccshelf = 6 + k_lItemControlMask
    ccGasket = 7 + k_lItemControlMask
    ccItemSpecialOrder = 8 + k_lItemControlMask
    ccItemSPOBasicInfo = 9 + k_lItemControlMask
    ccPricing = 10 + k_lItemControlMask
    ccpurchaseorder = 11
    ccorderedby = 12
    
    'SMR Intl 01/13/2006 - so that I can enable/disable these two controls
        '- Includes txtShipToContact(0) & txtShipToContact(1)
        '- ship to contact name and ship to contact phone number.
    ccShipToData = 13
End Enum

Private Enum ItemView
    ivList
    ivComponent
    ivKit
    ivGasket
    ivShelf
    ivWire
    ivSpecialOrder
End Enum

Private Enum TabMainIndexes
    tmiExistingCustomer = 1
    tmiExistingOrder = 2
    tmiCustomer = 3
    tmiOrder = 4
    tmiLines = 5
    tmiOrderStatus = 6
    tmiRmaLines = 7
    tmiOrderHistory = 8
End Enum

'Order Status tabs
Private Enum TabOrderStatus
    tosLineItem = 1
    tosShipment = 2
    tosInvoice = 3
End Enum


Private WithEvents m_oBrokenRules As BrokenRules
Attribute m_oBrokenRules.VB_VarHelpID = -1
Private m_oValidateItem As ValidateManual
Private m_oValidateDescr As ValidateManual

Private m_oOrder As Order
Attribute m_oOrder.VB_VarHelpID = -1
Private m_oCustomer As Customer
Attribute m_oCustomer.VB_VarHelpID = -1
Private WithEvents m_oItems As Items
Attribute m_oItems.VB_VarHelpID = -1

'Used for caching keystrokes in cboContact. Maintained in LostFocus event.
Private m_Name As String

' This is the current Line tab View
Private m_View As ItemView

' These are managed by ChangeViewMode
' Each has a corresponding Change event
Private WithEvents m_oFinGood As ItemFinGood
Attribute m_oFinGood.VB_VarHelpID = -1
Private WithEvents m_oBTOKit As ItemBTOKit
Attribute m_oBTOKit.VB_VarHelpID = -1
Private WithEvents m_oWarmerWire As ItemWWire
Attribute m_oWarmerWire.VB_VarHelpID = -1
Private WithEvents m_oShelf As ItemShelf
Attribute m_oShelf.VB_VarHelpID = -1
Private WithEvents m_oGasket As ItemGasket
Attribute m_oGasket.VB_VarHelpID = -1

'These objects wrap the GridEX controls to give a more convenient
'interface for getting the events we're interested in.
Private WithEvents m_gwCustOrders As GridEXWrapper
Attribute m_gwCustOrders.VB_VarHelpID = -1
Private WithEvents m_gwOrders As GridEXWrapper
Attribute m_gwOrders.VB_VarHelpID = -1
Private WithEvents m_gwItems As GridEXWrapper
Attribute m_gwItems.VB_VarHelpID = -1
Private WithEvents m_gwStatusLine As GridEXWrapper
Attribute m_gwStatusLine.VB_VarHelpID = -1

Private WithEvents m_gwShipments As GridEXWrapper
Attribute m_gwShipments.VB_VarHelpID = -1
Private WithEvents m_gwShipItems As GridEXWrapper
Attribute m_gwShipItems.VB_VarHelpID = -1

Private WithEvents m_gwInvoice As GridEXWrapper
Attribute m_gwInvoice.VB_VarHelpID = -1
Private WithEvents m_gwOSLineItems As GridEXWrapper
Attribute m_gwOSLineItems.VB_VarHelpID = -1
Private WithEvents m_gwRMALine As GridEXWrapper
Attribute m_gwRMALine.VB_VarHelpID = -1

Private m_colWCGridPrefs As Collection

Private m_bLoading As Boolean

Private m_bFindMode As Boolean

Private m_ePreviousTab As TabMainIndexes

Private m_bRecommit As Boolean

Private m_bNewItem As Boolean
Private m_bCanCancel As Boolean
Private m_bFindOrderFlag As Boolean
Private m_bWareHouseLoading As Boolean
Private m_bPromptedForShipComplete As Boolean

Private m_sDefaultShipMeth As String

Private m_lCustKey As Long
Private m_lDefaultBillAddrKey As Long
Private m_lDefaultShipAddrKey As Long

Private m_bChooseItem As Boolean
Private m_bRestoreDefPrice As Boolean
Private m_bDeleteItem As Boolean

Private m_sLineRemark() As String
Private m_oOSItemList As OSItemList
Private m_lSelectedIndex As Long
Private m_lRMAKey As Long
Private m_oRMALine As RMAList

Private m_sDfltCntctName As String
Private m_sDfltCntctPhone As String
Private m_sDfltCntctPhoneExt As String
Private m_sDfltCntctFax As String
Private m_sDfltCntctFaxExt As String


'***************************************************************************************
'Public Properties
'***************************************************************************************

Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property

Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property


Public Property Get Order() As Order
    Set Order = m_oOrder
End Property


Public Property Get Customer() As Customer
    Set Customer = m_oCustomer
End Property

Public Property Let Customer(oNewValue As Customer)
    Set m_oCustomer = oNewValue
End Property


Public Property Get Items() As Items
    Set Items = m_oItems
End Property

Public Property Let Items(oNewValue As Items)
    Set m_oItems = oNewValue
End Property


Public Property Get BrokenRules() As BrokenRules
    Set BrokenRules = m_oBrokenRules
End Property


Public Property Get RecommitSage() As Boolean
    RecommitSage = m_bRecommit
End Property


Public Property Get StatusCode() As String
    StatusCode = StatusCodeString(m_oOrder.StatusCode)
End Property


Public Property Get FindMode() As Boolean
    FindMode = m_bFindMode
End Property

Public Property Let FindMode(bNewValue As Boolean)
    m_bFindMode = bNewValue
End Property


' ViewMode controls the view on the Line tab
'
' Note: this Property is private to FOrder
'
' Enum ItemView
'    ivList
'    ivComponent
'    ivKit
'    ivGasket
'    ivShelf
'    ivWire
'    ivSpecialOrder
' summarized as ListView and ItemViews

Private Property Get ViewMode() As ItemView
    ViewMode = m_View
End Property

Private Property Let ViewMode(ByVal iView As ItemView)
    m_View = iView
    ChangeViewMode iView
End Property


Private Sub ChangeViewMode(ByVal iView As ItemView)
    Dim bLoading As Boolean
            
    ' clear the reference variables
    Set m_oFinGood = Nothing
    Set m_oBTOKit = Nothing
    Set m_oWarmerWire = Nothing
    Set m_oShelf = Nothing
    Set m_oGasket = Nothing

    m_oBrokenRules.EnableClass ccWWireLength, False
    m_oBrokenRules.EnableClass ccWWireDoorDim, False
    m_oBrokenRules.EnableClass ccWWireType, False
    m_oBrokenRules.EnableClass ccshelf, False
    m_oBrokenRules.EnableClass ccGasket, False
    m_oBrokenRules.EnableClass ccItemSpecialOrder, False
    m_oBrokenRules.EnableClass ccItemSPOBasicInfo, False

    Select Case iView
        Case ivComponent:
            Set m_oFinGood = m_oItems.SelectedItem
            imgType.Picture = imglType64.ListImages(1).Picture
            
            If m_bChooseItem = True Then
                m_oFinGood.Load m_oFinGood.IItem_ItemKey, m_oOrder.WhseKey
            End If
            
            cmdViewCat(0).Enabled = (m_oFinGood.CatPage > 0)
            cmdViewCat(0).caption = "View Catalog (pg " & m_oFinGood.CatPage & ")..."
            
            cboWarehouse(2).ListIndex = cboWarehouse(1).ListIndex
            
            UpdateInventoryInfo
            cmdResearchPO(1).Enabled = True

            If m_oFinGood.BaseClass.ItemInventoryStatus = iisDiscontinued Then
                cmdVendorDetails(2).Enabled = False
            Else
                cmdVendorDetails(2).Enabled = True
            End If
            cmdInvFinder.Enabled = True
        
        Case ivKit:
            Set m_oBTOKit = m_oItems.SelectedItem
            imgType.Picture = imglType64.ListImages(2).Picture

            cmdViewCat(1).Enabled = (m_oBTOKit.CatPage > 0)
            cmdViewCat(1).caption = "View Catalog (pg " & m_oBTOKit.CatPage & ")..."
            
            cboWarehouse(2).ListIndex = cboWarehouse(1).ListIndex
    
            UpdateInventoryInfo
            cmdResearchPO(0).Enabled = True
            
            If m_oBTOKit.BaseClass.ItemInventoryStatus = iisDiscontinued Then
                cmdVendorDetails(1).Enabled = False
            Else
                cmdVendorDetails(1).Enabled = True
            End If
            
            cmdInvFinder.Enabled = True
        
        Case ivGasket
            Set m_oGasket = m_oItems.SelectedItem
            imgType.Picture = imglType64.ListImages(3).Picture

            m_oBrokenRules.EnableClass ccGasket, True
            m_oGasket.LoadMaterialCombo cboGasket
        
        Case ivShelf
            Set m_oShelf = m_oItems.SelectedItem
            imgType.Picture = imglType64.ListImages(4).Picture

            m_oBrokenRules.EnableClass ccshelf, True
            m_oItems.SelectedItem.VendorKey = GetShelfVendKeyFromWhseKey(m_oOrder.WhseKey)
        
        Case ivWire
            Set m_oWarmerWire = m_oItems.SelectedItem
            imgType.Picture = imglType64.ListImages(5).Picture

            bLoading = m_bLoading
            m_bLoading = True
            
            With m_oWarmerWire
                m_oBrokenRules.EnableClass ccWWireType, True
                
                .lWhseKey = m_oOrder.WhseKey

                If .DoorHeight + .DoorWidth = 0 Then
                    m_oBrokenRules.EnableClass ccWWireLength, True
                    m_oBrokenRules.EnableClass ccWWireDoorDim, False
                    lenWireLength.Enabled = (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
                    lenDoorHeight.Enabled = False
                    lenDoorWidth.Enabled = False
                    frmWirePasses.Visible = True
                    frmDoorStyle.Visible = False
                    
                    optDoorStyle(0).value = True
                    optLengthAlgorithm(0).value = True
                    If .IsSinglePass Then
                        optWirePasses(0).value = True
                    Else
                        optWirePasses(1).value = True
                    End If
                Else
                    m_oBrokenRules.EnableClass ccWWireLength, False
                    m_oBrokenRules.EnableClass ccWWireDoorDim, True
                    lenWireLength.Enabled = False
                    lenDoorHeight.Enabled = (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
                    lenDoorWidth.Enabled = (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
                    frmWirePasses.Visible = False
                    frmDoorStyle.Visible = True
                    
                    optWirePasses(0).value = True
                    optLengthAlgorithm(1).value = True
                    If .IsThreeSided Then
                        optDoorStyle(1).value = True
                    Else
                        optDoorStyle(0).value = True
                    End If
                End If
            End With
            m_bLoading = bLoading
            
        Case ivSpecialOrder
            imgType.Picture = imglType64.ListImages(6).Picture
            
            m_oBrokenRules.EnableClass ccItemSpecialOrder, True
            m_oBrokenRules.EnableClass ccItemSPOBasicInfo, True
    
    End Select
    

    frmFindPart.Visible = (iView = ivList)
    frmSpecifyPart.Visible = (iView = ivList)
    frmItemList.Visible = (iView = ivList)
            
    'these controls are shared between views
    cmdItemOK.Visible = Not (iView = ivList)
    cmdItemCancel.Visible = Not (iView = ivList)
    cmdItemDelete.Visible = Not (iView = ivList)
    
    frmBasicInfo.Visible = Not (iView = ivList)
    frmPricing.Visible = Not (iView = ivList)
    
    frmStock.Visible = (iView = ivComponent)
    frmAssembly.Visible = (iView = ivKit)
    frmInventory.Visible = (iView = ivComponent Or iView = ivKit)
    frmGasket.Visible = (iView = ivGasket)
    frmShelf.Visible = (iView = ivShelf)
    frmWire.Visible = (iView = ivWire)
    frmSpecialOrder.Visible = (iView = ivSpecialOrder)
    
    cmdNextGasket.Visible = (iView = ivGasket And m_bNewItem)
    
    If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then
        cmdItemOK.Enabled = True
        cmdItemCancel.Enabled = True
    Else
        If iView = ivComponent Or iView = ivKit Then
            cmdItemOK.Enabled = True
        End If
    End If

    txtCost.Enabled = ((iView = ivSpecialOrder) Or (iView = ivShelf)) And (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
    
    If iView = ivList Then
        txtItemSearch.text = ""

        m_bNewItem = False 'reset this flag on return to list view
        m_oBrokenRules.EnableClass ccPricing, False

        'Update summary pricing
        txtTotalPrice.amount = m_oItems.TotalPrice
        txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)
        
        TryToSetFocus txtItemSearch
        SyncItemList
        gdxItems.Row = -1
    Else
        m_oItems.SelectedItem.Backup
        m_oItems.SelectedItem.BackNegotiatedPrice = m_oItems.SelectedItem.EffectivePrice
        ItemUpdateControls
        txtItemPartNbr.Enabled = (iView = ivSpecialOrder And m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
        txtItemDescr.Enabled = (iView = ivSpecialOrder And m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
        LoadLineItemRemark
        m_oBrokenRules.EnableClass ccPricing, True
        
        Select Case iView
            Case ivGasket
                If m_oGasket.IsMagnetic Then
                    TryToSetFocus optGasketType(0)
                Else
                    TryToSetFocus optGasketType(1)
                End If
            Case ivWire
                TryToSetFocus lenWireLength
            Case ivShelf
                TryToSetFocus cboFrame
            Case ivSpecialOrder
                TryToSetFocus txtItemPartNbr
                cmdVendorDetails(3).Enabled = True
            Case Else
                TryToSetFocus cmdItemOK
        End Select
    End If

    'Force an evaluation of all broken rules
    m_oBrokenRules.Validate
        
    MDIMain.UpdateToolbarStatus
    
    'Do this MM RemarkViewer update after everything else
    
    If iView <> ivList And Not m_oItems.SelectedItem Is Nothing Then
        rvOrderLine(0).RemarkContext = m_oItems.SelectedItem.RemarkContext
    End If
    
    Select Case iView
        Case ivComponent
            rvFinGood.OwnerID = ""
            rvFinGood.OwnerID = m_oItems.SelectedItem.ItemID
        Case ivKit
            rvAssembly.OwnerID = ""
            rvAssembly.OwnerID = m_oItems.SelectedItem.ItemID
    End Select
    
End Sub


'***************************************************************************************
'Public Methods
'***************************************************************************************

Public Sub SetCaption(ByRef i_sTitle As String)
    Me.caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub


Public Sub Repaint()
    tabMain.Refresh
End Sub


Public Sub Init()
    Form_Load
End Sub


Public Sub DoShowHelp()
    Select Case tabMain.SelectedTab.Index
        Case tmiExistingCustomer:   ShowHelp ("FindCust")
        Case tmiExistingOrder:      ShowHelp ("FindOrder")
        Case tmiCustomer:           ShowHelp ("Customer")
        Case tmiOrder:              ShowHelp ("Order")
        Case tmiLines
            Select Case ViewMode
                Case ivList:            ShowHelp ("Lines")
                Case ivComponent:       ShowHelp ("FinGood")
                Case ivKit:             ShowHelp ("BTOKit")
                Case ivGasket:          ShowHelp ("Gasket")
                Case ivShelf:           ShowHelp ("Shelf")
                Case ivWire:            ShowHelp ("WWire")
                Case ivSpecialOrder:    ShowHelp ("SPO")
                Case Else:              ShowHelp ("TOC")
            End Select
        Case tmiOrderStatus:        ShowHelp ("OrderStatus")
        Case tmiRmaLines:           ShowHelp ("RMALines")
        Case Else:                  ShowHelp ("TOC")
    End Select
End Sub


Public Function CancelButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    Dim vbValue As VbMsgBoxResult
    Dim lOPKey As Long

    If Not m_bCanCancel Then Exit Function
    
    If Not FindMode And ViewMode <> ivList Then Exit Function
    
    CancelButton = True
    
    If i_bDoIt Then
        ForceLostFocus
        SetWaitCursor True
        If FindMode = True Then
            lOPKey = m_oOrder.OPKey
            Set m_oOrder = New Order
            m_oOrder.Load lOPKey
            Set m_oCustomer = m_oOrder.Customer
            TransitionTabs False
        Else
            If ViewMode <> ivList Then
                cmdItemCancel_Click 'exit detail view, if necessary
            End If
            
            If m_oOrder.IsDirty And CanSaveOrder Then
                vbValue = msg("Do you want to save the changes you made?", vbExclamation + vbYesNoCancel, Me.caption)
                If vbValue = vbCancel Then
                    SetWaitCursor False
                    Exit Function
                ElseIf vbValue = vbYes Then
                    If Not SaveButton(True) Then
                        SetWaitCursor False
                        Exit Function
                    End If
                Else
                    m_oOrder.Restore
                    Set m_oCustomer = m_oOrder.Customer
                    Set m_oItems = m_oOrder.Items
                End If
            End If
            If m_lCustKey = 0 Then  'FillSelectCustTab m_lCustKey
                cmdNewSearch_Click  'reset Existing Customer tab
            Else
                FillSelectCustTab m_lCustKey, False
            End If
'            cmdFindOrders2_Click 'refresh existing orders tab
            FindOrdersByCriteria
            TransitionTabs True
        End If
        m_bRecommit = False
        SetWaitCursor False
    End If
    Exit Function

ErrorHandler:
    m_bRecommit = False
    ErrorUI.DisplayWarning "Cancel Failed"
End Function


'Called by MDIMain's DoExit function

Public Function ExitCheck() As Boolean
    Dim vbValue As VbMsgBoxResult
    
    ExitCheck = True
    If g_bConfirmExit And Not g_bExitNow Then
        If m_oOrder.IsDirty And CanSaveOrder Then
            vbValue = msg("Did you make any changes you'd like to save before exiting?", vbExclamation + vbYesNoCancel, "Exit " & Me.caption & "?")
            If vbValue = vbCancel Then
                ExitCheck = False
            ElseIf vbValue = vbYes Then
                If Not SaveButton(True) Then
                    ExitCheck = False
                End If
            End If
        End If
    End If

End Function


Private Sub cboFindStatus_Click()
    If cboFindStatus.text = "Open RMA" Then
        txtFindCust.text = ""
        txtFindText.text = ""
        txtFindCust.Enabled = False
        txtFindText.Enabled = False
        txtFindCust.BackColor = &H8000000F
        txtFindText.BackColor = &H8000000F
    Else
        txtFindCust.Enabled = True
        txtFindText.Enabled = True
        txtFindCust.BackColor = &H80000005
        txtFindText.BackColor = &H80000005
    End If
End Sub





'************************************************************************************
'Form Events and subroutines
'************************************************************************************

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Load()
    SetCaption "[New Order]"
    tabMain.Tabs(tmiExistingCustomer).Selected = True
    gdxItems.ItemCount = 0
    
    Set m_oOrder = New Order
    Set m_oCustomer = m_oOrder.Customer
    Set m_oItems = m_oOrder.Items
    
    '11/10/04 LR
    'Note: since no userkey or id is specified this will load combo and select index=0
    '(first User in list). This calls cboCSR_click which writes selected User to Order.UserKey.
    
    'User.SetUpUsers cboCSR, g_rstUsers
    User.LoadActiveCSRs cboCSR
    
    'User.SetUpUsers cboFindCSR, g_rstUsers
    User.LoadActiveCSRs cboFindCSR
    cboFindCSR.AddItem "<Any>"
    cboFindCSR.AddItem "<MPK>"
    cboFindCSR.AddItem "<SEA>"
    cboFindCSR.AddItem "<STL>"
    
    Set m_oBrokenRules = New BrokenRules
    m_oBrokenRules.Form = Me
    LoadValidationRules

    'assign images
    LoadImageList imglStatus16, gdxCustOrders
    LoadImageList imglStatus16, gdxOrders
    LoadImageList imglTypeAndStatus32, gdxItems
    LoadImageList imglRemarks, gdxOSLineItems
    picItem.Picture = imglType64.ListImages(2).Picture
    cmdSpecifyItem(btnGasket).Picture = imglTypeAndStatus32.ListImages(3).Picture
    cmdSpecifyItem(btnShelf).Picture = imglTypeAndStatus32.ListImages(4).Picture
    cmdSpecifyItem(btnWire).Picture = imglTypeAndStatus32.ListImages(5).Picture
    cmdSpecifyItem(btnSPO).Picture = imglTypeAndStatus32.ListImages(6).Picture
    
    'Initialize grid wrappers
    Set m_gwOrders = New GridEXWrapper
    m_gwOrders.Grid = gdxOrders
    Set m_gwCustOrders = New GridEXWrapper
    m_gwCustOrders.Grid = gdxCustOrders
    
    Set m_gwItems = New GridEXWrapper
    m_gwItems.Grid = gdxItems
    Set m_gwOSLineItems = New GridEXWrapper
    m_gwOSLineItems.Grid = gdxOSLineItems
    Set m_gwRMALine = New GridEXWrapper
    m_gwRMALine.Grid = gdxRMALine
    Set m_gwStatusLine = New GridEXWrapper
    m_gwStatusLine.Grid = gdxOSLine
    
    Set m_gwShipments = New GridEXWrapper
    m_gwShipments.Grid = gdxOSShipments
    Set m_gwShipItems = New GridEXWrapper
    m_gwShipItems.Grid = gdxOSShipItems
     
    Set m_gwInvoice = New GridEXWrapper
    m_gwInvoice.Grid = gdxOSInvoice
    
    m_gwOrders.InitGridLayout GetUserKey, g_OrderGridRev
    m_gwCustOrders.InitGridLayout GetUserKey, g_CustOrderGridRev

    If InStr(1, GetUserName, "WillCall", vbTextCompare) > 0 Then
        LoadWCGridPrefs
    End If

    SetSearchDefaults

    SetComboByText cboTimeInterval, "30 days"

    'We have three copies of this combo because it appears on three frames
    Helpers.LoadCombo cboMake(1), g_rstMakes, "MakeText", "MakeID"
    Helpers.LoadCombo cboMake(2), g_rstMakes, "MakeText", "MakeID"
    Helpers.LoadCombo cboMake(3), g_rstMakes, "MakeText", "MakeID"
    
    Helpers.LoadCombo cboVendor, g_rstVendors, "VendName", "VendKey", , True

    txtPrice.ATMMode = g_bATMMode
    txtCost.ATMMode = g_bATMMode
        
    Set m_oShelf = New ItemShelf 'we just need one for a second to init these combos
    m_oShelf.LoadFinishCombo cboFinish
    m_oShelf.LoadFrameCombo cboFrame
    Set m_oShelf = Nothing

    With icbItemStatus
        Dim i As Long
        Set .ImageList = imglStatus16
        For i = 1 To imglStatus16.ListImages.Count
            .ComboItems.Add i, , imglStatus16.ListImages(i).Tag, i, i
        Next
    End With

    'support a Click event handler on cboSearchType
    m_bLoading = True
    cboSearchType.ListIndex = 0
    m_bLoading = False

    cmdUpdateRMALine.Enabled = False

    chkShowOrdersForShipAddr.Visible = False
    
    lblOrderStatus.Visible = False
    cboOrderStatus.Visible = False
    lblFilterByPart.Visible = False
    txtFilterByPart.Visible = False
    cmdFilterByPart.Visible = False
    
    cmdSelectCustomer(1).Visible = Not g_bWillCallUser
    cmdSelectCustomer(3).Visible = Not g_bWillCallUser
    lblExplain(0).Visible = Not g_bWillCallUser
    lblExplain(2).Visible = Not g_bWillCallUser
    
    'disable Load Order button on Select Order grid at start up because grid will be empty
    cmdLoadOrder(1).Enabled = False
    
    m_bFindMode = True

End Sub

'*********************************************************************************************
' Form Resize Logic
'*********************************************************************************************

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If FindMode Then
        ResizeFindMode
    Else
        ResizeOrderMode
    End If
End Sub


' Resizes Select Customer and Select Order tab for maximizing and minimizing

Private Sub ResizeFindMode()

    tabMain.width = Me.width - 255
'    tabMain.Height = Me.Height - 550
     tabMain.Height = Me.Height - 760  'allow room for status bar
     
    tpFind.width = tabMain.width
    tpFind.Height = tabMain.Height - 390
    tpFindOrder.width = tpFind.width
    tpFindOrder.Height = tpFind.Height
    
    'Top is wherever it's placed
    gdxCustOrders.width = tabMain.width - 250
    gdxCustOrders.Height = tabMain.Height - 2640
    
    gdxOrders.width = tabMain.width - 250
    gdxOrders.Height = tabMain.Height - 2340

    chkShowOrdersForShipAddr.Top = lblCustAddress.Top + lblCustAddress.Height + 60

    cmdLoadOrder(0).Top = gdxCustOrders.Top + gdxCustOrders.Height + 75
    cmdNewOrder.Top = cmdLoadOrder(0).Top + 360

    cmdContactMgr(0).Top = cmdLoadOrder(0).Top
    cmdNewSearch.Top = cmdNewOrder.Top
    lblOrderStatus.Top = cmdNewOrder.Top
    cboOrderStatus.Top = cmdNewOrder.Top
    lblFilterByPart.Top = cmdNewOrder.Top
    txtFilterByPart.Top = cmdNewOrder.Top
    cmdFilterByPart.Top = cmdNewOrder.Top

    cmdLoadOrder(1).Top = cmdNewOrder.Top - 60
    
    DoEvents
    MDIMain.DoRefresh
End Sub


'Called by
'   Form_Resize
'   TransitionTabs
'
'Resizes Customer, Order and Line tabs for maximizing and minimizing

Public Sub ResizeOrderMode()

    tabMain.width = Me.width - 255
    tabMain.Height = Me.Height - 780
    
    tpCustomer.width = tabMain.width
    tpCustomer.Height = tabMain.Height - 390
    tpOrder.width = tpCustomer.width
    tpOrder.Height = tpCustomer.Height
    tpItems.width = tpCustomer.width
    tpItems.Height = tpCustomer.Height

    frmItemList.width = tabMain.width - 250
    frmItemList.Height = tabMain.Height - 1920
    gdxItems.width = frmItemList.width
    gdxItems.Height = tabMain.Height - 2640
    cmdAuthorizeAll.Top = frmItemList.Height - cmdAuthorizeAll.Height - 120
    cmdUnAuthorize.Top = cmdAuthorizeAll.Top
    Label1(69).Top = cmdAuthorizeAll.Top
    Label1(62).Top = cmdAuthorizeAll.Top
    txtTotalTax.Top = cmdAuthorizeAll.Top
    txtTotalPrice.Top = cmdAuthorizeAll.Top
    Label1(2).Top = cmdAuthorizeAll.Top
    cboWarehouse(1).Top = cmdAuthorizeAll.Top
    
    If tabMain.Tabs(tmiOrderStatus).Visible = True Then ResizeOS
    If tabMain.Tabs(tmiRmaLines).Visible = True Then ResizeRMA
    If tabMain.Tabs(tmiOrderHistory).Visible = True Then ResizeOH
    
    DoEvents
    MDIMain.DoRefresh
End Sub


'Resizes Order Status tab for maximizing and minimizing

Private Sub ResizeOS()
    Dim lIndex As Integer
    tpStatus.width = tpCustomer.width
    tpStatus.Height = tpCustomer.Height
    
    SSAOrderDetails.width = tabMain.width - 250
    SSAOrderDetails.Height = tabMain.Height - 2280
    gdxOSLineItems.width = SSAOrderDetails.width - 250
    gdxOSInvoice.width = gdxOSLineItems.width
    gdxOSLine.width = gdxOSLineItems.width
    gdxOSShipItems.width = gdxOSLineItems.width
    gdxOSShipments.width = gdxOSLineItems.width
    gdxOSInvoiceItem.width = gdxOSLineItems.width
    
    gdxOSShipments.Height = ((SSAOrderDetails.Height - 1000) * 0.35)
    gdxOSShipItems.Height = ((SSAOrderDetails.Height - 1000) * 0.65)
    gdxOSShipItems.Top = gdxOSShipments.Height + 490
    lblOSLGridCaption.Item(1).Top = gdxOSShipments.Height + 240
    
    gdxOSInvoice.Height = ((SSAOrderDetails.Height - 1000) * 0.35)
    gdxOSInvoiceItem.Height = ((SSAOrderDetails.Height - 1000) * 0.65)
    gdxOSInvoiceItem.Top = gdxOSInvoice.Height + 490
    lblOSLGridCaption.Item(2).Top = gdxOSInvoice.Height + 240
    
    gdxOSLine.Height = ((SSAOrderDetails.Height - 1000) * 0.5)
    gdxOSLineItems.Height = ((SSAOrderDetails.Height - 1000) * 0.5)
    gdxOSLineItems.Top = gdxOSLine.Height + 490
    lblOSLGridCaption.Item(0).Top = gdxOSLine.Height + 240
End Sub


'Resizes RMA Line tab for maximizing and minimizing

Private Sub ResizeRMA()
    tpRMA.width = tpCustomer.width
    tpRMA.Height = tpCustomer.Height
    
    tabRMADetail.width = tabMain.width - 250
    tabRMADetail.Height = tabMain.Height - 720
    
    gdxRMALine.width = tabRMADetail.width - 250
    gdxRMALineStatus.width = gdxRMALine.width

    gdxRMALine.Height = (tabRMADetail.Height - 2025) / 2
    gdxRMALineStatus.Height = (tabRMADetail.Height - 2025) / 2

    gdxRMALineStatus.Top = tabRMADetail.Height - gdxRMALineStatus.Height - 550
    lblRMALineStatus.Top = gdxRMALineStatus.Top - lblRMALineStatus.Height - 120
    gdxRMALineStatus.Visible = True

End Sub


'resize the Order History tab

Private Sub ResizeOH()
    gdxOrderEvent.width = tabMain.width - 300
    gdxOrderEvent.Height = tabMain.Height - 1000
    cmdPrint.Top = gdxOrderEvent.Top + gdxOrderEvent.Height + 60
    cmdPrint.Left = tabMain.width - 300 - cmdPrint.width
End Sub

'*********************************************************************************************

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "Form_Unload"
    Dim vbValue As VbMsgBoxResult
    
    If g_bConfirmExit And Not g_bExitNow Then
        If m_oOrder.IsDirty And CanSaveOrder Then
            vbValue = msg("Did you make any changes you'd like to save before closing?", vbExclamation + vbYesNoCancel, "Close " & Me.caption & "?")
            If vbValue = vbCancel Then
                Cancel = 1
                Exit Sub
            ElseIf vbValue = vbYes Then
                If SaveButton(True) Then Exit Sub
            End If
        End If
    End If

    m_oBrokenRules.Destroy
    Set m_oBrokenRules = Nothing
    
    '4/7/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwCustOrders = Nothing
    Set m_gwOrders = Nothing
    Set m_gwItems = Nothing
    Set m_gwStatusLine = Nothing
    Set m_gwShipments = Nothing
    Set m_gwShipItems = Nothing
    Set m_gwInvoice = Nothing
    Set m_gwOSLineItems = Nothing
    Set m_gwRMALine = Nothing

    Set m_colWCGridPrefs = Nothing
    
    With MDIMain
        .UnloadTool Me.WindowID
        .UpdateToolbarStatus
    End With
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyW And Shift = vbCtrlMask Then
        'cmdWalkup_Click
        Call cmdSelectCustomer_Click(2)
    Else
        MDIMain.GlobalKeyDownProcessing KeyCode, Shift
    End If
End Sub


Private Sub LoadValidationRules()
    Dim oCtlWrapper As ControlWrapper

    With m_oBrokenRules
        'Customer search fields
        Set oCtlWrapper = .AddControl(txtCustSearch, k_sCustNameOrID, True, False)
        oCtlWrapper.AddRuleRequired "", ccCustomer, True, "Enter a value to search for a customer."
        .EnableClass ccCustomer, True

        'Order Tab
        Set oCtlWrapper = .AddControl(txtPO, "Purchase Order number", True)
        oCtlWrapper.AddRuleRequired "", ccpurchaseorder, False, "This customer requires a PO on all orders."

'***CONTACT
'This is intended to guarantee a receiving contact, not an ordering contact
'added 8/17/05 LR
'        Set oCtlWrapper = .AddControl(txtOrderedBy, "Ordered By", True)
'        oCtlWrapper.AddRuleAMDelivery ccorderedby

        'SMR Intl 01/13/2006 - 2nd param becomes part of the broken rule msg.
        Set oCtlWrapper = .AddControl(txtShipToContact(0), "Ship To Contact Name", True, True)
        oCtlWrapper.AddRuleRequired "", ccShipToData
        'SMR Intl 01/16/2006
        Set oCtlWrapper = .AddControl(txtShipToContact(1), "Ship To Contact Phone Number", True, True)
        oCtlWrapper.AddRuleRequired "", ccShipToData
        
        'Items Tab: List View
        Set oCtlWrapper = .AddControl(txtItemSearch, k_sPartNbr, True)
        Set m_oValidateItem = oCtlWrapper.AddRuleManual(True, "No item matches this criteria", ccItem)

        'Items Tab: Item Detail Views
        Set oCtlWrapper = .AddControl(txtQtyOrdered, "Quantity", True)
        oCtlWrapper.AddRuleNumeric 1, , ccPricing

        'Items Tab: SPO View
        Set oCtlWrapper = .AddControl(txtItemPartNbr, "Part Number", True)
        oCtlWrapper.AddRuleRequired "", ccItemSPOBasicInfo, , "You must enter a part number and/or description to add a SPO to an order."
    
        Set oCtlWrapper = .AddControl(txtItemDescr, "Part Description", True)
        oCtlWrapper.AddRuleRequired "", ccItemSPOBasicInfo, , "You must enter a part number and/or description to add a SPO to an order."
        Set m_oValidateDescr = oCtlWrapper.AddRuleManual(True, "You must shorten this description", ccItemSpecialOrder)
        
        'JJC BUGBUG: These rules should be enabled if the item status is past Research
        'but currently, the Item logic won't allow the status to be moved past Research
        'until these fields are entered, so for the time being those rules are commented out.
        'Set oCtlWrapper = .AddControl(txtCost, "Item Cost", False)
        'oCtlWrapper.AddRuleNumeric 0.01, , ccItemSpecialOrder, , "Item cost must be at least $0.01"
    
         'Set oCtlWrapper = .AddControl(cboVendor, "Vendor", False)
        'oCtlWrapper.AddRuleRequired "", ccItemSpecialOrder
       
        'Items Tab: Warmer Wire View
        Set oCtlWrapper = .AddControl(cboWires, "Available Wires", True)
        oCtlWrapper.AddRuleRequired "", ccWWireType
        
        Set oCtlWrapper = .AddControl(lenWireLength, "Wire Length", True)
        oCtlWrapper.AddRuleInchWorm ccWWireLength
        
        Set oCtlWrapper = .AddControl(lenDoorHeight, "Door Height", True)
        oCtlWrapper.AddRuleInchWorm ccWWireDoorDim
        
        Set oCtlWrapper = .AddControl(lenDoorWidth, "Door Width", True)
        oCtlWrapper.AddRuleInchWorm ccWWireDoorDim
    
        'Items Tab: Wire Shelf View
        Set oCtlWrapper = .AddControl(lenShelfWidth, "Shelf Width", True)
        oCtlWrapper.AddRuleInchWorm ccshelf
    
        Set oCtlWrapper = .AddControl(lenShelfDepth, "Shelf Depth", True)
        oCtlWrapper.AddRuleInchWorm ccshelf
    
        'Items Tab: Gasket View
        Set oCtlWrapper = .AddControl(lenGasket(0), "Gasket Height", True)
        oCtlWrapper.AddRuleInchWorm ccGasket
        
        Set oCtlWrapper = .AddControl(lenGasket(1), "Gasket Width", True)
        oCtlWrapper.AddRuleInchWorm ccGasket
        
    
        Set oCtlWrapper = .AddControl(cboGasket, "Gasket Material", True)
        oCtlWrapper.AddRuleRequired "", ccGasket
    End With
End Sub


'***********************************************************************************
'Control Arrays
'***********************************************************************************

Private Sub cmdPrint_Click()
    gdxOrderEvent.PrintGrid True
End Sub

'Right-click menu to support UPS tracking

Private Sub mnuTrackUPS_Click()
    Dim sURL As String
    Dim sTrackingNumber As String
    
    'sTrackingNumber = Trim(m_gwShipments.value("ShipTrackNo"))
    sTrackingNumber = Trim(m_gwShipItems.value("ShipTrackNo"))
    
    'If Len(sTrackingNumber) = 0 Then
    If UCase(Left$(sTrackingNumber, 2)) <> "1Z" Then
        MsgBox "This shipment does not have a valid UPS tracking number."
        Exit Sub
    End If
    
    sURL = "http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=5&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=" & sTrackingNumber & "&InquiryNumber2=&InquiryNumber3=&InquiryNumber4=&InquiryNumber5=&track.x=33&track.y=13"
    
    Dim R As Long
    R = ShellExecute(0, "open", sURL, 0, 0, 1)
    
End Sub

                            

'************************************************************************************
'
'   The Tab Control's Event Handlers
'
'************************************************************************************

Private Sub tabMain_BeforeTabClick(ByVal NewTab As ActiveTabs.SSTab, ByVal Cancel As ActiveTabs.SSReturnBoolean)
    Static bInitialized As Boolean
    
    Select Case NewTab.Index
        Case tmiExistingOrder
            'Do nothing, all the action is in cmdFindOrders
            
        Case tmiCustomer, tmiOrder, tmiOrderStatus, tmiRmaLines
        
            'What does tmiLines ViewMode have to do with anything?
            
            If ViewMode <> ivList Then
    
            '***!!! 10/26/04 LR This is certainly an arbitrary way to determine if the order is open in Edit mode.
            '       How is this determined elsewhere?
            '       Let's standardize this within the limitations of the "design".
                ' PRN#178 AVH - HACK HACK
                ' If this order is in edit mode, cmdSpecifyItem(btnGasket) would be enabled
                
                If cmdSpecifyItem(btnGasket).Enabled = True Then
                ' PRN#178 HACK HACK
                    
                    If m_oBrokenRules.MaskedCount(k_lItemControlMask) > 0 Then
                        If vbOK = msg("Discard changes to this item?", vbOKCancel, "Discard changes?") Then
                            cmdItemCancel_Click
                        Else
                            Cancel = True
                        End If
                    Else
                        Select Case msg("Save changes to this item?", vbYesNoCancel, "Save changes?")
                            Case vbYes
                                cmdItemOK_Click
                            Case vbNo
                                cmdItemCancel_Click
                            Case Else
                                Cancel = True
                        End Select
                    End If
                    
                End If ' PRN#178 HACK HACK
            End If
                        
        Case tmiLines
        
            gdxItems.Refetch
            If m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit Then
                cmdSpecifyItem(btnGasket).Enabled = (Not m_oOrder.IsDropShip)
                cmdSpecifyItem(btnShelf).Enabled = ShelfEnabled
                cmdSpecifyItem(btnWire).Enabled = (Not m_oOrder.IsDropShip)
            End If
            txtTotalPrice.amount = m_oItems.TotalPrice
            txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)

    End Select
End Sub


Private Function ShelfEnabled() As Boolean
    If Not m_oOrder.IsDropShip Then
        ShelfEnabled = True
    Else
        If m_oOrder.whseid = "STL" Then
            ShelfEnabled = True
        Else
            ShelfEnabled = False
        End If
    End If
End Function


Private Sub tabMain_TabClick(ByVal NewTab As ActiveTabs.SSTab)
    Select Case NewTab.Index
        
        Case tmiExistingCustomer
            If frmCustSearch.Visible Then
                TryToSetFocus txtCustSearch
            End If
            
        Case tmiExistingOrder
            TryToSetFocus txtFindOrder
            
        Case tmiCustomer
           
            Dim sAlerts As String
            
            If m_oCustomer.Hold Then
                sAlerts = "Is On Hold" & vbCrLf
            End If
     
            If WhseHasCatalogs And m_oCustomer.QueryForCatalog Then
                sAlerts = sAlerts & "Has not been asked about the Catalog"
            End If
    
            lblCustHold.caption = sAlerts
            If Len(sAlerts) > 0 Then
                picCustHold.Visible = True
            Else
                picCustHold.Visible = False
            End If
    
            TryToSetFocus rvCustomer
            
        Case tmiOrderHistory
            GetOrderHistory
        
        Case tmiOrder
    
            LoadContactCombo
    
            'if the customer's default terms are CrCard and
            'the order's terms are CrCard and
            'there's no CrCard Key assigned to the order then
    
            If m_oCustomer.BillAddr.DefaultPmtTerms.ID = "CrCard" And _
                m_oOrder.PmtTerms.ID = "CrCard" And _
                m_oOrder.CreditCard Is Nothing Then
                
                Dim oFrm As FCreditCardEditor
                Set oFrm = New FCreditCardEditor
                
                If oFrm.Init(m_oCustomer, Nothing, True, m_oOrder) = vbCancel Then
                    'if the CSR didn't enter a credit card
                    'restore terms to Customer default
                    SetComboByText cboTerms, m_oOrder.Customer.BillAddr.DefaultPmtTerms.ID
                    'hide the info frame, but leave the edit button visible and enabled
                    'order can't be committed (though there's no good customer feedback)
                    frmCreditCard.Visible = False
                    cmdCCEdit.Visible = True
                Else
                    m_oOrder.CreditCard = oFrm.SelCC
                    UpdateCCDisplay
                End If
                
                Unload oFrm
                Set oFrm = Nothing
            End If
            
        Case tmiLines
            TryToSetFocus txtItemSearch
            
        Case tmiOrderStatus
            'automatically point to the first row of upper grid and populate the lower grid accordingly
            With SSAOrderDetails.SelectedTab
                If .Index = tosLineItem Then
                    gdxOSLine.Row = 1
                    UpdateOSLineItem
                ElseIf .Index = tosShipment Then
                    gdxOSShipments.Row = 1
                    UpdateShipmentItem
                ElseIf .Index = tosInvoice Then
                    gdxOSInvoice.Row = 1
                    UpdateInvoiceItem
                End If
            End With
            
        Case tmiRmaLines
            If Not cmdAddMoreItem.Enabled Then
                cmdAddMoreItem.Enabled = True
                cmdUpdateRMALine.Enabled = False
                cmdRMARefresh.Enabled = True
                cmdRMAVendor.Enabled = True
            End If
    End Select

    With m_oBrokenRules
        .EnableClass ccCustomer, (NewTab.Index = tmiExistingCustomer)
        .Validate
    End With
End Sub


'************************************************************************************
'   Select Customer (Create Order) Tab
'************************************************************************************

'This is a control array (we've run out of control name resources)

Private Sub cmdSelectCustomer_Click(Index As Integer)
    Dim bLoading As Boolean

    Select Case Index
        'Select a customer with an existing account.
        'Find a customer account and load a list of the customer's past orders
        'This subroutine loads customer orders if finding customer succeeds.
        'Otherwise, set focus to cust search textbox for new search
       Case 0
            If Len(Trim$(txtCustSearch.text)) > 0 Then
                SetWaitCursor True
                Set m_oCustomer = New Customer
                m_lCustKey = Search.FindCustomer(txtCustSearch.text, cboSearchType.ListIndex, m_oCustomer)
                If m_lCustKey = 0 Then
                    TryToSetFocus txtCustSearch
                Else
                    FillSelectCustTab m_lCustKey
                    
                    If m_oCustomer.ShipAddr.CountryID <> "USA" Then Call DisplayIntrlCaution
                End If
                
                'If g_bXmasGifts Then
                If g_QueryForGift Then
                    If m_oCustomer.QueryForGift Then
                        msg "You'll be prompted about including a Christmas gift when committing orders for this customer.", vbExclamation, "Alert"
                    End If
                End If
    
                SetWaitCursor False
            Else
                'get rid of any spaces and restore brokenrule
                txtCustSearch.text = vbNullString
            End If
            
        'Create an order for a customer that will open an account with this order.
        Case 1
            With m_oCustomer
                .Clear
                .IsTemp = True
            End With
            
            m_oOrder.Create
            m_oOrder.Customer = m_oCustomer
            
            bLoading = m_bLoading
            m_bLoading = True
            
            txtCustID.text = ""
            txtCustName.text = ""
            
            TransitionTabs False
            
            m_bLoading = bLoading

        'Walkup customer
        Case 2
            m_oCustomer.InitWalkup CreateMISC_CustID
            
            m_oOrder.Create
            m_oOrder.Customer = m_oCustomer
            
            m_oOrder.ShipMethKey = GetWillCallShipMethodKey
            m_oOrder.IsWalkup = True
            m_oOrder.SalesTax.Init m_oCustomer
            If m_oOrder.IsWillCall Then
                m_oOrder.SalesTax.WillCallTaxOverride m_oOrder.whseid
            End If
        
            bLoading = m_bLoading
            m_bLoading = True
            
            TransitionTabs False
            
            m_bLoading = bLoading

        'Miscellaneous customer
        Case 3

            m_oCustomer.InitMiscCustomer CreateMISC_CustID
                    
            m_oOrder.Create
            m_oOrder.Customer = m_oCustomer

            'NOTE: Contact ownerkey = 0, Contact ownrtype = 0
            m_oOrder.isMisc = True
            m_oOrder.SalesTax.Init m_oCustomer
            If m_oOrder.IsWillCall Then
                m_oOrder.SalesTax.WillCallTaxOverride m_oOrder.whseid
            End If
            
            bLoading = m_bLoading
            m_bLoading = True
            
            TransitionTabs False
            
            m_bLoading = bLoading
    End Select
    
End Sub


' index 0 is on the CreateOrder tab
' index 1 is on the Order's Customer tab

Private Sub cmdContactMgr_Click(Index As Integer)

    If Not m_oCustomer.Contacts Is Nothing Then
        m_oCustomer.Contacts.Edit GetUserName

        'From Customer tab.
        'In case the order contact was modified.
        If Index = 1 Then
            UpdateShipContactInfo
        End If
    End If
    
End Sub


'this event handler returns focus to txtCustSearch after selecting a new Search type
'Added Loading flag logic to Form_Load because the SetFocus method throws an error during Form_Load
Private Sub cboSearchType_Click()
    If m_bLoading Then Exit Sub
    txtCustSearch.SetFocus
End Sub


Private Sub txtCustSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtCustSearch.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            Call cmdSelectCustomer_Click(0)
        End If
    End If
End Sub


'Always show upper case for input characters

Private Sub txtCustSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub


'The search type in the combo box changes according to text in the textbox

Private Sub txtCustSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    'this event fires *after* updating the control
    cboSearchType.ListIndex = GetSearchType(txtCustSearch.text)
    m_oBrokenRules.Validate txtCustSearch
End Sub


'This event lets the grid only show customer orders from the selected shipping address or not

Private Sub chkShowOrdersForShipAddr_Click()
    If m_bLoading Then Exit Sub
    m_bLoading = True
    If chkShowOrdersForShipAddr.value = vbUnchecked Then
        ShowOrderAddress vbNullString
    End If
    LoadCustomerOrders (chkShowOrdersForShipAddr.value = vbChecked), cboOrderStatus.text
    m_bLoading = False
End Sub


Private Sub cboOrderStatus_Click()
    If m_bLoading Then Exit Sub
    m_bLoading = True
    LoadCustomerOrders (chkShowOrdersForShipAddr.value = vbChecked), cboOrderStatus.text
    m_bLoading = False
End Sub


Private Sub cmdFilterByPart_Click()
    If cmdFilterByPart.caption = "Filter" Then
        If Len(Trim$(txtFilterByPart.text)) > 0 Then
            cmdFilterByPart.caption = "Reset"
            txtFilterByPart.Enabled = False
            'apply the filter
            LoadCustomerOrders (chkShowOrdersForShipAddr.value = vbChecked), cboOrderStatus.text, Trim$(txtFilterByPart)
        End If
    Else
        txtFilterByPart.text = ""
        txtFilterByPart.Enabled = True
        cmdFilterByPart.caption = "Filter"
        'remove filter
        LoadCustomerOrders (chkShowOrdersForShipAddr.value = vbChecked), cboOrderStatus.text
    End If
    
End Sub


'Create a new order for a Customer with an Account

Private Sub cmdNewOrder_Click()
    
    With m_oCustomer
        .Load .Key     'reload default customer info
        .ShipAddr.Load m_lDefaultShipAddrKey
        .BillAddr.Load m_lDefaultBillAddrKey
    End With
    
    m_oOrder.Create
    m_oOrder.Customer = m_oCustomer
    
    If g_bWillCallUser Then m_oOrder.ShipMethKey = GetWillCallShipMethodKey
    
    m_oOrder.SalesTax.Init m_oCustomer
    
    If m_oOrder.IsWillCall Then
        m_oOrder.SalesTax.WillCallTaxOverride m_oOrder.whseid
    End If
    
    m_bLoading = True

    m_oOrder.PricePackList = SetPPL(m_oCustomer)
    chkPricePackList.value = IIf(m_oOrder.PricePackList, vbChecked, vbUnchecked)
   
    TransitionTabs False
    
    LoadUPSCtrl m_oCustomer.Key
    
    m_bLoading = False
End Sub

    
Private Function SetPPL(oCustomer As Customer)
    With oCustomer
        If .PricePackList = True And .ShipAddr.AddrType = Default Then
            SetPPL = True
        ElseIf .PricePackList = True And .ShipAddr.AddrType = CSA Then
                If MsgBox("You've selected a Common Shipping Address and the customer's " & vbCrLf & _
                            "preference is to show prices on their packing slip. " & vbCrLf & vbCrLf & _
                            "Select 'Yes' to show prices or 'No' to not show prices.", _
                            vbYesNo, "Price Pack List") = vbYes Then
                SetPPL = True
            Else
                SetPPL = False
            End If
        Else
            SetPPL = False
        End If
    End With
End Function

'Click new search button to return to Select Customer search mode

Private Sub cmdNewSearch_Click()
    Set m_oCustomer = New Customer
    
    m_bLoading = True
    SetCaption "OrderPad"
    ClearCustomerSearch
    m_bLoading = False

End Sub


Private Sub ShowOrderAddress(sAddr As String)
    If Len(sAddr) = 0 Then
        lblCustInfo(1).Visible = False
        lblOrderAddress.caption = ""
        lblOrderAddress.Visible = False
    Else
        lblCustInfo(1).Visible = True
        lblOrderAddress.caption = sAddr
        lblOrderAddress.Visible = True
    End If
End Sub


'After user changes selection in the grid, update the header address accordingly.

Private Sub gdxCustOrders_SelectionChange()
    Dim lShipAddrKey As Long

    'Exit if the grid is empty
    If m_gwCustOrders.value("OPKey") = Empty Then Exit Sub
    
    lShipAddrKey = m_gwCustOrders.value("ShipAddrKey")
    
    If lShipAddrKey = m_oCustomer.ShipAddr.AddrKey And lShipAddrKey <> 0 Then
    Else
        Dim oAddr As Address
        On Error GoTo LoadFailed
        
        Set oAddr = New Address
        If lShipAddrKey <> 0 Then
            oAddr.Load lShipAddrKey
        Else
            'If this address is shipping only address, retrieve the address from XML in tcpSO
            oAddr.Import ImportString(m_gwCustOrders.value("ShipAddrXML"))
        End If
        ShowOrderAddress oAddr.CompleteAddr
    End If
    Exit Sub

LoadFailed:
    ShowOrderAddress "No Shipping Address specified for this order"
End Sub


'********************************************************************
' Right click menu on the gdxCustOrders grid
'********************************************************************

Private Sub gdxCustOrders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bTemp As Boolean
    
    If Button = vbRightButton Then
        bTemp = gdxCustOrders.GroupByBoxVisible
        mnuCustOrderGroup.Checked = bTemp
        mnuCustOrderExpand.Enabled = bTemp
        mnuCustOrderCollapse.Enabled = bTemp
        Me.PopupMenu mnugdxCustOrder
    End If
End Sub


Private Sub mnuCustOrderAutofit_Click()
    m_gwCustOrders.GridAutoFit
End Sub


Private Sub mnuCustOrderFont_Click()
    ChangeGridFont gdxCustOrders
End Sub


Private Sub mnuOrderFont_Click()
    ChangeGridFont gdxOrders
End Sub


Private Sub mnuCustOrderCollapse_Click()
    DoEvents
    gdxCustOrders.CollapseAll
    gdxCustOrders.Refresh
End Sub


Private Sub mnuCustOrderExpand_Click()
    DoEvents
    gdxCustOrders.ExpandAll
    gdxCustOrders.Refresh
End Sub


Private Sub mnuCustOrderGroup_Click()
    Dim bCheck As Boolean
    Dim lGroups As Long
    Dim lIndex As Long
    
    bCheck = mnuCustOrderGroup.Checked
    
    mnuCustOrderGroup.Checked = Not bCheck
    mnuCustOrderExpand.Enabled = Not bCheck
    mnuCustOrderCollapse.Enabled = Not bCheck
    gdxCustOrders.GroupByBoxVisible = Not bCheck
    
    lGroups = gdxCustOrders.Groups.Count
    
    If bCheck And lGroups > 0 Then
        For lIndex = 1 To lGroups
            gdxCustOrders.Groups.Remove 1
        Next
    End If
End Sub


Private Sub mnuCustOrderPrint_Click()
    gdxCustOrders.PrinterProperties.Orientation = jgexPPLandscape
    gdxCustOrders.PrintGrid
End Sub


Private Sub mnuCustOrderRefresh_Click()
    If m_lCustKey > 0 Then FillSelectCustTab m_lCustKey
End Sub


Private Sub mnuCustOrderSaveLayout_Click()
    SetWaitCursor True
    Call m_gwCustOrders.GridSaveLayout(GetUserKey)
    SetWaitCursor False
End Sub


Private Sub mnuCustOrdersExpand_Click()
    gdxCustOrders.ExpandAll
End Sub


Private Sub mnuCustOrderDetail_Click()
    Dim sMsg As String
    
    Dim orst As ADODB.Recordset
    Set orst = GetInvoiceDetail(m_gwCustOrders.value("OPKey"))

    If orst.EOF Then
        GlobalFunctions.DisplayInfo "OP " & m_gwCustOrders.value("OPKey") & " Ship Detail", "This order hasn't shipped yet.", 3000, 1100
    Else
        sMsg = "Freight = " & Format(orst.Fields("ShipAmt").value, "#,##0.00") & vbCrLf & "Invoice Total = " & Format(orst.Fields("TranAmt").value, "#,##0.00")
        GlobalFunctions.DisplayInfo "OP " & m_gwCustOrders.value("OPKey") & " Ship Detail", sMsg, 3000, 1100
    End If
    MDIMain.DoRefresh
End Sub

'*** End of Right-Click menu code


'This procedure clears control fields of Select Customer tab for a new search
'Called by
'    cmdNewSearch_Click
'    FindOrderByNumber
'    cmdLoadOrder_Click

Private Sub ClearCustomerSearch()
    m_lCustKey = 0
    txtCustSearch.text = ""
    ClearCustomerOrders
    SetSelCustCtrls True
    
    If tabMain.SelectedTab.Index = tmiExistingCustomer Then
        TryToSetFocus txtCustSearch
        m_oBrokenRules.Validate txtCustSearch
    End If
End Sub


Private Sub ClearCustomerOrders()
    With gdxCustOrders
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = CreateEmptyOrderSummaryRecordset
        cmdNewOrder.Enabled = False
        cmdLoadOrder(0).Enabled = False
        .Refetch
    End With
End Sub


'Load the customer with the found key.
'Load the controls on Select Customer tab.
'
'Called by:
'   cmdCustSearch_Click()           FindMode = True
'   mnuCustOrderRefresh_Click()     FindMode = True
'   SaveButton()                    FindMode = False
'   Commit()                        FindMode = False
'   CancelButton()                  FindMode = False
'
'   Note: the FindMode parameter has a different meaning than m_bFindMode
'
'Calls:
'   SetSelCustCtrls()
'   m_oCustomer.Load()
'   LoadCustomerOrders()

Private Sub FillSelectCustTab(ByVal lCustKey As Long, Optional bFindMode As Boolean = True)
    SetWaitCursor True
    
    'set controls on Select Customer for load customer orders mode.
    SetSelCustCtrls False
    
    With m_oCustomer
        .Load lCustKey
        
        'Set the customer header information
        SetCaption .Name
        lblCustID(1).caption = .ID
        lblCustType(1).caption = .CustType
        
        If bFindMode Then
            m_lDefaultBillAddrKey = .BillAddr.AddrKey
            m_lDefaultShipAddrKey = .ShipAddr.AddrKey
        Else
            .BillAddr.Load m_lDefaultBillAddrKey
            .ShipAddr.Load m_lDefaultShipAddrKey
        End If
        
        lblCustAddress.caption = .ShipAddr.CompleteAddr
    End With

    'change the default value of the checkbox
    'Load orders to grid
    LoadCustomerOrders (chkShowOrdersForShipAddr.value = vbChecked), cboOrderStatus.text

    cmdNewOrder.Enabled = HasRight(k_sRightShowToolOP)

    'if the customer has orders
    If gdxCustOrders.RowCount > 0 Then
        cmdLoadOrder(0).Enabled = True
        If FindMode Then TryToSetFocus gdxCustOrders    '.SetFocus
    Else
        cmdLoadOrder(0).Enabled = False
        If bFindMode Then
            TryToSetFocus cmdNewOrder
            If vbYes = msg("Would you like to create a new order for " & m_oCustomer.Name & "?", vbYesNo, "New Order?") Then
                cmdNewOrder_Click
            End If
        End If
    End If
    
    SetWaitCursor False
End Sub


'Called by:
'   chkShowOrdersForShipAddr_Click()    bShowAllShipAddr = value of chkShowOrdersForShipAddr
'   cboOrderStatus_Click()              bShowAllShipAddr = value of chkShowOrdersForShipAddr
'   FillSelectCustTab()                 bShowAllShipAddr = True

Private Sub LoadCustomerOrders(ByVal bShowAllShipAddr As Boolean, ByVal sOrderStatus As String, Optional ByVal sPartNo As String = "")
    Dim orst As ADODB.Recordset
    Dim sSQL As String
    Dim sWhere As String
    
    SetWaitCursor True

    If sOrderStatus = "Open RMA" Then
    
        sSQL = "SELECT OPKey, StatusCode, CreateDate, RTRIM(UserID) as UserID, WhseKey, CustKey, OrderedBy, SOKey, Summary, PurchOrd AS CustPO, ShipAddrKey, Info as Note " & _
           "FROM tcpSO WHERE 1 = 2"
        Set orst = LoadDiscRst(sSQL)
        
    Else
    
        sWhere = "WHERE o.CustKey=" & m_oCustomer.Key & " AND StatusCode < " & iscDeleted

        'Add keyword filter
        
        Dim vKeywords As Variant
        Dim i As Long

        vKeywords = Split(Trim$(txtFilterByPart), " ")
        For i = LBound(vKeywords) To UBound(vKeywords)
            AppendClause "o.Keywords LIKE '%" & PrepSQLText(CStr(vKeywords(i))) & "%'", sWhere
        Next
            
        If Not bShowAllShipAddr Then
            sWhere = sWhere & " AND ShipAddrKey=" & m_oCustomer.ShipAddr.AddrKey
        End If

        AddStatusFilterClause sWhere, cboOrderStatus

        'Load special research status to the grid instead of general research status if applicable
        'Add SOID to Customer order recordset
        'Add tsoSalesOrder to retrieve Open/Closed status

        sSQL = "SELECT o.OPKey as OPID, o.TranKey as SOID, o.CreateDate, o.ShipAddrXML, rma.rmakey, o.Info as Note, " _
             & "Case StatusCode when 0 then 0 when 1 then " _
             & "Case when ResearchStatus is null then 1001 " _
             & "when ResearchStatus = 0 then 1001 " _
             & "Else 1000+ResearchStatus End " _
             & "Else 2000+StatusCode End As StatusCode, " _
             & "(CASE WHEN o.flags&0x01 = 0x01 THEN -1 ELSE 0 END) AS Dropship, " _
             & "o.UpdateDate, RTRIM(o.UserID) as UserID, o.WhseKey, o.CustKey, " _
             & "RTRIM(ISNULL(c.Name, '')) as OrderedBy, o.OPKey, " _
             & "o.SOKey, o.Summary, RTRIM(o.PurchOrd) as CustPO, ShipAddrKey, " _
             & "(CASE tsoSalesOrder.Status WHEN 1 THEN 'Open' when 4 then 'Closed' ELSE 'Other' END) StatusDesc, " _
             & "tsoSalesOrder.Status"

        sSQL = sSQL & " FROM tcpSO o LEFT OUTER JOIN " & _
        "tsoSalesOrder ON o.SOKey = dbo.tsoSalesOrder.SOKey LEFT OUTER JOIN " & _
        "tcpRMA rma on rma.OPKey = o.OPKey LEFT OUTER JOIN " & _
        "tciContact c ON o.CntctKey = c.CntctKey " & _
        sWhere & " ORDER BY o.updatedate Desc"

        Set orst = LoadDiscRst(sSQL, , , g_MaxCustOrders)
    End If
    
    With gdxCustOrders
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
        cmdLoadOrder(0).Enabled = (.RowCount > 0)
        .FmtConditions("CustAddr").value1 = m_oCustomer.ShipAddr.AddrKey
    End With

    SetWaitCursor False

    If gdxCustOrders.RowCount = g_MaxCustOrders Then
        lblOrderCount = "The last " & g_MaxCustOrders & " orders"
    Else
        lblOrderCount = gdxCustOrders.RowCount & " orders"
    End If
    
    Set orst = Nothing
    
End Sub

' Display the intenational country caution based on the customer's shipping address

Private Sub DisplayIntrlCaution()
    Dim orst As ADODB.Recordset
    Dim sMsg As String
    Dim sMsgTitle As String
    
    sMsgTitle = m_oCustomer.Name

    Set orst = GetCautionsForCountry(m_oCustomer.ShipAddr.CountryID)

    If Not orst.EOF Then
        If IsNull(orst.Fields("Cautions").value) Then
            'there are no cautions for this country
        Else
            sMsgTitle = sMsgTitle & " - Order Shipping To: " & Replace(orst.Fields("CSWCountrySymbol").value, "_", " ")
            sMsg = orst.Fields("Cautions").value
        End If
    End If

    If Len(Trim$(sMsg)) = 0 Then
        sMsg = "This country is not set up for shipping.  Please contact IT."
    End If
        
    MsgBox sMsg, vbOKOnly, sMsgTitle
End Sub


'************************************************************************************
' Events and subroutines on Select Order tab
'************************************************************************************

Private Sub cmdFindOrders_Click(Index As Integer)
    Select Case Index
        Case 0:
            FindOrderByNumber
        Case 1:
            FindOrdersByCriteria
    End Select
End Sub

'Find order via OPID, SOID(Trankey), RMA Key, or Customer PO
'Called by:
'   cmdFindOrders_Click
'   mnuOrderRefresh_Click
'   txtFindOrder_KeyDown
' This relates to changing from OrderMode to FindMode
' It'll come out when separating FOrderManager
'   OrderDelete

Private Sub FindOrderByNumber()
    Dim sInput As String
    Dim orst As ADODB.Recordset
    
    sInput = Trim$(PrepSQLText(txtFindOrder.text))
    If Len(sInput) = 0 Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
            
    If InStr(LCase(sInput), "p") = 1 Then
'        If Not IsNumeric(Mid$(sInput, 2)) Then
'            ReportLookupFailure "Input must be numeric, though if preceeded by P/p it will look by associated Customer PO, " & _
'                "or if preceeded by A/a it will look by associated CaseParts PO"
'            Exit Sub
'        End If

        m_bFindOrderFlag = True
        
        SetWaitCursor True
        Set orst = CallSP("spcpcFindOrderByCustPO", "@_iCustPO", Mid$(sInput, 2))
        SetWaitCursor False
    
    ElseIf InStr(LCase(sInput), "a") = 1 Then
        If Not IsNumeric(Mid$(sInput, 2)) Then
            ReportLookupFailure "Input must be numeric, though if preceeded by P/p it will look by associated Customer PO, " & _
                "or if preceeded by A/a it will look by associated CaseParts PO"
            Exit Sub
        End If
        
        m_bFindOrderFlag = True
        
        SetWaitCursor True
        Set orst = CallSP("spcpcFindOrderByCPCPO", "@_iPOTranNo", Mid$(sInput, 2))
        SetWaitCursor False
        
    Else
        If Not IsNumeric(sInput) Then
            ReportLookupFailure "Input must be numeric, though if preceeded by P/p it will look by associated Customer PO, " & _
                "or if preceeded by A/a it will look by associated CaseParts PO"
            Exit Sub
        End If
           
        m_bFindOrderFlag = True

        SetWaitCursor True
        Set orst = CallSP("spcpcFindOrderByNumber", "@_iOrderNumber", CLng(sInput))
        SetWaitCursor False
    End If
    
    With gdxOrders
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
    End With
    
    Select Case orst.RecordCount
        Case 0:
            ReportLookupFailure "No records satisfy this request"
            
        Case 1:
            
            'SOStatusCode = 0 for now
            If LoadOrder(orst("OPKey").value, _
                            orst("StatusCode").value, _
                            orst("SOKey").value, _
                            orst("SOID").value, _
                            orst("Status").value) Then
                
                m_bLoading = True
            
                ClearCustomerSearch  'invoked from Select Order tab
            
                'loading a MISC order with OrderedBy value
                'Contact state went from New to New + Dirty
                Set m_oCustomer = m_oOrder.Customer
                Set m_oItems = m_oOrder.Items
                
                TransitionTabs False
            
                LoadUPSCtrl m_oCustomer.Key
            
                m_bLoading = False
            
            End If
            
        Case Else
            TryToSetFocus gdxOrders
            cmdLoadOrder(1).Enabled = True
            
    End Select

    Set orst = Nothing
        
    Exit Sub
    
ErrorHandler:
    SetWaitCursor False
    msg "Error while finding order." & vbCrLf & _
        Err.Description & vbCrLf & "Check the number you entered."
    txtFindOrder.SelStart = 0
    txtFindOrder.SelLength = Len(txtFindOrder.text)
    TryToSetFocus txtFindOrder
End Sub


Private Sub ReportLookupFailure(message As String)
    msg message, vbInformation, "Input Error"
    txtFindOrder.SelStart = 0
    txtFindOrder.SelLength = Len(txtFindOrder.text)
    TryToSetFocus txtFindOrder
End Sub


'Search orders via combination of many criteria: Customer, CSR, Status, etc.
'Called by
'   cmdFindOrders_Click
'   txtFindCust_KeyDown
'   txtFindText_KeyDown
'   mnuOrderRefresh_Click
' These relate to changing from OrderMode to FindMode
' They'll come out when separating FOrderManager
'   CancelButton
'   SaveButton
'   CommitButton

Private Sub FindOrdersByCriteria()

    Dim sWhere As String
    Dim sInput As String
    Dim lResult As Long
    Dim sTemp As String
    Dim rst As ADODB.Recordset
    
    SetWaitCursor True
    
    m_bFindOrderFlag = False
    
    If cboFindStatus.text = "Open RMA" Then
    
        Select Case cboFindCSR.text
            Case "<Any>"
                Set rst = CallSP("spCPCOpenRMAOrder")
            Case "<STL>"
                Set rst = CallSP("spCPCOpenRMAOrder", "@_iWhseKey", 25)
            Case "<MPK>"
                Set rst = CallSP("spCPCOpenRMAOrder", "@_iWhseKey", 23)
            Case "<SEA>"
                Set rst = CallSP("spCPCOpenRMAOrder", "@_iWhseKey", 24)
            Case Else
                Set rst = CallSP("spCPCOpenRMAOrder", "@_iUserID", cboFindCSR.text)
        End Select
    
    Else
    
        sInput = PrepSQLText(txtFindCust.text)

        'Search on Customer Name or ID
        If Len(sInput) > 0 Then
            AppendClause "ta.CustID LIKE '" & sInput & "%' OR ta.CustName LIKE '" & sInput & "%'", sWhere
        End If

        'Search for Keywords
        sInput = Trim(txtFindText.text)
        If Len(sInput) > 0 Then
            Dim vKeywords As Variant
            Dim i As Long
            
            vKeywords = Split(sInput, " ")
            For i = LBound(vKeywords) To UBound(vKeywords)
                AppendClause "o.Keywords LIKE '%" & PrepSQLText(CStr(vKeywords(i))) & "%'", sWhere
            Next
        End If
    
        Select Case cboFindCSR.text
            Case "<Any>"
            Case "<STL>"
                AppendClause "WhseKey = 25", sWhere
            Case "<MPK>"
                AppendClause "WhseKey = 23", sWhere
            Case "<SEA>"
                AppendClause "WhseKey = 24", sWhere
            Case Else
                sInput = cboFindCSR.text
                AppendClause "UserID = '" & sInput & "'", sWhere
        End Select

        AddStatusFilterClause sWhere, cboFindStatus

        On Error GoTo EH
        
' Use the new time interval combo in our query
' NOTE: we've coded the data directly into the control's List and ItemData properties using the design-time property pages

        If Len(sWhere) > 0 Then
            sWhere = sWhere & " and o.UpdateDate > '" & Now - cboTimeInterval.ItemData(cboTimeInterval.ListIndex) & "'"
        Else
            sWhere = " WHERE o.UpdateDate > '" & Now - cboTimeInterval.ItemData(cboTimeInterval.ListIndex) & "'"
        End If
        
        Set rst = SearchOrder(sWhere)
        
    End If
    
    LoadResultGrid rst
    
    Set rst = Nothing
    SetWaitCursor False
    
    Exit Sub

EH:
    SetWaitCursor False

    If Err.Number = -2147217871 Then  'ODBC SQL Server Driver  timeout expired
        LogEvent "FOrder", "FindOrdersByCriteria", "Query timed out (" & sWhere & ")"
        msg "Your query timed out." & vbCrLf & "Please try again.", vbExclamation, "OrderPad Warning"
    Else
        LogError "FOrder", "FindOrdersByCriteria", "", Err.Source, Err.Number, Err.Description
        msg "FOrder.FindOrdersByCriteria" & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "Call the computer guys.", vbCritical, "OrderPad Error"
    End If
End Sub


Private Sub LoadResultGrid(rst As ADODB.Recordset)
    With gdxOrders
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = rst
    End With

    If InStr(1, GetUserName, "WillCall", vbTextCompare) > 0 Then
        SetWCGridPrefs
    End If
    
    If gdxOrders.RowCount = 0 Then
        txtFindCust.SelStart = 0
        txtFindCust.SelLength = Len(txtFindCust.text)
        TryToSetFocus txtFindCust
        cmdLoadOrder(1).Enabled = False
    Else
        cmdLoadOrder(1).Enabled = True
        TryToSetFocus gdxOrders
    End If
End Sub
    
Private Sub SetSearchDefaults()
    Dim rst As ADODB.Recordset
    Dim sSQL As String
    
    txtFindCust.text = ""
    txtFindText.text = ""

    If InStr(1, GetUserName, "LAWillCall", vbTextCompare) > 0 Then
        SetComboByText cboFindCSR, "<MPK>"
        SetComboByText cboFindStatus, "<Any>"
    Else
        SetComboByText cboFindCSR, GetUserName
        SetComboByText cboFindStatus, "All Quotes"
    End If
    
    TryToSetFocus txtFindCust
    
    'generate a zero-record recordset to clear the grid
    sSQL = "SELECT OPKey FROM tcpSO WHERE 1 = 2"
    Set rst = LoadDiscRst(sSQL)
    LoadResultGrid rst
End Sub


Private Sub LoadWCGridPrefs()
    Set m_colWCGridPrefs = New Collection
    Select Case GetUserName
        Case "LAWillCall", "LAWillCall2"
            m_colWCGridPrefs.Add 1, "MPK"
            m_colWCGridPrefs.Add 1, "MPK-Will Call"
            m_colWCGridPrefs.Add 1, "Committed"
            m_colWCGridPrefs.Add 1, "Ready to Commit"
            m_colWCGridPrefs.Add 1, "Authorize"
            
            gdxOrders.Groups.Clear
            gdxOrders.GroupByBoxVisible = True
            'Maximum number of groups that can be added is 4.
            gdxOrders.Groups.Add gdxOrders.Columns("WhseKey").Index, jgexSortAscending
            gdxOrders.Groups.Add gdxOrders.Columns("ShipMethID").Index, jgexSortDescending
            gdxOrders.Groups.Add gdxOrders.Columns("StatusCode").Index, jgexSortDescending
            
            gdxOrders.Columns("WhseKey").Visible = False
            gdxOrders.Columns("ShipMethID").Visible = False
            gdxOrders.Columns("StatusCode").Visible = False

    End Select
End Sub

'This is used to expand only the columns needed for Will Call users.
'If the collection is not populated on form load, all rows will be expanded.
Private Sub SetWCGridPrefs()
    Dim i As Long

    gdxOrders.Redraw = False
    For i = 1 To gdxOrders.RowCount
        If gdxOrders.IsGroupItem(i) Then
            On Error Resume Next
            gdxOrders.RowExpanded(i) = m_colWCGridPrefs(Trim$(gdxOrders.GetRowData(i).GroupCaption))
            If Err.Number > 0 Then
                gdxOrders.RowExpanded(i) = False
            End If
        End If
    Next i
    gdxOrders.Redraw = True

End Sub

'This event is used to reset searching criteria
Private Sub cmdClearOrder_Click()
'*** 7/23/08 DH
    SetSearchDefaults
End Sub

'Right click on order grid will load context menu
Private Sub gdxOrders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bTemp As Boolean
    If Button = vbRightButton Then
        bTemp = gdxOrders.GroupByBoxVisible
        mnuOrderGroup.Checked = bTemp
        mnuOrderExpand.Enabled = bTemp
        mnuOrderCollapse.Enabled = bTemp
        Me.PopupMenu mnugdxOrder
    End If
End Sub

'When grid is sorted by column header.
Private Sub gdxOrders_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    Dim SortOrder As Integer
    
    If Column.IsGrouped Then Exit Sub

    SortOrder = Column.SortOrder
    'Remove the existing sorting criteria, if any
    gdxOrders.SortKeys.Clear
    If SortOrder = jgexSortAscending Then
        gdxOrders.SortKeys.Add Column.Index, jgexSortDescending
    Else 'if SortOrder is none or is descending
        gdxOrders.SortKeys.Add Column.Index, jgexSortAscending
    End If
    
End Sub

'Shows or hides columns as they are added or removed from grouping.
Private Sub gdxOrders_BeforeGroupChange(ByVal group As GridEX20.JSGroup, ByVal ChangeOperation As GridEX20.jgexGroupChange, ByVal GroupPosition As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    gdxOrders.Columns(group.ColIndex).Visible = Not gdxOrders.Columns(group.ColIndex).Visible
End Sub


Private Sub mnuOrderCollapse_Click()
    DoEvents
    gdxOrders.CollapseAll
    gdxOrders.Refresh
End Sub


Private Sub mnuOrderExpand_Click()
    DoEvents
    gdxOrders.ExpandAll
    gdxOrders.Refresh
End Sub


Private Sub mnuOrderGroup_Click()
    Dim bCheck As Boolean
    Dim oCol As JSColumn
    Dim lGroups As Long
    Dim lIndex As Long
    
    bCheck = mnuOrderGroup.Checked
    
    mnuOrderGroup.Checked = Not bCheck
    mnuOrderExpand.Enabled = Not bCheck
    mnuOrderCollapse.Enabled = Not bCheck
    gdxOrders.GroupByBoxVisible = Not bCheck
    
    lGroups = gdxOrders.Groups.Count
    If lGroups > 0 And bCheck Then
        For lIndex = 1 To lGroups
            gdxOrders.Groups.Remove 1
        Next
    End If
End Sub


Private Sub mnuOrderHelp_Click()
    msg "Sorry. Not available"
End Sub


Private Sub mnuOrderPrint_Click()
    gdxOrders.PrinterProperties.Orientation = jgexPPLandscape
    gdxOrders.PrintGrid
End Sub


Private Sub mnuOrderRefresh_Click()
    If m_bFindOrderFlag Then
        'cmdFindOrders_Click
        FindOrderByNumber
    Else
        'cmdFindOrders2_Click
        FindOrdersByCriteria
    End If
End Sub


Private Sub mnuOrderRestore_Click()
    Call m_gwOrders.RestoreGridLayout(GetUserKey)
    gdxOrders.Refresh
End Sub


Private Sub mnuOrdersAutofit_Click()
    Call m_gwOrders.GridAutoFit
End Sub


Private Sub mnuOrderChangeColumns_Click()
    Dim oFrm As FSelectCols
    Set oFrm = New FSelectCols
    Call oFrm.Init(Me.gdxOrders)
End Sub


Private Sub mnuCustOrderChangeColumns_Click()
    '12/09/2004 - smr
    Dim oFrm As FSelectCols
    
    Set oFrm = New FSelectCols
    Call oFrm.Init(Me.gdxCustOrders)
End Sub


Private Sub mnuOrderSaveLayout_Click()
    SetWaitCursor True
    Call m_gwOrders.GridSaveLayout(GetUserKey)
    SetWaitCursor False
End Sub


'Get the order status from combo box and add it to the where clause of the search query

Private Sub AddStatusFilterClause(ByRef i_sWhere As String, oCtrl As ComboBox)
    If oCtrl.text <> "<Any>" Then
        Select Case oCtrl.text
        Case "New":                 AppendClause "StatusCode = 0", i_sWhere
        Case "Research":            AppendClause "StatusCode = 1", i_sWhere
        Case "Quote":               AppendClause "StatusCode = 2", i_sWhere
        Case "Contact Factory":     AppendClause "StatusCode = 1 and ResearchStatus = 2", i_sWhere
        Case "Contact Customer":    AppendClause "StatusCode = 1 and ResearchStatus = 3", i_sWhere
        Case "Wait Factory":        AppendClause "StatusCode = 1 and ResearchStatus = 4", i_sWhere
        Case "Wait Customer":       AppendClause "StatusCode = 1 and ResearchStatus = 5", i_sWhere
        Case "In Our Court":        AppendClause "StatusCode = 2 or (StatusCode=1 and (ResearchStatus = 2 or ResearchStatus=3))", i_sWhere
        Case "Need Authorization":  AppendClause "StatusCode = 3", i_sWhere
        Case "Ready to Commit":     AppendClause "StatusCode = 4", i_sWhere
        Case "On Hold":             AppendClause "StatusCode = 5 or StatusCode = 6 or StatusCode = 7", i_sWhere
        Case "All Quotes":          AppendClause "StatusCode <= 7", i_sWhere
        
        Case "Committed":           AppendClause "StatusCode = 256 OR StatusCode = 128", i_sWhere
        Case "RMA":                 AppendClause "StatusCode = 128", i_sWhere
        Case "Deleted":             AppendClause "StatusCode >= 512", i_sWhere
        Case "Not Deleted":         AppendClause "StatusCode < 512", i_sWhere
        Case Else
            ErrorUI.DisplayError "Invalid search type: " & oCtrl.text
        End Select
    End If
End Sub


'************************************************************************************
'All events and subroutines on Customer tab
'************************************************************************************

'Edit billing address and shipping address

Private Sub cmdEditAddr_Click(Index As Integer)
    Dim frmThisOrderOnly As FThisOrderOnlyAddress
    Dim frmChangeAddr As FChangeAddr
    Dim ShipAddr As Address
    
    Select Case Index
    Case 0: 'Create a This Order Only Address
        Set frmThisOrderOnly = New FThisOrderOnlyAddress
        
        Set ShipAddr = m_oCustomer.ShipAddr
        
        If frmThisOrderOnly.EditShipAddress(ShipAddr) = VbMsgBoxResult.vbOK Then
            If m_oCustomer.ShipAddr.AddrType <> TOO Then
                m_oCustomer.ShipAddr = ShipAddr
                'This is only used for event logging purposes via Addr.Export
                m_oCustomer.ShipAddr.OPKey = m_oOrder.OPKey
            End If
            
            UpdateShipAddrInfo
            
            'Load Sales Tax
            m_oOrder.SalesTax.Init m_oCustomer
            If m_oOrder.IsWillCall Then
                m_oOrder.SalesTax.WillCallTaxOverride m_oOrder.whseid
            End If
            
            TestShipComplete
            chkPricePackList.value = vbUnchecked

            If m_oCustomer.IsTemp Then
                m_oCustomer.Name = m_oCustomer.BillAddr.AddrName
            End If
        End If
        Set frmThisOrderOnly = Nothing
        
    Case 1: 'Select a different CSA for the current customer
        Set frmChangeAddr = New FChangeAddr
        If frmChangeAddr.Load(m_oCustomer) = VbMsgBoxResult.vbOK Then
            'Load Sales Tax
            m_oOrder.SalesTax.Init m_oCustomer
            If m_oOrder.IsWillCall Then
                m_oOrder.SalesTax.WillCallTaxOverride m_oOrder.whseid
            End If
        
            '***Setting PPL chk based on Cust.PPLDefault and Addr.AddrType***
            'Reset PricePackList Flag; AddrType may have changed.
            If m_oCustomer.PricePackList = True And m_oCustomer.ShipAddr.AddrType = Default Then
                chkPricePackList.value = vbChecked
                m_oOrder.PricePackList = True
            ElseIf m_oCustomer.PricePackList = True And m_oCustomer.ShipAddr.AddrType = CSA Then
                If MsgBox("You've selected a Common Shipping Address and the customer's " & vbCrLf & "preference is to show prices on their packing slip. " & vbCrLf & vbCrLf & "Select 'Yes' to show prices or 'No' to not show prices.", vbYesNo, "Price Pack List") = vbYes Then
                    chkPricePackList.value = vbChecked
                    m_oOrder.PricePackList = True
                Else
                    chkPricePackList.value = vbUnchecked
                    m_oOrder.PricePackList = False
                End If
                Repaint 'To refresh the tab when the message box is displayed
            Else
                chkPricePackList.value = vbUnchecked
                m_oOrder.PricePackList = False
            End If
            
            UpdateShipAddrInfo
        End If
        Set frmChangeAddr = Nothing
    End Select
End Sub


'Update billing address information on Customer tab
Private Sub UpdateBillAddrInfo()
    With m_oCustomer.BillAddr
        lblBillAddr.caption = .CompleteAddr
    End With
End Sub


'Update shipping address information on Customer tab
Private Sub UpdateShipAddrInfo()
    With m_oCustomer.ShipAddr
        lblShipAddr.caption = .CompleteAddr
    End With
End Sub


Private Sub UpdateShipContactInfo()
    Dim orst As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tciContact.Name,tciContact.Phone,tciContact.PhoneExt,tciContact.Fax,tciContact.FaxExt " _
    & "FROM tarCustomer (nolock) INNER JOIN tarCustAddr (nolock) ON tarCustomer.DfltShipToAddrKey = tarCustAddr.AddrKey INNER JOIN " _
    & "tciContact (nolock) ON tarCustAddr.DfltCntctKey = tciContact.CntctKey WHERE tarCustomer.CustKey=" & m_oCustomer.Key

    'clear the display
    lblShipContact.caption = vbNullString
    lblShipPhone.caption = vbNullString
    lblShipFax.caption = vbNullString
    lblCellPhone.caption = vbNullString

'assertion: if m_oOrder.Contact is not nothing the .Key>0

    'if this Order does not have a contact
    If (m_oOrder.contact Is Nothing) Then
        If m_oCustomer.HasAccount Then
            'Look for a Default Shipping Contact
            Set orst = LoadDiscRst(sSQL)
            If Not orst.EOF Then
                With orst
                    lblShipContact.caption = FixNullField(.Fields("Name").value)
                    lblShipPhone.caption = FormatPhoneNumber(FixNullField(.Fields("Phone").value), FixNullField(.Fields("PhoneExt").value))
                    lblShipFax.caption = FormatPhoneNumber(FixNullField(.Fields("Fax").value), FixNullField(.Fields("FaxExt").value))
                    lblCellPhone.caption = vbNullString
                End With
            End If
            orst.Close
            Set orst = Nothing
        End If
    'otherwise display the information in the Contact
    Else
        With m_oOrder.contact
            lblShipContact.caption = .Name
            lblShipPhone.caption = FormatPhoneNumber(.Phone, .PhoneExt)
            lblShipFax.caption = FormatPhoneNumber(.Fax, .FaxExt)
            lblCellPhone.caption = FormatPhoneNumber(.CellPhone)
        End With
    End If
End Sub


Private Function FixNullField(field As Variant) As String
    FixNullField = IIf(IsNull(field), vbNullString, field)
End Function


'Update order's ShipComplete property after editing billing and shipping address
Private Sub TestShipComplete()
    Dim bShipComplete As Boolean
    
    With m_oCustomer
        bShipComplete = (.BillAddr.ShipComplete Or .ShipAddr.ShipComplete)
        If bShipComplete <> m_oOrder.ShipComplete Then
            If vbYes = msg("Would you like to update this order's ShipComplete from " & m_oOrder.ShipComplete & " to " & _
                            bShipComplete & "?", vbYesNo + vbExclamation, "Update ShipComplete") Then
                m_bLoading = True
                If bShipComplete Then
                    chkShipComplete.value = vbChecked
                Else
                    chkShipComplete.value = vbUnchecked
                End If
                m_bLoading = False
            End If
        End If
    End With
End Sub


'************************************************************************************
'All events and subroutines on Order tab
'************************************************************************************

'This event assigns CSR to the order
Private Sub cboCSR_Click()
    If m_bLoading Then Exit Sub
    
    m_oOrder.UserKey = cboCSR.ItemData(cboCSR.ListIndex)
    If Len(cboWarehouse(0).text) < 1 Then Exit Sub
    UpdateWarehouseColor (0)
End Sub


'Choose Ship Via for the order
Private Sub cboShipVia_Click()
    Dim lShipMethKey As Long
    
    If m_bLoading Then Exit Sub
    
    m_bLoading = True
    
    lShipMethKey = m_oOrder.ShipMethKey
    
    With cboShipVia
        m_oOrder.ShipMethKey = .ItemData(.ListIndex)
        
        'Update default ship method check box accordingly
        If Trim(cboShipVia.text) = Trim(m_sDefaultShipMeth) And chkDefaultShipMeth.Enabled Then
            chkDefaultShipMeth.value = vbChecked
        Else
            chkDefaultShipMeth.value = vbUnchecked
        End If
        
        
        'Update Bill recipient accordingly
        
        'If current payment terms is COD or COD-cash, prevent the user from specifying Bill Recipient
      
        If Not m_oOrder.PmtTerms.IsCOD Then
            If Left(Trim(cboShipVia.text), 3) = "UPS" Then
                If m_oOrder.Customer.HasAccount Then
                'NOTE: The Bill Recipient logic is disabled for Build 61, LR 3/8/02
                    chkBillRecipient.Enabled = True
                Else
                    DisableBillRecipient
                End If
            
            ElseIf Not m_bWareHouseLoading Then
                If chkBillRecipient.value = vbChecked Then
                    If vbYes = msg("Are you sure you want to change ship method to NON-UPS method? ", _
                                    vbYesNo + vbExclamation, "Ship Method Change?") Then
                            DisableBillRecipient
                    Else
                        m_oOrder.ShipMethKey = lShipMethKey
                        SetComboByKey cboShipVia, m_oOrder.ShipMethKey
                    End If
                Else
                    DisableBillRecipient
                End If
            
            Else
                'The cascading event of changing warehouse maybe cause Bill Recipient to be disabled
                DisableBillRecipient
            End If
            
        Else
            DisableBillRecipient
        End If
        
        MDIMain.UpdateToolbarStatus
        
    End With
    
    UpdateOrderInfo
        
    m_bLoading = False
End Sub


Private Sub EnableShipToContactCtrls()
    If cboShipVia.text = "UPS STND" Or cboShipVia.text = "UPS Red AM" Then
        txtShipToContact(0).Visible = True
        txtShipToContact(1).Visible = True
        Label1(12).Visible = True
        Label1(13).Visible = True
        
        m_oBrokenRules.EnableClass ccShipToData, True
    Else
        txtShipToContact(0).Visible = False
        txtShipToContact(1).Visible = False
        Label1(12).Visible = False
        Label1(13).Visible = False
        
        m_oBrokenRules.EnableClass ccShipToData, False
    End If
    
    m_oBrokenRules.Validate txtShipToContact(0)
    m_oBrokenRules.Validate txtShipToContact(1)
End Sub


'Set related controls in Order tab if Bill Recipient is disabled.

Private Sub DisableBillRecipient()
    chkBillRecipient.Enabled = False
    chkBillRecipient.value = vbUnchecked
    txtUPSAcct.text = ""
    cmdUPSUpdate.Enabled = False
    m_oOrder.UPSAcct = ""
End Sub


Private Sub txtFilterByPart_Change()
    cmdFilterByPart.Default = True
End Sub


Private Sub txtInfo_LostFocus()
    m_oOrder.Info = txtInfo.text
End Sub

Private Sub txtInfo_Change()
    m_oOrder.Info = txtInfo.text
End Sub

Private Sub txtShipToContact_LostFocus(Index As Integer)

    If Index = 0 Then
        m_oOrder.ShipToName = txtShipToContact(Index).text
    ElseIf Index = 1 Then
        m_oOrder.ShipToPhone = txtShipToContact(Index).text
    End If

    m_oBrokenRules.Validate txtShipToContact(Index)

End Sub

Private Sub txtShipToNote_Change()
    m_oOrder.ShipToNote = txtShipToNote.text
End Sub

'Set default ship method for current CustID
Private Sub chkDefaultShipMeth_Click()
    Dim sTemp As String
    Dim bLoading As Boolean
    
    If m_bLoading Then Exit Sub
    bLoading = m_bLoading

    If m_bLoading Then Exit Sub
    m_bLoading = True
 
    If chkDefaultShipMeth.value Then
        sTemp = "Set the default shipping method for " & m_oCustomer.ID & " to" & vbCrLf & cboShipVia.text
            
        sTemp = sTemp & "?"

        If vbNo = MsgBox(sTemp, vbYesNo + vbQuestion, "Confirm Default Shipping Method") Then
            chkDefaultShipMeth.value = vbUnchecked
            m_bLoading = bLoading
            Exit Sub
        End If
        
        If Trim(cboShipVia.text) <> Trim(m_sDefaultShipMeth) Then
            UpdateShipRemarks "CustPrefs", "Cust.Pref.ShipMeth", cboShipVia.text, m_oCustomer.ID
            m_sDefaultShipMeth = cboShipVia.text
        End If

    Else
        If vbNo = MsgBox("Clear the default shipping method for " & m_oCustomer.ID & "?", _
                   vbYesNo + vbQuestion, "Confirm Default Shipping Method") Then
            chkDefaultShipMeth.value = vbChecked
            m_bLoading = bLoading
            Exit Sub
        End If
        DeleteShipRemarks "CustPrefs", "Cust.Pref.ShipMeth", m_oCustomer.ID
        m_sDefaultShipMeth = ""
    End If

    m_bLoading = bLoading
End Sub


'Show any/all UPS accounts for this customer

Private Sub cmdUPSUpdate_Click()
    Dim sNewUPSAcct As String
    Dim oFrm As FChangeUPSAcct
    Dim lNewAddrKey As Long
    
    Set oFrm = New FChangeUPSAcct
    sNewUPSAcct = oFrm.SearchUPSAcct(m_oCustomer.ID, m_oCustomer.Key, Trim(txtUPSAcct.text), m_oOrder.Customer.ShipAddr.AddrKey)
    
    If Trim(sNewUPSAcct) = "" Then Exit Sub
    
    SetWaitCursor True
    If vbYes = msg("Are you sure that you want to change UPS Account for this order?" _
                , vbYesNo + vbExclamation, "Change UPS Account") Then
            txtUPSAcct.text = Trim(sNewUPSAcct)
            m_oOrder.UPSAcct = txtUPSAcct.text
            chkBillRecipient.value = vbChecked
            chkBillRecipient.Enabled = True
    End If
    SetWaitCursor False
End Sub


'Called by
'   cmdNewOrder_Click
'   cmdLoadOrder_Click
'   FindOrderByNumber

Private Sub LoadUPSCtrl(ByVal CustKey As Long)
    
    txtUPSAcct.text = BillRecipientStat
    If txtUPSAcct.text <> "" Then
        chkBillRecipient.Enabled = True
        chkBillRecipient.value = vbChecked
    Else
        chkBillRecipient.Enabled = False
        chkBillRecipient.value = vbUnchecked
    End If
    
    Dim rst As Recordset
    Set rst = LoadDiscRst("Select distinct tcpUPSAcct.UPSAcct " & _
                            "from tarCustAddr inner join tcpUPSAcct on tcpUPSAcct.CustAddrKey = tarCustAddr.AddrKey " & _
                            "inner join tciAddress on tciAddress.Addrkey = tarCustAddr.AddrKey  " & _
                            "inner join tarCustomer on tarCustomer.CustKey = tarCustAddr.CustKey  " & _
                            "where tcpUPSAcct.UPSAcct <> '' and tarCustAddr.CustKey = " & CustKey & _
                            "and ((tciAddress.AddrKey = tarCustomer.DfltShipToAddrKey) or (tciAddress.AddrKey = " & m_oOrder.Customer.ShipAddr.AddrKey & "))")
    If rst.EOF Then
        ShowUPSControls False
    Else
        ShowUPSControls True
        '***HACK
        ' Several, if not most use cases were not persisting the UPSAcct number in
        ' the order object (or tcpSO).
        m_oOrder.UPSAcct = txtUPSAcct.text

        'Check the state of the cboShipVia control to determine whether we're in
        'edit or view mode. If it's enabled we're in edit mode.
        'If we're in edit mode & there's more than one choice, enable the update button
        If cboShipVia.Enabled Then
            cmdUPSUpdate.Enabled = rst.RecordCount > 1
        End If
    End If
    
    rst.Close
    Set rst = Nothing
End Sub


'Called by:
'   LoadUPSCtrl

Private Sub ShowUPSControls(ByVal showMe As Boolean)
    Label1(8).Visible = showMe
    txtUPSAcct.Visible = showMe
    cmdUPSUpdate.Visible = showMe
    chkBillRecipient.Visible = showMe
End Sub


'Change the shipping warehouse for this order

Private Sub cboWarehouse_Click(Index As Integer)
    On Error Resume Next
    If m_bLoading Then Exit Sub
    
    If Index = 2 Then 'Inventory frame on details panel
        UpdateInventoryInfo
    Else
        With cboWarehouse(Index)
            m_oOrder.WhseKey = .ItemData(.ListIndex)
            LoadShelfVendKey
            m_bLoading = True
            If Index = 1 Then
                If ViewMode = ivList Then gdxItems.Refetch
                cboWarehouse(0).ListIndex = cboWarehouse(1).ListIndex
            Else
                cboWarehouse(1).ListIndex = cboWarehouse(0).ListIndex
            End If
            m_bLoading = False
        End With
        
        UpdateOrderCaption
        
        m_bWareHouseLoading = True
        m_oOrder.ShipMethKey = RecalcShipVia(cboShipVia, g_rstShipVia, g_rstWhses, m_oOrder.WhseKey)
        m_bWareHouseLoading = False
    End If
    UpdateWarehouseColor Index
End Sub


'Get the latest shelf vendor key when ware house is changed for current order
Private Sub LoadShelfVendKey()
    Dim oItem As IItem
    Dim lShelfVendKey As Long
    
    lShelfVendKey = GetShelfVendKeyFromWhseKey(m_oOrder.WhseKey)
    
    For Each oItem In m_oItems
        If oItem.OPItemType = itWireShelf Then
            oItem.VendorKey = lShelfVendKey
        End If
    Next
End Sub


Private Sub chkShipComplete_Click()
    m_oOrder.ShipComplete = (chkShipComplete.value = vbChecked)
     
    If m_bLoading Then Exit Sub
    PromptShipComplete
    
    If tabMain.Tabs(tmiOrderStatus).Visible = True Then
        If Not m_oOrder.ShipComplete Then
            lblShipComplete.Visible = False
            lblShipComplete.Visible = True
            lblShipComplete.caption = "Ship Complete"
        End If
    End If
End Sub


Private Sub chkBillRecipient_Click()
    If m_bLoading Then Exit Sub
    
    Dim sTemp As String
    
    m_bLoading = True

    If chkBillRecipient.value = vbChecked Then
'        sTemp = BillRecipientStat
'        If Trim(sTemp) = "" Then
'            'Msg "This customer has not been set-up to support 'Bill Recipient' freight terms", vbOKOnly + vbExclamation, "Bill Recipient"
'            'cmdUPSUpdate.Enabled = False
'            'txtUPSAcct.Text = ""
'            'chkBillRecipient.Value = vbUnchecked
'        Else
'            'cmdUPSUpdate.Enabled = True
'            'txtUPSAcct.Text = Trim(sTemp)
'        End If
    Else
        cmdUPSUpdate.Enabled = True
        txtUPSAcct.text = ""
        chkBillRecipient.Enabled = False
    End If
    m_oOrder.UPSAcct = txtUPSAcct.text
    m_bLoading = False
End Sub


Private Sub txtPO_LostFocus()
    m_oOrder.PurchOrd = txtPO.text
    m_oBrokenRules.Validate txtPO
End Sub


'JJC: This function is never called because the property is read-only,
'but we left the logic here in case we ever change our mind
Private Sub chkReqPO_Click()
    m_oCustomer.ReqPO = (chkReqPO.value = vbChecked)
End Sub


'If this button is enabled, a credit card has already been assigned to the order.
'Except in the case of default terms = CrCard (there's nowhere to go but here)

Private Sub cmdCCEdit_Click()
    Dim oFrm As FCreditCardEditor
    
    Set oFrm = New FCreditCardEditor
    'pass in the currently selected credit card
    If oFrm.Init(m_oCustomer, m_oOrder.CreditCard, True, m_oOrder) = vbCancel Then
        'the user cancelled the editor
        If oFrm.SelCC Is Nothing Then
            'the user deleted the previously selected card
            m_oOrder.CreditCard = Nothing
            'restore terms to Customer default
            SetComboByText cboTerms, m_oOrder.Customer.BillAddr.DefaultPmtTerms.ID
            frmCreditCard.Visible = False
'TODO*** handle the default terms = crcard case
            cmdCCEdit.Visible = False
        Else
            UpdateCCDisplay  'not really necessary
        End If
    Else
        m_oOrder.CreditCard = oFrm.SelCC
        UpdateCCDisplay
    End If
    
    Unload oFrm
    Set oFrm = Nothing
End Sub



Private Sub cboTerms_Click()
    Dim oFrm As FCreditCardEditor
    
    If m_bLoading Then Exit Sub
    
    'suppress control event handlers which we're going to interact with
    m_bLoading = True
    
    With cboTerms
        'If PaymentTerms is COD or COD-CASH
        If Left$(UCase(.text), 3) = "COD" Then
            'prevent the user from specifying Bill Recipient
            DisableBillRecipient
            
            'if dropship is checked, don't allow COD. Throw a warning.
            If chkDropShip.value = vbChecked Then
                msg "You can't dropship to a COD customer."
                chkDropShip.value = vbUnchecked
                'I suspect I need to do these 4 things when clearing the dropship setting.
                '(see chkDropShip_Click)
                m_oOrder.IsDropShip = False
                
                'sbOrderStatus.Panels(7).Picture = Nothing
                sbOrderStatus.Panels(7).text = ""
                
                UndoMorphBTOItems
                gdxItems.Refetch
            End If
            
        Else
            If Left$(Trim$(cboShipVia.text), 3) = "UPS" Then
                If m_oOrder.Customer.HasAccount Then
                    chkBillRecipient.Enabled = True
                Else
                    DisableBillRecipient
                End If
            End If
        End If

        'Credit Card processing
        
        If .text = "CrCard" Then
        
            frmCreditCard.Visible = True
            cmdCCEdit.Visible = True

            '***if the order doesn't have a card assigned***
            If m_oOrder.CreditCard Is Nothing Then
                Set oFrm = New FCreditCardEditor
                If Not m_oOrder.Customer.HasAccount Then
                
                'force them into the visual editor to enter a CC
                    If oFrm.Init(m_oCustomer, Nothing, True, m_oOrder) = vbCancel Then
                        'if the CSR didn't enter a credit card
                        'restore terms to Customer default
                        SetComboByText cboTerms, m_oOrder.Customer.BillAddr.DefaultPmtTerms.ID

'??? 2/22/05 LR  shouldn't I change the Order's PmtTerms too?
'I'm doing that below

                        frmCreditCard.Visible = False
                        cmdCCEdit.Visible = False
                    Else
                        m_oOrder.CreditCard = oFrm.SelCC
                        UpdateCCDisplay
                    End If
                    
                'customer with account
                Else
                    'load the editor form as a class (hidden)
                    'get back the preferred credit card if there is one
                    If oFrm.Init(m_oCustomer, Nothing, False, m_oOrder) = vbCancel Then
                        'the customer has no credit cards on record
                        'show the editor
                                                   
                        If oFrm.Init(m_oCustomer, Nothing, True, m_oOrder) = vbCancel Then
                            'if the CSR didn't enter a credit card
                            'restore terms to Customer default
                            SetComboByText cboTerms, m_oOrder.Customer.BillAddr.DefaultPmtTerms.ID

'??? 2/22/05 LR  shouldn't I change the Order's PmtTerms too?
'I'm doing that below

                            frmCreditCard.Visible = False
                            cmdCCEdit.Visible = False
                        Else
                            m_oOrder.CreditCard = oFrm.SelCC
                            UpdateCCDisplay
                        End If
                    Else
                        m_oOrder.CreditCard = oFrm.SelCC
                        m_oOrder.CreditCard.Order = m_oOrder
                        UpdateCCDisplay
                    End If
                End If
                Unload oFrm
                Set oFrm = Nothing
            Else
                UpdateCCDisplay
            End If
        Else
            frmCreditCard.Visible = False
            cmdCCEdit.Visible = False
        End If

        'when closing the editor window, OP sometimes needs to be refreshed
        MDIMain.DoRefresh

        m_oOrder.PmtTerms.Key = .ItemData(.ListIndex)

    End With
        
    m_bLoading = False

End Sub


Private Sub UpdateCCDisplay()
    If m_oOrder.CreditCard Is Nothing Then
        lblCCType = vbNullString
        lblCCNo = vbNullString
        lblExpireDate = vbNullString
        lblHolderName = vbNullString
        lblCCStreet = vbNullString
        lblCCZipCode = vbNullString
        lblCCStatus = vbNullString
    Else
        With m_oOrder.CreditCard
            lblCCType = .TypeID
            
            'added 6/10/10 LR
            If HasRight(k_sRightARViewCCNo) Then
                lblCCNo = .CardNo
            Else
                lblCCNo = .MaskedCCNo
            End If
            
            lblExpireDate = .ExpireDate
            lblHolderName = .CardHolderName
            lblCCStreet = .StreetNbr
            lblCCZipCode = .ZipCode
            lblCCStatus = .Status
        End With
    End If
End Sub

'***********************************************************************************************
' DropShip Logic

' Question: what does it mean to UnMorph and Morph BTO Items?

' Toggle the dropship property of the order

Private Sub chkDropShip_Click()
    
    If m_bLoading Then Exit Sub
    
    Dim oItem As IItem
    Dim lDropShipVendKey As Long
    Dim lNonDropShipItemCount As Long
    
    m_oOrder.IsDropShip = (chkDropShip.value = vbChecked)
        
    'There are two exit points from this handler.
    'If the checkbox has been checked by the user, we drop out the bottom.
    'There are several points in the logic below that uncheck the checkbox.
    'This causes this event to fire again.
    'If the user or this routine has cleared the checkbox we exit from the Else clause below.
    'We want this event handler to clear the Order object's dropship property (see above),
    
    If m_oOrder.IsDropShip Then
        'sbOrderStatus.Panels(7).Picture = imgDrop.ListImages(1).Picture
        sbOrderStatus.Panels(7).text = "Drop Ship"
        'this is temporary
        'cmdManageDropShips.Enabled = True
        cmdManageDropShips.Enabled = False
    Else
        'sbOrderStatus.Panels(7).Picture = Nothing
        sbOrderStatus.Panels(7).text = ""
        cmdManageDropShips.Enabled = False
        UndoMorphBTOItems
        gdxItems.Refetch
        Exit Sub
    End If
    
    If Left(UCase(cboTerms.text), 3) = "COD" Then
        msg "You can't dropship to a COD customer."
        chkDropShip.value = vbUnchecked
    Else
        
        lDropShipVendKey = GetDropShipVendorKey(m_oOrder.whseid)
    
        If lDropShipVendKey = 0 And m_oOrder.Items.Count > 0 Then
            'we couldn't find a vendor to ship this
            msg "You can't Drop Ship this order because either" & vbCrLf & "there are no drop-shippable items or" & vbCrLf & _
            "the items come from multiple vendors.", vbExclamation + vbOKOnly, "Drop Ship Order"
            chkDropShip.value = vbUnchecked
        Else
            If HasManufacturedItems(m_oOrder.whseid) Then
                If vbOK <> msg("This order includes one or more items that can not be included" & vbCrLf _
                     & "on a drop ship order.  If you choose to continue, this order" & vbCrLf _
                     & " will be split and those items will be moved to the other order." & vbCrLf & vbCrLf _
                     & "Continue?", vbOKCancel, "Split Order?") Then
                    chkDropShip.value = vbUnchecked
                Else
                    m_oOrder.DropShipVendKey = lDropShipVendKey
                    SplitOrder m_oOrder.whseid
                End If
            Else
                '1/30/2014 LR
                'If the customer has requested 3rd party billing then display an alert
                If Len(txtUPSAcct.text) > 0 Then
                    MsgBox "Let the customer know that we can't use their UPS account if we're drop-shipping from the vendor.", vbOKOnly + vbInformation, "Drop Ship Alert"
                End If
        
                m_oOrder.DropShipVendKey = lDropShipVendKey
                MorphBTOItems
                gdxItems.Refetch
            End If
        End If
    End If
End Sub

' this returns 0 when
' 1. there are no lines on the order
' 2. there are no drop-shippable items on the order (only manufactured stuff)
' 3. there are items from more than 1 vendor
' else it returns the common vendor key (including STL shelves)

Private Function GetDropShipVendorKey(whseid As String) As Integer
    Dim oItem As IItem
    GetDropShipVendorKey = 0
    
    For Each oItem In m_oItems
        'if it's a drop-shippable item
        If Not (TypeOf oItem Is ItemGasket _
            Or TypeOf oItem Is ItemWWire _
            Or (TypeOf oItem Is ItemShelf And whseid <> "STL")) Then

            'assess its vendor
            If oItem.VendorKey <> 0 Then
                If GetDropShipVendorKey = 0 Then
                    GetDropShipVendorKey = oItem.VendorKey
                Else
                ' oops! one too many vendors
                    If GetDropShipVendorKey <> oItem.VendorKey Then
                        GetDropShipVendorKey = 0
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function HasManufacturedItems(whseid As String) As Boolean
    Dim oItem As IItem
    HasManufacturedItems = False
    
    For Each oItem In m_oItems
        If TypeOf oItem Is ItemGasket _
            Or TypeOf oItem Is ItemWWire _
            Or (TypeOf oItem Is ItemShelf And whseid <> "STL") Then
            HasManufacturedItems = True
            Exit Function
        End If
    Next
End Function

' This subroutine splits order after user changes it to drop ship order

Private Sub SplitOrder(whseid As String)
    Dim oFrm As FOrder
    Dim oItem As IItem
    Dim i As Long
    
    Set oFrm = CommandHandler.DoSplitOrder
    oFrm.chkDropShip.value = False
    
    MorphBTOItems
    
    With oFrm.Order.Items
        i = 1
        For Each oItem In m_oItems
            If TypeOf oItem Is ItemGasket _
            Or TypeOf oItem Is ItemWWire _
            Or TypeOf oItem Is ItemBTOKit _
            Or (TypeOf oItem Is ItemShelf And whseid <> "STL") Then
                .ImportItem oItem.Export, oFrm.Order.WhseKey
                m_oItems.Remove i
            Else
                i = i + 1
            End If
        Next
    End With
    
    oFrm.gdxItems.ItemCount = oFrm.Order.Items.Count
    oFrm.gdxItems.Refetch
    gdxItems.ItemCount = m_oItems.Count
    gdxItems.Refetch
    tabMain.SelectedTab = tmiLines
End Sub


'***********************************************************************************************



'Update Price Pack property for current order
Private Sub chkPricePackList_Click()
    m_oOrder.PricePackList = (chkPricePackList.value = vbChecked)
End Sub


'Get Order History
Private Sub GetOrderHistory()
    Dim sRst As ADODB.Recordset
    
    SetWaitCursor True
    
    Set sRst = CallSP("spcpcEventGetOrder", "@_iOPKey", m_oOrder.OPKey)
    
    With gdxOrderEvent
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = sRst
    End With
    
    SetWaitCursor False
End Sub


'Create RMA for Sage order
Private Sub cmdCreateRMA_Click()
    Dim oFrm As FRMACreate
    
    Set oFrm = New FRMACreate
    If oFrm.CreateRMA(m_oOrder) Then
        SetWaitCursor True
        
        cmdCreateRMA.Enabled = False
        m_oOrder.bRMA = True
        m_oOrder.Save
        
        m_lRMAKey = GetRMAKey
        LoadRMALine m_lRMAKey
        
        'turn on the RMA Line tab
        tabMain.Tabs(TabMainIndexes.tmiRmaLines).Visible = True
        SetCaption m_oCustomer.Name & "   OP " & m_oOrder.OPKey & " - RMA " & m_lRMAKey & "  " & cboWarehouse(1).text
        
        rvRMA.OwnerID = ""
        rvRMA.OwnerID = m_oOrder.OPKey
       
        sbOrderStatus.Panels(5).text = "RMA " & m_lRMAKey
        
'        sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(10).Picture
'        sbOrderStatus.Panels(1).text = StatusCode
        SetOrderStatusBar
    
        SetWaitCursor False
    End If
End Sub


'TODO: replace this with a Command returning a scalar value
'This function gets the RMA Key for current order
'The current design allows only on RMA record to be assigned to an order

Private Function GetRMAKey() As Long
    Dim rst As ADODB.Recordset
    
    Set rst = LoadDiscRst("Select * from tcpRMA where OPKey = " & m_oOrder.OPKey)
    If Not rst.EOF Then
        GetRMAKey = rst.Fields("RMAKey").value
        Set rst = Nothing
    Else
        Set rst = Nothing
        Err.Raise -1, , "There is no RMA info available for Order# " & m_oOrder.OPKey
    End If
End Function


Private Sub cmdSpecialHandling_Click()
    Dim oFrm As FSpecialHandling
    
    Set oFrm = New FSpecialHandling
    oFrm.ShipMethod = m_oOrder.ShipMethod
    oFrm.Load m_oOrder
    
    UpdatePartsNoCharge

    If m_oOrder.HasSpecialHandling Then
        sbOrderStatus.Panels(6).text = "Special Handling"
    Else
        sbOrderStatus.Panels(6).text = ""
    End If
    
End Sub


'Event handler for the Special Edit button

Private Sub cmdLiteEdit_Click()
    Dim oFrm As FLiteEdit
    
    Set oFrm = New FLiteEdit
    If oFrm.Load(m_oOrder) Then
        'if changes were made
        UpdateOPAfterLiteEdit
    End If
End Sub

'update the controls on the form based on changes made to the order object by FLiteEdit
'add logging to order history here

Private Sub UpdateOPAfterLiteEdit()
    m_bLoading = True
    With m_oOrder
        txtPO.text = .PurchOrd
        If .ShipComplete Then
            chkShipComplete.value = vbChecked
        Else
            chkShipComplete.value = vbUnchecked
        End If

        SetUpShipVia cboShipVia, .WhseKey, .ShipMethKey
        
        SetComboByKey cboTerms, .PmtTerms.Key
        
        txtUPSAcct.text = Trim(.UPSAcct)
        
        If txtUPSAcct.text <> "" Then
            chkBillRecipient.value = vbChecked
        Else
            chkBillRecipient.value = vbUnchecked
        End If
        
        txtShipToContact(0).text = .ShipToName
        txtShipToContact(1).text = .ShipToPhone
    End With
    m_bLoading = False
End Sub


'Update the Order after editing special handling of parts no charge

Private Sub UpdatePartsNoCharge()
    Dim lIndex As Long
    
    If m_oItems.Count > 0 Then
        If m_oOrder.HasPartsNoCharge Then
            For lIndex = 1 To m_oItems.Count
                m_oItems.Item(lIndex).BackNegotiatedPrice = m_oItems.Item(lIndex).EffectivePrice
                m_oItems.Item(lIndex).NegotiatedPrice = 0
            Next
        Else
'-----------------------------------------------------------------------------------------------------
'PRN 223 11/10/03 LR
'This is a major kludge. I don't fully understand this code.
'It's called every time the Special Handling dialog is closed.
'I think this Else clause is intended to restore the prices of parts that were set
'PartsNoCharge in a prior Special Handling edit cycle.
'I do know that for custom gaskets, NegotiatedPrice = -1 and BackNegotiatedPrice = 0
'if they haven't first been toggled through the True clause above.
'This assignment is what causes the custom gasket price to get set to 0.
'So I've simply added the guarding IF stmt to ensure that the cached negotiated price is > 0.
'-----------------------------------------------------------------------------------------------------
            For lIndex = 1 To m_oItems.Count
                If m_oItems.Item(lIndex).BackNegotiatedPrice > 0 Then
                    m_oItems.Item(lIndex).NegotiatedPrice = m_oItems.Item(lIndex).BackNegotiatedPrice
                End If
            Next
            
        End If
        gdxItems.Refetch
        txtTotalPrice.amount = m_oItems.TotalPrice
'7/25/05 LR
        txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)
    End If
End Sub


'************************************************************************************
'All events and subroutines on Line tab
'************************************************************************************

'Search for part number

Private Sub cmdSearch_Click()
    Dim lItemKey As Long
    Dim eItemType As ItemTypeCode
    Dim sOriginalItemID As String
    Dim sRefSource As String
    Dim frmItemSearch As FItemSearch
    Dim bCancelSearch As Boolean

    If Len(txtItemSearch.text) = 0 Then Exit Sub
    
    Set frmItemSearch = New FItemSearch
    Load frmItemSearch

'NOTE: xref logic in here

    If chkSearchItemDescr.value = vbChecked Then
        frmItemSearch.Find k_sPartDescr, txtItemSearch.text, lItemKey, eItemType, sOriginalItemID, sRefSource, bCancelSearch, cboWarehouse(1).ItemData(cboWarehouse(1).ListIndex), False, m_oOrder.Customer.CustTypeKey
    Else
        frmItemSearch.Find k_sPartNbr, txtItemSearch.text, lItemKey, eItemType, sOriginalItemID, sRefSource, bCancelSearch, cboWarehouse(1).ItemData(cboWarehouse(1).ListIndex), False, m_oOrder.Customer.CustTypeKey
    End If
    
    If bCancelSearch Then Exit Sub
    
    cmdPriceHistory.Visible = True

    'Morph BTO kit to SPO if this is a dropship order
    If eItemType = itBTOKit And m_oOrder.IsDropShip Then
        If VerifyBTOKit(lItemKey, sOriginalItemID, sRefSource) Then Exit Sub
        lItemKey = 0
    End If
        
    If lItemKey <> 0 Then
        'Check if the stocked item comes from the same vendor as DropShip Vendor
        'if this order is a dropship order
        
'NOTE: in the nominal case everything occurs in CheckandLoadItem
        If CheckandLoadItem(lItemKey, eItemType, sOriginalItemID, sRefSource) Then
        
            chkCGMPN.Visible = True
            m_oValidateItem.Valid = True
        End If
    Else
        m_oValidateItem.Valid = False
        m_oBrokenRules.Validate txtItemSearch
        TryToSetFocus txtItemSearch
    End If
End Sub


'Called By
'   cmdSearch_Click()
'
'Verify that the BTO kit can be added to the order

Private Function VerifyBTOKit(ByVal i_lItemKey As Long, ByVal i_sOriginalItemID As String, ByVal i_sRefSource As String) As Boolean
    Dim oBTOItem As ItemBTOKit
    Dim oDropShipBTOItem As ItemBTOKit
    Dim oSPOItem As IItem

    VerifyBTOKit = False
    
    If vbYes = msg("A matching part number was found but it refers to a" & vbCrLf _
               & "BTO Kit which cannot be included on a drop ship order." & vbCrLf _
               & "Would you like to load it and convert to SPO for dropship?", vbYesNo + vbQuestion + vbDefaultButton2, "Convert BTOKit to SPO for dropship?") Then
                
            'Check DropShip Vendor first before morphing BTO Kit to SPO item. 09/12/02 TeddyX
            Set oDropShipBTOItem = New ItemBTOKit
            oDropShipBTOItem.Load i_lItemKey
            
            If m_oOrder.DropShipVendKey > 0 And m_oOrder.DropShipVendKey <> oDropShipBTOItem.IItem_VendorKey Then
                g_rstVendors.Filter = "VendKey = " & m_oOrder.DropShipVendKey
    
                msg "This is a dropship order and the dropship vendor is " & RTrim(g_rstVendors.Fields("VendName").value) & ". " & _
                "You can't order an item from " & vbCrLf & "different vendor for dropship order. You can either order item(s) " & _
                "from dropship vendor or split this dropship order" & vbCrLf & " if you want to order items from different vendors."
    
                g_rstVendors.Filter = adFilterNone
                Set oDropShipBTOItem = Nothing
            Else
                Call cmdSpecifyItem_Click(btnSPO)
                m_oItems.SelectedItem.MorphBTOKey = i_lItemKey
                
                Set oBTOItem = MorphSPOToBTO(m_oItems.SelectedItem)
                Set oSPOItem = MorphBTOtoSPO(oBTOItem)
                
                m_oItems.Remove m_oItems.SelectedIndex
                
                oSPOItem.OriginalItemID = i_sOriginalItemID
                oSPOItem.RefSource = i_sRefSource
                m_oItems.Add oSPOItem
                
                m_oItems.Item(m_oItems.Count).Backup
                
                rvOrderLine(3).RemarkContext = m_oItems.SelectedItem.RemarkContext
                ItemUpdateControls
                txtItemSearch.text = ""
                chkSearchItemDescr.value = vbUnchecked
                Set oDropShipBTOItem = Nothing
                
                VerifyBTOKit = True
            End If
    End If
End Function


'2/2/05 LR created to consolidate control names
'replaces click handlers for cmdGasket, cmdShelf, cmdWire, cmdItemSpecial

Private Sub cmdSpecifyItem_Click(Index As Integer)
    Select Case Index
                                                                    
        'Create new Gasket Item
        Case btnGasket:
            cmdPriceHistory.Visible = False
            If m_oOrder.WhseKey = 24 Then
                msg "Sorry. The current order warehouse doesn't support gasket.", vbOKOnly + vbExclamation, "Can't add gasket"
            Else
                chkCGMPN.Visible = False
                cmdNextGasket.Visible = True
                cmdNextGasket.Enabled = False
                AddItem m_oItems.CreateItem(itMoldedGasket)
            End If
            
        'Create new Shelf item
        Case btnShelf:
            cmdPriceHistory.Visible = False
            chkCGMPN.Visible = False
            AddItem m_oItems.CreateItem(itWireShelf)
            
        'Create new Warmer Wire item
        Case btnWire:
            cmdPriceHistory.Visible = False
            g_rstWarmerWire.Filter = "WhseKey=" & m_oOrder.WhseKey
            If g_rstWarmerWire.EOF Then
                msg "Sorry. The current order warehouse doesn't support warmer wire.", vbOKOnly + vbExclamation, "Can't add warmer wire"
            Else
                chkCGMPN.Visible = False
                AddItem m_oItems.CreateItem(itWarmerWire)
            End If
            g_rstWarmerWire.Filter = adFilterNone

        'Create new SPO Item
        Case btnSPO:
            cmdPriceHistory.Visible = False
            chkCGMPN.Visible = True
            
            AddItem m_oItems.CreateItem(itSpecialOrder)
            
            'Set DropShipVendor for new added dropship order items. 09/12/02 TeddyX
            If m_oOrder.IsDropShip And m_oOrder.DropShipVendKey <> 0 Then
                SetVendorKey m_oOrder.DropShipVendKey
            End If

    End Select

End Sub


'Mark all line items as ready to commit
Private Sub cmdAuthorizeAll_Click()
    m_oItems.AuthorizeAll
    gdxItems.Refetch
    MDIMain.UpdateToolbarStatus
    SetOrderStatusBar

End Sub

Private Sub cmdUnAuthorize_Click()
    m_oItems.UnAuthorizeAll
    gdxItems.Refetch
    MDIMain.UpdateToolbarStatus
    SetOrderStatusBar
End Sub

'Load Inventory Finder for the current selected item
Private Sub cmdInvFinder_Click()
    Dim oFrm As FInvFinder
    
    Set oFrm = New FInvFinder
    MDIMain.AddNewWindow oFrm
    oFrm.LoadItem m_oItems.SelectedItem, m_oOrder.WhseKey
End Sub


'Cancel editing item. If it's new item, delete it. Otherwise, return back to List mode
Private Sub cmdItemCancel_Click()

    'Bit of a kludge, added 2/26/15 LR
    If m_oItems.SelectedItem.OPItemType = itSpecialOrder Then
        If m_oItems.SelectedItem.MakeKey > 0 Or Len(m_oItems.SelectedItem.ModelNbr) Or Len(m_oItems.SelectedItem.SerialNbr) Then
            If MsgBox("You've entered Make/Model/SerialNbr info. Sure you want to Cancel?", vbExclamation + vbYesNo, "Cancelling Item") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    If m_bNewItem Then
        cmdItemDelete_Click
    Else
        m_oItems.SelectedItem.Restore
        ViewMode = ivList
    End If
    
    SetOrderStatusBar
End Sub


'Delete selected item from order
Private Sub cmdItemDelete_Click()
    If m_oOrder.IsDropShip And m_oOrder.DropShipVendKey > 0 And CheckDropShipSelectedItem Then
        m_oOrder.DropShipVendKey = 0
    End If
    m_oItems.Remove m_oItems.SelectedIndex
    SyncItemList
    
    ViewMode = ivList
    
    SetOrderStatusBar
End Sub


'Save changes after editing selected item

Private Sub cmdItemOK_Click()
    Dim bSave As Boolean
    
    ' If item is NOT active, do not allow order more than Qty Available
    If Not m_oItems.SelectedItem Is Nothing Then
        Select Case m_oItems.SelectedItem.ItemInventoryStatus
            Case iisInactive, iisDiscontinued, iisDeleted
                If IsNumeric(txtQtyOrdered.text) And IsNumeric(lblQtyAvail.caption) Then
                    If CLng(txtQtyOrdered.text) > CLng(lblQtyAvail.caption) Then
                        MsgBox "This part is no longer available for purchase. Please limit Quantity to " & lblQtyAvail.caption & " or less.", vbInformation, caption
                        TryToSetFocus txtQtyOrdered
                        Exit Sub
                    End If
                End If
        End Select
    End If

    Select Case ViewMode
        Case ivComponent, ivKit
        'if the Order tab's Ship From Warehouse is not the same as the Warehouse on the summary Line tab
            If cboWarehouse(2).ListIndex <> cboWarehouse(1).ListIndex Then
                If vbYes = msg("Would you like to change the warehouse processing this order from " _
                         & cboWarehouse(1).text & " to " & cboWarehouse(2).text & "?", vbYesNo, "Change Warehouse?") Then
                    cboWarehouse(1).ListIndex = cboWarehouse(2).ListIndex
                End If
            End If
    End Select
    
    '9/12/02 TX  Assign DropShipVendKey for DropShip order
    If m_oOrder.IsDropShip Then
        If CheckDropShipSelectedItem Then
            m_oOrder.DropShipVendKey = m_oItems.SelectedItem.VendorKey
        End If
    End If
    
    AddLineItemsRemark
    ViewMode = ivList
    PromptShipComplete
    
    gdxItems.Refetch
    
    '7/31/2012 Added to get the grid to repaint if there are more items than grid will show at once - LR
    ForceRefresh Me.hwnd
End Sub


'Continue editing the next gasket
Private Sub cmdNextGasket_Click()
'***GMOD
    'm_oItems.SelectedItem.StatusCode = ItemStatusCode.iscReadyToCommit
    'cache and restore these (override ItemGasket.Class_Initialize())
    Dim prevMaterialId As Long
    Dim prevIsMagnetic As Boolean
    prevMaterialId = m_oGasket.materialId
    prevIsMagnetic = m_oGasket.IsMagnetic
'***

    gdxItems.Refetch
    PromptShipComplete
    
    SetOrderStatusBar

'***GMOD
    Dim newItem As ItemGasket
    Set newItem = m_oItems.CreateItem(itMoldedGasket)
    newItem.materialId = prevMaterialId
    newItem.IsMagnetic = prevIsMagnetic
    AddItem newItem
'***
End Sub


'Display vendor information for selected item
Private Sub cmdVendorDetails_Click(Index As Integer)
'2/3/05 LR replaced
'    FVendor.DisplayInfo cboVendor.ItemData(cboVendor.ListIndex)
    DisplayVendorInfo cboVendor.ItemData(cboVendor.ListIndex)
End Sub


'View catlog information for selected items.

Private Sub cmdViewCat_Click(Index As Integer)
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    With m_oItems.SelectedItem
        ViewCatPage .ItemID, .CustType
    End With
End Sub


'This subroutine views the Catalog page for the selected item based on customer type

Private Sub ViewCatPage(ByVal i_sItemID As String, ByVal i_sCustType As String)
    Dim frmCatPage As FCatPage
    
    Set frmCatPage = New FCatPage
    MDIMain.AddNewWindow frmCatPage
    With frmCatPage
        .Show
        .PartNo = i_sItemID
        Select Case i_sCustType
            Case "EndUser"
                .CustType = 1
            Case Is = "Dealer"
                .CustType = 2
            Case Is = "Wholesale"
                .CustType = 3
            Case Else
                .CustType = 1
        End Select
        .ShowPage
    End With
End Sub


'Set gasket type for gasket item

Private Sub optGasketType_Click(Index As Integer)
    With m_oGasket
        If .IsMagnetic <> CBool(Index = 0) Then
            .IsMagnetic = (Index = 0)
            .materialId = 0 'material type is undefined after type change
            .LoadMaterialCombo cboGasket
            m_oBrokenRules.Validate cboGasket
        End If
        
        '8/14/13 MOD to change max height for compression gaskets
        'this kludge overcomes how the control's Value setter & internal validation works
        Dim cacheheight As Integer
        cacheheight = lenGasket(0).value
        lenGasket(0).value = 0
        lenGasket(0).value = cacheheight
        m_oBrokenRules.Validate lenGasket(0)
        '***
        
    End With

End Sub


'Get gasket materials

Private Sub cboGasket_Click()
    m_oGasket.materialId = cboGasket.ItemData(cboGasket.ListIndex)
End Sub


'Get Gasket Options

Private Sub chkGasketOptions_Click(Index As Integer)
    Dim lOptions As Long

    If m_bLoading Then Exit Sub

'    Dim lOptions As Long
'    Dim i As Long
'    'loop through all the options each time one changes
'    For i = 3 To 0 Step -1
'        lOptions = lOptions * 2 'logical shift left
'        If chkGasketOptions(i).value = vbChecked Then
'            lOptions = lOptions + 1 'logically "OR" in bit if checked
'        End If
'    Next
    
    If optGasketSides(0).value = True Then
        'clear three-sided
         lOptions = lOptions And Not k_lGasketThreeSided
    Else
        'set three-sided
        lOptions = lOptions Or k_lGasketThreeSided
    End If
    
    If chkGasketOptions(0).value = vbChecked Then
        lOptions = lOptions Or k_lGasketDartToDart
    End If
    
    If chkGasketOptions(1).value = vbChecked Then
        lOptions = lOptions Or k_lGasketInverted
    End If
    
    If chkGasketOptions(2).value = vbChecked Then
        lOptions = lOptions Or k_lGasketNoMagLHS
    End If
    
    If chkGasketOptions(3).value = vbChecked Then
        lOptions = lOptions Or k_lGasketNoMagRHS
    End If
    
    m_oGasket.Options = lOptions
    
End Sub


Private Sub optGasketSides_Click(Index As Integer)
    If optGasketSides(0).value = True Then
        'clear three-sided
         m_oGasket.Options = m_oGasket.Options And Not k_lGasketThreeSided
    Else
        'set three-sided
        m_oGasket.Options = m_oGasket.Options Or k_lGasketThreeSided
    End If
End Sub


'Set Gasket Height and width
'index 0 - height
'index 1 - width

Private Sub lenGasket_Change(Index As Integer, ByVal i_bValid As Boolean)
    Select Case Index
        Case 0:
            '8/14/13 MOD to change max height for compression gaskets
            If m_oGasket.IsMagnetic Then
                lenGasket(0).MaxValue = g_MaxHeightMagnetic
            Else
                lenGasket(0).MaxValue = g_MaxHeightCompression
            End If
            '***
            m_oBrokenRules.Validate lenGasket(0)
            m_oGasket.Height = lenGasket(0).value
        Case 1:
            m_oBrokenRules.Validate lenGasket(1)
            m_oGasket.width = lenGasket(1).value
    End Select
End Sub


'Set Frame for shelf item
Private Sub cboFrame_Click()
    With cboFrame
        'Only FrameID forces refresh so call it AFTER FrameText
        m_oShelf.FrameText = .text
        m_oShelf.FrameID = .ItemData(.ListIndex)
    End With
End Sub


'Set finish property for shelf item
Private Sub cboFinish_Click()
    With cboFinish
        'Only FinishID forces refresh so call it AFTER FinishText
        m_oShelf.FinishText = .text
        m_oShelf.FinishID = .ItemData(.ListIndex)
    End With
End Sub


'Set shelf width
Private Sub lenShelfWidth_Change(ByVal i_bValid As Boolean)
    m_oBrokenRules.Validate lenShelfWidth
    m_oShelf.width = lenShelfWidth.value
End Sub


'Set shelf depth
Private Sub lenShelfDepth_Change(ByVal i_bValid As Boolean)
    m_oBrokenRules.Validate lenShelfDepth
    m_oShelf.Depth = lenShelfDepth.value
End Sub


'Set Wire Length
Private Sub lenWireLength_Change(ByVal i_bValid As Boolean)
    m_oBrokenRules.Validate lenWireLength
    m_oWarmerWire.TotalInches = lenWireLength.value
End Sub


'Set door height for wire length
Private Sub lenDoorHeight_Change(ByVal i_bValid As Boolean)
    m_oBrokenRules.Validate lenDoorHeight
    m_oWarmerWire.DoorHeight = lenDoorHeight.value
End Sub


'set door width for warmer wire
Private Sub lenDoorWidth_Change(ByVal i_bValid As Boolean)
    m_oBrokenRules.Validate lenDoorWidth
    m_oWarmerWire.DoorWidth = lenDoorWidth.value
End Sub


'Set shelf options
Private Sub chkShelfOpt_Click(Index As Integer)
    Dim lOptions As Long
    Dim i As Long

    If Not m_bLoading Then
        'loop through all the options each time one changes
        For i = 4 To 0 Step -1
            lOptions = lOptions * 2 'logical shift left
            If chkShelfOpt(i).value = vbChecked Then
                lOptions = lOptions + 1 'logically "OR" in bit if checked
            End If
        Next
        m_oShelf.Options = lOptions
    End If
End Sub


'Set wire passes property
Private Sub optWirePasses_Click(Index As Integer)
    If m_bLoading Then Exit Sub

    m_oWarmerWire.IsSinglePass = (Index = 0)
End Sub


'Set Door style for wire
Private Sub optDoorStyle_Click(Index As Integer)
    If m_bLoading Then Exit Sub
    
    m_oWarmerWire.IsThreeSided = (Index = 1)
End Sub


'Set voltage from combo box for wire
Private Sub cboVoltage_Change()
    Dim lVoltage As Long
    On Error Resume Next 'ignore error for invalid numeric value
    lVoltage = CLng(cboVoltage.text)

'***DH 6/25/09
    If lVoltage = 230 Then
        If MsgBox("Warning! 230V is potentially very dangerous and not recommend. Continue?", _
                    vbExclamation + vbYesNo, "Voltage warning.") = vbNo Then
            'Suppress the click event.
            m_bLoading = True
            cboVoltage.ListIndex = 1
            m_bLoading = False
            lVoltage = CLng(cboVoltage.text)
        End If
    End If
    
    m_oWarmerWire.Voltage = lVoltage
    
End Sub


'Click on Voltage combo box to get the voltage for wire
Private Sub cboVoltage_Click()
'***DH 6/25/09
    If m_bLoading Then Exit Sub
    cboVoltage_Change
End Sub


'If user presses invalid key, let the input value is 0
Private Sub cboVoltage_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack
        Exit Sub
    Case Is < vbKey0, Is > vbKey9
        KeyAscii = 0
    End Select
End Sub

Private Sub cmdPriceHistory_Click()
    FPriceHistory.ShowPriceHistory txtItemPartNbr.text, m_oOrder.Customer.CustTypeKey
End Sub


'************************************************************************************
'All events and subroutines on Order Status tab
'************************************************************************************

'This button refresh Order Status tab
Private Sub cmdOSRefresh_Click()
    LoadOrderStatus m_oOrder.OPKey
End Sub


'************************************************************************************
' Order Status Grid Events
'   gdxOSLine
'   gdxOSLineItems
'   gdxOSShipments
'   gdxOSShipItems
'   gdxOSInvoice
'   gdxInvoiceItem
'************************************************************************************

'Change the event from double click to single click
Private Sub gdxOSLine_Click()
    UpdateOSLineItem
End Sub


'When User presses Up and Down key, the lower grid is updated accordingly
Private Sub gdxOSLine_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        UpdateOSLineItem
    End If
End Sub


Private Sub gdxOSLine_SelectionChange()
    UpdateOSLineItem
End Sub


'Private Sub gdxOSLine_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
'    gdxOSLine.Update
'End Sub


'Load data from OS Item List to gdxOSLine grid
Private Sub gdxOSLine_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_oOSItemList Is Nothing Then Exit Sub
    If RowIndex > m_oOSItemList.Count Then Exit Sub
    
    With m_oOSItemList.Item(RowIndex)
        Values(1) = GetLineStatus(.Status)
        Values(2) = .ItemID
        Values(3) = .Description
        Values(4) = .QtyOrdered
        Values(5) = .UnitPrice
        Values(6) = .ItemKey
        Values(7) = .SOLineKey
        Values(8) = .QtyOpenToShip
        Values(9) = .QtyInvcd
    End With
End Sub


'Load Line Items PO remarks in double click event
Private Sub gdxOSLineItems_DblClick()
    Dim oRC As RemarkContext
    
    Set oRC = New RemarkContext
    oRC.Edit "ViewPO", m_gwOSLineItems.value("PONumber")
End Sub


'When User clicks on shipment grid, the lower grid is updated accordingly.
Private Sub gdxOSShipments_Click()
    UpdateShipmentItem
End Sub

'Private Sub gdxOSShipments_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    'Dim bTemp As Boolean
'    If Button = vbRightButton Then
'        'bTemp = gdxOSShipments.GroupByBoxVisible
'        'mnuOrderGroup.Checked = bTemp
'        'mnuOrderExpand.Enabled = bTemp
'        'mnuOrderCollapse.Enabled = bTemp
'        Me.PopupMenu mnugdxOSShipment
'    End If
'End Sub

Private Sub gdxOSShipItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnugdxOSShipItems
    End If
End Sub

'If user wants to copy the ShipTracking number from the grid.
'He/she can choose the row where the ShipTracking number is located in
'the grid, use 'Ctrl + C' to copy the ShipTracking number to Clipboard.
'Then user can use 'Ctrl + P' to paste the ShipTracking number to where he wants.

'Private Sub gdxOSShipments_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not (gdxOSShipments.RowCount > 0) Then Exit Sub
'
'    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
'        Clipboard.Clear
'        Clipboard.SetText (Trim(m_gwShipments.value("ShipTrackNo")))
'    End If
'End Sub

Private Sub gdxOSShipItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (gdxOSShipItems.RowCount > 0) Then Exit Sub
    
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        Clipboard.Clear
        Clipboard.SetText (Trim(m_gwShipItems.value("ShipTrackNo")))
    End If
End Sub


'When User presses Up and Down key, the lower grid is updated accordingly
Private Sub gdxOSShipments_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        UpdateShipmentItem
    End If
End Sub


'Single click event updates the lower grid accordingly
Private Sub gdxOSInvoice_Click()
    UpdateInvoiceItem
End Sub


'When user presses up and down key, the lower grid is updated accordingly
Private Sub gdxOSInvoice_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        UpdateInvoiceItem
    End If
End Sub


'This subroutines updates the lower grid of Line Items tab in Order Status
Private Sub UpdateOSLineItem()
    Dim sSQL As String
    Dim rst As ADODB.Recordset
    
    If gdxOSLine.value(6) = Empty Then Exit Sub
    
    SetWaitCursor True

    'Mod to display the dropship's PO order status correctly (10/28/02 TX)
    If m_oOrder.IsDropShip Then
        If gdxOSLine.value(7) = Empty Then
            SetWaitCursor False
            Exit Sub
        End If

        Set rst = CallSP("spOPOrdStatGetDropShipSpoPOs", "@i_SOLineKey", gdxOSLine.value(7))
    Else
        If Left(gdxOSLine.value(2), 4) = "SPO-" Or Left(gdxOSLine.value(2), 4) = "SHF-" Then
            If gdxOSLine.value(7) = Empty Then
                SetWaitCursor False
                Exit Sub
            End If
            Set rst = CallSP("spOPOrdStatGetSpoPOs1", "@i_SOLineKey", gdxOSLine.value(7))
        Else
            Set rst = CallSP("spOPOrdStatGetStkPOs1", "@i_ItemKey", gdxOSLine.value(6))
        End If
    End If
    
    With gdxOSLineItems
        If m_oOrder.IsDropShip Then
            .Columns("Status").Visible = True
        Else
            .Columns("Status").Visible = False
        End If
        .HoldFields
        Set .ADORecordset = rst
    End With
    
    lblOSLGridCaption(0).caption = "Open Purchase Order(s) for Part Number " & gdxOSLine.value(2) & ":"
    SetWaitCursor False
    Set rst = Nothing
End Sub


'This subroutine updates the lower grid of shipment tab
Private Sub UpdateShipmentItem()
    Dim lShipKey As Long
    Dim rst As ADODB.Recordset
    
    If IsNull(m_gwShipments.value("Shipkey")) Then Exit Sub
    
    SetWaitCursor True
    lShipKey = m_gwShipments.value("ShipKey")
    
    If m_gwShipments.value("TranNo") = "Provisional" Then
        Set rst = CallSP("spOPOrdStatShipDtl1", "@i_ShipKey", lShipKey, "@b_Provisional", True)
    Else
        Set rst = CallSP("spOPOrdStatShipDtl1", "@i_ShipKey", lShipKey)
    End If

    gdxOSShipItems.HoldFields
    
    lblOSLGridCaption(1).caption = "Item(s) contained on Shipment " & m_gwShipments.value("TranNo") & " dated " & Format(m_gwShipments.value("ShipDate"), "MM/DD/YY") & ":"
    Set gdxOSShipItems.ADORecordset = rst
    Set rst = Nothing
    SetWaitCursor False
End Sub


'This subroutine updates the lower grid of Invoice tab
Private Sub UpdateInvoiceItem()
    Dim lInvcKey As Long
    Dim rst As ADODB.Recordset
    
    If IsNull(m_gwInvoice.value("invckey")) Then Exit Sub
    
    SetWaitCursor True
    lInvcKey = m_gwInvoice.value("InvcKey")
    Set rst = CallSP("spOPOrdStatGetInvDtl1", "@i_InvcKey", lInvcKey)
    gdxOSInvoiceItem.HoldFields
    
    lblOSLGridCaption(2).caption = "Item(s) Invoiced on " & m_gwInvoice.value("TranID") & " dated " & Format(m_gwInvoice.value("InvoiceDate"), "MM/DD/YY") & ":"
    
    Set gdxOSInvoiceItem.ADORecordset = rst
    Set rst = Nothing
    SetWaitCursor False
End Sub


'This is the nested tab control
'Add this event here to set focus to upper grid

Private Sub SSAOrderDetails_TabClick(ByVal NewTab As ActiveTabs.SSTab)
    Select Case NewTab.Index
        Case tosLineItem:
            If gdxOSLine.Row = 1 Then
                UpdateOSLineItem
            End If
        Case tosShipment:
            If gdxOSShipments.Row = 1 Then
                UpdateShipmentItem
            End If
        Case tosInvoice:
            If gdxOSInvoice.Row = 1 Then
                UpdateInvoiceItem
            End If
    End Select
End Sub


'This subroutine loads header information for Order Status
Private Sub LoadHeadInfo(ByRef oRstOS As ADODB.Recordset)
    With oRstOS
        lblOS(0).caption = .Fields("CustID").value
        lblOS(1).caption = .Fields("CustName").value
        lblOS(6).caption = .Fields("FirstName").value + " " + .Fields("LastName").value
        
        lblOS(4).caption = .Fields("UserID").value
        lblOS(3).caption = .Fields("OPKey").value
        lblOS(5).caption = .Fields("SOID").value
        lblOS(7).caption = .Fields("CreateDate").value
        If .Fields("ShipComplete") = 0 Then
            lblShipComplete.Visible = False
        Else
            lblShipComplete.Visible = True
            lblShipComplete.caption = "Ship Complete"
        End If
        lblOS(2).caption = .Fields("PurchOrd").value
    End With
End Sub


'This subroutine loads info for all Order Status tab
Private Sub LoadOrderStatus(ByVal lOPKey As Long)
    Dim oRstOrderStatus As ADODB.Recordset

    Set oRstOrderStatus = CallSP("spcpcGetOPStatusHeader", "@i_OPKey", lOPKey)
    
    SetOSTabsVisible
    ClearOSGrids
    
    If Not oRstOrderStatus.EOF Then
        LoadHeadInfo oRstOrderStatus
        LoadListStatus
        LoadShipments
        LoadInvoice
    End If
    
    Set oRstOrderStatus = Nothing
End Sub


'This subroutine loads shipment info for Sage order

Private Sub LoadShipments()
    Dim rstShipments As ADODB.Recordset
    Dim i As Integer
    
    Set rstShipments = CallSP("spOPOrdStatGetShipment", "@_iOPKey", m_oOrder.OPKey)
    
    If Not rstShipments.EOF Then
        SSAOrderDetails.Tabs(tosShipment).Visible = True
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
    Else

        'Make the tab of shipment invisible if there are no shipment
        SSAOrderDetails.Tabs(tosShipment).Visible = False
    End If
End Sub


' Load Invoice information for Sage Order
' Grid: gdxOSInvoice

Private Sub LoadInvoice()
    Dim rstInvoice As ADODB.Recordset
    Dim i As Integer
    
    Set rstInvoice = CallSP("spOPOrdStatGetInvoice", "@_iOPKey", m_oOrder.OPKey)

    If Not rstInvoice.EOF Then
        SSAOrderDetails.Tabs(tosInvoice).Visible = True
       With gdxOSInvoice
            .HoldFields
            Set .ADORecordset = rstInvoice
            
            For i = 1 To .Columns.Count
                .Columns(i).AutoSize
            Next
       End With
    Else
        'Make the tab of invoice invisible if there has no Invoice info
        SSAOrderDetails.Tabs(tosInvoice).Visible = False
    End If
End Sub


'This subroutine loads list items to the grid
Private Sub LoadListStatus()
    Set m_oOSItemList = New OSItemList
    
    m_oOSItemList.Load m_oOrder.OPKey, m_oOrder.IsDropShip
    If m_oOSItemList.Count > 0 Then
        With gdxOSLine
            .HoldFields
            .ItemCount = m_oOSItemList.Count
            .Refetch
        End With
    End If
    
    SSAOrderDetails.Tabs(tosLineItem).Selected = True
End Sub


'Clear grids on the Order Status tab for future updating
Private Sub ClearOSGrids()
    gdxOSShipments.HoldFields
    gdxOSInvoice.HoldFields
    gdxOSLineItems.HoldFields
    gdxOSShipItems.HoldFields
    gdxOSInvoiceItem.HoldFields
    Set gdxOSLineItems.ADORecordset = Nothing
    Set gdxOSShipItems.ADORecordset = Nothing
    Set gdxOSInvoiceItem.ADORecordset = Nothing
    Set gdxOSShipments.ADORecordset = Nothing
    Set gdxOSInvoice.ADORecordset = Nothing
End Sub


'*************************************************************************************
'Toolbar commands
'These are custom methods of the form.
'They're invoked by MDIMain.
'*************************************************************************************

Public Function DeleteButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean

    On Error GoTo ErrorHandler

    If Not i_bDoIt Then
    
        If Not FindMode And m_oOrder.StatusCode <> ItemStatusCode.iscDeleted And m_oOrder.StatusCode <> ItemStatusCode.iscHasRMA And m_oOrder.StatusCode <> ItemStatusCode.iscCommitted Then
            DeleteButton = True
        End If

    Else
    
        SetWaitCursor True
        
        If vbYes = msg("Delete " & Me.caption & "? Are you sure?", vbYesNo, "Confirm Delete") Then
            With m_oOrder
                .Delete
                .Save
            End With
            m_oOrder.Clear False
            
            cmdNewSearch_Click  'reset Existing Customer tab
            FindOrderByNumber
            TransitionTabs True, sCaption & "was deleted from OrderPad"
        Else
            TransitionTabs True, sCaption & "was not deleted from OrderPad"
        End If
    
        SetWaitCursor False

    End If

    Exit Function

ErrorHandler:
    ClearWaitCursor
    ErrorUI.DisplayWarning "Delete Failed"
End Function


Public Function PrintButton(ByVal i_bPrintOnly As Boolean, Optional ByVal i_bDoIt As Boolean = True) As Boolean
    Dim oFrm As FViewOrder

    If FindMode Then
        PrintButton = False
    Else
        PrintButton = True
        If i_bDoIt Then
            m_oOrder.Save
            Set oFrm = New FViewOrder
            oFrm.ShowOrder m_oOrder, i_bPrintOnly
        End If
    End If
End Function


'Called By:
'MDIMain.UpdateToolbarStatus        False
'CommandHandler.DoSave              True
'FOrder.CancelButton                True
'FOrder.ExitCheck                   True
'FOrder.Form_Unload                 True


Public Function SaveButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean
                           
    On Error GoTo ErrorHandler

    If Not CanSaveOrder Then Exit Function

    SaveButton = True
    
    If i_bDoIt Then
        ForceLostFocus
        
        'If HasPartOnMultipleLines("Save") Then Exit Function
    
        If m_oOrder.StatusCode = ItemStatusCode.iscHasRMA Then
            If vbCancel = msg("Are you sure you want to make those changes for this RMA order in OA?", vbExclamation + vbOKCancel, "Saving Changes for RMA Order?") Then
                Exit Function
            End If
        ElseIf m_oOrder.soKey <> 0 Then
            If vbCancel = msg("This order has already been committed." & vbCrLf & "Are you sure you want to save your remark changes?", vbExclamation + vbOKCancel, "Saving Changes for Committed Order?") Then
                Exit Function
            End If
        End If

        If m_oCustomer.ReqPO And Trim(txtPO.text) = "" Then
            txtPO.text = "Temp"
            m_oBrokenRules.Validate txtPO
            txtPO.text = ""
        End If
        
        'If there are validation issues
        If m_oBrokenRules.Count > 0 Then
        
            With m_oBrokenRules
                Dim oCtrl As Control
                Dim lRuleIndex As Long

                lRuleIndex = .FindBrokenRuleIndex(Me.ActiveControl)
                If lRuleIndex <= 0 Then lRuleIndex = 1
                .ForceFocus .Item(lRuleIndex).Rule.CtlWrapper.Ctl, .Item(lRuleIndex).ErrorMsg
                
                Set oCtrl = Me.ActiveControl
                msg .Item(lRuleIndex).ErrorMsg, , "Correct the following error, then try again"
                oCtrl.SetFocus
            End With
            
            'set the ReqPO back
            If m_oCustomer.ReqPO And Trim(txtPO.text) = "" Then
                m_oBrokenRules.Validate txtPO
            End If
        
        'else all looks good
        Else
            'What's this all about?
            'Set the ReqPO back before saving in case saving does not succeed
            If m_oCustomer.ReqPO And Trim(txtPO.text) = "" Then
                m_oBrokenRules.Validate txtPO
            End If
            
            With m_oOrder
            
                'If the order's status code is Research and there is more than one
                'item, the user has to decide how to set the order's research status.
                
                If m_oOrder.StatusCode = iscResearch Then
                    Dim oFrm As FResearchStatus
                    Set oFrm = New FResearchStatus
                    If Not oFrm.LoadOrderResearchStatus(m_oOrder) Then
                        SaveButton = False
                        Set oFrm = Nothing
                        Exit Function
                    End If
                    Set oFrm = Nothing
                End If
                
                ' If saving fails then exit the entire save process

                ' If this is being saved by a WillCall user then
                ' check the order's ship method.
                ' It should be "WillCall"
                ' This method is Warehouse specific.
                
                If g_bWillCallUser And Not ShipMethIsWillCall Then
                    If vbYes = msg("This order will ship " & m_oOrder.ShipMethod & ". Do you want to change it to WillCall?", vbYesNo, "OrderPad") Then
                        m_oOrder.ShipMethKey = WhseWillCallShipMethKey
                    End If
                End If
                
                If Not .Save(, (m_oOrder.soKey > 0)) Then
                    SaveButton = False
                    Exit Function
                End If

                ' (for testing & debugging) Write out a snapshot of the order
                ' SaveOrderAsXML m_oOrder
                
                m_bRecommit = False
            End With
            

            If m_oOrder.soKey <> 0 Or (g_bWillCallUser) Then
                DoEvents
                If m_lCustKey = 0 Then
                    'reset Existing Customer tab
                    cmdNewSearch_Click
                Else
                    FillSelectCustTab m_lCustKey, False
                End If
                'refresh existing orders tab
                FindOrdersByCriteria
                TransitionTabs True
            End If
            
        End If
    End If
    Exit Function
    
ErrorHandler:
    If m_oCustomer.ReqPO And Trim(txtPO.text) = "" Then
        m_oBrokenRules.Validate txtPO
    End If
    ErrorUI.DisplayWarning "Save Failed"
End Function

'DECIDED NOT TO USE
'*********************************************************************************************
' This validation function is used by both SaveButton & CommitButton
' action = "Save" or "Commit" and is used to create the MessageBox

' Prevent an order being saved (and/or committed) with more than one line for the same item

Private Function HasPartOnMultipleLines(action As String) As Boolean
    Dim oItem As IItem
    
    HasPartOnMultipleLines = False
    
    For Each oItem In m_oItems
        If CountItemInList(oItem) > 1 Then
            msg "Part# " & oItem.ItemID & " is on more than one order line." & vbCrLf & _
                    "You need to consolidate this into one line before you can " & action & " your order.", vbExclamation + vbOKCancel, "Order " & action
           HasPartOnMultipleLines = True
           Exit For
        End If
    Next
End Function

Private Function CountItemInList(Item As IItem) As Integer
    Dim oItem As IItem
    For Each oItem In m_oItems
        If Item.ItemID = oItem.ItemID Then 'and is FinGood or BTOKit
            CountItemInList = CountItemInList + 1
        End If
    Next
End Function

'*********************************************************************************************

'Determine whether the current ship method is willcall or not.

Private Function ShipMethIsWillCall() As Boolean
    If InStr(1, m_oOrder.ShipMethod, "Will Call") Then
        ShipMethIsWillCall = True
    Else
        ShipMethIsWillCall = False
    End If
End Function


'Return the key for the WillCall ShipMethod for the User's warehouse

Private Function WhseWillCallShipMethKey() As Long
    Dim sWillCallShipMethID As String
    sWillCallShipMethID = GetUserWhseID & "-Will Call"
    WhseWillCallShipMethKey = ShipMethIDtoKey(sWillCallShipMethID)
End Function


'**************************************************************************************
' MDI Tool Bar button handlers (Called by MDIMain)
'**************************************************************************************

Public Function SplitOrderButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean
    If Not CanSaveOrder Then Exit Function
    
    SplitOrderButton = Not (m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit)
    
    'TODO: The DoSplitOrder function in MDIMain should be moved here.
    '      The code works correctly as is, but violates the design
    '      concept of MDIMain not needing to know details of the forms
    '      that it manages.
End Function


'Calls
'    CanSaveOrder
'    CheckReleaseStatus
'    AllItemsOnHand
'    CommitCheck
'    CommitOrder
'    TransitionTabs

Public Function CommitButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean
    Dim bOnARHold As Boolean
    
'Should MDIMain enabled the Commit button?

    CommitButton = False

    'If the Save button can't be enabled, neither can Commit
    
    If Not CanSaveOrder Then Exit Function

    'Allow the user Save but not Commit an order when PmtTerms='CrCard' but Order.CreditCard = nothing
    'in the event that the Customer has default PmtTerms of 'CrCard'
    
    If m_oOrder.PmtTerms.ID = "CrCard" And m_oOrder.CreditCard Is Nothing Then Exit Function
    
    'If the order has already been committed
    
    If m_oOrder.soKey <> 0 Then Exit Function
    
    'If the Quote's status is not ReadyToCommit (this should capture m_oOrder.SOKey <> 0
    
    If m_oOrder.StatusCode <> ItemStatusCode.iscReadyToCommit Then Exit Function
    
    CommitButton = True

    If Not i_bDoIt Then Exit Function

'***DoIt
'It's possible to commit the order, but let's do some validation first

    On Error GoTo ErrorHandler
    
    'Had to duplicate this from SaveButton in the event the order is dirty and unsaved
    'CommitButton should use the SaveButton logic as appropriate
    
    'If HasPartOnMultipleLines("Commit") Then Exit Function
        
    If m_oOrder.Customer.lStatusKey <> CustomerStatusCode.cscActive Then
        msg "This customer is not active customer. Can't commit the order." & vbCrLf _
            & "Please contact the computer guys for more details.", _
            vbCritical + vbOKOnly, "Can't Commit"
        Exit Function
    End If

    If HasBeenCommitted(m_oOrder.OPKey) Then
        msg "This order has been committed by someone else while you had it open. " & vbCrLf _
            & "Close this order and reload it.", _
            vbCritical + vbOKOnly, "Can't Commit"
        Exit Function
    End If
    
    'Not sure why this is needed.
    'Does it have something to do with the preceding CheckReleaseStatus call?
    ForceLostFocus

    If m_oBrokenRules.Count > 0 Then
        With m_oBrokenRules
            Dim oCtrl As Control
            Dim lRuleIndex As Long

            lRuleIndex = .FindBrokenRuleIndex(Me.ActiveControl)
            If lRuleIndex <= 0 Then lRuleIndex = 1
            
            .ForceFocus .Item(lRuleIndex).Rule.CtlWrapper.Ctl, .Item(lRuleIndex).ErrorMsg
            Set oCtrl = Me.ActiveControl
            msg .Item(lRuleIndex).ErrorMsg, , "Correct the following error, then try again"
            oCtrl.SetFocus
        End With
        
        Exit Function
    End If
        
    With m_oOrder
        
        'No commit permitted if order contains a line item invalid for selected warehouse
        If AllItemsOnHand Then

            If m_oCustomer.IsCOD And g_bWillCallUser Then
                If vbNo = msg("This customer does NOT have an open account. Are you sure you want to proceed?", vbExclamation + vbYesNo, "Commit Will Call Order") Then
                    Exit Function
                End If
            End If
            
            '2/7/2012 LR
            'call attention to PO Boxes in shipping address
            If PossiblePOB(.Customer.ShipAddr) Then
                Dim msgstr As String
                msgstr = "You are shipping by " & .ShipMethod & " to" & vbCrLf & _
                    .Customer.ShipAddr.Addr1 & vbCrLf & _
                    .Customer.ShipAddr.Addr2 & vbCrLf & _
                    vbCrLf & "which is possibly a PO Box." & vbCrLf & _
                    "Do you really want to commit?"
                If vbNo = msg(msgstr, vbExclamation + vbYesNo, "Possible PO Box") Then
                    Exit Function
                End If
            End If
            
            '9/24/03 LR. If this is being committed by a WillCall user,
            'then check the order's ship method. It should be "WillCall"
            'This method is Warehouse specific.
            '12/29/03 AVH. This is a kludge to resolve PRN312. Adding the OR
            'clause to detect SLWillCall user. If true, WillCall ship method is preferred.
            If (g_bWillCallUser And Not ShipMethIsWillCall) Or (UCase(GetUserName) = UCase("SLWillCall")) Then
                If vbYes = msg("This order will ship " & m_oOrder.ShipMethod & ". Do you want to change it to WillCall?", vbYesNo, "OrderPad") Then
                    m_oOrder.ShipMethKey = WhseWillCallShipMethKey
                End If
            End If
            
            '11/7/03 LR, PRN 271
            'if the Order's Ship From warehouse is not equal to the User's warehouse then
            'do not allow the shipmethod = WillCall
            If GetUserWhseID <> cboWarehouse(0).text Then
                If ShipMethIsWillCall Then
                    msg "You can't commit a Will Call order sourced from a foreign warehouse (" & cboWarehouse(0).text & ")."
                    Exit Function
                End If
            End If

            If m_oOrder.PmtTerms.IsCOD Then
                If InStr(m_oOrder.ShipMethod, "UPS") Then
                    If Not (m_oOrder.ShipComplete) Then
                        msg "COD orders must be Ship Complete."
                        Exit Function
                    End If
                End If
            End If
            
'*** The remaining vbCancel case is provisionally replaced by
'    (Something is still not right.  Why go on from here?  Is this a bug or by design?)
            If Not CommitCheck Then
                CommitButton = False
            End If

'*** CATALOG Request begin
            'You can't send a catalog with a dropship order.
            'Check to see if the Customer is eligable for a catalog.
            If WhseHasCatalogs And (Not m_oOrder.IsDropShip) And m_oCustomer.QueryForCatalog Then
                'show the dialog box
                Dim oFrm1 As FCatalogRequest
                
                Set oFrm1 = New FCatalogRequest
                With m_oCustomer
                    oFrm1.Init .Key, .ShipAddr.AddrKey, .CustType, Me
                End With
                Set oFrm1 = Nothing
                DoEvents 'let the dialog box clear before commit logic progresses
            End If
'*** CATALOG Request end

'*** STL Xmas gift processor
            'If g_bXmasGifts And (Not m_oOrder.IsDropShip) And m_oCustomer.QueryForGift Then
            If g_QueryForGift And (Not m_oOrder.IsDropShip) And m_oCustomer.QueryForGift Then
                Dim oForm As FXmasGift
                
                Set oForm = New FXmasGift
                With m_oCustomer
                    oForm.Init .Key, .ShipAddr.AddrKey, Me
                End With
                Set oForm = Nothing
                DoEvents 'let the dialog box clear before commit logic progresses
            End If
'*** End Xmas gift processor
            
            'Commit the order
            'if it fails, exit
            
            If Not CommitOrder(bOnARHold) Then
                Exit Function
            End If
            
        'all items are not on hand
        Else
            If vbYes = msg("Would you like the opportunity to change the warehouse?", vbQuestion + vbYesNo, "Continue Editting?") Then
                Exit Function
            Else
                .Save
                CommitButton = False
            End If
        End If
        
        DoEvents 'what's this for?

'What's going on here?
        If m_lCustKey = 0 Then  'FillSelectCustTab m_lCustKey
            cmdNewSearch_Click  'reset Existing Customer tab
        Else
            FillSelectCustTab m_lCustKey, False
        End If
        
        'cmdFindOrders2_Click 'refresh existing orders tab
        FindOrdersByCriteria

        If CommitButton Then
            If bOnARHold Then
                TransitionTabs True, sCaption & "was sent to A/R"
            Else
                TransitionTabs True, sCaption & "was committed"
            End If
        Else
            TransitionTabs True, sCaption & "was saved. It's not ready to commit"
        End If

    End With
    
    Exit Function

ErrorHandler:
    ClearWaitCursor
    ErrorUI.DisplayWarning "Commit Failed"
    
End Function




'*********************************************************************************************
' Transition Logic
'*********************************************************************************************

'Called by:
'   CommitButton
'   SaveButton
'   CancelButton
'   CheckReleaseStatus
'   cmdNewCust_Click
'   cmdWalkup_Click
'   cmdMiscCustomer_Click
'   cmdNewOrder_Click
'   cmdFindOrders_Click
'   cmdLoadOrder_Click
'
'Called from outside FOrder by:
'   FPhoneFlagger.cmdLoadOrder_Click
'   CommandHandler.DoSplitOrder
'   FBilling.LoadOrder
'   FConflictOrders.cmdViewOrder_Click
'   FWillCallTool.gdxWillCall_DblClick
'
'Calls:
'   UpdateOrderInfo
'   LogOAEvent
'   UpdateOrderCaption
'   gdxItems.Refresh
'   Form_Resize
'   FOrder.UpdateStatusBar
'   MDIMain.UpdateToolbarStatus
'
'Global SideEffect:
'   m_bRecommit = False
'   FindMode = bFindMode
'   m_bCanCancel = True
'   m_bPromptedForShipComplete = True

Public Sub TransitionTabs(ByVal bFindMode As Boolean, Optional strSBText As String = vbNullString)

    SetWaitCursor True
    
    If bFindMode Then
        EnteringFindMode
        ResizeFindMode
        UpdateFindStatusBar strSBText
    Else
        EnteringOrderMode
        ResizeOrderMode
        UpdateOrderStatusBar
    End If

    FindMode = bFindMode
    
    MDIMain.UpdateToolbarStatus
    
    SetWaitCursor False
    
End Sub


'Called by
'   TransitionTabs

Private Sub EnteringFindMode()
    Dim oCtrl As Control
    
    m_bRecommit = False
    
    txtFindOrder.SelStart = 0
    txtFindOrder.SelLength = Len(txtFindOrder.text)
    
    For Each oCtrl In Controls
        If TypeOf oCtrl Is TextBox _
            Or TypeOf oCtrl Is CheckBox _
            Or TypeOf oCtrl Is ComboBox _
            Or TypeOf oCtrl Is SOTAMaskedEdit _
            Or TypeOf oCtrl Is SOTANumber _
            Or TypeOf oCtrl Is CommandButton _
            Or TypeOf oCtrl Is ImageCombo _
            Or TypeOf oCtrl Is CurrencyInput _
            Or TypeOf oCtrl Is LengthEntry _
            Or TypeOf oCtrl Is OptionButton _
            Or TypeOf oCtrl Is SOTACalendar _
        Then
            oCtrl.Enabled = True
        End If
    Next
    
    With tabMain
        .Tabs(m_ePreviousTab).Visible = True
        .Tabs(m_ePreviousTab).Selected = True
        If m_ePreviousTab = tmiExistingCustomer And Len(m_oCustomer.Name) > 0 Then
            If m_oCustomer.IsTemp Or m_oCustomer.IsWalkup Then
                SetCaption "OrderPad"
            Else
                SetCaption m_oCustomer.Name
            End If
        Else
            SetCaption "OrderPad"
        End If
        .Tabs(tmiExistingCustomer).Visible = True
        .Tabs(tmiExistingOrder).Visible = True
        .Tabs(tmiCustomer).Visible = False
        .Tabs(tmiOrder).Visible = False
        .Tabs(tmiOrderHistory).Visible = False
        .Tabs(tmiLines).Visible = False
        .Tabs(tmiOrderStatus).Visible = False
        .Tabs(tmiRmaLines).Visible = False
    End With
        
End Sub

Public Sub EnteringOrderMode()
    Dim oCtrl As Control
    
    m_bCanCancel = True 'Latches true the first time we begin editing an order
    
    If m_oOrder.StatusCode = ItemStatusCode.iscCommitted Then
        cmdEditAddr(1).Enabled = False
        cmdEditAddr(0).Enabled = False
        cmdCreateRMA.Enabled = True
    Else
        cmdEditAddr(1).Enabled = True
        cmdEditAddr(0).Enabled = True
        cmdCreateRMA.Enabled = False
        If m_oOrder.StatusCode <> ItemStatusCode.iscHasRMA Then
            cmdLiteEdit.Enabled = False
        End If
    End If
    
    'ensure no warnings while loading order (cleared below)
    m_bPromptedForShipComplete = True
    
    UpdateOrderInfo
    
    'Log the View event into order history
    If Not m_oOrder.IsNewOrder Then
        LogOAEvent "Order", GetUserID, m_oOrder.OPKey, , m_oOrder.StatusCode, "Viewed."
    End If

    UpdateOrderCaption
    
    gdxItems.Refresh
    
    rvOrder.RemarkContext = m_oOrder.RemarkContext
        
'5/12/09 LR BUG?
'we really want establish a context if the cust has an account
'        If Not m_oCustomer.IsTemp And Not m_oCustomer.IsWalkup Then

'5/14/09 LR
'For GeorgeT
    If m_oCustomer.HasAccount Then
        rvCustomer.OwnerID = m_oCustomer.ID
    'else it's Misc, Walkup (or Temp)
    ElseIf GetUserWhseID = "SEA" Then
        rvCustomer.OwnerID = "SEA-MISC"
    Else
        rvCustomer.OwnerID = ""
    End If

    m_bPromptedForShipComplete = False
    
    If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then
         For Each oCtrl In Controls
            If TypeOf oCtrl Is TextBox _
                Or TypeOf oCtrl Is CheckBox _
                Or TypeOf oCtrl Is ComboBox _
                Or TypeOf oCtrl Is SOTAMaskedEdit _
                Or TypeOf oCtrl Is SOTANumber _
                Or TypeOf oCtrl Is CommandButton _
                Or TypeOf oCtrl Is ImageCombo _
                Or TypeOf oCtrl Is CurrencyInput _
                Or TypeOf oCtrl Is LengthEntry _
                Or TypeOf oCtrl Is OptionButton _
                Or TypeOf oCtrl Is SOTACalendar _
            Then
                oCtrl.Enabled = False
            End If
        Next
        
        '11/6/09 LR Kludge! enable the email quote button when viewing a committed order
        cmdEmailQuote.Enabled = True
    End If
    
    cmdCreateRMA.Enabled = (m_oOrder.StatusCode = ItemStatusCode.iscCommitted)
    cmdLiteEdit.Enabled = (m_oOrder.StatusCode = ItemStatusCode.iscCommitted Or m_oOrder.StatusCode = ItemStatusCode.iscHasRMA)
    cmdOSRefresh.Enabled = (m_oOrder.StatusCode = ItemStatusCode.iscCommitted Or m_oOrder.StatusCode = ItemStatusCode.iscHasRMA)
    cmdSpecialHandling.Enabled = True
    cmdPrint.Enabled = True

    If Order.Customer.HasAccount Then
        cmdContactMgr(1).Enabled = True
    Else
        cmdContactMgr(1).Enabled = False
    End If
    
    With tabMain
        m_ePreviousTab = .SelectedTab.Index
        .Tabs(tmiCustomer).Visible = True
        .Tabs(tmiCustomer).Selected = True
        .Tabs(tmiExistingCustomer).Visible = False
        .Tabs(tmiExistingOrder).Visible = False
        .Tabs(tmiCustomer).Visible = True
        .Tabs(tmiOrder).Visible = True
        .Tabs(tmiOrderHistory).Visible = True
        .Tabs(tmiLines).Visible = True
        .Tabs(tmiOrderStatus).Visible = (m_oOrder.soKey > 0) And (m_oOrder.StatusCode <> iscDeleted)
        .Tabs(tmiRmaLines).Visible = (m_oOrder.StatusCode = ItemStatusCode.iscHasRMA)
    End With
    
End Sub


Private Sub UpdateFindStatusBar(strSBText As String)
    sbOrderStatus.Panels.Clear
    If strSBText <> "" Then
        sbOrderStatus.Panels.Add 1, , strSBText
        sbOrderStatus.Panels.Item(1).width = sbOrderStatus.width
    End If
End Sub


'Called by
'   TransitionTabs

'all statusbar panel text is cleared prior to calling this routine

Private Sub UpdateOrderStatusBar()
    sbOrderStatus.Enabled = True
    
    AddStatusBarPanel
    
    SetOrderStatusBar

    If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then
        sbOrderStatus.Panels(2).text = "Read Only"
    Else
        sbOrderStatus.Panels(2).text = "Edit"
    End If
    
    sbOrderStatus.Panels(3).text = "OP " & m_oOrder.OPKey
    If m_oOrder.StatusCode = ItemStatusCode.iscHasRMA Then
        sbOrderStatus.Panels(5).text = "RMA " & m_lRMAKey
    End If
    If m_oOrder.soKey > 0 Then
        sbOrderStatus.Panels(4).text = "SO " & CStr(m_oOrder.TranNo)
    End If

    If m_oOrder.HasSpecialHandling Then
        sbOrderStatus.Panels(6).text = "Special Handling"
    End If
        
    If m_oOrder.IsDropShip Then
        'sbOrderStatus.Panels(7).Picture = imgDrop.ListImages(1).Picture
        sbOrderStatus.Panels(7).text = "Drop Ship"
        'this is temporary
        'cmdManageDropShips.Enabled = True
        cmdManageDropShips.Enabled = False
    Else
        'sbOrderStatus.Panels(7).Picture = Nothing
        sbOrderStatus.Panels(7).text = ""
        cmdManageDropShips.Enabled = False
    End If
End Sub



'*********************************************************************************************
' End Transition Logic
'*********************************************************************************************


Private Sub cboCustType_Click()
    Dim oItem As IItem
    Dim sCustType As String
    
    sCustType = cboCustType.text
    m_oCustomer.CustType = sCustType
    If cboCustType.text = "EndUser" Then
        SetComboByText cboTerms, "COD"
    Else
        SetComboByText cboTerms, "N30"
    End If

    For Each oItem In m_oItems
        oItem.CustType = sCustType
    Next
    gdxItems.Refetch
End Sub


Private Sub cboMake_LostFocus(Index As Integer)
    
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    Dim rst As ADODB.Recordset
    Dim lPrefVendorKey As Long
    
    With cboMake(Index)
        m_oItems.SelectedItem.MakeKey = .ItemData(.ListIndex)
    
        'Decide DropShipVendKey here
        If Not (m_oOrder.IsDropShip And m_oOrder.DropShipVendKey <> 0) Then
            If Index = 3 Then
                Set rst = CallSP("spcpcGetPrefVendor", "@_iMakeID", .ItemData(.ListIndex), _
                                "@_iWhseKey", m_oOrder.WhseKey)
                If rst.RecordCount > 0 Then
                    lPrefVendorKey = rst.Fields("PrefVendKey").value
              
                    If cboVendor.ItemData(cboVendor.ListIndex) = 0 Then
                        SetVendorKey lPrefVendorKey
                    ElseIf cboVendor.ItemData(cboVendor.ListIndex) <> lPrefVendorKey Then
                        If vbYes = msg("We generally buy " & .list(.ListIndex) & " parts for this warehouse from " _
                                    & rst.Fields("VendName").value & "." & vbCrLf & vbCrLf _
                                    & "Would you like to change the selected vendor from " _
                                    & cboVendor.list(cboVendor.ListIndex) & " to " & rst.Fields("VendName").value & "?", vbYesNo + vbExclamation, "Update Preferred Vendor") Then
                            SetVendorKey lPrefVendorKey
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub SetVendorKey(ByVal i_lPrefVendKey As Long)
    SetComboByKey cboVendor, i_lPrefVendKey
    m_oItems.SelectedItem.VendorKey = cboVendor.ItemData(cboVendor.ListIndex)
    ItemUpdateControls
End Sub


'Check if the dropship VendKey is decided only by selected item.
'If it is, we can change the vendkey of selected items.

'Called by
'   cboVendor_LostFocus()
'   gdxItems_OLECompleteDrag(Effect As Long)
'   cmdItemDelete_Click()
'   cmdItemOK_Click()

Private Function CheckDropShipSelectedItem() As Boolean
    Dim oItem As IItem
    Dim Count As Integer
    Dim i As Integer

    CheckDropShipSelectedItem = False
    
    Count = 0
    For i = 1 To m_oItems.Count
        If i <> m_oItems.SelectedIndex Then
            If m_oItems(i).VendorKey = m_oOrder.DropShipVendKey And m_oOrder.DropShipVendKey <> 0 Then
                Count = Count + 1
            End If
        End If
    Next i
    
    If Count = 0 Then CheckDropShipSelectedItem = True
    
End Function


Private Sub cboVendor_LostFocus()

    'When would there be no SelectedItem in the collection?
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    Dim bLoad As Boolean
    bLoad = m_bLoading
    m_bLoading = True
    
    With cboVendor
        'Check DropShipVendKey first while setting vendor for selected items.
        If m_oOrder.IsDropShip And _
           m_oOrder.DropShipVendKey <> 0 And _
           .ItemData(.ListIndex) <> 0 And _
           Not CheckDropShipSelectedItem And _
           m_oOrder.DropShipVendKey <> .ItemData(.ListIndex) Then
           
            g_rstVendors.Filter = "VendKey = " & m_oOrder.DropShipVendKey
            
            msg "This is a dropship order and the dropship vendor is " & RTrim(g_rstVendors.Fields("VendName").value) & ". " & _
            "You can't order an item from " & vbCrLf & "different vendor for dropship order. You can either order item(s) " & _
            "from dropship vendor or split this dropship order" & vbCrLf & " if you want to order items from different vendors."
            
            g_rstVendors.Filter = adFilterNone
            
            SetVendorKey m_oOrder.DropShipVendKey
            
        Else
            If .ItemData(.ListIndex) <> 0 Then
                rvVendor.Visible = True
                rvVendor.OwnerID = ""
                g_rstVendors.Filter = "VendKey = " & .ItemData(.ListIndex)
                rvVendor.OwnerID = g_rstVendors.Fields("VendID").value
                g_rstVendors.Filter = adFilterNone
            End If
            m_oItems.SelectedItem.VendorKey = .ItemData(.ListIndex)
            ItemUpdateControls
        End If
    End With
    m_bLoading = bLoad
End Sub


Private Sub cboVendor_Click()
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    If m_bLoading Then Exit Sub
    
    m_bLoading = True
    With cboVendor
        If .ItemData(.ListIndex) = 0 Then
            rvVendor.Visible = False
            rvVendor.OwnerID = ""
        End If
    End With
    m_bLoading = False
End Sub


Private Function BillRecipientStat() As String
    Dim rst As ADODB.Recordset
    Dim dfltRst As ADODB.Recordset
    
    If m_oOrder.Customer.ShipAddr.AddrKey > 0 Then
        BillRecipientStat = SetUPSAcct(m_oOrder.Customer.ShipAddr.AddrKey)
        If BillRecipientStat <> "" Then Exit Function
    End If
        
    Set dfltRst = LoadDiscRst("Select DfltShipToAddrKey from tarCustomer where CustKey = " & m_oOrder.Customer.Key)
    If Not dfltRst.EOF Then
        BillRecipientStat = SetUPSAcct(dfltRst.Fields("DfltShipToAddrKey"))
    End If
    Set dfltRst = Nothing
End Function


'Since we're only looking for the first record, if any, this should use a parameterized
'command rather than a recordset.
'What if there's more than one record in the table for a custaddrkey?

Private Function SetUPSAcct(lAddrKey As Long) As String
    Dim rst As ADODB.Recordset
    
    Set rst = LoadDiscRst("Select * from tcpUPSAcct where CustAddrKey = " & lAddrKey)
    If Not rst.EOF Then
        SetUPSAcct = Trim(rst.Fields("UPSAcct").value)
    End If
    Set rst = Nothing
End Function


Private Sub chkCGMPN_Click()
    If m_bChooseItem Then Exit Sub
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    m_oItems.SelectedItem.IsCGMPN = (chkCGMPN.value = vbChecked)
End Sub


Private Sub cbowires_Scroll()
    m_oWarmerWire.OhmsPerFoot = cboWires.list(cboWires.ListIndex)
End Sub


Private Sub mnuCustOrderCollapseAll_Click()
    gdxCustOrders.CollapseAll
End Sub


Private Sub lenShelfWidth_LostFocus()
    If lenShelfWidth.value > 36 Then
        msg "A shelf of this width probably requires anti-sway support", vbOKOnly, "Shelf Width > 36"
    End If
End Sub


'What is this??
Private Sub sbOrderStatus_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 6 And sbOrderStatus.Panels(6).text <> "" Then
        cmdSpecialHandling_Click
    End If
End Sub


Private Sub txtCost_GotFocus()
    If Not txtCost.Enabled Then Exit Sub
    
    If cboVendor.Visible = True Then
        If rvVendor.Visible = False And cboVendor.ListIndex > 0 Then
            rvVendor.Visible = True
            rvVendor.OwnerID = ""
            
            g_rstVendors.Filter = "VendKey = " & cboVendor.ItemData(cboVendor.ListIndex)
            rvVendor.OwnerID = g_rstVendors.Fields("VendID").value
            g_rstVendors.Filter = adFilterNone
        End If
    End If
End Sub


Private Sub txtFindCust_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub


Private Function GetShipRemarksMemo(sContext As String, sTypeID As String, sOwnerID As String) As String
    Dim ShipRemarkContext As RemarkContext
    Dim lTemp As Long
    
    Set ShipRemarkContext = New RemarkContext
    ShipRemarkContext.Load sContext, sOwnerID
    
    With ShipRemarkContext
        lTemp = CheckShipRemarks(ShipRemarkContext, sTypeID)
        If lTemp > 0 Then
            GetShipRemarksMemo = .RemarkList(lTemp).MemoText
            Exit Function
        Else
            GetShipRemarksMemo = ""
        End If
    End With
End Function


Private Sub DeleteShipRemarks(sContext As String, sTypeID As String, sOwnerID As String)
    Dim ShipRemarkContext As RemarkContext
    Dim lTemp As Long
    
    Set ShipRemarkContext = New RemarkContext
    ShipRemarkContext.Load sContext, sOwnerID
    
    If ShipRemarkContext Is Nothing Then Exit Sub
    
    With ShipRemarkContext
        lTemp = CheckShipRemarks(ShipRemarkContext, sTypeID)
        If lTemp > 0 Then
            .RemarkList(lTemp).Delete
        End If
    End With
End Sub


Private Sub UpdateShipRemarks(sContext As String, sTypeID As String, sNewMemo As String, sOwnerID As String)
    Dim ShipRemarkContext As RemarkContext
    Dim lTemp As Long
    
    Set ShipRemarkContext = New RemarkContext
    ShipRemarkContext.Load sContext, sOwnerID
    
    If ShipRemarkContext Is Nothing Then Exit Sub
    
    With ShipRemarkContext
        lTemp = CheckShipRemarks(ShipRemarkContext, sTypeID)
        If lTemp > 0 Then
            .RemarkList(lTemp).Dirty = True
            .RemarkList(lTemp).MemoText = sNewMemo
        Else
             .AddRemark GetRemarkTypeIndex(ShipRemarkContext, sTypeID), sNewMemo
        End If
         .Save True
    End With
End Sub


Private Function CheckShipRemarks(ByRef SRContext As RemarkContext, sTypeID As String) As Long
    Dim ShipRemarks As remark
    Dim lIndex As Long
    
    If SRContext.RemarkList.Count > 0 Then
        For Each ShipRemarks In SRContext.RemarkList
            lIndex = lIndex + 1
            If ShipRemarks.RemarkType.TypeID = sTypeID Then
                CheckShipRemarks = lIndex
                Exit Function
            End If
        Next
    End If
    
    CheckShipRemarks = 0
End Function


'When user checks dropship checkbox, OA will morph all BTO kits to SPO items automatically.
'If user unchecks dropship checkbox, OA will morph those SPO items back to BTO kits automatically.
'Order line remarks are lost in these two scenarios.
'
'The detailed process for morphing BTO kits to SPO item is:
'1. OA morph all BTO kit info to a SPO item, including line remarks. The function used
'for morphing is FOrder.MorphBTOtoSPO().
'2. OA delete the BTO kit from m_oItems.
'3. OA add the SPO item to m_oItems to replace the BTO kit. The error happens here. OA
'uses Items.Add to add new item to m_oItems. Items.Add tries to init and load order line
'remarks based on OPLineKey. But at this moment, there is nothing in tciMemo if this order
'is still not saved. Therefore, the new item's remarkcontext turns to be empty and those line remarks cloned are lost.
'
'The same thing for un-morphing SPO items back to BTO kits.
'
'We could make changes in Items.Add() to fix this problem. But I am not sure if this will
'cause other problems.  I'd need time to investigate.
'For now, I've added ClonesRemarks() in MorphBTOItems() and UndoMorphBTOItems() to clone
'line remarks again after adding items.

Private Sub MorphBTOItems()
    Dim i As Long
    Dim oItem As IItem
    Dim oSPOItem As IItem
    Dim sTemp As String
    
    i = 1
    For Each oItem In m_oItems
        If TypeOf oItem Is ItemBTOKit Then
                'lItemKey = oItem.ItemKey
                If Trim(sTemp) = "" Then
                    sTemp = oItem.ItemID
                Else
                    sTemp = sTemp & ", " & oItem.ItemID
                End If
                Set oSPOItem = MorphBTOtoSPO(oItem)
                m_oItems.Remove (i)
                m_oItems.Add oSPOItem
                '09/18/02 TeddyX
                'During m_oItems.Add, order line remarks lost. Therefore,
                'use CloneRemarks to clone line remarks to new added BTO kits.
                'It's necessary to refactor Items.Add later to fix this problem
                CloneRemarks oItem.RemarkContext, m_oItems.Item(m_oItems.Count).RemarkContext
                m_oItems.Item(m_oItems.Count).Backup
        Else
            i = i + 1
        End If
    Next
    
    If Trim(sTemp) <> "" Then
        msg "BTO Kit " & Trim(sTemp) & " has been converted to a SPO for dropship. " & vbCrLf & _
            "You will need to enter vendor cost and verify sales price before " & vbCrLf & _
            "you can submit this order.", vbOKOnly + vbExclamation, "Morph BTO to SPO"
    End If
End Sub


Private Function MorphSPOToBTO(ByRef oSPO As IItem) As ItemBTOKit
    Dim oBTOKit As ItemBTOKit
    
     Set oBTOKit = New ItemBTOKit
    oBTOKit.Load oSPO.MorphBTOKey
    With oBTOKit 'now the new item
        .IItem_OPLineKey = oSPO.OPLineKey
        .IItem_MakeKey = oSPO.MakeKey
        .IItem_ModelNbr = oSPO.ModelNbr
        .IItem_Qty = oSPO.Qty
        .IItem_IsTaxable = oSPO.IsTaxable
        .IItem_IsCGMPN = oSPO.IsCGMPN
        .IItem_SerialNbr = oSPO.SerialNbr
        .IItem_BackNegotiatedPrice = oSPO.BackNegotiatedPrice
        .IItem_NegotiatedPrice = oSPO.NegotiatedPrice
        '.IItem_Cost = oSPO.Cost
        .IItem_RemarkContext.Load "ViewOrderLine"
        CloneRemarks oSPO.RemarkContext, .IItem_RemarkContext
    End With
    
    Set MorphSPOToBTO = oBTOKit
End Function


Private Sub UndoMorphBTOItems()
    Dim i As Long
    Dim oItem As IItem
    Dim oBTOKit As ItemBTOKit
    Dim sTemp As String
    
    i = 1
    For Each oItem In m_oItems
        If oItem.ItemKey = 0 And oItem.MorphBTOKey > 0 Then
            m_oItems.Remove i
            Set oBTOKit = MorphSPOToBTO(oItem)
            m_oItems.Add oBTOKit
            '09/18/02 TX
            'During m_oItems.Add, order line remarks lost.
            'Therefore, use cloneremarks to clone line remarks to new added BTO kits.
            CloneRemarks oItem.RemarkContext, m_oItems.Item(m_oItems.Count).RemarkContext
            m_oItems.Item(m_oItems.Count).Backup
            m_bNewItem = False 'do not delete new morphed item on cancel
        Else
            i = i + 1
        End If
    Next
End Sub


Private Sub chkSearchItemDescr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSearch_Click
    End If
End Sub


Private Sub cmdCollapseAll_Click()
    gdxOrders.CollapseAll
End Sub


Private Sub cmdCopyModel_Click(Index As Integer)
    Dim oItem As IItem
    Dim oCurrentItem As IItem
    Dim lIndex As Long

    Set oCurrentItem = m_oItems.SelectedItem
    For lIndex = m_oItems.Count To 1 Step -1
        Set oItem = m_oItems(lIndex)
        With oItem
            If .ModelNbr <> "" And Not oCurrentItem Is oItem Then
                oCurrentItem.ModelNbr = .ModelNbr
                oCurrentItem.SerialNbr = .SerialNbr
                oCurrentItem.MakeKey = .MakeKey
                oCurrentItem.VendorKey = .VendorKey
                GoTo Cleanup
            End If
        End With
    Next
    msg "No model information to copy", , "Nothing to Copy"

Cleanup:
    Set oItem = Nothing
    ItemUpdateControls
End Sub


'Set controls on Select Customer Tab
'Called by:
'   ClearCustomerSearch()   IsNewSearch = True
'   FillSelectCustTab()     IsNewSearch = False

Private Sub SetSelCustCtrls(ByVal b_IsNewSearch As Boolean)
    Dim bLoading As Boolean
    bLoading = m_bLoading
    m_bLoading = True

'    cmdCustSearch.Visible = b_IsNewSearch
    cmdSelectCustomer(0).Visible = b_IsNewSearch
    
    txtCustSearch.Visible = b_IsNewSearch
    cboSearchType.Visible = b_IsNewSearch
    'PRN 231 fix 3/4/04 LR
    lblExplain(0).Visible = b_IsNewSearch And Not g_bWillCallUser
    lblExplain(1).Visible = b_IsNewSearch
    lblExplain(2).Visible = b_IsNewSearch And Not g_bWillCallUser

'9/27/05 changed initial value to Checked
    chkShowOrdersForShipAddr.Visible = Not b_IsNewSearch
    chkShowOrdersForShipAddr.value = vbChecked
    
    lblOrderStatus.Visible = Not b_IsNewSearch
    cboOrderStatus.Visible = Not b_IsNewSearch
    lblFilterByPart.Visible = Not b_IsNewSearch
    txtFilterByPart.Visible = Not b_IsNewSearch
    cmdFilterByPart.Visible = Not b_IsNewSearch
    
    If b_IsNewSearch Then
        txtFilterByPart.text = ""
        cmdFilterByPart_Click
    End If
    
    SetComboByText cboOrderStatus, "<Any>"
    
    frmCustInfo.Visible = Not b_IsNewSearch
    frmCustSearch.Visible = b_IsNewSearch
    
    cmdNewSearch.Visible = Not b_IsNewSearch
    cmdNewOrder.Visible = Not b_IsNewSearch
    cmdLoadOrder(0).Visible = Not b_IsNewSearch
    cmdContactMgr(0).Visible = Not b_IsNewSearch
    
'    cmdNewCust.Visible = b_IsNewSearch And Not g_bWillCallUser
'    cmdWalkup.Visible = b_IsNewSearch
'    cmdMiscCustomer.Visible = b_IsNewSearch And Not g_bWillCallUser
    cmdSelectCustomer(1).Visible = b_IsNewSearch And Not g_bWillCallUser
    cmdSelectCustomer(2).Visible = b_IsNewSearch
    cmdSelectCustomer(3).Visible = b_IsNewSearch And Not g_bWillCallUser
    
    gdxCustOrders.Visible = Not b_IsNewSearch
    
    m_bLoading = bLoading
End Sub


'Called by:
'   UpdateOrderInfo
'   txtCustID_LostFocus
'
'Notes:
'   IsTemp is a property of the Customer object.
'   It's value drives the input parameter here in the single case that it's True.
'   If IsTemp is True, the routine basically enables a bunch of input controls
'   on the Customer tab.
'   If False, they're turned off.
'   This is probably for a New Customer order only.

Private Sub SetCustCtrlVisible(ByVal b_IsTemp As Boolean)
    txtCustName.Visible = b_IsTemp
    lblCustName.Visible = Not b_IsTemp
    txtCustID.Visible = b_IsTemp
    lblCustID(0).Visible = Not b_IsTemp
    cboCustType.Visible = b_IsTemp
    lblCustType(0).Visible = Not b_IsTemp
    rvCustomer.Visible = Not b_IsTemp

    If b_IsTemp = True Then
        cboCustType.ListIndex = -1
        SetComboByText cboCustType, "EndUser"
    End If
End Sub


'Set the shipping controls settings on Order tab based on order type
Private Sub SetShippingCtrl(ByVal b_ExistingCust As Boolean)
    
    chkDefaultShipMeth.Enabled = b_ExistingCust
    chkDefaultShipMeth.value = vbUnchecked
  
    '***465 SMR 04-12-2006***
    'Only reset the UPS Account controls if the customer is not Bill Recipient.
    If Len(txtUPSAcct.text) = 0 Then
        chkBillRecipient.Enabled = False
        chkBillRecipient.value = vbUnchecked
            
        txtUPSAcct.text = ""
        cmdUPSUpdate.Enabled = False
    End If

End Sub


'Called By
'   ContinueEditSageOrder
'   CheckSageOrder

Private Sub ResetToQuote(ByVal i_OPKey As Long)
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spCPCUpdateOAOrder")
    cmd.Parameters("@_OPKey").value = i_OPKey
    cmd.Execute
    Set cmd = Nothing
End Sub


'Load an existing order
'
' There are 2 Load Order buttons in FOrder
'   0 - Customer Order tab
'   1 - Select Order tab
'
' Called By:
'   m_gwCustOrders_RowChosen
'   m_gwOrders_RowChosen
'
' Assumes the bound recordset for each of the above grids contains
' OPKey, StatusCode, SOKey, SOID

'This is the calling hierarchy.
' cmdLoadOrder_Click()
'     LoadOrder()
'         ContinueEditSageOrder()
'             CheckSageOrder()
'                 EligibleForCancel()
'
' Where do we determine view/edit mode?
' How is the state represented?

Private Sub cmdLoadOrder_Click(Index As Integer)
    Dim oGW As GridEXWrapper
    Dim lStatusCode As Long
    
    If Index = 0 Then
        Set oGW = m_gwCustOrders
    ElseIf Index = 1 Then
        Set oGW = m_gwOrders
    End If
    
    'If the grid is empty, then exit
    If IsEmpty(oGW.value("OPKey")) Then Exit Sub
    
    'SOStatusCode = 0 for now
    If LoadOrder(oGW.value("OPKey"), _
                    oGW.value("StatusCode"), _
                    oGW.value("SOKey"), _
                    oGW.value("SOID"), _
                    oGW.value("Status")) Then
        
        m_bLoading = True
    
        'How about keeping current customer search result. (?)
        
        'Clear customer search if order selected from Select Order tab
        If Index = 1 Then
            ClearCustomerSearch
        End If
        
        'cache module level references to the order's customer and item collection objects
        Set m_oCustomer = m_oOrder.Customer
        Set m_oItems = m_oOrder.Items
        
        TransitionTabs False
    
        LoadUPSCtrl m_oCustomer.Key
    
        m_bLoading = False
    End If

End Sub


'Called by:
'   cmdLoadOrder_Click
'   FindOrderByNumber

'i_SOStatusCode is null if the OP does not have an SO

Private Function LoadOrder(ByVal i_OPKey As Long, _
                            ByVal i_OPStatusCode As Long, _
                            ByVal i_SOKey As Long, _
                            i_SOID As String, _
                            i_SOStatusCode As Variant) As Boolean

    LoadOrder = False

    'control the caption of toolbar's commit button
    m_bRecommit = False
    
    'i_OPStatusCode was generated by the SQL select
    'it is not the same as tcpSO.StatusCode (aka ItemStatusCode)
    Dim lOPStatusCode As ItemStatusCode
    lOPStatusCode = RestoreItemStatusCode(i_OPStatusCode)
    
    Select Case lOPStatusCode
    
        Case ItemStatusCode.iscPendingCommit
            msg "This order is in the process of being saved to Sage and can't be opened at this time."
            Exit Function
    
        Case ItemStatusCode.iscCommitted
            If Not ContinueEditSageOrder(i_OPKey, i_SOKey, i_SOID, i_SOStatusCode) Then
                Exit Function
            End If
        
        Case ItemStatusCode.iscHasRMA
            Set m_oOrder = New Order        'Why ?!
            m_oOrder.Load i_OPKey
            
        Case Else
            m_oOrder.Load i_OPKey
    
    End Select
    
    LoadOrder = True

End Function


'The SQL queries that populate the grids map the OP statuscode (of type ItemStatusCode)
'to new values in oder to expand the Research status.
'To compare statuscode to ItemStatusCode values, you need to remap it to its original value.

Private Function RestoreItemStatusCode(i_StatusCode As Long) As ItemStatusCode
    If i_StatusCode > 2000 Then
        RestoreItemStatusCode = i_StatusCode - 2000
    ElseIf i_StatusCode > 1000 Then
        RestoreItemStatusCode = ItemStatusCode.iscResearch
    Else
        RestoreItemStatusCode = i_StatusCode
    End If
End Function


Private Function ContinueEditSageOrder(ByVal i_OPKey As Long, ByVal i_SOKey As Long, i_SOID As String, i_SOStatusCode As Variant) As Boolean
    Dim cmd As ADODB.Command
    Dim emailbody As String
    
    ContinueEditSageOrder = True

    ' are either of the first two cases likely or possible?
    ' instrument this to find out
    ' if not, simply put the call to CheckSageOrder() into LoadOrder
    
    'SO associated with the OP doesn't exist
    'i_SOStatusCode will be null if there is no SO tied to the OP
    If IsEmpty(i_SOStatusCode) Then
    
        msg "This order doesn't exist in Sage any more, please contact the computer guys for details", vbOKOnly + vbExclamation
        EMail.Send "operations@caseparts.com", "lennyr@caseparts.com", "ContinueEditSageOrder, Case 1", "", False
        ContinueEditSageOrder = False
    
    'SO associated with the OP has already been cancelled
    ElseIf i_SOStatusCode = SOStatusCode.sscCancelled Then
    
        emailbody = "Order has already been cancelled. Recommit?" & vbCrLf & "OP " & i_OPKey & ", SO " & i_SOID & ", Status " & i_SOStatusCode
        
        'the order is cancelled in Sage
        If vbYes = msg("This order has already been cancelled in Sage." & vbCrLf & "Would you like to recommit it?", vbExclamation + vbYesNo) Then
            EMail.Send "operations@caseparts.com", "lennyr@caseparts.com", "ContinueEditSageOrder, Case 2a", emailbody, False
            
            ResetToQuote i_OPKey
            'open in edit mode
            m_oOrder.Load i_OPKey
            
            'NOTE: The warehouse remarktype triggers its display on the Pick Report downstream
            m_oOrder.RemarkContext.AddRemark "Order.Warehouse", "This order was previously committed as SO " & i_SOID
            
            m_oOrder.Recommit = True
            m_oOrder.Save
            m_oOrder.Recommit = False
            m_bRecommit = True
        Else
            EMail.Send "operations@caseparts.com", "lennyr@caseparts.com", "ContinueEditSageOrder, Case 2b", emailbody, False
            m_oOrder.Load i_OPKey
        End If
    
    'This is the most likely scenario
    Else
        If Not CheckSageOrder(i_OPKey, i_SOKey, i_SOID) Then ContinueEditSageOrder = False
    End If
    
    Set cmd = Nothing
End Function



'Check if the Sage order is qualified for re-editing

'Called By:
'   ContinueEditSageOrder
'Calls:
'   EligibleForCancel

Private Function CheckSageOrder(ByVal i_OPKey As Long, ByVal i_SOKey As Long, i_SOID As String) As Boolean
    Dim RetVal As Integer
    
    On Error GoTo ErrorHandler
    CheckSageOrder = True
    
    If Not EligibleForCancel(i_SOKey) Then
        If vbCancel = msg("This order has items on shipments and is not eligible for editing." & vbCrLf & "It will be loaded for viewing only," & vbCrLf _
                        & "though you can edit Order and LineItem remarks and create RMAs." & vbCrLf & vbCrLf & "Continue Loading?", vbExclamation + vbOKCancel) Then
            CheckSageOrder = False
            Exit Function
        End If
        m_oOrder.Load i_OPKey
    Else
        Dim frm As FLoadOrder

        ClearWaitCursor     '???
        
        Set frm = New FLoadOrder
        frm.LoadSageOrder i_SOKey
        DoEvents
        
        SetWaitCursor True
        
        Select Case frm.Result
        
            Case "View"
                m_oOrder.Load i_OPKey

            Case "Edit"
                            
                RetVal = CancelSO(i_SOKey)
                
                Select Case RetVal
                    Case 1:
                        ResetToQuote i_OPKey
                        
                        LogDB.LogActivity "SA", "Cancelling SO for edit", _
                            i_OPKey, i_SOKey, i_SOID, , , , , m_oOrder.WhseKey
                    
                        If OrderHasReceipt(i_OPKey) Then
                            ClearReceiptPrinted i_OPKey
                            LogOAEvent "Order", GetUserID, i_OPKey, , , "Cleared 2-Copy Receipt"
                        End If
                        
                        'open in edit mode
                        m_oOrder.Load i_OPKey
                        
                        'NOTE: The warehouse remarktype triggers its display on the Pick Report downstream
                        m_oOrder.RemarkContext.AddRemark "Order.Warehouse", "This order was previously committed as SO " & i_SOID
                        
                        m_oOrder.Recommit = True
                        m_oOrder.OldSOID = i_SOID
                        m_oOrder.Save
                        m_oOrder.Recommit = False
                        m_bRecommit = True
                    Case 0:
                        msg "Failed to cancel Sage's SO-" & i_SOKey & vbCrLf & "Unspecified error." & vbCrLf & "Opening in Read-only mode.", _
                            vbOKOnly, "Edit Order"
                        m_oOrder.Load i_OPKey
                    Case 2:
                        msg "Failed to cancel Sage's SO-" & i_SOKey & vbCrLf & "Order has lines on a ahipment." & vbCrLf & "Opening in Read-only mode.", _
                            vbOKOnly, "Edit Order"
                        m_oOrder.Load i_OPKey
                End Select

            Case "Cancel"
                CheckSageOrder = False
                
        End Select
        
        Unload frm
        Set frm = Nothing
        SetWaitCursor False
    End If
    
    Exit Function
    
ErrorHandler:
    msg Err.Description, vbCritical, Err.Source
End Function


Private Sub PromptShipComplete()
    Dim bBackOrder As Boolean
    Dim bNotBackOrder As Boolean
    Dim oItem As IItem
    
    If m_bPromptedForShipComplete Then Exit Sub
    If Not chkShipComplete.value = vbChecked Then Exit Sub
    
    For Each oItem In m_oItems
        With oItem
            If .OPItemType = itSpecialOrder Or .OPItemType = itWireShelf Then
                bBackOrder = True
            ElseIf .OPItemType = itWarmerWire Or .OPItemType = itMoldedGasket Then
                bNotBackOrder = True
            Else
                If .Qty > .QtyAvail(m_oOrder.WhseKey) Then
                    bBackOrder = True
                Else
                    bNotBackOrder = True
                End If
            End If
        End With
    Next
    
    If bBackOrder And bNotBackOrder Then
        If vbYes = msg("This order contains some items that can ship immediately and some that can't." & vbCrLf _
                    & "Are you sure that you want hold the whole order to Ship Complete?", vbQuestion + vbYesNo, _
                      "Confirm Ship Complete") Then
               chkShipComplete.value = vbChecked
         End If
         m_bPromptedForShipComplete = True
    End If
End Sub


Private Sub cmdExpandAll_Click()
    gdxOrders.ExpandAll
End Sub


'Called by
'   FindOrderByNumber (eliminated 1/31/05 LR)
'   FindOrdersByCriteria

Private Function SearchOrder(sSearchingText As String) As ADODB.Recordset
    Dim sSQL As String

    'Search the orders based on OPKey, TransKey, or RMA Key
    'PRN#3 Add tsoSalesOrder to retrieve Open/Closed status

    sSQL = "SELECT o.OPKey, o.TranKey as SOID, o.CreateDate, " _
         & "Case StatusCode when 0 then 0 when 1 then " _
         & "Case when ResearchStatus is null then 1001 " _
         & "when ResearchStatus = 0 then 1001 " _
         & "Else 1000+ResearchStatus End " _
         & "Else 2000+StatusCode End As StatusCode, " _
         & "(CASE WHEN o.flags&0x01 = 0x01 THEN -1 ELSE 0 END) AS Dropship, " _
         & "o.UpdateDate, RTRIM(o.UserID) as UserID, RTRIM(ta.UserFld1) as Collector, o.WhseKey,o.CustKey, " _
         & "RTRIM(ISNULL(tciContact.Name, '')) as OrderedBy, " _
         & "RTRIM(ta.CustID) as CustID, RTRIM(o.CustType) as CustType, " _
         & "o.CustName, o.SOKey, o.Summary, RTRIM(o.PurchOrd) as CustPO, s.Hold, rma.RMAKey, " _
         & "(CASE tsoSalesOrder.Status WHEN 1 THEN 'Open' when 4 then 'Closed' ELSE 'Other' END) as StatusDesc, tsoSalesOrder.Status, " _
         & "tciShipMethod.ShipMethID, tciPaymentTerms.PmtTermsID, o.Info as Note " _
         & "FROM tcpSO o (nolock) " _
         & "LEFT OUTER JOIN tciShipMethod (nolock) ON o.ShipMethKey = tciShipMethod.ShipMethKey " _
         & "LEFT OUTER JOIN tciPaymentTerms (nolock) ON o.PmtTermsKey = tciPaymentTerms.PmtTermsKey " _
         & "LEFT OUTER JOIN tarCustomer ta (nolock) ON o.CustKey = ta.CustKey " _
         & "LEFT OUTER JOIN tarCustClass cc (nolock) ON ta.CustClassKey = cc.CustClassKey " _
         & "LEFT OUTER JOIN tcpCustHold h (nolock) ON o.CustKey = h.CustKey " _
         & "LEFT OUTER JOIN tcpHoldStatus s (nolock) ON h.HoldStatusKey = s.HoldStatusKey " _
         & "LEFT OUTER JOIN tcpRMA rma (nolock) ON rma.OPKey = o.OPKey " _
         & "LEFT OUTER JOIN tsoSalesOrder (nolock) ON o.SOKey = tsoSalesOrder.SOKey " _
         & "LEFT OUTER JOIN tciContact (nolock) ON tciContact.CntctKey = o.CntctKey " _
         & sSearchingText _
         & vbCrLf & "ORDER BY o.UpdateDate DESC"
    
    Set SearchOrder = LoadDiscRst(sSQL)
End Function


Private Sub AppendClause(ByVal i_sClause As String, ByRef io_sWhere As String)
    If Len(io_sWhere) = 0 Then
        io_sWhere = vbCrLf & "WHERE (" & i_sClause & ")"
    Else
        io_sWhere = io_sWhere & vbCrLf & "  AND (" & i_sClause & ")"
    End If
End Sub


Private Sub cmdPartsWiz_Click()
    Dim oFrm As FPartsWiz
    Dim sPartNbr As String
    Dim lMake As Long
    Dim sDescr As String
    
    Set oFrm = New FPartsWiz
    If oFrm.FindPart(txtItemDescr.text, _
                     txtModel(3).text, _
                   txtSerial(3).text, _
                     cboMake(3).ItemData(cboMake(3).ListIndex), _
                    m_oOrder.OPKey, sPartNbr, lMake, sDescr) Then
        'Get more information from parts wiz
        With m_oItems.SelectedItem
            '.ModelNbr = sModal
            '.SerialNbr = sSerialNbr
            .MakeKey = lMake
            .Descr = sDescr
            '.VendorKey = lVendor
        End With

        ItemUpdateControls
        
        With txtItemPartNbr
            .text = sPartNbr
            TryToSetFocus txtItemPartNbr
        End With
    End If
    UpdateOrderInfo
End Sub


Private Function CanSaveSageOrder() As Boolean
    On Error Resume Next

    Dim oItem As IItem
    CanSaveSageOrder = True
    
    If m_oOrder.RemarkContext.RemarkList.Dirty Then
        CanSaveSageOrder = False
        Exit Function
    End If
        
    For Each oItem In m_oOrder.Items
        If oItem.RemarkContext.RemarkList.Dirty Then
            CanSaveSageOrder = False
            Exit Function
        End If
    Next
End Function


'We can save the order if
'   FOrder is in Order Mode (not Find Mode) AND
'   we're not editing a lineitem (ViewMode is not List) AND
'   the user has save permission

'Called by:
'   CancelButton
'   ExitCheck
'   Unload
'   SaveButton
'   SplitOrderButton
'   Commit

Private Function CanSaveOrder() As Boolean
    CanSaveOrder = False
    
    'Save is only available while an order is being edited.
    If FindMode Then Exit Function

    'No Save while editing an item
    If ViewMode <> ivList Then Exit Function

    'If the user does not have permission to Save, get out
    If Not HasRight(k_sRightOPSaveOrder) Then Exit Function
    
    CanSaveOrder = True
End Function


'??? Why?
'If the active control is in the select list, fire its LostFocus event.

Private Sub ForceLostFocus()
    On Error Resume Next
    If Not (TypeOf Me.ActiveControl Is TextBox Or TypeOf Me.ActiveControl Is SOTAMaskedEdit) Then Exit Sub
    
    Select Case Me.ActiveControl.Name
        Case "txtCustName":
            txtCustName_LostFocus
        Case "txtCustID":
            txtCustID_LostFocus
        Case "txtPO":
            txtPO_LostFocus
        Case "txtInfo":
            txtInfo_LostFocus
        'SMR 01/03/2006 - This function is called before the order is Saved.
            'Without this call, the lost_focus event on the last field modified will not get triggered.
            'Therefore, saving the old value.  This call ensures that if the activecontrol is one of
            'these fields that the lost focus event is triggered. (This could probably be done in a different way.)
        Case "txtShipToContact"
            Call txtShipToContact_LostFocus(Me.ActiveControl.Index)
    End Select
End Sub


'Called by CommitButton

Private Function HasBeenCommitted(ByVal lOPKey As Long) As Boolean
    Dim orst As ADODB.Recordset
    
    HasBeenCommitted = False
    
    'Get the current status of in-memory order from the database
    Set orst = LoadDiscRst("Select tcpSO.StatusCode as OPStatus, isnull(tsoSalesOrder.Status, 0) as SOStatus " _
                        & "from tcpSO left outer join tsoSalesOrder on " _
                        & "tsoSalesOrder.SOKey = tcpSO.SOKey where tcpSO.OPKey = " & lOPKey)
    
    'if the OP status > ReadyToCommit and the corresponding SO hasn't been cancelled, return true, else false
    If Not orst.EOF Then
        HasBeenCommitted = (orst.Fields("OPStatus").value > ItemStatusCode.iscReadyToCommit) And (orst.Fields("SOStatus").value <> SOStatusCode.sscCancelled)
    End If
End Function


Private Sub cmdResearchPO_Click(Index As Integer)
    Dim oFrm As FOnPurchOrder
    
    Set oFrm = New FOnPurchOrder
    With m_oItems.SelectedItem
        oFrm.ShowPurchaseOrders .ItemID, .Descr, .ItemKey
    End With
End Sub


Private Sub gdxItems_GotFocus()
    'JJC: gdxItems receives a GotFocus event as a consequence of the
    'processing we do loading an item to edit.  This causes big problems
    'because the SyncItemList logic results in a SelectionChange event
    'which then causes the m_oItems.SelectedIndex property to be set to
    'the wrong value.  Consequently, any user edits applied to
    'm_oItems.SelectedItem are applied to the wrong IItem object.
    'So, let's just cut this off at the knees by ignoring GotFocus
    'unless the items list is the visible.
    If ViewMode = ivList Then
        SyncItemList
    End If
End Sub


Private Sub gdxItems_LostFocus()
    gdxItems.Row = -1
End Sub


Private Sub gdxItems_SelectionChange()
    Dim lIndex As Long

    If m_bLoading Then Exit Sub 'don't change SelectedItem
    If m_bNewItem Then Exit Sub
    
    If ViewMode <> ivList Then Exit Sub
    
    With gdxItems
        lIndex = .RowIndex(.Row)
        If lIndex > 0 And lIndex <= m_oItems.Count Then
            m_oItems.SelectedIndex = lIndex
        End If
    End With
End Sub


'Private Sub gdxOrders_AfterGroupChange()
'    With gdxOrders.Groups
'        cmdExpandAll.Enabled = (.Count > 0)
'        cmdCollapseAll.Enabled = (.Count > 0)
'    End With
'End Sub


Private Sub icbItemStatus_Change()
    '09/25/02 TeddyX
    'When icbItemStatus changes, we also have to consider what happend
    'after adding research status to the combo box.
    Dim eResearchStatus As ItemResearchStatus
    Dim eStatusCode As ItemStatusCode
    On Error GoTo Cleanup 'ignore error if nothing selected
    If Left(icbItemStatus.SelectedItem.Key, 1) = "*" Then
        'If m_oOrder.ResearchStatus <> irsResearchEmpty Then
        eStatusCode = iscResearch
       ' End If
        
        Select Case CLng(Mid(icbItemStatus.SelectedItem.Key, 2))
'            Case 1
'                eResearchStatus = irsNeedResearch
            Case 2
                eResearchStatus = irsContactFactory
            Case 3
                eResearchStatus = irsContactCustomer
            Case 4
                eResearchStatus = irsWaitFactory
            Case 5
                eResearchStatus = irsWaitCustomer
            Case Else
                msg "Unexpected item research status"
        End Select
        m_oItems.SelectedItem.ResearchStatus = eResearchStatus
    Else
        Select Case CLng(Mid(icbItemStatus.SelectedItem.Key, 2))
        Case 1
            eStatusCode = ItemStatusCode.iscResearch
        Case 2
            eStatusCode = ItemStatusCode.iscQuote
        Case 3
            eStatusCode = ItemStatusCode.iscAuthorize
        Case 4
            eStatusCode = ItemStatusCode.iscReadyToCommit
        Case Else
            msg "Unexpected item status"
        End Select
    End If
    
    m_oItems.SelectedItem.StatusCode = eStatusCode
    SetOrderStatusBar
Cleanup:
End Sub


Private Sub icbItemStatus_Click()
    icbItemStatus_Change
End Sub


Private Sub icbItemStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub m_oFinGood_Change()
    ItemUpdateControls
End Sub

Private Sub m_oBTOKit_Change()
    ItemUpdateControls
End Sub

Private Sub m_oWarmerWire_Change()
    ItemUpdateControls
End Sub

Private Sub m_oShelf_Change()
    ItemUpdateControls
End Sub

Private Sub m_oGasket_Change()
    ItemUpdateControls
End Sub


Private Sub cbowires_Click()
    m_oWarmerWire.OhmsPerFoot = cboWires.list(cboWires.ListIndex)
End Sub


Private Sub optLengthAlgorithm_Click(Index As Integer)
    If m_bLoading Then Exit Sub

    m_oBrokenRules.EnableClass ccWWireLength, (Index = 0)
    lenWireLength.Enabled = (Index = 0)
    
    m_oBrokenRules.EnableClass ccWWireDoorDim, (Index = 1)
    lenDoorHeight.Enabled = (Index = 1)
    lenDoorWidth.Enabled = (Index = 1)

    frmWirePasses.Visible = (Index = 0)
    frmDoorStyle.Visible = (Index = 1)
    
    m_oBrokenRules.EnableClass ccWWireType, True
    
    m_oBrokenRules.Validate
    cmdItemOK.Enabled = (m_oBrokenRules.MaskedCount(k_lItemControlMask) = 0)
End Sub


Private Sub SetOSTabsVisible()
    'Set all tabs' initial visible to be true
        
    Dim i As Integer
    
    For i = tosLineItem To tosInvoice
        SSAOrderDetails.Tabs(i).Visible = True
    Next
End Sub


Private Function sCaption() As String
    sCaption = m_oOrder.Customer.Name & "  OP " & m_oOrder.OPKey & " "
End Function


Private Sub UpdateOrderCaption()
    If m_oOrder.StatusCode = ItemStatusCode.iscHasRMA Then
        SetCaption sCaption & "- RMA " & m_lRMAKey & "  " & cboWarehouse(1).text
    Else
        SetCaption sCaption & cboWarehouse(1).text
    End If
End Sub


'Called by:
'   TransitionTabs()
'   cmdPartsWiz_Click
'   txtCustID_LostFocus
'   cboShipVia_Click

'there's a valid Customer object on entry

Private Sub UpdateOrderInfo()
    Dim dCurrentTime As Date
    
    Dim bLoading As Boolean
    bLoading = m_bLoading
    
    m_bLoading = True
    
    Set m_oItems = m_oOrder.Items
    
    With m_oOrder
        If m_oCustomer.IsTemp Then
            SetCustCtrlVisible True
            SetShippingCtrl False
        Else
            SetCustCtrlVisible False
            SetShippingCtrl Not (.isMisc Or .IsWalkup)
            
            If m_oCustomer.HasAccount Then
                rvCustomer.ContextID = "ViewCustomer"
                rvCustomer.Visible = True
            'else it's Misc or Walkup
            ElseIf GetUserWhseID = "SEA" Then
                rvCustomer.ContextID = "SEAMisc"
                If UCase(GetUserName) = "GEORGET" Or UCase(GetUserName) = "LENNYR" Then
                    rvCustomer.Visible = True
                Else
                    rvCustomer.Visible = False
                End If
            Else
                rvCustomer.Visible = False
            End If
            
        End If
    
        SetComboByKey cboCSR, .UserKey
        lblDate.caption = Format$(.CreateDate, "mm/dd/yy")
        lblLastUpdate.caption = Format$(.UpdateDate, "mm/dd/yy")
        
        txtShipToContact(0) = .ShipToName
        txtShipToContact(1) = .ShipToPhone
        
        txtShipToNote.text = .ShipToNote
                
        txtInfo.text = .Info
                
        txtPO.text = .PurchOrd
        SetCheckbox chkDropShip, .IsDropShip

        SetUpWarehouses cboWarehouse(0), g_rstWhses, .WhseKey
        SetUpWarehouses cboWarehouse(1), g_rstWhses, .WhseKey
        SetUpWarehouses cboWarehouse(2), g_rstWhses, .WhseKey
        SetUpShipVia cboShipVia, .WhseKey, .ShipMethKey

        SetComboByKey cboWarehouse(0), .WhseKey
        SetComboByKey cboWarehouse(1), .WhseKey
        SetComboByKey cboWarehouse(2), .WhseKey
        SetComboByKey cboShipVia, .ShipMethKey
                
        '09/12/02 TX
        'Error scenario: If you are an MPK CSR and you create an order for SEA,
        'then create an order for MPK, the color remains yellow until you unselect and
        'reselect the MPK warehouse.
        UpdateWarehouseColor 0
        
        'If the order's been committed
        '   Initialize the Order Status tab
        If m_oOrder.soKey > 0 Then
            LoadOrderStatus .OPKey
        End If
        
        If .StatusCode = ItemStatusCode.iscHasRMA Then
            m_lRMAKey = GetRMAKey
            LoadRMALine m_lRMAKey
            rvRMA.OwnerID = ""
            rvRMA.OwnerID = .OPKey
        End If
        
        '***465 SMR 04-10-2006 (note: looked at & no changed needed)***
        SetCheckbox chkPricePackList, .PricePackList
        
        If m_oOrder.IsNewOrder Then
            SetCheckbox chkShipComplete, m_oCustomer.BillAddr.ShipComplete Or m_oCustomer.ShipAddr.ShipComplete
        Else
            SetCheckbox chkShipComplete, .ShipComplete
        End If
        
        chkCGMPN.Enabled = (.soKey = 0)

    End With

    With m_oCustomer
        If .IsTemp Then
            txtCustName.text = Helpers.FormatCaption(.Name)
            SetComboByText cboCustType, .CustType
        Else
            lblCustName.caption = Helpers.FormatCaption(.Name)
            lblCustType(0).caption = .CustType
        End If
        
        lblCustID(0).caption = .ID
        'lblVaxAcct.Caption = GetVaxAcct(.BillAddr.AddrKey)
        'lblPassword.Caption = .Password
        
        SetCheckbox chkReqPO, .ReqPO
            
'load and filter the combobox based on Customer Default
'TODO: PRN 562 3/3/05 LR  a quick fix
        If m_oCustomer.IsTemp Then
            m_oCustomer.BillAddr.DefaultPmtTerms.Key = 36  'COD
        End If
        
        .BillAddr.DefaultPmtTerms.LoadFilteredComboBox cboTerms, .BillAddr.DefaultPmtTerms.Key
        
'****   When does this condition occur?  New order? New customer?
        If m_oOrder.PmtTerms.Key = 0 Then
            SetComboByKey cboTerms, .BillAddr.DefaultPmtTerms.Key
            m_oOrder.PmtTerms.Key = .BillAddr.DefaultPmtTerms.Key
        Else
            SetComboByKey cboTerms, m_oOrder.PmtTerms.Key
        End If

'11/9/04 LR
'When loading an order in edit mode
'(we got here via cmdLoadOrder_Click > StartTransition > TransitionTabs > UpdateOrderInfo)
'    Compare the order's pmtterms to the customer's default
'    if the customer's default is now more restricted then
'        adjust the order's terms and display a warning to the user

'the Not m_oOrder.IsNewOrder is a HACK to fix a problem in Order.Clear
'TODO: Is this needed since I put in the fix above? 11/10/04 LR
'
'loading existing MISC CrCard orders will force them back to COD because
'CrCard DueDayOrMonth > COD DueDayOrMonth

'is InStr(1, m_oCustomer.ID, "-MISC") equivalent to Order.IsMisc ?

        'this condition was added 7/30/2012
        'to handle the case where a CC order for a customer with default COD terms
        'is getting reset from CC to COD
        'The fix is - if it's a CC order, leave it a CC order, no matter what the default terms are
        If m_oOrder.PmtTerms.ID <> "CrCard" Then
            If m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit And Not m_oOrder.IsNewOrder Then
                If InStr(1, m_oCustomer.ID, "-MISC") = 0 Then
                    If m_oOrder.PmtTerms.DueDayOrMonth > m_oCustomer.BillAddr.DefaultPmtTerms.DueDayOrMonth Then
                        m_oOrder.PmtTerms.Key = m_oCustomer.BillAddr.DefaultPmtTerms.Key
                        SetComboByKey cboTerms, m_oCustomer.BillAddr.DefaultPmtTerms.Key
                        msg "The customer's terms have been restricted since you loaded this order last. The order's terms have been adjusted accordingly.", vbInformation
                    End If
                End If
            End If
        End If
        
        If m_oOrder.PmtTerms.ID = "CrCard" Then
            frmCreditCard.Visible = True
            'HACK: for the time being I'll use the same test the Status bar employs
            If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then
                cmdCCEdit.Visible = False 'view mode
            Else
                cmdCCEdit.Visible = True 'edit mode
            End If
            UpdateCCDisplay
        Else
            frmCreditCard.Visible = False
            cmdCCEdit.Visible = False
        End If
        
        m_oBrokenRules.EnableClass ccpurchaseorder, .ReqPO
        m_oBrokenRules.EnableClass ccorderedby, True

    End With

'These should be called when the Customer tab is selected
    UpdateBillAddrInfo
    UpdateShipAddrInfo

    LoadContactCombo
    UpdateShipContactInfo

    If InStr(1, m_oCustomer.ID, "-MISC") Then
        cmdEditAddr(1).Enabled = False
    Else
        cmdEditAddr(1).Enabled = True
    End If

    UpdateDefaultShipRemarks
    
    'SMR Intl 01/17/2006 - cboShipVia can change if default shipping method is reset.
    'Set/Remove the Ship To Contact Info controls (visible property & required validation).
    Call EnableShipToContactCtrls

    chkBillRecipient.Enabled = False
    chkBillRecipient.value = vbUnchecked
    txtUPSAcct.text = ""
    cmdUPSUpdate.Enabled = False

    If Not m_oOrder.PmtTerms.IsCOD Then
        If Left(Trim(cboShipVia.text), 3) = "UPS" And m_oOrder.Customer.HasAccount Then
            chkBillRecipient.Enabled = True
            If m_oOrder.UPSAcct <> "" Then
                txtUPSAcct.text = m_oOrder.UPSAcct
                chkBillRecipient.value = vbChecked
                cmdUPSUpdate.Enabled = True
            End If
        End If
   End If
   
   m_oOrder.UPSAcct = txtUPSAcct.text
    
    With gdxItems
        If .ItemCount = m_oItems.Count Then
            .Refetch
        Else
            .ItemCount = m_oItems.Count
            .Refetch
        End If
    End With
    txtTotalPrice.amount = m_oItems.TotalPrice
    txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)
      
    m_oBrokenRules.Validate
    m_bLoading = bLoading
End Sub


Private Sub LoadContactCombo()
    Dim oContact As contact
    
    'This is the logic that intializes the UI
    'If the customer has an account, enable cbobox
    'else enable txtbox
    
    'load the Contact combobox

    cboContact.Clear
    cmdEditContact.Enabled = False
    
    'if the customer has one or more contacts on record
    '(do customers without an account have contacts?)

    If m_oCustomer.Contacts.Count > 0 Then
        'add the customer's contact(s)
        For Each oContact In m_oCustomer.Contacts
            cboContact.AddItem oContact.Name
            cboContact.ItemData(cboContact.NewIndex) = oContact.Key
        Next
    Else
        'if the order has a contact, add it
        If Not m_oOrder.contact Is Nothing Then
            cboContact.AddItem m_oOrder.contact.Name
            cboContact.ItemData(cboContact.NewIndex) = m_oOrder.contact.Key
        End If
    End If
    
    ' If there's a contact, select it
    'what if name is not in the combobox?
    If Not m_oOrder.contact Is Nothing Then
        SetComboByText cboContact, m_oOrder.contact.Name
'***20110518 LR
'color combo background to indicate contact email status
        SetContactColorByEaddr cboContact, m_oOrder.contact
        
        AssignContactToOrder
    Else
        UpdateShipContactInfo
    End If
End Sub


'color combo background to indicate contact email status
Private Sub SetContactColorByEaddr(cbo As ComboBox, contact As contact)

    If contact.DeclinedEmailAddr Then
        cbo.BackColor = &H80FFFF
        cbo.ToolTipText = "declined to provide email addr"
    ElseIf Trim(contact.emailaddr) = "" Then
        cbo.BackColor = &HC0C0FF
        cbo.ToolTipText = "needs an email address"
    Else
        cbo.BackColor = &HC0FFC0
        cbo.ToolTipText = "has email address"
    End If

End Sub

Private Sub ClearContactColor(cbo As ComboBox)
    cbo.BackColor = &H80000005
    cbo.ToolTipText = ""
End Sub


Private Function HasOtherShipAddresses() As Boolean
    Dim oCmd As ADODB.Command

    If InStr(1, m_oCustomer.ID, "-MISC") Then
        'MISC accounts have no other viable addresses
        HasOtherShipAddresses = False
    Else
        Set oCmd = CreateCommandSP("spcpcHasOtherShipAddr")
        oCmd.Parameters("@_iCustKey").value = m_oCustomer.Key
        oCmd.Parameters("@_iAddrKey").value = m_oCustomer.ShipAddr.AddrKey
        oCmd.Execute
        If oCmd.Parameters("@_oCount").value > 0 Then
            HasOtherShipAddresses = True
        Else
            HasOtherShipAddresses = False
        End If
        Set oCmd = Nothing
    End If
End Function


'Private Function GetVaxAcct(lAddrKey As Long) As String
'    Dim rstVaxAcct As ADODB.Recordset
'    Dim sSQL As String
'
'    sSQL = "Select VaxAcct from tcpVaxAcct where addrkey = " & lAddrKey
'    Set rstVaxAcct = LoadDiscRst(sSQL)
'
'    If rstVaxAcct.RecordCount > 0 Then
'        GetVaxAcct = rstVaxAcct.Fields("VaxAcct").value
'    End If
'
'    Set rstVaxAcct = Nothing
'End Function


Private Sub UpdateDefaultShipRemarks()
    Dim lTemp As Long
    Dim sMethTemp As String
    Dim bLoading As Boolean
    
    bLoading = m_bLoading
    m_bLoading = True

    If Not m_oOrder.Customer.HasAccount Then
        chkDefaultShipMeth.Enabled = False
        chkDefaultShipMeth.value = vbUnchecked
    Else
        chkDefaultShipMeth.Enabled = True
        chkDefaultShipMeth.value = vbUnchecked
    End If

    If chkDefaultShipMeth.Enabled = True Then
        sMethTemp = Trim(GetShipRemarksMemo("CustPrefs", "Cust.Pref.ShipMeth", m_oCustomer.ID))
        With chkDefaultShipMeth
            If Len(sMethTemp) > 0 Then
                If m_oOrder.IsNewOrder And Not g_bWillCallUser Then
                    If CheckDefaultShipVia(sMethTemp, lTemp) Then
                        SetComboByKey cboShipVia, lTemp
                        m_oOrder.ShipMethKey = cboShipVia.ItemData(cboShipVia.ListIndex)
                    Else
                        MsgBox "Customer default shipping method " & sMethTemp & " is not found in the warehouse!", vbCritical + vbOKOnly, "Default Shipping Method"
                    End If
                End If
            End If
        End With
    
        If sMethTemp = Trim(cboShipVia.text) Then
            chkDefaultShipMeth.value = vbChecked
        Else
            chkDefaultShipMeth.value = vbUnchecked
        End If
        
         m_sDefaultShipMeth = sMethTemp
    End If
    
    m_bLoading = bLoading
End Sub


Private Function CheckDefaultShipVia(ByVal sDefaultShipMeth As String, ByRef lItemData As Long) As Boolean
    Dim lTemp As Long
    
    With cboShipVia
        For lTemp = 0 To .ListCount - 1
            If Trim(.list(lTemp)) = Trim(sDefaultShipMeth) Then
                lItemData = .ItemData(lTemp)
                CheckDefaultShipVia = True
                Exit Function
            End If
        Next
    End With
End Function


Private Sub txtCost_LostFocus()
    Dim dCost As Double

'***what's going on here?

    'added in case that setFocus fails 9/30/02 TeddyX
    On Error Resume Next
    
    With m_oItems.SelectedItem
        dCost = .Cost
        If dCost <> txtCost.amount Then
            If checkMargin(txtCost.amount, txtPrice.amount) Then
                TryToSetFocus txtCost
                'txtCost.SetFocus
            Else
                .Cost = txtCost.amount
                ItemUpdateControls
                txtPrice.SelStart = 0
                txtPrice.SelLength = Len(txtPrice.text)
                TryToSetFocus txtPrice
'                txtPrice.SetFocus
            End If
        End If
    End With
End Sub


Private Function checkMargin(dCost As Double, dPrice As Double) As Boolean
    If dCost = 0 Or dPrice = 0 Then Exit Function
    
    Dim dMarkup As Double
    
    dMarkup = (dPrice * 100) / dCost - 100
    If (dCost * minMarkUp(m_oCustomer.CustType)) > dPrice Then
        If vbNo = msg("Warning:the mark-up on this item is " & Format(dMarkup / 100, "0.00%") & " (low). Continue?", vbYesNo + vbExclamation, "Margin Warning") Then
            checkMargin = True
        End If
    ElseIf (dCost * maxMarkUp(m_oCustomer.CustType)) < dPrice Then
        If vbNo = msg("Warning:the mark-up on this item is " & Format(dMarkup / 100, "0.00%") & " (high). Continue?", vbYesNo + vbExclamation, "Margin Warning") Then
            checkMargin = True
        End If
    End If
End Function


Private Function minMarkUp(sCustType As String) As Double
    Select Case Trim(sCustType)
        Case "EndUser":
            minMarkUp = 1.5
        Case "Dealer":
            minMarkUp = 1.3
        Case "Wholesale":
            minMarkUp = 1.2
    End Select
End Function


Private Function maxMarkUp(sCustType As String) As Long
    Select Case Trim(sCustType)
        Case "EndUser":
            maxMarkUp = 10
        Case "Dealer":
            maxMarkUp = 5
        Case "Wholesale":
            maxMarkUp = 4
    End Select
End Function


'The only time txtCustID is enabled is on a NewCustomer order
'If the textbox is empty, we exit immediately

Private Sub txtCustID_LostFocus()
    Dim rst As Recordset
    Dim sCustID As String
    Dim oItem As IItem

'Is there a value in the control?
    sCustID = Trim(txtCustID.text)
    If Len(sCustID) = 0 Then Exit Sub

'Does it match an existing CustID in the database?
    Set rst = LoadDiscRst("Select CustKey, DfltBillToAddrKey, DfltShipToAddrKey from tarCustomer Where CustID = '" & sCustID & "' and status = 1")
    If rst.RecordCount <> 1 Then
        msg sCustID & " is not a valid customer account.", , "Invalid CustID"
        Set rst = Nothing
        Exit Sub
    End If

'morph the Order's Contact into a Cust Contact
    If Not m_oOrder.contact Is Nothing Then
        If m_oOrder.contact.OwnerType = opOrder Then
            m_oOrder.contact.OwnerType = opCustomer
            m_oOrder.contact.OwnerKey = rst.Fields("CustKey")
            m_oOrder.contact.Update
        End If
    End If

'Load the Customer and Address objects
    With m_oCustomer
        .Load rst.Fields("CustKey")
        .BillAddr.Load rst.Fields("DfltBillToAddrKey")
        .ShipAddr.Load rst.Fields("DfltShipToAddrKey")
        rvCustomer.OwnerID = .ID
        rvCustomer.Visible = True
    End With
    
'now relink us to the contact object in the Customer's collection
    If Not m_oOrder.contact Is Nothing Then
        Set m_oOrder.contact = m_oCustomer.Contacts.GetContactByKey(m_oOrder.contact.Key)
    End If

'Mark each item in the Items collection with the CustomerType
    For Each oItem In m_oItems
        oItem.CustType = m_oCustomer.CustType
    Next

    'Load Sales Tax
    m_oOrder.SalesTax.Init m_oCustomer
    If m_oOrder.IsWillCall Then
        m_oOrder.SalesTax.WillCallTaxOverride m_oOrder.whseid
    End If
        
'Update the OP Window caption to show the Customer Name
    UpdateOrderCaption
        
'Disable the input controls on the Customer tab
    SetCustCtrlVisible False

'    cmdEditAddr(0).Enabled = HasRight(k_sRightUpdateBillingAddr)
    
    UpdateOrderInfo
    m_oBrokenRules.Validate

'Indicate that the order has been assigned a CustomerID/Key
'This is the only place in the entire project that sets this property to True.
'It's the only place the property is used at all within the project.
'Various places within Order.cls itself clear it.
    m_oOrder.UpdateCustomer = True
    
    gdxItems.Refresh
    Set rst = Nothing
End Sub


Private Sub txtCustName_LostFocus()
    m_oCustomer.Name = txtCustName.text
End Sub


Private Sub txtFindCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtFindCust.text) > 0 Then
        If KeyCode = vbKeyReturn Then
'            cmdFindOrders2_Click
            FindOrdersByCriteria
        End If
    End If
End Sub


Private Sub txtFindText_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtFindText.text) > 0 Then
        If KeyCode = vbKeyReturn Then
'            cmdFindOrders2_Click
            FindOrdersByCriteria
        End If
    End If
End Sub


Private Sub txtFindOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'cmdFindOrders_Click
        FindOrderByNumber
    End If
End Sub


Private Sub txtItemDescr_LostFocus()
    'Sometimes, when the user deletes a SPO item, after returning
    'to ivList Mode, LostFocus was still called and causes a cascading error.
    
    If ViewMode = ivList Then Exit Sub
    
    If m_oItems.SelectedItem Is Nothing Then Exit Sub

    With m_oItems.SelectedItem
        If .Descr = txtItemDescr.text Then Exit Sub
        .Descr = txtItemDescr.text
        ValidateDescrLength
    End With
    ItemUpdateControls
End Sub


Private Sub ValidateDescrLength()
    With m_oValidateDescr
        .Valid = True
        If ViewMode = ivSpecialOrder Then
            Dim lCharsAvailable As Long
            
            With m_oItems.SelectedItem
                lCharsAvailable = 39 - Len(.Descr) - Len(.ItemID)
            End With
            
            If lCharsAvailable < 0 Then
                .ErrorMsg = "You must shorten this description by " _
                          & Abs(lCharsAvailable) & " characters."
                .Valid = False
            End If
        End If
    End With
End Sub


Private Sub txtItemPartNbr_LostFocus()
    Dim lItemKey As Long
    Dim eItemType As ItemTypeCode
    Dim sOriginalItemID As String
    Dim sRefSource As String
    Dim frmItemSearch As FItemSearch
    Dim oBTOItem As ItemBTOKit
    Dim oDropShipBTOItem As ItemBTOKit
    Dim oSPOItem As IItem
    
    Dim sMsg As String
    Dim sargs As String
    Dim Result As VbMsgBoxResult

    'Ignore the event if the text hasn't changed value
    'If Not txtItemPartNbr.Enabled Then Exit Sub
    If m_oItems.SelectedIndex <> m_lSelectedIndex Then Exit Sub
    If Trim(m_oItems.SelectedItem.ItemID) = Trim(txtItemPartNbr.text) Then Exit Sub
    '10/07/02 TeddyX
    'Sometimes, when the user delete a SPO item from the item. After returning
    'to ivList Mode, txtItemPartNbr_LostFocus was still called and caused cascading
    'error.
    If ViewMode = ivList Then Exit Sub

    'do this ahead of the len check so the user can erase a part number
    m_oItems.SelectedItem.ItemID = txtItemPartNbr.text
    ValidateDescrLength 'special check for SPO items
    ItemUpdateControls

    'if the field is blank, no look-up is needed
    If Len(txtItemPartNbr.text) = 0 Then
        Exit Sub 'no error on empty field
    End If

    Set frmItemSearch = New FItemSearch
    Load frmItemSearch

    Dim bCancelSearch As Boolean
    frmItemSearch.Find k_sPartNbr, txtItemPartNbr.text, lItemKey, eItemType, sOriginalItemID, sRefSource, bCancelSearch, cboWarehouse(1).ItemData(cboWarehouse(1).ListIndex), False, m_oOrder.Customer.CustTypeKey
    
    If bCancelSearch Then
'***DH 7/24/09
        'cmdItemCancel_Click
        Exit Sub
    End If

    If lItemKey > 0 Then
        Dim sType As String

        If eItemType = itFinishedGood Then
            sType = "stock part"
        Else
            sType = "build to order kit"
            If m_oOrder.IsDropShip Then
                eItemType = itSpecialOrder
                'Msg "A matching part number was found but it refers to a" & vbCrLf _
                   '& "BTO Kit which cannot be included on a drop ship order.", _
                   'vbOKOnly + vbExclamation, "BTO Kit"
                If vbYes = msg("A matching part number was found but it refers to a" & vbCrLf _
                   & "BTO Kit which cannot be included on a drop ship order." & vbCrLf _
                   & "Would you like to load it and convert to SPO for dropship? ", _
                    vbYesNo + vbQuestion, "Convert BTOKit to SPO for dropship?") Then
                        Set oDropShipBTOItem = New ItemBTOKit
                        oDropShipBTOItem.Load lItemKey
                    
                        '09/12/02 TeddyX
                        'Check if the BTOKit comes from the same vendor as DropShip vendor if it's a dropship order
                        If m_oOrder.DropShipVendKey > 0 And m_oOrder.DropShipVendKey <> oDropShipBTOItem.IItem_VendorKey Then
                            g_rstVendors.Filter = "VendKey = " & m_oOrder.DropShipVendKey
            
                            msg "This is a dropship order and the dropship vendor is " & RTrim(g_rstVendors.Fields("VendName").value) & ". " & _
                            "You can't order an item from " & vbCrLf & "different vendor for dropship order. You can either order item(s) " & _
                            "from dropship vendor or split this dropship order" & vbCrLf & " if you want to order items from different vendors."
            
                            g_rstVendors.Filter = adFilterNone
                        
                            Set oDropShipBTOItem = Nothing
                            Exit Sub
                        Else
                            '09/17/02 TeddyX
                            'Before morphing this BTO kit into a SPO item. Add remark to the selected item
                            'from remark ejector for later cloning remark purpose if user agrees.
'                            If CheckLineItemsRemark Then
'                                If vbYes = Msg("There are line item remarks on this SPO item." & vbCrLf _
'                                        & "Would you like to keep those remarks before morphing?", vbYesNo, "Keep line remarks?") Then
'                                        AddLineItemsRemark
'                                End If
'                            End If
                            
                            m_oItems.SelectedItem.MorphBTOKey = lItemKey
                            Set oBTOItem = MorphSPOToBTO(m_oItems.SelectedItem)
                            Set oSPOItem = MorphBTOtoSPO(oBTOItem)
                            oSPOItem.OriginalItemID = sOriginalItemID
                            oSPOItem.RefSource = sRefSource
                            m_oItems.Remove m_oItems.SelectedIndex
                            m_oItems.Add oSPOItem
                            m_oItems.Item(m_oItems.Count).Backup
                            ItemUpdateControls
                            rvOrderLine(3).RemarkContext = m_oItems.SelectedItem.RemarkContext
                            Exit Sub
                        End If
                Else
                    Exit Sub
                End If
            End If
        End If
    
        sMsg = "A matching part number was found." & vbCrLf _
                & "Would you like to convert this SPO into a " & sType & "?"
                    
        If m_oOrder.IsDropShip Then
            sMsg = sMsg & vbCrLf & "(This is a drop ship order)"
            Result = msg(sMsg, vbYesNo + vbDefaultButton2, "Convert SPO to " & sType & "?")
        Else
            Result = msg(sMsg, vbYesNo + vbDefaultButton1, "Convert SPO to " & sType & "?")
        End If
    
        If Result = vbYes Then
            'Check if the BTOKit comes from the same vendor as DropShip vendor if it's a dropship order
            Dim oItem As IItem
            
            'Before converting this SPO into a stocked item. Add remark to the selected item
            'from remark ejector for later cloning remark purpose if user agrees.
            If CheckLineItemsRemark Then
                If vbYes = msg("There are line item remarks on this SPO item." & vbCrLf _
                    & "Would you like to keep those remarks after converting?", vbYesNo, "Keep line remarks?") Then
                    AddLineItemsRemark
                End If
            End If
        
            Set oItem = m_oItems.SelectedItem
            m_oItems.Remove m_oItems.SelectedIndex
            If CheckandLoadItem(lItemKey, eItemType, sOriginalItemID, sRefSource) Then
                With m_oItems.SelectedItem 'now the new item
                    .MakeKey = oItem.MakeKey
                    .ModelNbr = oItem.ModelNbr
                    CloneRemarks oItem.RemarkContext, .RemarkContext
                    .Qty = oItem.Qty
                    .IsTaxable = oItem.IsTaxable
                    .SerialNbr = oItem.SerialNbr
                    .IsCGMPN = oItem.IsCGMPN
                    .BackNegotiatedPrice = oItem.BackNegotiatedPrice
                    .NegotiatedPrice = oItem.NegotiatedPrice
                    .Backup
                    ItemUpdateControls
                    'Update stocked item remark control if there are item remarks cloned from SPO item
                     rvOrderLine(0).RemarkContext = m_oItems.SelectedItem.RemarkContext
                    'txtCost.Enabled = False
                    'MDIMain.DoRefresh
                End With
            
                Set oItem = Nothing
                m_bNewItem = False 'do not delete new morphed item on cancel
            Else
                m_oItems.Add oItem
            End If
        End If
    End If
End Sub


Private Function MorphBTOtoSPO(ByRef oBTOItem As ItemBTOKit) As IItem
    Dim oItem As IItem
    
    Set oItem = New ItemFinGood
    With oItem
        .SageItemType = 1
        .Cost = 0
        .BackNegotiatedPrice = oBTOItem.IItem_BackNegotiatedPrice
        .CustType = oBTOItem.IItem_CustType
        .DealerPrice = oBTOItem.IItem_DealerPrice
        .Descr = oBTOItem.IItem_Descr
        .IsCGMPN = oBTOItem.IItem_IsCGMPN
        .IsTaxable = oBTOItem.IItem_IsTaxable
        .ItemID = oBTOItem.IItem_ItemID
        .ItemKey = 0
        .LineKey = oBTOItem.IItem_LineKey
        .ListPrice = oBTOItem.IItem_ListPrice
        .MakeKey = oBTOItem.IItem_MakeKey
        .ModelNbr = oBTOItem.IItem_ModelNbr
        .NegotiatedPrice = oBTOItem.IItem_NegotiatedPrice
        .OPKey = oBTOItem.IItem_OPKey
        .OPLineKey = oBTOItem.IItem_OPLineKey
        .Qty = oBTOItem.IItem_Qty
        .SerialNbr = oBTOItem.IItem_SerialNbr
        .StatusCode = oBTOItem.IItem_StatusCode
        .VendorKey = oBTOItem.IItem_VendorKey
        .WholesalePrice = oBTOItem.IItem_WholesalePrice
        .MorphBTOKey = oBTOItem.IItem_ItemKey
        .RemarkContext.Load "ViewOrderLine"
        CloneRemarks oBTOItem.IItem_RemarkContext, .RemarkContext
    End With
    
    Set MorphBTOtoSPO = oItem
End Function


Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSearch_Click
    End If
End Sub


Private Sub txtModel_LostFocus(Index As Integer)
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    With m_oItems.SelectedItem
        .ModelNbr = txtModel(Index)
    End With
End Sub


Private Sub txtSerial_LostFocus(Index As Integer)
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    With m_oItems.SelectedItem
        .SerialNbr = txtSerial(Index)
    End With
End Sub


Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    HandlePriceKeyCode txtPrice, KeyCode, Shift
    m_bRestoreDefPrice = (KeyCode = vbKeyDelete Or KeyCode = vbKeyBack) And (txtPrice.SelLength = Len(txtPrice.text) Or Len(txtPrice.text) = 2)
End Sub


Private Sub txtCost_KeyDown(KeyCode As Integer, Shift As Integer)
    HandlePriceKeyCode txtCost, KeyCode, Shift
End Sub


Private Sub txtPrice_LostFocus()
    Dim dPrice As Double
        
    '09/30/02 TeddyX
    'Add error handler here in case that setFocus fails
    On Error Resume Next
    
    With m_oItems.SelectedItem
        dPrice = .EffectivePrice
        
        If m_oOrder.HasPartsNoCharge Then
            .NegotiatedPrice = 0
        Else
            If m_bRestoreDefPrice Or txtPrice.amount < 0 Then
                .NegotiatedPrice = -1
            Else
                .NegotiatedPrice = txtPrice.amount
            End If
        End If
        
        m_bRestoreDefPrice = False
        txtPrice.amount = .EffectivePrice
        If dPrice <> txtPrice.amount Then
            If checkMargin(txtCost.amount, txtPrice.amount) Then
                TryToSetFocus txtPrice
'                txtPrice.SetFocus
                Exit Sub
            End If
        End If
        .BackNegotiatedPrice = txtPrice.amount
        txtExtPrice.amount = .ExtendedPrice
    End With
    txtTotalPrice.amount = m_oItems.TotalPrice
'7/25/05 LR
    txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)
    'txtQtyOrdered.SetFocus  ' removed 7/16/02 LR
End Sub


Private Sub txtQtyOrdered_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lIncrement As Long

    Select Case KeyCode
    Case vbKeyUp
        lIncrement = 1
    Case vbKeyDown
        lIncrement = -1
    End Select

    If lIncrement <> 0 Then
        KeyCode = 0
        With m_oItems.SelectedItem
            .Qty = .Qty + lIncrement
            If .Qty < 0 Then
                .Qty = 0
            End If
            txtQtyOrdered.value = .Qty
            txtExtPrice.amount = .ExtendedPrice
        End With
        m_oBrokenRules.Validate txtQtyOrdered
        ItemUpdateControls
    End If
End Sub


Private Sub txtQtyOrdered_LostFocus()
    If m_oItems.SelectedItem Is Nothing Then Exit Sub

    With m_oItems.SelectedItem
        .Qty = txtQtyOrdered.value
        txtExtPrice.amount = .ExtendedPrice
    End With
    m_oBrokenRules.Validate txtQtyOrdered
    
    txtTotalPrice.amount = m_oItems.TotalPrice
'7/25/05 LR
    txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)

    m_oBrokenRules.Validate
    cmdItemOK.Enabled = (m_oBrokenRules.MaskedCount(k_lItemControlMask) = 0)
    'ItemUpdateControls
End Sub


'Private Sub txtShfRemarks_LostFocus()
'    m_oShelf.remarks = txtShfRemarks.Text
'End Sub


'===================================================
' BrokenRule events

Private Sub m_oBrokenRules_HaveBrokenRules()
    MDIMain.UpdateToolbarStatus
End Sub

Private Sub m_oBrokenRules_NoBrokenRules()
    MDIMain.UpdateToolbarStatus
End Sub



'***********************************************************************************************
' Item summary grid
'***********************************************************************************************

Private Sub gdxItems_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    msg "Didn't expect this event, Add should be disabled", , "gdxItems_UnboundAddNew"
End Sub


Private Sub gdxItems_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    msg "Didn't expect this event, Delete should be disabled", , "gdxItems_UnboundDelete"
End Sub


Private Sub gdxItems_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim oItem As IItem

    If RowIndex <= 0 Then Exit Sub

    Set oItem = m_oItems(RowIndex)
    If oItem Is Nothing Then Exit Sub
    
    With oItem
        Values(1) = Trim(.ItemID)
        Values(2) = Trim(.Descr)
        Values(3) = .Qty
        Values(4) = .EffectivePrice
        Values(5) = .ExtendedPrice
        
        'VL 091015 Commented out
'        If Not IsEmpty(.QtyAvail(m_oOrder.WhseKey)) Then
'
'
'            If m_oOrder.StatusCode = ItemStatusCode.iscCommitted Or m_oOrder.StatusCode = ItemStatusCode.iscHasRMA Then
'                Values(6) = .QtyOnHand(m_oOrder.WhseKey) - .QtyOnSO(m_oOrder.WhseKey) + (m_oOSItemList.Item(RowIndex).QtyOpenToShip)
'            Else
'                Values(6) = .QtyAvail(m_oOrder.WhseKey)
'            End If
'        End If
        
        'VL 091015 Go directly to the property
        Values(6) = .QtyAvail(m_oOrder.WhseKey)
        
        Values(7) = .QtyOnHand(m_oOrder.WhseKey)
        Values(8) = .QtyOnPO(m_oOrder.WhseKey)
        Values(9) = .OPItemType
        '09/26/02 TeddyX
        'If this order has special research status,
        'use this special status to replace 'Need Research status'
        If .StatusCode = iscResearch Then
            If .ResearchStatus > irsResearchEmpty Then
                Values(10) = "*" & .ResearchStatus
            Else
                Values(10) = "*" & irsNeedResearch
            End If
        Else
            Values(10) = "#" & .StatusCode
        End If
        If .RemarkContext.RemarkList.Count > 0 Then
            Values(11) = -1
        Else
            Values(11) = 2 'use zero to see light off or something to Suppress lightbulb
        End If
        If Not IsEmpty(Values(6)) Then
            If Values(3) > Values(6) Then
                Values(12) = Values(3) - Values(6)
            Else
                Values(12) = 0
            End If
        Else
            Values(12) = 0
        End If
    End With
End Sub


Private Sub gdxItems_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim oItem As IItem

    If RowIndex <= 0 Then
        Exit Sub
    End If

    On Error Resume Next
    
    Set oItem = m_oItems(RowIndex)
    With oItem
        .Qty = Values(3)
        .NegotiatedPrice = Values(4)
    End With
    txtTotalPrice.amount = m_oItems.TotalPrice
'7/25/05 LR
    txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)
End Sub


Private Sub gdxItems_OLEStartDrag(Data As JSDataObject, AllowedEffects As Long)
    Dim sBuffer As String
    Dim oRemark As remark

    'If order is committed, do not allow drag
    If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then Exit Sub
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    Stat "OLEStartDrag (Enter)"
    
    With m_oItems.SelectedItem
        sBuffer = .Export.ExportString
        For Each oRemark In .RemarkContext.RemarkList
            sBuffer = sBuffer & k_sTextDelimiter & oRemark.TypeID & k_sTextDelimiter & oRemark.MemoText
            Stat "Exporting " & oRemark.MemoText
        Next
    End With
    
    Data.SetData sBuffer, vbCFText
    AllowedEffects = jgexDropEffectMove + jgexDropEffectCopy
    Stat "OLEStartDrag (Exit)"
End Sub


Private Sub gdxItems_OLESetData(Data As JSDataObject, DataFormat As Integer)
    Dim sBuffer As String
    Dim oRemark As remark
    
    Stat "OLESetData (Enter)"
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    With m_oItems.SelectedItem
        sBuffer = .Export.ExportString
        For Each oRemark In .RemarkContext.RemarkList
            sBuffer = sBuffer & k_sTextDelimiter & oRemark.TypeID & k_sTextDelimiter & oRemark.MemoText
            Stat "Exporting " & oRemark.MemoText
        Next
    End With
    
    If DataFormat = vbCFText Then
        Data.SetData sBuffer, vbCFText
    End If
    Stat "OLESetData (Exit)"
End Sub


Private Sub gdxItems_RowDrag(ByVal Button As Integer, ByVal Shift As Integer)
    Debug.Print "gdxItems_RowDrag " & Shift
    gdxItems.OLEDrag
End Sub


Private Sub gdxItems_OLEDragOver(Data As JSDataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    'Do not allow committed order to accept dragging & dropping order line items
    If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then Exit Sub
    
    If Shift = 0 Then
        Effect = jgexDropEffectMove
    Else
        Effect = jgexDropEffectCopy
    End If
End Sub


Private Sub gdxItems_OLEDragDrop(Data As JSDataObject, Effect As Long, _
    Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim sBuffer As String
    Dim vSegments As Variant
    Dim i As Long
    Dim oRemark As remark
    Dim sXML As String
    Dim oXML As JDMPDXML.XMLNode
    Dim sErrMsg As String
    Dim lVendorKey As String
    Dim lVendorBegin As Long
    Dim lVendorOver As Long
    
    'Do not allow drag/drop for committed orders
    If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then Exit Sub
    
    Stat "OLEDragDrop (Enter)"
    
    sBuffer = Data.GetData(vbCFText)
    vSegments = Split(sBuffer, k_sTextDelimiter)
    
    sXML = vSegments(0)
    lVendorBegin = InStr(1, sXML, "<VendorKey>", vbTextCompare)
    lVendorOver = InStr(1, sXML, "</VendorKey>", vbTextCompare)
    lVendorKey = CLng(Mid(sXML, lVendorBegin + 11, lVendorOver - lVendorBegin - 11))
    
    Set oXML = New JDMPDXML.XMLNode
    oXML.ImportString sXML

    'Deal with drop ship restrictions
    If m_oOrder.IsDropShip Then
        Select Case oXML.Tag
        Case "Gasket":      sErrMsg = "custom gasket"
        Case "Shelf":       sErrMsg = "wire shelf"
        Case "WWire":       sErrMsg = "warmer wire"
        Case "BTOKit":      sErrMsg = "BTO Kit"
        Case Else:          sErrMsg = ""
        End Select

        If Len(sErrMsg) > 0 And sErrMsg <> "BTO Kit" Then
            msg "Sorry, OP " & m_oOrder.OPKey _
              & " is set up for drop-shipment and may not include manfactured goods such as a " _
              & sErrMsg & ".", , "Selected item may not be added to a drop-ship order"
            Effect = jgexDropEffectNone
            Stat "OLEDragDrop (Exit 1)"
            Exit Sub
        Else
            'Check if order is dropship order and
            'the dragged item's vendor key is the same as dropship vendor
            'if not, throw out  9/13/02 TX
            If m_oOrder.DropShipVendKey > 0 And lVendorKey > 0 And lVendorKey <> m_oOrder.DropShipVendKey Then
                g_rstVendors.Filter = "VendKey = " & m_oOrder.DropShipVendKey
                msg "This is a dropship order and the dropship vendor is " & RTrim(g_rstVendors.Fields("VendName").value) & ". " & _
                "You can't add an item from " & vbCrLf & "different vendor for dropship order."
                g_rstVendors.Filter = adFilterNone
                Effect = jgexDropEffectNone
                Stat "OLEDragDrop (Exit 1)"
                Exit Sub
            End If
        End If
    
        'ask user if he wants to morph BTO kit to SPO items during drag and drop  9/16/02 TX
        If sErrMsg = "BTO Kit" Then
            If vbNo = msg("This item refers to a " _
                & "BTO Kit which cannot be included on a drop ship order." & vbCrLf _
                & "Would you like to load it and convert to SPO for dropship? ", _
                vbYesNo + vbQuestion, "Convert BTOKit to SPO for dropship?") Then
                Effect = jgexDropEffectNone
                Stat "OLEDragDrop (Exit 1)"
                Exit Sub
            End If
        End If
    End If
                
    m_oItems.ImportItem oXML, m_oOrder.WhseKey
    
    'morph BTO kit to SPO items during drag and drop  9/16/02 TX
    If m_oOrder.IsDropShip And sErrMsg = "BTO Kit" Then
        Dim oSPOItem As IItem
        Dim oItem As IItem
        Dim lIndex As Long
        
        Set oItem = m_oItems.SelectedItem
        lIndex = m_oItems.SelectedIndex
        
        Set oSPOItem = MorphBTOtoSPO(oItem)
        m_oItems.Remove (lIndex)
        m_oItems.Add oSPOItem
        CloneRemarks oItem.RemarkContext, m_oItems.Item(m_oItems.Count).RemarkContext
        m_oItems.Item(m_oItems.Count).Backup
    End If
    
    With m_oItems.SelectedItem
        .OPLineKey = database.GetSurrogateKey("tcpSOLine")
        .RemarkContext.RemarkList.Clear
        .RemarkContext.OwnerKey = .OPLineKey
        For i = 1 To UBound(vSegments) Step 2
            .RemarkContext.AddRemark vSegments(i), CStr(vSegments(i + 1))
            Stat "Importing " & vSegments(i + 1)
        Next
    End With
    
    If m_oOrder.IsDropShip Then
        m_oOrder.DropShipVendKey = m_oItems.SelectedItem.VendorKey
    End If
    
    gdxItems.ItemCount = m_oItems.Count
    gdxItems.Refetch
    SelectItemRow m_oItems.Count
    If Shift = 0 Then
        Effect = jgexDropEffectMove
    Else
        Effect = jgexDropEffectCopy
    End If
    
    'Update Price and Tax amount on the new form after item has been dropped onto the grid.
    txtTotalPrice.amount = m_oItems.TotalPrice
    txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)
    
    Me.SetFocus
    Stat "OLEDragDrop (Exit 2)"
    PromptShipComplete
End Sub


Private Sub gdxItems_OLECompleteDrag(Effect As Long)
    'Do not allow committed order to accept dragging & dropping order line items
    If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then Exit Sub
    
    Stat "OLECompleteDrag (Enter)"
    If (Effect And jgexDropEffectMove) = jgexDropEffectMove Then
        With m_oItems
            'Set dropship vendor to nothing after the item who decides dropship vendor is dropped from the order
            If m_oOrder.IsDropShip And CheckDropShipSelectedItem Then
                m_oOrder.DropShipVendKey = 0
            End If
            .Remove .SelectedIndex
            SyncItemList
        End With
    End If
    
    'Update Price and Tax amount on the old form after item has been dragged from the grid.
    txtTotalPrice.amount = m_oItems.TotalPrice
    txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)

    Stat "OLECompleteDrag (Exit)"
End Sub


Private Sub ItemUpdateControls()
    Dim eStatusCode As ItemStatusCode
    Dim eCurrentStatus As ItemStatusCode
    Dim eResearchStatus As ItemResearchStatus
    Dim oComboItem As ComboItem
    Dim lIndex As Long
    Dim lCurrentIndex As Long
        
    With m_oItems.SelectedItem
        Debug.Print "The Item ID is " & .ItemID
        txtItemPartNbr = .ItemID
        txtItemDescr = .Descr
        txtQtyOrdered = .Qty
        If m_oOrder.HasPartsNoCharge Then
            .NegotiatedPrice = 0
            If Not m_bChooseItem And .OPItemType <> itWireShelf And .OPItemType <> itSpecialOrder Then .BackNegotiatedPrice = -1
        Else
            If Not m_bChooseItem And .OPItemType <> itWireShelf And .OPItemType <> itSpecialOrder Then
                .NegotiatedPrice = -1
            Else
                .NegotiatedPrice = .BackNegotiatedPrice
            End If
        End If
        
        txtPrice.amount = .EffectivePrice
        txtExtPrice = .ExtendedPrice
        txtListPrice = .ListPrice
        txtCost.amount = .Cost

        If .IsCGMPN Then
            chkCGMPN.value = vbChecked
        Else
            chkCGMPN.value = vbUnchecked
        End If

        Select Case .OPItemType
        Case itFinishedGood
            FinGoodUpdateControls
        Case itSpecialOrder
            SpecialOrderUpdateControls
        Case itMoldedGasket
            GasketUpdateControls
        Case itWireShelf
            ShelfUpdateControls
        Case itWarmerWire
            WWireUpdateControls
        Case itBTOKit
            BTOKitUpdateControls
        End Select
    
        icbItemStatus.ComboItems.Clear
        
        .StatusCode = .StatusCode 'this will ensure current status is valid (Huh?)
        
        eCurrentStatus = .StatusCode
        For eStatusCode = iscResearch To ItemStatusCode.iscReadyToCommit
            If .IsValidStatusCode(eStatusCode) Then

                'Add new Research status to the combo box
                If eStatusCode = iscResearch Then
                    'Get rid of Need Research status from the status combo box
                    For eResearchStatus = irsContactFactory To irsWaitCustomer
'                    For eResearchStatus = irsNeedResearch To irsWaitCustomer
                       ' Set oComboItem = icbItemStatus.ComboItems.Add(, _
                            "#" + CStr(eResearchStatus), _
                            ResearchStatusString(eResearchStatus), _
                            eResearchStatus, _
                            eResearchStatus)
                        
                        lIndex = lIndex + 1
                        Set oComboItem = icbItemStatus.ComboItems.Add(lIndex, _
                            "*" + CStr(eResearchStatus), _
                            ResearchStatusString(eResearchStatus), _
                            iscResearch, _
                            iscResearch)
                        
                        'If the current item research status is general Need Research. Set it
                        'to Contact Factory as default
                        If .ResearchStatus = irsNeedResearch And eResearchStatus = irsContactFactory Then
                            oComboItem.Selected = True
                            lCurrentIndex = lIndex
                        ElseIf eResearchStatus = .ResearchStatus Then
                            oComboItem.Selected = True
                            lCurrentIndex = lIndex
                        End If
                    Next
                Else
                    lIndex = lIndex + 1
                    Set oComboItem = icbItemStatus.ComboItems.Add(lIndex, _
                            "#" + CStr(eStatusCode), _
                            StatusCodeString(eStatusCode), _
                            eStatusCode, _
                            eStatusCode)
                
                    If eStatusCode = .StatusCode Then
                        oComboItem.Selected = True
                        lCurrentIndex = lIndex
                    End If
                End If
            End If
        Next
        
        DoEvents
        
    End With
    
    txtTotalPrice.amount = m_oItems.TotalPrice
    txtTotalTax.amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)

    m_oBrokenRules.Validate
    If m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit Then
        cmdItemOK.Enabled = (m_oBrokenRules.MaskedCount(k_lItemControlMask) = 0) And Not (m_oOrder.StatusCode = ItemStatusCode.iscHasRMA)
        cmdNextGasket.Enabled = cmdItemOK.Enabled
    End If
End Sub


Private Sub FinGoodUpdateControls()
    UpdateMakeControls 2
End Sub


Private Sub BTOKitUpdateControls()
    UpdateMakeControls 1
End Sub


Private Sub SpecialOrderUpdateControls()
    m_oBrokenRules.EnableClass ccItemSPOBasicInfo, _
        Len(Trim(txtItemPartNbr.text)) = 0 And Len(Trim(txtItemDescr.text)) = 0
    m_oBrokenRules.Validate
    UpdateMakeControls 3
End Sub


Private Sub UpdateMakeControls(Index As Integer)
    With m_oItems.SelectedItem
        txtModel(Index).text = Trim(.ModelNbr)
        txtSerial(Index).text = Trim(.SerialNbr)
        SetComboByKey cboMake(Index), .MakeKey, True
        SetComboByKey cboVendor, .VendorKey, True
        lblVendor(1).caption = cboVendor.text
        lblVendor(2).caption = cboVendor.text
    End With
End Sub


'Called By
'  ItemUpdateControls

Private Sub GasketUpdateControls()
    With m_oGasket
        lenGasket(1).value = .width
        lenGasket(0).value = .Height
        
'        'update gasket type, if necessary
        If .IsMagnetic <> optGasketType(0).value Then
            If .IsMagnetic Then
                optGasketType(0).value = True
            Else
                optGasketType(1).value = True
            End If
        End If

        If .materialId <> cboGasket.ItemData(cboGasket.ListIndex) Then
            SetComboByKey cboGasket, .materialId, True
        End If

        If .materialId = 0 Then
            optGasketSides(1).Enabled = False
            optGasketSides(1).value = False
            optGasketSides(0).value = True

            DisableCheckbox chkGasketOptions(0) 'dart-to-dart
            DisableCheckbox chkGasketOptions(1) 'inverted
            DisableCheckbox chkGasketOptions(2) 'no magnet LHS
            DisableCheckbox chkGasketOptions(3) 'no magnet RHS
            txtItemDescr.text = ""
        Else
            ' 3 or 4 sided
            optGasketSides(1).Enabled = (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
            If (.Options And k_lGasketThreeSided) > 0 Then
                optGasketSides(1).value = True
            Else
                optGasketSides(0).value = True
            End If
            
            ' Dart-to-Dart
            If .IsDart Then
                chkGasketOptions(0).Enabled = (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
            Else
                chkGasketOptions(0).Enabled = True
                chkGasketOptions(0).value = vbUnchecked
                chkGasketOptions(0).Enabled = False
            End If
            
            ' Inverted
            chkGasketOptions(1).Enabled = (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
            
            If .IsMagnetic Then
                chkGasketOptions(2).Enabled = (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
                chkGasketOptions(3).Enabled = (m_oOrder.StatusCode <= ItemStatusCode.iscReadyToCommit)
            Else
                chkGasketOptions(2).Enabled = True
                chkGasketOptions(2).value = vbUnchecked
                chkGasketOptions(2).Enabled = False
                chkGasketOptions(3).Enabled = True
                chkGasketOptions(3).value = vbUnchecked
                chkGasketOptions(3).Enabled = False
            End If
        End If
                
        lblGasketMatlUsed.caption = .MoldedBy
       
        Dim bLoading As Boolean
        bLoading = m_bLoading
        m_bLoading = True

        If (m_oGasket.Options And k_lGasketDartToDart) > 0 Then
            chkGasketOptions(0).value = vbChecked
        Else
            chkGasketOptions(0).value = vbUnchecked
        End If
        
        If (m_oGasket.Options And k_lGasketInverted) > 0 Then
            chkGasketOptions(1).value = vbChecked
        Else
            chkGasketOptions(1).value = vbUnchecked
        End If
        
        If (m_oGasket.Options And k_lGasketNoMagLHS) > 0 Then
            chkGasketOptions(2).value = vbChecked
        Else
            chkGasketOptions(2).value = vbUnchecked
        End If

        If (m_oGasket.Options And k_lGasketNoMagRHS) > 0 Then
            chkGasketOptions(3).value = vbChecked
        Else
            chkGasketOptions(3).value = vbUnchecked
        End If

        m_bLoading = bLoading
    End With

    m_oBrokenRules.Validate cboGasket
End Sub


Private Sub ShelfUpdateControls()
    Dim i As Long
    Dim bLoading As Boolean
    
    With m_oShelf
'        txtShfRemarks.Text = .remarks
        lenShelfDepth.value = .Depth
        lenShelfWidth.value = .width
        SetComboByText cboFrame, .FrameText
        SetComboByText cboFinish, .FinishText

        bLoading = m_bLoading
        m_bLoading = True
        For i = 0 To 4
            If (2 ^ i) And .Options Then
                chkShelfOpt(i).value = vbChecked
            Else
                chkShelfOpt(i).value = vbUnchecked
            End If
        Next
        m_bLoading = bLoading
    End With
End Sub


Private Sub WWireUpdateControls()
    Dim i As Long
    Dim bIsValid As Boolean
        
    With m_oWarmerWire
        lenWireLength.value = .TotalInches
        lenDoorHeight.value = .DoorHeight
        lenDoorWidth.value = .DoorWidth

        cboVoltage.text = CStr(.Voltage)
        
        lblAmperage.caption = Format$(.Amperage, k_sAmpMask)
        lblWattsPerFoot.caption = Format$(.WattsPerFoot, k_sWPFMask)

        cboWires.Clear
        On Error GoTo Cleanup 'ignore error if list is empty
        
        For i = 1 To UBound(.SafeWires)
            cboWires.AddItem .SafeWires(i)
        Next
        
        SetComboByText cboWires, .OhmsPerFoot
    End With
    m_oBrokenRules.Validate cboWires

Cleanup:
    'nothing to do
End Sub
    

'Who reads and writes these variables?
'   m_bNewItem
'   m_lSelectedIndex

Private Sub AddItem(ByVal i_oItem As IItem)
    
    m_bNewItem = True
    
    i_oItem.CustType = m_oCustomer.CustType     'estab pricing?

    m_oItems.Add i_oItem
    
    m_lSelectedIndex = m_oItems.SelectedIndex

    'select the appropriate wizard
    
    Select Case i_oItem.OPItemType
        Case itFinishedGood
            ViewMode = ivComponent
        Case itBTOKit
            ViewMode = ivKit
        Case itMoldedGasket
            ViewMode = ivGasket
        Case itWireShelf
            ViewMode = ivShelf
        Case itWarmerWire
            ViewMode = ivWire
        Case itSpecialOrder
            ViewMode = ivSpecialOrder
        Case Else
            msg "Unexpected item type", , "Error in AddItem"
    End Select
End Sub


'Called By
'   cmdSearch_Click()
'   txtItemPartNbr_LostFocus()

Private Function CheckandLoadItem(ByVal i_lItemKey As Long, ByVal i_eItemType As ItemTypeCode, ByVal i_sOriginalItemID As String, ByVal i_sRefSource As String) As Boolean
    Dim oItem As IItem
    
    Select Case i_eItemType
        Case itFinishedGood
            Dim oFinGood As ItemFinGood
            Set oFinGood = New ItemFinGood
    
            oFinGood.Load i_lItemKey, m_oOrder.WhseKey
            
            'Check if this item's vend is different from DropShipVendor
            If m_oOrder.IsDropShip And m_oOrder.DropShipVendKey > 0 And oFinGood.IItem_VendorKey <> m_oOrder.DropShipVendKey Then
                g_rstVendors.Filter = "VendKey = " & m_oOrder.DropShipVendKey
                
                msg "This is a dropship order and the dropship vendor is " & RTrim(g_rstVendors.Fields("VendName").value) & ". " & _
                "You can't order an item from " & vbCrLf & "different vendor for dropship order. You can either order item(s) " & _
                "from dropship vendor or split this dropship order" & vbCrLf & " if you want to order items from different vendors."
                
                g_rstVendors.Filter = adFilterNone
                Exit Function
            Else
                AddItem oFinGood
            End If
    
        Case itBTOKit
            Dim oBTOKit As ItemBTOKit
            Set oBTOKit = New ItemBTOKit
            oBTOKit.Load i_lItemKey, m_oOrder.WhseKey
            AddItem oBTOKit
        
        Case Else
            msg "Error - Unsupported Item Type in LoadItem: " & i_eItemType
            Exit Function
    End Select

    Set oItem = m_oItems.SelectedItem
    
    oItem.OriginalItemID = i_sOriginalItemID
    oItem.RefSource = i_sRefSource
    
    If oItem.VendorKey = 0 Or oItem.Cost = 0 Then
        msg "This item is missing vendor and/or cost information." & vbCrLf _
          & "Please report this problem to the IT department so we can correct" & vbCrLf _
          & "the database.  You will not be able to commit this order" & vbCrLf _
          & "to Acuit until the database is corrected.", vbExclamation + vbOKOnly, "Incomplete Database Information"
    End If
    
    CheckandLoadItem = True
End Function


Private Function HandlePriceKeyCode(Ctl As Control, KeyCode As Integer, Shift As Integer) As Double
    Dim dblIncrement As Double

    Select Case KeyCode
    Case vbKeyUp
        If Shift = 0 Then
            dblIncrement = 1
        Else
            dblIncrement = 0.01
        End If
        Ctl.amount = Ctl.amount + dblIncrement
        KeyCode = 0
    Case vbKeyDown
        If Shift = 0 Then
            dblIncrement = 1
        Else
            dblIncrement = 0.01
        End If
        If Ctl.amount > dblIncrement Then
            Ctl.amount = Ctl.amount - dblIncrement
        Else
            Ctl.amount = 0
        End If
        KeyCode = 0
    End Select
End Function


Private Sub UpdateWarehouseColor(Index As Integer)
    Dim lColor As Long
    
    With cboWarehouse(Index)
        If .ItemData(.ListIndex) = GetWhseKeyFromBranchID(GetBranchIDFromUserKey(m_oOrder.UserKey)) Then
            lColor = RGB(255, 255, 255)
        Else
            lColor = RGB(255, 255, 0)
        End If
    End With
    
    cboWarehouse(0).BackColor = lColor
    cboWarehouse(1).BackColor = lColor
    cboWarehouse(2).BackColor = lColor
End Sub


Private Sub UpdateInventoryInfo()
    Dim lWhseKey As Long
        
    With cboWarehouse(2)
        lWhseKey = .ItemData(.ListIndex)
    End With

    'This routine is called as a side-effect of initializing
    'the warehouse combo boxes while getting ready to edit an order.
    'Since there is no selected item at that time,
    'we need to get out before checking QtyAvail properties.
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    With m_oItems.SelectedItem
    
'*** VL 091015 changed QtyAvail source
'        If m_oOrder.StatusCode = ItemStatusCode.iscCommitted Or m_oOrder.StatusCode = ItemStatusCode.iscHasRMA Then
'            lblQtyAvail.caption = .QtyOnHand(lWhseKey) - .QtyOnSO(lWhseKey) + (m_oOSItemList.Item(m_oItems.SelectedIndex).QtyOpenToShip)
'        Else
'            lblQtyAvail.caption = .QtyAvail(lWhseKey)
'        End If
        'Always go direct to the property since the calculation happens in the SQL
        lblQtyAvail.caption = .QtyAvail(lWhseKey)
'***
'*** LR 101915 change count displays and formatting
'removed these
'        lblQtyOH.caption = .QtyOnHand(lWhseKey)
'        lblQtySO.caption = .QtyOnSO(lWhseKey)
'        lblQtyBO.caption = .QtyOnBO(lWhseKey)

'       use these to cast the item
        Dim component As ItemFinGood
        Dim kit As ItemBTOKit
        
        'default values
        lblQtyAvail.ToolTipText = ""
        lblQtyAvail.ForeColor = &H228B22    'Green
        
        If m_oItems.SelectedItem.QtyAvail(lWhseKey) = 0 Then
            lblQtyAvail.ForeColor = &HFF&   'Red
        ElseIf m_oItems.SelectedItem.OPItemType = ItemTypeCode.itFinishedGood Then
            Set component = m_oItems.SelectedItem
            If component.IItem_InConflict(lWhseKey) Then
                lblQtyAvail.ForeColor = &HFF&
                lblQtyAvail.ToolTipText = "Item is in conflict"
            End If
            Set component = Nothing
        ElseIf m_oItems.SelectedItem.OPItemType = ItemTypeCode.itBTOKit Then
            Set kit = m_oItems.SelectedItem
            If kit.IItem_InConflict(lWhseKey) Then
                lblQtyAvail.ForeColor = &HFF&
                lblQtyAvail.ToolTipText = "Kit has a component in conflict"
            End If
            Set kit = Nothing
        End If
        
        lblQtyPO.caption = .QtyOnPO(lWhseKey)
        
    End With
End Sub


Private Sub SelectItemRow(ByVal i_lRow As Long)
    Dim lRowIndex As Long
    Dim i As Long
    
    For i = m_oItems.Count To 1 Step -1
        lRowIndex = gdxItems.RowIndex(i)
        If lRowIndex = i_lRow Then
            gdxItems.Row = i
            Exit Sub
        End If
    Next
End Sub


Private Sub SyncItemList()
    Dim i As Long
    
    With gdxItems
        If .ItemCount <> m_oItems.Count Then
            .ItemCount = m_oItems.Count
        End If
        .Refetch

        'if no item is selected but a row is selected, sync to that
        If m_oItems.SelectedIndex = 0 Then
            m_oItems.SelectedIndex = .RowIndex(.Row)
        End If

        'Ensure that the grid row is in sync with the items collection
        For i = 1 To gdxItems.RowCount
            If .RowIndex(i) = m_oItems.SelectedIndex Then
                .Row = i
            End If
        Next
    End With
End Sub


Private Sub Stat(sContext As String)
    With m_oItems
        Debug.Print sContext & " (" & m_lWindowID & "): " & .SelectedIndex & " of " & .Count
    End With
End Sub


Private Sub m_gwItems_RowChosen()
    If m_oItems.SelectedIndex <= 0 Then
        Exit Sub
    End If
    
    m_lSelectedIndex = m_oItems.SelectedIndex
    SetWaitCursor True
    m_bChooseItem = True
    chkCGMPN.Visible = False
    
    ' Only Finished Good and BTO Kit have price history
    cmdPriceHistory.Visible = False
    Select Case m_oItems.SelectedItem.OPItemType
    Case itFinishedGood
        cmdPriceHistory.Visible = True
        chkCGMPN.Visible = True
        ViewMode = ivComponent
    Case itBTOKit
        cmdPriceHistory.Visible = True
        chkCGMPN.Visible = True
        ViewMode = ivKit
    Case itWarmerWire
        ViewMode = ivWire
    Case itWireShelf
        ViewMode = ivShelf
    Case itMoldedGasket
        ViewMode = ivGasket
    Case itSpecialOrder
        chkCGMPN.Visible = True
        ViewMode = ivSpecialOrder
    Case Else
        msg "Unexpected Item subclass"
    End Select
    m_bChooseItem = False
    SetWaitCursor False
End Sub


Private Sub m_gwCustOrders_RowChosen()
    SetWaitCursor True
    cmdLoadOrder_Click (0)
    SetWaitCursor False
End Sub


Private Sub m_gwOrders_RowChosen()
    SetWaitCursor True
    cmdLoadOrder_Click (1)
    SetWaitCursor False
End Sub


Private Function AllItemsOnHand() As Boolean
    Dim oItem As IItem
    
    AllItemsOnHand = True
    
    For Each oItem In m_oItems
        If oItem.SageItemType = 5 Or oItem.SageItemType = 7 Then
            If IsEmpty(oItem.QtyOnHand(m_oOrder.WhseKey)) Then
                AllItemsOnHand = False
                msg "Item " & oItem.ItemID & " is invalid for this warehouse.", vbExclamation + vbOKOnly, "Invalid Item"
            End If
        End If
    Next
        
End Function


Private Function PossiblePOB(addr As Address) As Boolean
    Dim regex As RegExp
    PossiblePOB = False

    Set regex = New RegExp
    regex.Pattern = "\b[P|p]?(OST|ost)?\.?\s*[O|o|0]?(ffice|FFICE)?\.?\s*[B|b][O|o|0]?[X|x]?\.?\s+[#]?(\d+)\b"

    If regex.Test(addr.Addr1 & " " & addr.Addr2) Then
        PossiblePOB = True
    End If
End Function


Private Sub AddStatusBarPanel()
    Dim lIndex
    
    sbOrderStatus.Panels.Clear
    
    With sbOrderStatus.Panels
        For lIndex = 1 To 7
            .Add lIndex
            .Item(lIndex).Bevel = sbrNoBevel
            
            If lIndex = 1 Then
                .Item(lIndex).width = 2400
            ElseIf lIndex = 2 Then
                .Item(lIndex).width = 1000
            ElseIf lIndex = 3 Then
                .Item(lIndex).width = 1100
            ElseIf lIndex = 4 Then
                .Item(lIndex).width = 1100
            ElseIf lIndex = 6 Then
                .Item(lIndex).width = 1400
            Else
                .Item(lIndex).width = 1400
            End If
        Next
    End With
End Sub


Private Sub SetOrderStatusBar()
    Dim lStatus As Long
    lStatus = m_oOrder.StatusCode
    
    Select Case lStatus
            Case ItemStatusCode.iscResearch: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(1).Picture
            Case ItemStatusCode.iscQuote: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(2).Picture
            Case ItemStatusCode.iscAuthorize: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(3).Picture
            Case ItemStatusCode.iscReadyToCommit: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(4).Picture
            Case ItemStatusCode.iscEmpty: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(5).Picture
            Case ItemStatusCode.iscCommitted: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(6).Picture
            Case ItemStatusCode.iscDeleted: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(7).Picture
            Case ItemStatusCode.iscARHold: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(9).Picture
            Case ItemStatusCode.iscPendingCommit: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(9).Picture
            Case ItemStatusCode.iscHasRMA: sbOrderStatus.Panels(1).Picture = imglStatus16.ListImages(10).Picture
            Case Else: sbOrderStatus.Panels(1).Picture = Nothing
    End Select
    sbOrderStatus.Panels(1).text = StatusCode
End Sub


Private Sub LoadLineItemRemark()
    Dim oRemarkType As RemarkType
    Dim lIndex As Long
    Dim lRemarkTypeIndex As Long
    Dim bLoad As Boolean
    Dim lControlIndex As Long
    Dim lMFIndex As Long
    
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    bLoad = m_bLoading
    m_bLoading = True
    
    Select Case m_oItems.SelectedItem.OPItemType
        Case ItemTypeCode.itMoldedGasket:
            lControlIndex = 1
        Case ItemTypeCode.itSpecialOrder
            lControlIndex = 2
        Case ItemTypeCode.itWireShelf
            lControlIndex = 0
        Case Else
            lControlIndex = -1
    End Select
    rvOrderLine(0).Visible = (lControlIndex = -1)
    
    If lControlIndex = -1 Then
        m_bLoading = bLoad
        Exit Sub
    End If

    lIndex = 0
    cboMMType(lControlIndex).Clear
    For Each oRemarkType In m_oItems.SelectedItem.RemarkContext
        lIndex = lIndex + 1
        If oRemarkType.CanCreate Then
            cboMMType(lControlIndex).AddItem oRemarkType.caption
            cboMMType(lControlIndex).ItemData(cboMMType(lControlIndex).NewIndex) = lIndex
            If Left(Trim(oRemarkType.caption), 13) = "Manufacturing" Then
                lMFIndex = lIndex - 1
            End If
        End If
    Next
    
    If cboMMType(lControlIndex).ListCount > 0 Then
        If lControlIndex = 1 Then
            cboMMType(lControlIndex).ListIndex = lMFIndex
        Else
            If lControlIndex <> 2 Then
                cboMMType(lControlIndex).ListIndex = 0
            Else
                cboMMType(lControlIndex).ListIndex = 1 ' PRN#150 Private is default for SPO
            End If
        End If
        txtMMRemark(lControlIndex).Enabled = True
        cboMMType(lControlIndex).Enabled = True
    Else
        txtMMRemark(lControlIndex).Enabled = False
        cboMMType(lControlIndex).Enabled = False
    End If
    txtMMRemark(lControlIndex).text = ""
    rvOrderLine(lControlIndex + 1).RemarkContext = m_oItems.SelectedItem.RemarkContext
    rvOrderLine(0).Visible = False
    ReDim m_sLineRemark(lIndex) As String
    
    MDIMain.DoRefresh
    m_bLoading = bLoad
End Sub


Private Sub txtMMRemark_LostFocus(Index As Integer)
    If cboMMType(Index).ListCount > 0 Then
        m_sLineRemark(cboMMType(Index).ItemData(cboMMType(Index).ListIndex)) = txtMMRemark(Index).text
    End If
End Sub


Private Sub cboMMType_Click(Index As Integer)
    If m_bLoading Then Exit Sub
    
    Dim bLoad As Boolean
    bLoad = m_bLoading
    
    m_bLoading = True
    txtMMRemark(Index).text = m_sLineRemark(cboMMType(Index).ItemData(cboMMType(Index).ListIndex))
    m_bLoading = bLoad
End Sub


Private Sub AddLineItemsRemark()
    Dim lIndex As Long
    Dim bSave As Boolean
    
    If m_oItems.SelectedItem Is Nothing Then Exit Sub
    
    With m_oItems.SelectedItem
        If .OPItemType <> itSpecialOrder And _
            .OPItemType <> itMoldedGasket And _
            .OPItemType <> itWireShelf Then
            Exit Sub
        End If
    End With
    
    For lIndex = 1 To UBound(m_sLineRemark)
        If Trim(m_sLineRemark(lIndex)) <> "" Then
            m_oItems.SelectedItem.RemarkContext.AddRemark m_oItems.SelectedItem.RemarkContext(lIndex).TypeID, m_sLineRemark(lIndex)
        End If
    Next
     
End Sub


Private Function CheckLineItemsRemark()
    Dim lIndex As Long
    
    For lIndex = 1 To UBound(m_sLineRemark)
        If Trim(m_sLineRemark(lIndex)) <> "" Then
            CheckLineItemsRemark = True
            Exit Function
        End If
    Next
End Function


'This function is used to check if all requirements are met before OA sends the order
'to Enterprise Monitor for background commit

'Called By:
'   Commit

'Private Function CommitCheck() As VbMsgBoxResult
'    Dim oItem As IItem
'    Dim ballspo As Boolean
'    Dim bCreditCard As Boolean
'    Dim bCustomerHold As Boolean

Private Function CommitCheck() As Boolean

    'Nothing to do if user does not have permission to save
'    If Not HasRight(k_sRightOPSaveOrder) Then Exit Function

    'Just save without releasing if any of the following are true:
    '1. Order is not ready to be released
    '2. Order already has been released
    '3. User does not have permission to release order
    If Not (m_oOrder.Items.StatusCode = ItemStatusCode.iscReadyToCommit _
            And m_oOrder.soKey = 0 _
            And HasRight(k_sRightOPReleaseOrder)) Then
        m_oOrder.Save
        CommitCheck = False
        Exit Function
    End If
    
    'Check that this is not a Temp customer.  If so, display user note
    
    If m_oOrder.Customer.Key = 0 Then
        msg "Note: You must assign an valid CustID before you can commit order.", vbInformation + vbOKOnly, "Need CustID To Save Order"
        m_oOrder.Save
        CommitCheck = False
        Exit Function
    End If

    'otherwise, don't save the order here
    CommitCheck = True

End Function


'Send Order to AutoCommit or ARHold
'Called By
'   CommitButton

Private Function CommitOrder(ByRef OnARHold As Boolean) As Boolean
    
    OnARHold = False
    CommitOrder = True
    
    On Error GoTo EH
    
    SetWaitCursor True
    
    If m_oOrder.Customer.Hold Then

        'place the order on AR credit hold
        m_oOrder.bARCustHold = True
        m_oOrder.Save

        If GetUserWhseID(m_oOrder.UserKey) = "STL" And _
            m_oOrder.ShipMethod Like "*Will Call" Then
            
            EMail.Send _
                GetUserID & "@caseparts.com", _
                "joannar@caseparts.com", _
                "WillCall order " & m_oOrder.OPKey & " for " & m_oOrder.Customer.ID & " is On Hold [" & g_DB.server & "]", _
                GetUserName & " committed OP " & m_oOrder.OPKey & " for " & m_oOrder.Customer.ID & ": " & m_oOrder.Customer.Name, _
                False

        End If
        
        'check the collector assigned to the customer
        'if it's InCollections or WriteOff, notify accounting
        If m_oCustomer.Collector = "InCollection" Or _
            m_oCustomer.Collector = "WriteOff" Then
            
            EMail.Send _
                GetUserID & "@caseparts.com", _
                "pams@caseparts.com", _
                "Attention: Order on Hold", _
                m_oOrder.userid & " placed an order for " & m_oCustomer.ID & ": " & m_oCustomer.Name & " who's in collection or has been written off.", _
                False
        
        End If
        
        OnARHold = True

    Else
    
        m_oOrder.Commit
        
    End If

    ClearWaitCursor
    Exit Function
    
EH:
    ClearWaitCursor
    DisplayWarning "Unexpected error in queung the order to AutoCommit."
    CommitOrder = False

End Function


'************************************************************************************
' RMA tab
'************************************************************************************

'------------------------------------------------------------------------------------
' RMA Lines subtab
'------------------------------------------------------------------------------------

Private Sub cmdAddMoreItem_Click()
    Dim oFrm As FRMACreate
    
    Set oFrm = New FRMACreate
    oFrm.AddRMAItem m_oOrder, m_lRMAKey
    cmdRMARefresh_Click
End Sub


'Save changes after User modifies grid

Private Sub cmdUpdateRMALine_Click()
    Dim cmd As ADODB.Command
    Dim lIndex As Long

    'why would the collection be empty?

    If m_oRMALine.Count < 1 Then
        cmdUpdateRMALine.Enabled = False
        Exit Sub
    End If
    
    SetWaitCursor True
    
    'for each row in the collection, update the database

    With m_oRMALine
        For lIndex = 1 To .Count
            Set cmd = CreateCommandSP("spcpcRMAUpdateLine")
            cmd.Parameters("@_iRMALineKey").value = .Item(lIndex).RmaLineKey
            cmd.Parameters("@_iReasonCode").value = .Item(lIndex).Reason
            cmd.Parameters("@_iRestock").value = .Item(lIndex).Restock
            cmd.Parameters("@_iAuthQuantity").value = .Item(lIndex).QtyAuthorized
            cmd.Parameters("@_iCreditFreight").value = .Item(lIndex).CreditFreight
            cmd.Parameters("@_iReturnToVendor").value = .Item(lIndex).ReturnToVendor
            cmd.Parameters("@_iVendorRMANumber").value = .Item(lIndex).VendorRMANumber
            cmd.Parameters("@_iDaysNoPenalty").value = .Item(lIndex).DaysNoPenalty
            cmd.Execute
            Set cmd = Nothing
        Next
    End With
    
    'why reload the grid at this point?
    LoadRMALine m_lRMAKey

    'essentially clear the dirty state
    cmdUpdateRMALine.Enabled = False
    
    SetWaitCursor False
    
End Sub


Private Sub cmdRMARefresh_Click()
    LoadRMALine m_lRMAKey
End Sub


Private Sub cmdRMAVendor_Click()

    'Grid column 19 is VendKey
    If gdxRMALine.value(19) = Empty Or gdxRMALine.value(19) = 0 Then Exit Sub

    DisplayVendorInfo gdxRMALine.value(19)
    
End Sub


'Columns changes will enable the save button of RMA line
Private Sub gdxRMALine_AfterColEdit(ByVal ColIndex As Integer)
   gdxRMALine.Update
End Sub


'Load data from RMA line list to grid
Private Sub gdxRMALine_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    
    If m_oRMALine Is Nothing Then Exit Sub
    
    If RowIndex > m_oRMALine.Count Then Exit Sub
    
    With m_oRMALine.Item(RowIndex)
        Values(1) = .SOLineKey
        Values(2) = .RmaLineKey
        Values(3) = .ItemID
        Values(4) = .Cost
        Values(5) = .Price
        Values(6) = .QtyAuthorized
        Values(7) = .AuthBy
        Values(8) = .AuthDate
        Values(9) = .QtyPreRcvd
        Values(10) = .QtyPreCred
        Values(11) = .Disposition
        Values(12) = .Restock * 100
        Values(13) = .OPLineKey
        Values(14) = .Reason
        Values(15) = .CreditFreight
        'for RMA Vendor return
        Values(16) = .ReturnToVendor
        Values(17) = .VendorRMANumber
        Values(18) = .DaysNoPenalty
        Values(19) = .VendKey
    End With
End Sub


'By default the Janus grid does not update the recordset until you move off the record.
Private Sub gdxRMALine_AfterColUpdate(ByVal ColIndex As Integer)
    gdxRMALine.Update
End Sub


Private Sub gdxRMALine_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
       
    'update the recordset with the values in the grid
    With m_oRMALine.Item(RowIndex)
        
        'Values(12) is restocking fee
        If Not IsNumeric(Values(12)) Then
            .Restock = 0
        ElseIf Values(12) <= 100 And Values(12) >= 0 Then
            .Restock = Values(12) / 100
        End If
        
        .Reason = Values(14)
        .CreditFreight = Values(15)
        .ReturnToVendor = Values(16)
        .VendorRMANumber = Values(17)
        
        If IsNumeric(Values(18)) Then
            .DaysNoPenalty = Values(18)
        End If
        
        If Not IsNumeric(Values(6)) Then
            msg "The value in the Qty Authorized field is not valid.", vbOKOnly + vbExclamation, "Edit Authorization Quantity"
        Else
            If Values(6) > .lMaxQtyAuth Then
                msg "Qty Authorized must be less than or equal to " & .lMaxQtyAuth, vbExclamation, "Edit Authorization Quantity"
            ElseIf Values(6) = 0 Then
                msg "Qty Authorized cannot be 0.", vbExclamation, "Edit Authorization Quantity"
            Else
                .QtyAuthorized = Values(6)
                .ExtPrice = Values(6) * Values(5)
            End If
        End If
    End With
    
    cmdUpdateRMALine.Enabled = True
End Sub


Private Sub gdxRMALine_LostFocus()
    gdxRMALine.Update
End Sub


Private Sub gdxRMALine_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxRMALine.Update
End Sub


'Refresh the lower grid (gdxRMALineStatus)
'   1. after clicking on upper grid
'   2. after press and up and down key

Private Sub gdxRMALine_Click()
    RefreshRMALineStatus
End Sub


Private Sub gdxRMALine_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        RefreshRMALineStatus
    End If
End Sub


'This subroutines refreshes line status for RMA order
'Grid column 2 = RMALineKey
'Grid column 3 = ItemID

Private Sub RefreshRMALineStatus()
    Dim rst As ADODB.Recordset
    
    lblRMALineStatus = "RMA Line Status for " & gdxRMALine.value(3)
    'LoadRMALineStatus gdxRMALine.Value(2)
    Set rst = CallSP("spcpcRMALineSummary", "@_iRMALineKey", gdxRMALine.value(2))

    Dim i As Integer
    With gdxRMALineStatus
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = rst
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
End Sub



'This subroutine sets up RMA return reason list in the grid

Private Sub SetRMAReason()
    Dim colTemp As JSColumn
    Dim vl As JSValueList
    
    Set colTemp = gdxRMALine.Columns("Reason")
    If Not colTemp.HasValueList Then
        colTemp.HasValueList = True
        Set vl = colTemp.ValueList
        g_rstRMAReason.MoveFirst
        Do While Not g_rstRMAReason.EOF
            vl.Add g_rstRMAReason.Fields("RMAReasonKey").value, g_rstRMAReason.Fields("RMAReasonID").value
            g_rstRMAReason.MoveNext
        Loop
        colTemp.EditType = jgexEditDropDown
    End If
End Sub


' Load RMA line grid

Private Sub LoadRMALine(lRMAKey As Long)
    Dim i As Integer
    
    SetWaitCursor True
    
    'why do this every time we 'refresh' the grid?
    SetRMAReason

    'why reload the collection on a refresh?
    Set m_oRMALine = Billing.LoadRMALine(lRMAKey)
    
    With gdxRMALine
        .HoldFields
        .ItemCount = m_oRMALine.Count
        .Refetch
        .Row = 1
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With

    rvRMA.OwnerID = ""
    rvRMA.OwnerID = m_oOrder.OPKey
    
    RefreshRMALineStatus
    SetWaitCursor False
End Sub


'This public method is called by the CatalogRequest dialog box.

Public Sub AddMarketingItem(ByVal QtyBound As Integer, ByVal QtyNotebook As Integer, ByVal QtyPriceList As Integer)
    
    If QtyBound > 0 Then
        CreateMarketingItem "BC", QtyBound
    End If
    
    If QtyNotebook > 0 Then
        CreateMarketingItem "NC", QtyNotebook
    End If
    
    If QtyPriceList > 0 Then
        If m_oCustomer.CustType = "Wholesale" Then
            CreateMarketingItem "PLW", QtyPriceList
        ElseIf m_oCustomer.CustType = "Dealer" Then
            CreateMarketingItem "PLD", QtyPriceList
        End If
    End If
    m_bNewItem = True
    m_lSelectedIndex = m_oItems.SelectedIndex
    m_oValidateItem.Valid = True
    SyncItemList
End Sub


'Reads module-level variables
'   m_oOrder
'   m_oItems
'Writes module-level variables
'   m_bNewItem
'   m_oItems
'   m_lSelectedIndex
'   m_oValidateItem

Private Sub CreateMarketingItem(ByVal ItemID As String, ByVal Qty As Integer)
    Dim cmd As ADODB.Command
    Dim lItemKey As Long
    Dim oFinGood As ItemFinGood
    Dim oItem As IItem

    Set cmd = CreateCommandSP("spcpcItemIDtoKey")
    cmd.Parameters("@_iItemID").value = ItemID
    cmd.Execute
    lItemKey = cmd.Parameters("@_oRetVal").value
    
    Set oFinGood = New ItemFinGood
    oFinGood.Load lItemKey, m_oOrder.WhseKey
    Set oItem = oFinGood
    oItem.Qty = Qty
    oItem.StatusCode = ItemStatusCode.iscReadyToCommit
    m_oItems.Add oItem
End Sub


'Added 11/29/04 LR (PRN 499)

Public Sub AddGiftItem()
    Dim oCmd As ADODB.Command
    Dim lItemKey As Long
    Dim oFinGood As ItemFinGood
    Dim oItem As IItem
    Dim oRC As RemarkContext

    Set oCmd = CreateCommandSP("spcpcItemIDtoKey")
    oCmd.Parameters("@_iItemID").value = "Holiday Gift"
    oCmd.Execute
    lItemKey = oCmd.Parameters("@_oRetVal").value
    
    Set oFinGood = New ItemFinGood
    oFinGood.Load lItemKey, m_oOrder.WhseKey
    Set oItem = oFinGood
    oItem.Qty = 1
    
    oItem.StatusCode = ItemStatusCode.iscReadyToCommit
    m_oItems.Add oItem
    
    Set oRC = New RemarkContext
    'the Add method (above) creates the item's OPLineKey
    oRC.Load "ViewOrderLine", CStr(oItem.OPLineKey)
    oRC.AddRemark "OrderLine.Warehouse", "Send gift to: " & m_oCustomer.GiftRecipient
    oRC.Save True
    Set oRC = Nothing
        
    m_bNewItem = True
    m_lSelectedIndex = m_oItems.SelectedIndex
    m_oValidateItem.Valid = True
    SyncItemList
    
End Sub


'This replaces the DisplayInfo method of the FVendor form
'invoked by cmdRMAVendor_Click() and cmdVendorSetails_Click()

Private Sub DisplayVendorInfo(ByVal i_lVendorKey As Long)
    Dim orst As ADODB.Recordset
    Dim sSQL As String
    Dim sMsg As String

    If i_lVendorKey = 0 Then
        Exit Sub
    End If

    Set orst = CallSP("spcpcGetVendInfo", "@_iVendorKey", i_lVendorKey)
    With orst
        sMsg = "[" & Trim(.Fields("VendID").value) & "] " & Trim(.Fields("VendName").value) & vbCrLf & vbCrLf
        sMsg = sMsg & CompAddr(.Fields("AddrName").value, _
            .Fields("AddrLine1").value, _
            .Fields("AddrLine2").value, _
            .Fields("City").value, _
            .Fields("StateID").value, _
            .Fields("PostalCode").value, _
            .Fields("CountryID").value) & vbCrLf & vbCrLf
        sMsg = sMsg & "Contact: " & .Fields("Name").value & vbCrLf
        sMsg = sMsg & "Phone: " & FormatPhoneNumber(.Fields("Phone").value, .Fields("PhoneExt").value) & vbCrLf
        sMsg = sMsg & "Fax: " & FormatPhoneNumber(.Fields("Fax").value, .Fields("FaxExt").value) & vbCrLf
        sMsg = sMsg & "Email: " & .Fields("EmailAddr").value
    End With
    Set orst = Nothing
    GlobalFunctions.DisplayInfo "Vendor Information", sMsg, 5200, 2800

End Sub


'*********************************************************************************************
' Contact Manager
'*********************************************************************************************

Private Sub cboContact_GotFocus()
    cboContact.text = Trim$(cboContact.text)
    If cboContact.ListCount > 0 Then
        ComboBoxOpenList cboContact
    End If
End Sub

' Show or hide the list portion of a combobox
    
Private Sub ComboBoxOpenList(cbo As ComboBox, Optional showIt As Boolean = True)
    SendMessage cbo.hwnd, CB_SHOWDROPDOWN, showIt, ByVal 0&
End Sub

'if the user beings to type a new name into the combobox, disable the Edit button

Private Sub cboContact_KeyPress(KeyAscii As Integer)
    cmdEditContact.Enabled = False
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
        cmdEditContact.Enabled = True
    End If
End Sub

Private Sub cboContact_KeyDown(KeyCode As Integer, Shift As Integer)

    '38 = UpArrow, 40 = DownArrow
    'If you are on a non-existing contact and press the UpArrow...
    If KeyCode = 38 And cboContact.ListIndex = -1 Then
        KeyCode = 0
    'If you are on the first record in the list and press the up arrow...
    ElseIf KeyCode = 38 And cboContact.ListIndex = 0 Then
        KeyCode = 0
        If Len(m_Name) = 0 Then
            cboContact.text = vbNullString
        Else
            cboContact.text = m_Name
            cboContact.SelLength = Len(m_Name)
        End If
        cboContact.ListIndex = -1

        ClearContactColor cboContact
        
        'clear the display
        lblShipContact.caption = vbNullString
        lblShipPhone.caption = vbNullString
        lblShipFax.caption = vbNullString
        lblCellPhone.caption = vbNullString
        cmdEditContact.Enabled = False
    End If
    
End Sub

Private Sub cboContact_KeyUp(KeyCode As Integer, Shift As Integer)
    'if the contents of the box has just been deleted
    If Len(cboContact.text) = 0 Then
        ClearContactColor cboContact
    End If
End Sub


'color coding for emails
Private Sub cboContact_DropDown()
    ClearContactColor cboContact
End Sub


'The only time you'll be able to select a different item from the
'combobox is for an account customer with prior contacts on record.

Private Sub cboContact_Click()
    If m_bLoading Then Exit Sub

    If cboContact.ListIndex <> -1 Then
        If m_oCustomer.HasAccount Then
            m_oOrder.contact = m_oCustomer.Contacts.GetContactByKey(cboContact.ItemData(cboContact.ListIndex))
            AssignContactToOrder
        Else
            AssignContactToOrder
        End If
    End If
'***
End Sub


Private Sub cboContact_Change()
    If m_bLoading Then Exit Sub
    m_Name = Trim$(cboContact.text)
End Sub


Private Sub cboContact_LostFocus()
    Dim sName As String
    Dim i_oTempContact As contact
    Dim oType As OwnerType
    Dim lOwnerKey As Long
    
    If Len(m_Name) > 0 And cboContact.ListIndex = -1 Then
        cboContact.text = m_Name
        sName = m_Name
    Else
        sName = Trim$(cboContact.text)
    End If
    
    'Evaluate the customer to get OwnerType
    If Not m_oCustomer.HasAccount Then
        oType = opOrder
        lOwnerKey = m_oOrder.OPKey
    Else
        oType = opCustomer
        lOwnerKey = m_oCustomer.Key
    End If
    
    If Len(sName) > 0 Then
    
        'm_oOrder does not have a Contact object
        If m_oOrder.contact Is Nothing Then
            If FoundInCombo(sName) Then
                'The order contact was just set in FoundInCombo
                'Test to see if it is valid.
                If Not m_oOrder.contact.IsValid Then
                    'Prompt to fix
                    If vbOK = m_oOrder.contact.Edit(GetUserName, cboContact.text, oType, lOwnerKey) Then
                        RefreshContactList
                    End If
                End If

'TODO: Change the name of this routine to match it's function.
                AssignContactToOrder
            Else
                m_oOrder.contact = New contact
                'm_oOrder.contact.Connection = g_DB.Connection
                
                If vbCancel = m_oOrder.contact.Edit(GetUserName, cboContact.text, oType, lOwnerKey) Then
                    cboContact.text = vbNullString
                    TryToSetFocus cboContact
                    m_oOrder.contact = Nothing
                Else
                    'add the new contact to the customers collection
                    If m_oCustomer.HasAccount Then
                        m_oCustomer.Contacts.InsertNewContact m_oOrder.contact
                    End If
                    'reload the combobox and select the new contact
                    RefreshContactList
                    'update labels
                    AssignContactToOrder
                End If
            End If
            
        'm_oOrder has a Contact object
        Else
            If LCase(sName) = LCase(m_oOrder.contact.Name) Then 'Nothing has changed
                'Test to see if it is valid.
                If Not m_oOrder.contact.IsValid Then
                    'Prompt to fix
                    If vbOK = m_oOrder.contact.Edit(GetUserName, cboContact.text, oType, lOwnerKey) Then
                        RefreshContactList
                        AssignContactToOrder
                    End If
                End If
                GoTo Cleanup
                
            Else
                If FoundInCombo(sName) Then
                    'Test to see if it is valid.
                    If Not m_oOrder.contact.IsValid Then
                        'Prompt to fix
                        If vbOK = m_oOrder.contact.Edit(GetUserName, cboContact.text, oType, lOwnerKey) Then
                            RefreshContactList
                        End If
                    End If
                    AssignContactToOrder
                Else
                    Set i_oTempContact = m_oOrder.contact
                    m_oOrder.contact = New contact
                    'm_oOrder.contact.Connection = g_DB.Connection

                    If vbCancel = m_oOrder.contact.Edit(GetUserName, cboContact.text, oType, lOwnerKey) Then
                        'Restore the contact and reset the name.
                        m_oOrder.contact = i_oTempContact
                        cboContact.text = m_oOrder.contact.Name
                        'cmdContactMgr(2).Enabled = True
                        cmdEditContact.Enabled = True
                        TryToSetFocus cboContact
                    Else
                        'add the new contact to the customers collection
                        If m_oCustomer.HasAccount Then
                            m_oCustomer.Contacts.InsertNewContact m_oOrder.contact
                        Else
                            'Order contact was replaced for a non-account cust.
                            'Delete the previous contact if the order hasn't been committed.
                            If m_oOrder.soKey = 0 Then
                                On Error GoTo EH
                                Screen.MousePointer = vbHourglass
                                i_oTempContact.Delete bFullDelete:=True
                                Screen.MousePointer = vbDefault
                            End If
                        End If
                        'reload the combobox and select the new contact
                        RefreshContactList
                        AssignContactToOrder
                    End If
                    Set i_oTempContact = Nothing
                End If
            End If
        End If
    
    ' The combobox text is empty
    ' 1. Nothing has been typed or selected, or
    ' 2. What was there just got deleted
    Else
        If Not (m_oOrder.contact Is Nothing) Then
            'If the Order's Customer doesn't have an account and the Order
            'has not been committed, delete the contact from the database.
            If (Not m_oCustomer.HasAccount) Then
                If m_oOrder.soKey = 0 Then
                    On Error GoTo EH
                    Screen.MousePointer = vbHourglass
                    m_oOrder.contact.Delete bFullDelete:=True
                    Screen.MousePointer = vbDefault
                    cboContact.Clear
                    m_oOrder.contact = Nothing
                    'clear the tooltip
                Else
                    cboContact.Clear
                    m_oOrder.contact = Nothing
                    'clear the tooltip
                End If
            Else
                cboContact.text = vbNullString
                m_oOrder.contact = Nothing
            End If
        End If
        UpdateShipContactInfo
    End If
    
    m_Name = ""
    
Cleanup:
    If Not m_oOrder.contact Is Nothing Then
        SetContactColorByEaddr cboContact, m_oOrder.contact
    End If
    'This is to fix the case where the combobox become invisible or partially visible.
    ForceRefresh Me.hwnd
    Exit Sub
EH:
    MsgBox Err.Number & ": " & Err.Description
    Screen.MousePointer = vbDefault
End Sub


Private Function FoundInCombo(i_sName As String) As Boolean
    Dim i As Integer
    
    FoundInCombo = False
    For i = 0 To cboContact.ListCount - 1
        If LCase(i_sName) = LCase(cboContact.list(i)) Then
            FoundInCombo = True
            'This fires off the cboContact_Click event which sets Order.Contact
            cboContact.ListIndex = i
            Exit For
        End If
    Next
End Function


Private Sub AssignContactToOrder()
    UpdateShipContactInfo
    cmdEditContact.Enabled = True
End Sub

Private Sub cmdEditContact_Click()
    Dim oType As OwnerType
    Dim lOwnerKey As Long
    
    '*** Evaluate the customer to get OwnerType
    If Not m_oCustomer.HasAccount Then
        oType = opOrder
        lOwnerKey = m_oOrder.OPKey
    Else
        oType = opCustomer
        lOwnerKey = m_oCustomer.Key
    End If

    If vbOK = m_oOrder.contact.Edit(GetUserName, cboContact.text, oType, lOwnerKey) Then
        'was a change made, something added?
        RefreshContactList
        'update labels
        AssignContactToOrder
    End If
    
    SetContactColorByEaddr cboContact, m_oOrder.contact
End Sub


Private Sub RefreshContactList()
    Dim oContact As contact
    
    If m_oCustomer.HasAccount Then
        cboContact.Clear
        For Each oContact In m_oCustomer.Contacts
            cboContact.AddItem oContact.Name
            cboContact.ItemData(cboContact.NewIndex) = oContact.Key
        Next
    Else
        cboContact.Clear
        cboContact.AddItem m_oOrder.contact.Name
        cboContact.ItemData(cboContact.NewIndex) = m_oOrder.contact.Key
    End If
    
    SetComboByText cboContact, m_oOrder.contact.Name
End Sub


'*********************************************************************************
'Write all saved orders to XSL/XML files in the event of trouble on 10/3/05
'*********************************************************************************

Private Sub SaveOrderAsXML(i_oOrder As Order)
    Dim sXML As String
    Dim sXMLPath As String

    sXMLPath = g_SnapshotPath & i_oOrder.OPKey & ".xml"
    sXML = GetOrderXML(i_oOrder)
    SaveToFile sXMLPath, XslHeader(g_XsltPath & "Order.xsl") + sXML
End Sub


Private Function GetOrderXML(i_oOrder As Order) As String
    Dim oXMLNode As JDMPDXML.XMLNode

    Set oXMLNode = i_oOrder.Export(True, True, True, False, True, True)
    oXMLNode.IndentWidth = 2    'need to have a width >= 0 to make xml well-formed
    GetOrderXML = oXMLNode.ExportString
End Function


Private Function WhseHasCatalogs() As Boolean
    
    Select Case GetUserWhseID
        Case "MPK"
            WhseHasCatalogs = g_MPKHasCatalogs
        Case "STL"
            WhseHasCatalogs = g_STLHasCatalogs
        Case "SEA"
            WhseHasCatalogs = g_SEAHasCatalogs
    End Select
    
End Function


Private Sub cmdEmailQuote_Click()
    Dim frm As FEmailQuote
    Dim remarkText As String
    
    'serialize the order to XML
    SaveAsTextFile g_QuoteEmailPath & GetUserName & ".xml", SerializeOrder

    Set frm = New FEmailQuote
    
    frm.QuoteNumber = m_oOrder.OPKey
    frm.PONumber = m_oOrder.PurchOrd
    
    If Not m_oOrder.contact Is Nothing Then
        frm.emailaddr = m_oOrder.contact.emailaddr
        frm.IsHtml = m_oOrder.contact.EMailFormat
    End If
    
    frm.Init
    
    If frm.LogEvent Then
        remarkText = "Emailed quote to " & frm.emailaddr
        If frm.chkAddNotes.value Then
            remarkText = remarkText & vbCrLf & frm.txtNotes
        End If
        rvOrder.AddRemark "Order.Private", remarkText
    End If
    
    Unload frm
    Set frm = Nothing
End Sub


Private Function SerializeOrder() As String
    Dim oXMLNode As JDMPDXML.XMLNode
    Set oXMLNode = m_oOrder.Export(bPubNotes:=True)
    oXMLNode.IndentWidth = 2    'required?
    SerializeOrder = oXMLNode.ExportString
End Function



'**********************************************************************************************
' DATA ACCESS LAYER
'**********************************************************************************************

'Check if Sage order can be cancelled

Private Function EligibleForCancel(ByVal i_SOKey) As Boolean
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spcpcEligibleForDelete")

    With cmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@_iSOKey", adInteger, adParamInput, 0, i_SOKey)
        .Execute
        EligibleForCancel = IIf(.Parameters("RETURN_VALUE") = 0, True, False)
    End With
End Function


Private Function CancelPick(ByVal soKey As Long) As Integer
'To be written
    CancelPick = 0
End Function


' Return Values:
' 0 - Error
' 1 - OK
' 2 - lines on shipment

Private Function CancelSO(ByVal i_SOKey As Long) As Integer
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spcpcCancelSO")
    cmd.Parameters("@SOKey").value = i_SOKey
    cmd.Execute
    CancelSO = cmd.Parameters("@RetVal").value
End Function


Private Function GetInvoiceDetail(ByVal OPKey As Long) As ADODB.Recordset
    Set GetInvoiceDetail = CallSP("cpopGetInvoiceDetail", "@_iOPKey", m_gwCustOrders.value("OPKey"))
End Function


Private Function CreateEmptyOrderSummaryRecordset() As ADODB.Recordset
    'generate a zero-record recordset
    Dim sSQL As String
    sSQL = "SELECT OPKey, StatusCode, CreateDate, RTRIM(UserID) as UserID, WhseKey, CustKey, OrderedBy, SOKey, Summary, PurchOrd AS CustPO, ShipAddrKey, Info as Note " & _
           "FROM tcpSO WHERE 1 = 2"
    Set CreateEmptyOrderSummaryRecordset = LoadDiscRst(sSQL)
End Function


Private Function GetCautionsForCountry(CountryID As String) As ADODB.Recordset
    Dim sSQL As String
    sSQL = "select Cautions, CSWCountrySymbol from tcpcswcountry where countryid = '" & CountryID & "'"
    Set GetCautionsForCountry = LoadDiscRst(sSQL, , adLockBatchOptimistic)
End Function


Private Sub cmdManageDropShips_Click()
    Dim oFrm As FProvisionalShipment

    Set oFrm = New FProvisionalShipment
    oFrm.OPKey = m_oOrder.OPKey
    oFrm.CreateShipment Me
End Sub

    

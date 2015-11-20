VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CCDE1390-FFE8-11D4-8122-AA0004000604}#17.0#0"; "InchWorm.ocx"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "mmremark.ocx"
Object = "{0FA91D91-3062-44DB-B896-91406D28F92A}#54.0#0"; "SOTACalendar.ocx"
Object = "{DAE11CD2-4384-11D7-9DBD-000102499D33}#1.0#0"; "currcontrol.ocx"
Begin VB.Form FOrder2 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9375
   Begin ActiveTabs.SSActiveTabs tabMain 
      Height          =   5955
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   10504
      _Version        =   262144
      TabCount        =   6
      Tabs            =   "FOrder2.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   5565
         Left            =   -99969
         TabIndex        =   1
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder2.frx":014E
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print History"
            Height          =   375
            Index           =   0
            Left            =   7920
            TabIndex        =   2
            Top             =   5040
            Width           =   1095
         End
         Begin GridEX20.GridEX gdxOrderEvent 
            Height          =   4815
            Left            =   120
            TabIndex        =   3
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
            Column(1)       =   "FOrder2.frx":0176
            Column(2)       =   "FOrder2.frx":02BE
            Column(3)       =   "FOrder2.frx":03E2
            Column(4)       =   "FOrder2.frx":055E
            Column(5)       =   "FOrder2.frx":0C36
            Column(6)       =   "FOrder2.frx":0D82
            SortKeysCount   =   1
            SortKey(1)      =   "FOrder2.frx":0EEA
            FormatStylesCount=   6
            FormatStyle(1)  =   "FOrder2.frx":0F52
            FormatStyle(2)  =   "FOrder2.frx":108A
            FormatStyle(3)  =   "FOrder2.frx":113A
            FormatStyle(4)  =   "FOrder2.frx":11EE
            FormatStyle(5)  =   "FOrder2.frx":12C6
            FormatStyle(6)  =   "FOrder2.frx":137E
            ImageCount      =   0
            PrinterProperties=   "FOrder2.frx":145E
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpStatus 
         Height          =   5565
         Left            =   -99969
         TabIndex        =   4
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder2.frx":1636
         Begin VB.Frame Frame1 
            Height          =   1392
            Index           =   15
            Left            =   120
            TabIndex        =   6
            Top             =   60
            Width           =   8895
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "CustID"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   23
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Name"
               Height          =   255
               Index           =   28
               Left            =   360
               TabIndex        =   22
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "OP#"
               Height          =   252
               Index           =   27
               Left            =   3960
               TabIndex        =   21
               Top             =   240
               Width           =   672
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "SO#"
               Height          =   252
               Index           =   26
               Left            =   6360
               TabIndex        =   20
               Top             =   240
               Width           =   972
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "CSR"
               Height          =   252
               Index           =   29
               Left            =   4140
               TabIndex        =   19
               Top             =   600
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Order Date"
               Height          =   252
               Index           =   67
               Left            =   6360
               TabIndex        =   18
               Top             =   960
               Width           =   972
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ordered By"
               Height          =   252
               Index           =   49
               Left            =   6360
               TabIndex        =   17
               Top             =   600
               Width           =   972
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   16
               Top             =   240
               Width           =   2775
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   7
               Left            =   7440
               TabIndex        =   15
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   5
               Left            =   7440
               TabIndex        =   14
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Index           =   3
               Left            =   4800
               TabIndex        =   13
               Top             =   240
               Width           =   1452
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   4
               Left            =   4800
               TabIndex        =   12
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   6
               Left            =   7440
               TabIndex        =   11
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   10
               Top             =   600
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "PurchOrd"
               Height          =   255
               Index           =   54
               Left            =   120
               TabIndex        =   9
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lblOS 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   2
               Left            =   960
               TabIndex        =   8
               Top             =   960
               Width           =   2775
            End
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
               TabIndex        =   7
               Top             =   960
               Width           =   1455
            End
         End
         Begin VB.CommandButton cmdOSRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   7920
            TabIndex        =   5
            Top             =   1560
            Width           =   1095
         End
         Begin ActiveTabs.SSActiveTabs SSAOrderDetails 
            Height          =   3732
            Left            =   120
            TabIndex        =   24
            Top             =   1680
            Width           =   8892
            _ExtentX        =   15690
            _ExtentY        =   6588
            _Version        =   262144
            TabCount        =   3
            Tabs            =   "FOrder2.frx":165E
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
               Height          =   3345
               Left            =   -99969
               TabIndex        =   25
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   5900
               _Version        =   262144
               TabGuid         =   "FOrder2.frx":1712
               Begin GridEX20.GridEX gdxOSInvoice 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   26
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
                  Column(1)       =   "FOrder2.frx":173A
                  Column(2)       =   "FOrder2.frx":188E
                  Column(3)       =   "FOrder2.frx":19B6
                  Column(4)       =   "FOrder2.frx":1B42
                  Column(5)       =   "FOrder2.frx":1CC2
                  Column(6)       =   "FOrder2.frx":1E42
                  Column(7)       =   "FOrder2.frx":1FC6
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":2106
                  FormatStyle(2)  =   "FOrder2.frx":21E6
                  FormatStyle(3)  =   "FOrder2.frx":231E
                  FormatStyle(4)  =   "FOrder2.frx":23CE
                  FormatStyle(5)  =   "FOrder2.frx":2482
                  FormatStyle(6)  =   "FOrder2.frx":255A
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":2612
               End
               Begin GridEX20.GridEX gdxOSInvoiceItem 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   27
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
                  Column(1)       =   "FOrder2.frx":27EA
                  Column(2)       =   "FOrder2.frx":2932
                  Column(3)       =   "FOrder2.frx":2A6E
                  Column(4)       =   "FOrder2.frx":2B92
                  Column(5)       =   "FOrder2.frx":2CCA
                  Column(6)       =   "FOrder2.frx":2E52
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":2FDA
                  FormatStyle(2)  =   "FOrder2.frx":30BA
                  FormatStyle(3)  =   "FOrder2.frx":31F2
                  FormatStyle(4)  =   "FOrder2.frx":32A2
                  FormatStyle(5)  =   "FOrder2.frx":3356
                  FormatStyle(6)  =   "FOrder2.frx":342E
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":34E6
               End
               Begin VB.Label lblOSLGridCaption 
                  Caption         =   "Ordered By"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   28
                  Top             =   1560
                  Width           =   7695
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
               Height          =   3345
               Left            =   -99969
               TabIndex        =   29
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   5900
               _Version        =   262144
               TabGuid         =   "FOrder2.frx":36BE
               Begin GridEX20.GridEX gdxOSShipments 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   30
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
                  ColumnsCount    =   7
                  Column(1)       =   "FOrder2.frx":36E6
                  Column(2)       =   "FOrder2.frx":3836
                  Column(3)       =   "FOrder2.frx":3962
                  Column(4)       =   "FOrder2.frx":3ABE
                  Column(5)       =   "FOrder2.frx":3C02
                  Column(6)       =   "FOrder2.frx":3D9A
                  Column(7)       =   "FOrder2.frx":3EC6
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":4006
                  FormatStyle(2)  =   "FOrder2.frx":40E6
                  FormatStyle(3)  =   "FOrder2.frx":421E
                  FormatStyle(4)  =   "FOrder2.frx":42CE
                  FormatStyle(5)  =   "FOrder2.frx":4382
                  FormatStyle(6)  =   "FOrder2.frx":445A
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":4512
               End
               Begin GridEX20.GridEX gdxOSShipItems 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   31
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
                  ColumnsCount    =   6
                  Column(1)       =   "FOrder2.frx":46EA
                  Column(2)       =   "FOrder2.frx":4832
                  Column(3)       =   "FOrder2.frx":496E
                  Column(4)       =   "FOrder2.frx":4A92
                  Column(5)       =   "FOrder2.frx":4BCA
                  Column(6)       =   "FOrder2.frx":4D4E
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":4ED6
                  FormatStyle(2)  =   "FOrder2.frx":4FB6
                  FormatStyle(3)  =   "FOrder2.frx":50EE
                  FormatStyle(4)  =   "FOrder2.frx":519E
                  FormatStyle(5)  =   "FOrder2.frx":5252
                  FormatStyle(6)  =   "FOrder2.frx":532A
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":53E2
               End
               Begin VB.Label lblOSLGridCaption 
                  Caption         =   "Ordered By"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   32
                  Top             =   1560
                  Width           =   7095
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
               Height          =   3345
               Left            =   30
               TabIndex        =   33
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   5900
               _Version        =   262144
               TabGuid         =   "FOrder2.frx":55BA
               Begin GridEX20.GridEX gdxOSLine 
                  Height          =   1455
                  Left            =   120
                  TabIndex        =   34
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
                  Column(1)       =   "FOrder2.frx":55E2
                  Column(2)       =   "FOrder2.frx":572A
                  Column(3)       =   "FOrder2.frx":584E
                  Column(4)       =   "FOrder2.frx":598A
                  Column(5)       =   "FOrder2.frx":5AC6
                  Column(6)       =   "FOrder2.frx":5C4E
                  Column(7)       =   "FOrder2.frx":5D8E
                  Column(8)       =   "FOrder2.frx":5EDA
                  Column(9)       =   "FOrder2.frx":6022
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":6152
                  FormatStyle(2)  =   "FOrder2.frx":6232
                  FormatStyle(3)  =   "FOrder2.frx":636A
                  FormatStyle(4)  =   "FOrder2.frx":641A
                  FormatStyle(5)  =   "FOrder2.frx":64CE
                  FormatStyle(6)  =   "FOrder2.frx":65A6
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":665E
               End
               Begin GridEX20.GridEX gdxOSLineItems 
                  Height          =   1215
                  Left            =   120
                  TabIndex        =   35
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
                  Column(1)       =   "FOrder2.frx":6836
                  Column(2)       =   "FOrder2.frx":697E
                  Column(3)       =   "FOrder2.frx":6CD6
                  Column(4)       =   "FOrder2.frx":6E0A
                  Column(5)       =   "FOrder2.frx":6F4A
                  Column(6)       =   "FOrder2.frx":706E
                  Column(7)       =   "FOrder2.frx":719A
                  Column(8)       =   "FOrder2.frx":72C6
                  Column(9)       =   "FOrder2.frx":73E6
                  Column(10)      =   "FOrder2.frx":7656
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":7776
                  FormatStyle(2)  =   "FOrder2.frx":7856
                  FormatStyle(3)  =   "FOrder2.frx":798E
                  FormatStyle(4)  =   "FOrder2.frx":7A3E
                  FormatStyle(5)  =   "FOrder2.frx":7AF2
                  FormatStyle(6)  =   "FOrder2.frx":7BCA
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":7C82
               End
               Begin VB.Label lblOSLGridCaption 
                  Caption         =   "Ordered By"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   36
                  Top             =   1680
                  Width           =   8175
               End
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpRMA 
         Height          =   5565
         Left            =   -99969
         TabIndex        =   37
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder2.frx":7E5A
         Begin ActiveTabs.SSActiveTabs tabRMADetail 
            Height          =   5295
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   9340
            _Version        =   262144
            TabCount        =   2
            Tabs            =   "FOrder2.frx":7E82
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
               Height          =   4905
               Left            =   -99969
               TabIndex        =   39
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   8652
               _Version        =   262144
               TabGuid         =   "FOrder2.frx":7F01
               Begin VB.CommandButton cmdRefreshCallTag 
                  Caption         =   "Refresh"
                  Height          =   325
                  Left            =   120
                  TabIndex        =   46
                  Top             =   4440
                  Width           =   1335
               End
               Begin VB.CommandButton cmdPrint 
                  Caption         =   "Print"
                  Height          =   325
                  Index           =   1
                  Left            =   120
                  TabIndex        =   44
                  Top             =   4440
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.Frame Frame1 
                  Caption         =   "Add Weight for New Package"
                  Height          =   855
                  Index           =   19
                  Left            =   120
                  TabIndex        =   40
                  Top             =   120
                  Width           =   8535
                  Begin VB.CommandButton cmdAddWeight 
                     Caption         =   "Add"
                     Height          =   325
                     Left            =   4800
                     TabIndex        =   42
                     Top             =   360
                     Width           =   1095
                  End
                  Begin NEWSOTALib.SOTANumber txtWeight 
                     Height          =   330
                     Left            =   2520
                     TabIndex        =   41
                     Top             =   360
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   582
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
                     mask            =   "<ILH>##<ILp0>#<IRp0>|.##"
                     text            =   "  0.00"
                     sIntegralPlaces =   3
                     sDecimalPlaces  =   2
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Weight"
                     Height          =   255
                     Index           =   83
                     Left            =   1320
                     TabIndex        =   43
                     Top             =   360
                     Width           =   735
                  End
               End
               Begin GridEX20.GridEX gdxCallTag 
                  Height          =   2775
                  Left            =   120
                  TabIndex        =   45
                  Top             =   1560
                  Width           =   8535
                  _ExtentX        =   15055
                  _ExtentY        =   4895
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
                  ColumnsCount    =   3
                  Column(1)       =   "FOrder2.frx":7F29
                  Column(2)       =   "FOrder2.frx":8069
                  Column(3)       =   "FOrder2.frx":81B5
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":82F5
                  FormatStyle(2)  =   "FOrder2.frx":83D5
                  FormatStyle(3)  =   "FOrder2.frx":850D
                  FormatStyle(4)  =   "FOrder2.frx":85BD
                  FormatStyle(5)  =   "FOrder2.frx":8671
                  FormatStyle(6)  =   "FOrder2.frx":8749
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":8801
               End
               Begin VB.Label Label1 
                  Caption         =   "Call Tag Details"
                  Height          =   255
                  Index           =   84
                  Left            =   120
                  TabIndex        =   47
                  Top             =   1200
                  Width           =   1815
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
               Height          =   4905
               Left            =   30
               TabIndex        =   48
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   8652
               _Version        =   262144
               TabGuid         =   "FOrder2.frx":89D9
               Begin VB.CommandButton cmdRMARefresh 
                  Caption         =   "Refresh"
                  Height          =   325
                  Left            =   3000
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdUpdateRMALine 
                  Caption         =   "Save Changes"
                  Enabled         =   0   'False
                  Height          =   325
                  Left            =   1560
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdAddMoreItem 
                  Caption         =   "Add Items"
                  Height          =   325
                  Left            =   120
                  TabIndex        =   51
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdRMAVendor 
                  Caption         =   "Show Vendor"
                  Height          =   325
                  Index           =   0
                  Left            =   4440
                  TabIndex        =   50
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdRMAVendor 
                  Caption         =   "Packing List"
                  Height          =   325
                  Index           =   1
                  Left            =   5880
                  TabIndex        =   49
                  Top             =   240
                  Width           =   1215
               End
               Begin GridEX20.GridEX gdxRMALine 
                  Height          =   1635
                  Left            =   120
                  TabIndex        =   54
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
                  Column(1)       =   "FOrder2.frx":8A01
                  Column(2)       =   "FOrder2.frx":8B6D
                  Column(3)       =   "FOrder2.frx":8CE5
                  Column(4)       =   "FOrder2.frx":8E29
                  Column(5)       =   "FOrder2.frx":8FB5
                  Column(6)       =   "FOrder2.frx":9141
                  Column(7)       =   "FOrder2.frx":9265
                  Column(8)       =   "FOrder2.frx":93A9
                  Column(9)       =   "FOrder2.frx":9529
                  Column(10)      =   "FOrder2.frx":9699
                  Column(11)      =   "FOrder2.frx":9809
                  Column(12)      =   "FOrder2.frx":9979
                  Column(13)      =   "FOrder2.frx":9AFD
                  Column(14)      =   "FOrder2.frx":9C69
                  Column(15)      =   "FOrder2.frx":9DB5
                  Column(16)      =   "FOrder2.frx":9F41
                  Column(17)      =   "FOrder2.frx":A0D9
                  Column(18)      =   "FOrder2.frx":A22D
                  Column(19)      =   "FOrder2.frx":A369
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":A4A9
                  FormatStyle(2)  =   "FOrder2.frx":A5E1
                  FormatStyle(3)  =   "FOrder2.frx":A691
                  FormatStyle(4)  =   "FOrder2.frx":A745
                  FormatStyle(5)  =   "FOrder2.frx":A81D
                  FormatStyle(6)  =   "FOrder2.frx":A8D5
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":A9B5
               End
               Begin GridEX20.GridEX gdxRMALineStatus 
                  Height          =   1635
                  Left            =   120
                  TabIndex        =   55
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
                  Column(1)       =   "FOrder2.frx":AB8D
                  Column(2)       =   "FOrder2.frx":ACED
                  Column(3)       =   "FOrder2.frx":AE59
                  Column(4)       =   "FOrder2.frx":AFAD
                  Column(5)       =   "FOrder2.frx":B10D
                  Column(6)       =   "FOrder2.frx":B231
                  Column(7)       =   "FOrder2.frx":B391
                  Column(8)       =   "FOrder2.frx":B521
                  Column(9)       =   "FOrder2.frx":B681
                  Column(10)      =   "FOrder2.frx":B7A5
                  Column(11)      =   "FOrder2.frx":B8C9
                  Column(12)      =   "FOrder2.frx":BA29
                  Column(13)      =   "FOrder2.frx":BB6D
                  Column(14)      =   "FOrder2.frx":BCD1
                  Column(15)      =   "FOrder2.frx":BE11
                  Column(16)      =   "FOrder2.frx":BF69
                  Column(17)      =   "FOrder2.frx":C0A5
                  Column(18)      =   "FOrder2.frx":C1F1
                  Column(19)      =   "FOrder2.frx":C35D
                  FormatStylesCount=   6
                  FormatStyle(1)  =   "FOrder2.frx":C469
                  FormatStyle(2)  =   "FOrder2.frx":C5A1
                  FormatStyle(3)  =   "FOrder2.frx":C651
                  FormatStyle(4)  =   "FOrder2.frx":C705
                  FormatStyle(5)  =   "FOrder2.frx":C7DD
                  FormatStyle(6)  =   "FOrder2.frx":C895
                  ImageCount      =   0
                  PrinterProperties=   "FOrder2.frx":C975
               End
               Begin MMRemark.RemarkViewer rvRMA 
                  Height          =   810
                  Left            =   7920
                  TabIndex        =   56
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
                  TabIndex        =   57
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
            TabIndex        =   58
            Top             =   960
            Width           =   1095
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpItems 
         Height          =   5565
         Left            =   -99969
         TabIndex        =   59
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder2.frx":CB4D
         Begin VB.Frame frmInventory 
            Caption         =   "Inventory"
            ClipControls    =   0   'False
            Height          =   972
            Left            =   120
            TabIndex        =   221
            Top             =   2520
            Visible         =   0   'False
            Width           =   8952
            Begin VB.CommandButton cmdInvFinder 
               Caption         =   "Inventory Details..."
               Height          =   552
               Left            =   7620
               TabIndex        =   223
               Top             =   240
               Width           =   975
            End
            Begin VB.ComboBox cboWarehouse 
               Height          =   315
               Index           =   2
               Left            =   6480
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   222
               Top             =   510
               Width           =   852
            End
            Begin VB.Label lblQtyBO 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Left            =   5160
               TabIndex        =   234
               Top             =   510
               Width           =   732
            End
            Begin VB.Label lblQtyPO 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Left            =   3960
               TabIndex        =   233
               Top             =   510
               Width           =   732
            End
            Begin VB.Label lblQtySO 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Left            =   2760
               TabIndex        =   232
               Top             =   510
               Width           =   732
            End
            Begin VB.Label lblQtyOH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Left            =   1560
               TabIndex        =   231
               Top             =   510
               Width           =   732
            End
            Begin VB.Label lblQtyAvail 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Left            =   360
               TabIndex        =   230
               Top             =   510
               Width           =   732
            End
            Begin VB.Label Label1 
               Caption         =   "On PurchOrd"
               Height          =   252
               Index           =   38
               Left            =   3960
               TabIndex        =   229
               Top             =   270
               Width           =   1092
            End
            Begin VB.Label Label1 
               Caption         =   "On SalesOrd"
               Height          =   252
               Index           =   37
               Left            =   2760
               TabIndex        =   228
               Top             =   270
               Width           =   972
            End
            Begin VB.Label Label1 
               Caption         =   "On Hand"
               Height          =   252
               Index           =   36
               Left            =   1560
               TabIndex        =   227
               Top             =   270
               Width           =   732
            End
            Begin VB.Label Label1 
               Caption         =   "Available"
               Height          =   252
               Index           =   35
               Left            =   360
               TabIndex        =   226
               Top             =   270
               Width           =   732
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Warehouse"
               Height          =   252
               Index           =   33
               Left            =   6480
               TabIndex        =   225
               Top             =   270
               Width           =   852
            End
            Begin VB.Label Label1 
               Caption         =   "On BackOrd"
               Height          =   252
               Index           =   68
               Left            =   5160
               TabIndex        =   224
               Top             =   270
               Width           =   972
            End
         End
         Begin VB.Frame frmItemList 
            BorderStyle     =   0  'None
            Caption         =   "Frame13"
            ClipControls    =   0   'False
            Height          =   4095
            Left            =   120
            TabIndex        =   135
            Top             =   1560
            Width           =   8952
            Begin VB.ComboBox cboWarehouse 
               Height          =   315
               Index           =   1
               Left            =   7800
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   3600
               Width           =   972
            End
            Begin VB.CommandButton cmdAuthorizeAll 
               Caption         =   "&Authorize All"
               Height          =   312
               Left            =   0
               TabIndex        =   137
               Top             =   3600
               Width           =   1092
            End
            Begin VB.CommandButton cmdUnAuthorize 
               Caption         =   "UnAuthori&ze"
               Height          =   312
               Left            =   1140
               TabIndex        =   136
               Top             =   3600
               Width           =   1092
            End
            Begin NEWSOTALib.SOTACurrency txtTotalPrice 
               Height          =   315
               Left            =   5640
               TabIndex        =   139
               Top             =   3600
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1926
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
               Enabled         =   0   'False
               mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               sDecimalPlaces  =   2
            End
            Begin NEWSOTALib.SOTACurrency txtTotalTax 
               Height          =   315
               Left            =   3360
               TabIndex        =   140
               Top             =   3600
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1926
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
               Enabled         =   0   'False
               mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               sDecimalPlaces  =   2
            End
            Begin GridEX20.GridEX gdxItems 
               Height          =   3375
               Left            =   0
               TabIndex        =   141
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
               Column(1)       =   "FOrder2.frx":CB75
               Column(2)       =   "FOrder2.frx":CD31
               Column(3)       =   "FOrder2.frx":CF11
               Column(4)       =   "FOrder2.frx":D10D
               Column(5)       =   "FOrder2.frx":D2F9
               Column(6)       =   "FOrder2.frx":D545
               Column(7)       =   "FOrder2.frx":D7B5
               Column(8)       =   "FOrder2.frx":DA15
               Column(9)       =   "FOrder2.frx":DC8D
               Column(10)      =   "FOrder2.frx":E1A9
               Column(11)      =   "FOrder2.frx":E9BD
               Column(12)      =   "FOrder2.frx":EDA5
               FmtConditionsCount=   1
               FmtCondition(1) =   "FOrder2.frx":EF51
               FormatStylesCount=   7
               FormatStyle(1)  =   "FOrder2.frx":F09D
               FormatStyle(2)  =   "FOrder2.frx":F17D
               FormatStyle(3)  =   "FOrder2.frx":F2B5
               FormatStyle(4)  =   "FOrder2.frx":F365
               FormatStyle(5)  =   "FOrder2.frx":F419
               FormatStyle(6)  =   "FOrder2.frx":F4F1
               FormatStyle(7)  =   "FOrder2.frx":F5A9
               ImageCount      =   0
               PrinterProperties=   "FOrder2.frx":F65D
            End
            Begin VB.Label lblSalesTax 
               Caption         =   "Sales Tax"
               Height          =   255
               Left            =   2520
               TabIndex        =   144
               Top             =   3630
               Width           =   735
            End
            Begin VB.Label lblWarehouse 
               Alignment       =   1  'Right Justify
               Caption         =   "Warehouse"
               Height          =   252
               Left            =   6840
               TabIndex        =   143
               Top             =   3630
               Width           =   852
            End
            Begin VB.Label lblSalesAmt 
               Alignment       =   1  'Right Justify
               Caption         =   "Sales Amount"
               Height          =   210
               Left            =   4440
               TabIndex        =   142
               Top             =   3630
               Width           =   1095
            End
         End
         Begin VB.Frame frmAssembly 
            Caption         =   "Assembly Information"
            ClipControls    =   0   'False
            Height          =   1932
            Left            =   120
            TabIndex        =   121
            Top             =   3600
            Visible         =   0   'False
            Width           =   8952
            Begin VB.CommandButton cmdResearchPO 
               Caption         =   "Purchase Orders..."
               Height          =   312
               Index           =   0
               Left            =   4320
               TabIndex        =   126
               Top             =   1080
               Width           =   2052
            End
            Begin VB.CommandButton cmdVendorDetails 
               Caption         =   "Vendor Details..."
               Height          =   312
               Index           =   1
               Left            =   4320
               TabIndex        =   125
               Top             =   1440
               Width           =   2052
            End
            Begin VB.CommandButton cmdCopyModel 
               Caption         =   "Copy From Previous Item"
               Height          =   312
               Index           =   1
               Left            =   4320
               TabIndex        =   124
               Top             =   360
               Width           =   2052
            End
            Begin VB.CommandButton cmdViewCat 
               Caption         =   "View Catalog Page..."
               Height          =   312
               Index           =   1
               Left            =   4320
               TabIndex        =   123
               Top             =   720
               Width           =   2052
            End
            Begin VB.ComboBox cboMake 
               Height          =   315
               Index           =   1
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   360
               Width           =   3015
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtSerial 
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   127
               Top             =   1080
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   128
               Top             =   720
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   129
               Top             =   360
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewItem"
            End
            Begin VB.Label lblVendor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   134
               Top             =   1440
               Width           =   3015
            End
            Begin VB.Label Label1 
               Caption         =   "Vendor"
               Height          =   255
               Index           =   65
               Left            =   480
               TabIndex        =   133
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Serial #"
               Height          =   252
               Index           =   60
               Left            =   360
               TabIndex        =   132
               Top             =   1080
               Width           =   612
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Model"
               Height          =   255
               Index           =   93
               Left            =   360
               TabIndex        =   131
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Make"
               Height          =   255
               Index           =   94
               Left            =   360
               TabIndex        =   130
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.Frame frmGasket 
            BorderStyle     =   0  'None
            Caption         =   "Specify Gasket"
            ClipControls    =   0   'False
            Height          =   2892
            Left            =   120
            TabIndex        =   97
            Top             =   2640
            Visible         =   0   'False
            Width           =   8952
            Begin VB.Frame Frame1 
               Caption         =   "Gasket Specs"
               ClipControls    =   0   'False
               Height          =   2775
               Index           =   3
               Left            =   0
               TabIndex        =   102
               Top             =   0
               Width           =   5052
               Begin VB.ComboBox cboGasket 
                  Height          =   315
                  Left            =   3000
                  Style           =   2  'Dropdown List
                  TabIndex        =   113
                  Top             =   360
                  Width           =   1452
               End
               Begin VB.Frame Frame1 
                  ClipControls    =   0   'False
                  Height          =   735
                  Index           =   1
                  Left            =   180
                  TabIndex        =   110
                  Top             =   960
                  Width           =   1695
                  Begin VB.OptionButton optGasketSides 
                     Caption         =   "4-Sided"
                     Height          =   255
                     Index           =   0
                     Left            =   180
                     TabIndex        =   112
                     Top             =   180
                     Value           =   -1  'True
                     Width           =   1215
                  End
                  Begin VB.OptionButton optGasketSides 
                     Caption         =   "3-Sided"
                     Height          =   255
                     Index           =   1
                     Left            =   180
                     TabIndex        =   111
                     Top             =   420
                     Width           =   1215
                  End
               End
               Begin VB.CheckBox chkGasketOptions 
                  Caption         =   "No Magnet RHS"
                  Height          =   255
                  Index           =   3
                  Left            =   3000
                  TabIndex        =   109
                  Tag             =   "2"
                  Top             =   2220
                  Width           =   1695
               End
               Begin VB.CheckBox chkGasketOptions 
                  Caption         =   "No Magnet LHS"
                  Height          =   255
                  Index           =   2
                  Left            =   3000
                  TabIndex        =   108
                  Tag             =   "4"
                  Top             =   1920
                  Width           =   1932
               End
               Begin VB.CheckBox chkGasketOptions 
                  Caption         =   "Dart-to-Dart"
                  Height          =   255
                  Index           =   0
                  Left            =   3000
                  TabIndex        =   107
                  Tag             =   "8"
                  Top             =   1320
                  Width           =   1455
               End
               Begin VB.CheckBox chkGasketOptions 
                  Caption         =   "Inverted"
                  Height          =   255
                  Index           =   1
                  Left            =   3000
                  TabIndex        =   106
                  Tag             =   "1"
                  Top             =   1620
                  Width           =   1215
               End
               Begin VB.Frame Frame1 
                  ClipControls    =   0   'False
                  Height          =   735
                  Index           =   2
                  Left            =   180
                  TabIndex        =   103
                  Top             =   240
                  Width           =   1695
                  Begin VB.OptionButton optGasketType 
                     Caption         =   "Compression"
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   105
                     Top             =   420
                     Width           =   1335
                  End
                  Begin VB.OptionButton optGasketType 
                     Caption         =   "Magnetic"
                     Height          =   195
                     Index           =   0
                     Left            =   120
                     TabIndex        =   104
                     Top             =   180
                     Value           =   -1  'True
                     Width           =   1095
                  End
               End
               Begin InchWorm.LengthEntry lenGasket 
                  Height          =   285
                  Index           =   1
                  Left            =   780
                  TabIndex        =   114
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
                  TabIndex        =   115
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
                  Caption         =   "Material"
                  Height          =   255
                  Index           =   57
                  Left            =   2040
                  TabIndex        =   120
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label lblGasketMatlUsed 
                  Caption         =   "Los Angeles or St. Louis"
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   119
                  Top             =   780
                  Width           =   1935
               End
               Begin VB.Label Label1 
                  Caption         =   "Molded in"
                  Height          =   255
                  Index           =   92
                  Left            =   2040
                  TabIndex        =   118
                  Top             =   780
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Caption         =   "Width"
                  Height          =   255
                  Index           =   55
                  Left            =   180
                  TabIndex        =   117
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.Label Label1 
                  Caption         =   "Height"
                  Height          =   255
                  Index           =   56
                  Left            =   180
                  TabIndex        =   116
                  Top             =   2340
                  Width           =   735
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Line Remarks"
               Height          =   2895
               Index           =   17
               Left            =   5160
               TabIndex        =   98
               Top             =   0
               Width           =   3735
               Begin VB.TextBox txtMMRemark 
                  Height          =   2175
                  Index           =   1
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   100
                  Top             =   600
                  Width           =   2655
               End
               Begin VB.ComboBox cboMMType 
                  Height          =   315
                  Index           =   1
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   99
                  Top             =   240
                  Width           =   1935
               End
               Begin MMRemark.RemarkViewer rvOrderLine 
                  Height          =   810
                  Index           =   2
                  Left            =   2830
                  TabIndex        =   101
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
         Begin VB.Frame frmBasicInfo 
            Caption         =   "General"
            ClipControls    =   0   'False
            Height          =   1395
            Left            =   120
            TabIndex        =   88
            Top             =   120
            Visible         =   0   'False
            Width           =   7572
            Begin VB.CheckBox chkCGMPN 
               Caption         =   "Customer Gave Me Part Number"
               Height          =   255
               Left            =   1800
               TabIndex        =   89
               Top             =   1080
               Width           =   2775
            End
            Begin MSComctlLib.ImageCombo icbItemStatus 
               Height          =   330
               Left            =   3900
               TabIndex        =   90
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
               TabIndex        =   91
               Top             =   240
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   92
               Top             =   720
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7641
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
               lMaxLength      =   50
            End
            Begin MMRemark.RemarkViewer rvOrderLine 
               Height          =   810
               Index           =   0
               Left            =   6480
               TabIndex        =   93
               Top             =   240
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewOrderLine"
               Caption         =   "Line Remarks"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Part #"
               Height          =   204
               Index           =   40
               Left            =   1200
               TabIndex        =   96
               Top             =   280
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status"
               Height          =   216
               Index           =   42
               Left            =   3300
               TabIndex        =   95
               Top             =   288
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Descr"
               Height          =   210
               Index           =   41
               Left            =   1200
               TabIndex        =   94
               Top             =   720
               Width           =   495
            End
            Begin VB.Image imgType 
               Height          =   1050
               Left            =   120
               Top             =   240
               Width           =   1050
            End
         End
         Begin VB.Frame frmFindPart 
            Caption         =   "Find Item"
            ClipControls    =   0   'False
            Height          =   1395
            Left            =   120
            TabIndex        =   82
            Top             =   120
            Width           =   4095
            Begin VB.CommandButton cmdSearch 
               Caption         =   "Fi&nd"
               Height          =   360
               Left            =   3420
               TabIndex        =   85
               Top             =   600
               Width           =   555
            End
            Begin VB.PictureBox picItem 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   768
               Left            =   360
               ScaleHeight     =   765
               ScaleWidth      =   765
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   360
               Width           =   768
            End
            Begin VB.CheckBox chkSearchItemDescr 
               Caption         =   "Search Item Description"
               Height          =   375
               Left            =   1600
               TabIndex        =   83
               Top             =   960
               Width           =   2055
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtItemSearch 
               Height          =   315
               Left            =   1620
               TabIndex        =   86
               Top             =   600
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   87
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame frmSpecifyPart 
            Caption         =   "Specify Item"
            ClipControls    =   0   'False
            Height          =   1395
            Left            =   4320
            TabIndex        =   77
            Top             =   120
            Width           =   4752
            Begin VB.CommandButton cmdSpecifyItem 
               Caption         =   "W&ire"
               Height          =   855
               Index           =   2
               Left            =   2560
               Style           =   1  'Graphical
               TabIndex        =   81
               ToolTipText     =   "Add Warmer Wire to order"
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton cmdSpecifyItem 
               Caption         =   "Sh&elf"
               Height          =   855
               Index           =   1
               Left            =   1400
               Style           =   1  'Graphical
               TabIndex        =   80
               ToolTipText     =   "Add Wire Shelf to order"
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton cmdSpecifyItem 
               Caption         =   "&Gasket"
               Height          =   855
               Index           =   0
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   79
               ToolTipText     =   "Add Gasket to order"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   855
            End
            Begin VB.CommandButton cmdSpecifyItem 
               Caption         =   "SP&O"
               Height          =   855
               Index           =   3
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   78
               ToolTipText     =   "Add special order item to order"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame frmPricing 
            Caption         =   "Pricing"
            ClipControls    =   0   'False
            Height          =   792
            Left            =   120
            TabIndex        =   64
            Top             =   1680
            Visible         =   0   'False
            Width           =   8952
            Begin CurrControl.CurrencyInput txtPrice 
               Height          =   315
               Left            =   5760
               TabIndex        =   65
               Top             =   300
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
            End
            Begin NEWSOTALib.SOTACurrency txtPrice1 
               Height          =   315
               Left            =   5760
               TabIndex        =   66
               Top             =   300
               Visible         =   0   'False
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   67
               Top             =   300
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   68
               Top             =   300
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   69
               Top             =   300
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
            End
            Begin NEWSOTALib.SOTACurrency txtCost1 
               Height          =   315
               Left            =   4080
               TabIndex        =   70
               Top             =   300
               Visible         =   0   'False
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   71
               Top             =   300
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               Caption         =   "Quantity"
               Height          =   252
               Index           =   31
               Left            =   240
               TabIndex        =   76
               Top             =   300
               Width           =   612
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Price"
               Height          =   255
               Index           =   32
               Left            =   5160
               TabIndex        =   75
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Cost"
               Height          =   255
               Index           =   34
               Left            =   3480
               TabIndex        =   74
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ext. Price"
               Height          =   255
               Index           =   39
               Left            =   6960
               TabIndex        =   73
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "List Price"
               Height          =   255
               Index           =   96
               Left            =   1740
               TabIndex        =   72
               Top             =   300
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdItemCancel 
            Caption         =   "&Cancel"
            Height          =   312
            Left            =   7965
            TabIndex        =   63
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdItemDelete 
            Caption         =   "&Delete"
            Height          =   312
            Left            =   7965
            TabIndex        =   62
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdNextGasket 
            Caption         =   "&Next Gasket"
            Height          =   312
            Left            =   7965
            TabIndex        =   61
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdItemOK 
            Caption         =   "&OK"
            Height          =   312
            Left            =   7965
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Frame frmSpecialOrder 
            Caption         =   "Specify Special Order"
            ClipControls    =   0   'False
            Height          =   2892
            Left            =   120
            TabIndex        =   204
            Top             =   2640
            Visible         =   0   'False
            Width           =   8952
            Begin VB.Frame Frame1 
               Caption         =   "Line Remarks"
               Height          =   2655
               Index           =   18
               Left            =   5880
               TabIndex        =   213
               Top             =   120
               Width           =   2960
               Begin VB.TextBox txtMMRemark 
                  Height          =   1935
                  Index           =   2
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   215
                  Top             =   600
                  Width           =   1905
               End
               Begin VB.ComboBox cboMMType 
                  Height          =   315
                  Index           =   2
                  ItemData        =   "FOrder2.frx":F835
                  Left            =   120
                  List            =   "FOrder2.frx":F837
                  Style           =   2  'Dropdown List
                  TabIndex        =   214
                  Top             =   240
                  Width           =   1935
               End
               Begin MMRemark.RemarkViewer rvOrderLine 
                  Height          =   810
                  Index           =   3
                  Left            =   2080
                  TabIndex        =   216
                  ToolTipText     =   "Line Remarks"
                  Top             =   600
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   1429
                  ContextID       =   "ViewOrderLine"
                  Caption         =   ""
               End
            End
            Begin VB.CommandButton cmdVendorDetails 
               Caption         =   "Vendor Details..."
               Height          =   312
               Index           =   3
               Left            =   3840
               TabIndex        =   209
               Top             =   1440
               Width           =   1935
            End
            Begin VB.CommandButton cmdCopyModel 
               Caption         =   "Copy From Previous Item"
               Height          =   312
               Index           =   3
               Left            =   3840
               TabIndex        =   208
               Top             =   360
               Width           =   1935
            End
            Begin VB.CommandButton cmdPartsWiz 
               Caption         =   "Parts Wiz..."
               Height          =   312
               Left            =   3840
               TabIndex        =   207
               Top             =   720
               Width           =   1935
            End
            Begin VB.ComboBox cboVendor 
               Height          =   315
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   206
               Top             =   1440
               Width           =   2655
            End
            Begin VB.ComboBox cboMake 
               Height          =   315
               Index           =   3
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   205
               Top             =   360
               Width           =   2655
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtSerial 
               Height          =   315
               Index           =   3
               Left            =   1080
               TabIndex        =   210
               Top             =   1080
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   211
               Top             =   720
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   212
               Top             =   1920
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   1429
               ContextID       =   "ViewVendor"
               Caption         =   "Vendor Remarks"
            End
            Begin VB.Label Label1 
               Caption         =   "Vendor"
               Height          =   252
               Index           =   64
               Left            =   360
               TabIndex        =   220
               Top             =   1440
               Width           =   612
            End
            Begin VB.Label Label1 
               Caption         =   "Serial"
               Height          =   252
               Index           =   63
               Left            =   360
               TabIndex        =   219
               Top             =   1080
               Width           =   492
            End
            Begin VB.Label Label1 
               Caption         =   "Model"
               Height          =   255
               Index           =   90
               Left            =   357
               TabIndex        =   218
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Make"
               Height          =   255
               Index           =   91
               Left            =   360
               TabIndex        =   217
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame frmShelf 
            BorderStyle     =   0  'None
            Caption         =   "Specify Wire Shelf"
            ClipControls    =   0   'False
            Height          =   2892
            Left            =   120
            TabIndex        =   184
            Top             =   2640
            Visible         =   0   'False
            Width           =   8952
            Begin VB.Frame Frame1 
               Caption         =   "Options"
               ClipControls    =   0   'False
               Height          =   2772
               Index           =   12
               Left            =   3000
               TabIndex        =   198
               Top             =   0
               Width           =   1695
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Support"
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   203
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Product Stop"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   202
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Bent Leg"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   201
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Straight Leg"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   200
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.CheckBox chkShelfOpt 
                  Caption         =   "Cut-Out"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   199
                  Top             =   360
                  Width           =   1335
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Dimensions"
               ClipControls    =   0   'False
               Height          =   2772
               Index           =   14
               Left            =   0
               TabIndex        =   189
               Top             =   0
               Width           =   2772
               Begin VB.ComboBox cboFrame 
                  Height          =   315
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   191
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.ComboBox cboFinish 
                  Height          =   315
                  Left            =   1080
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   190
                  Top             =   960
                  Width           =   1452
               End
               Begin InchWorm.LengthEntry lenShelfWidth 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   192
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
                  TabIndex        =   193
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
                  Caption         =   "Width"
                  Height          =   252
                  Index           =   52
                  Left            =   480
                  TabIndex        =   197
                  Top             =   2040
                  Width           =   492
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Depth"
                  Height          =   252
                  Index           =   53
                  Left            =   240
                  TabIndex        =   196
                  Top             =   1560
                  Width           =   732
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Frame Diameter"
                  Height          =   492
                  Index           =   51
                  Left            =   240
                  TabIndex        =   195
                  Top             =   360
                  Width           =   732
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Finish"
                  Height          =   252
                  Index           =   59
                  Left            =   360
                  TabIndex        =   194
                  Top             =   960
                  Width           =   612
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Line Remarks"
               Height          =   2775
               Index           =   16
               Left            =   4800
               TabIndex        =   185
               Top             =   0
               Width           =   4095
               Begin VB.ComboBox cboMMType 
                  Height          =   315
                  Index           =   0
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   187
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.TextBox txtMMRemark 
                  Height          =   2055
                  Index           =   0
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   186
                  Top             =   600
                  Width           =   2895
               End
               Begin MMRemark.RemarkViewer rvOrderLine 
                  Height          =   810
                  Index           =   1
                  Left            =   3120
                  TabIndex        =   188
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
         Begin VB.Frame frmWire 
            BorderStyle     =   0  'None
            Caption         =   "Warmer Wire"
            ClipControls    =   0   'False
            Height          =   2892
            Left            =   120
            TabIndex        =   159
            Top             =   2640
            Visible         =   0   'False
            Width           =   8952
            Begin VB.Frame Frame1 
               Caption         =   "Electrical Properties"
               ClipControls    =   0   'False
               Height          =   2652
               Index           =   13
               Left            =   5040
               TabIndex        =   175
               Top             =   120
               Width           =   3732
               Begin VB.ComboBox cboWires 
                  Height          =   315
                  Left            =   2160
                  Style           =   2  'Dropdown List
                  TabIndex        =   177
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.ComboBox cboVoltage 
                  Height          =   315
                  ItemData        =   "FOrder2.frx":F839
                  Left            =   480
                  List            =   "FOrder2.frx":F846
                  TabIndex        =   176
                  Text            =   "cboVoltage"
                  Top             =   600
                  Width           =   852
               End
               Begin VB.Label Label1 
                  Caption         =   "Voltage"
                  Height          =   252
                  Index           =   50
                  Left            =   480
                  TabIndex        =   183
                  Top             =   360
                  Width           =   612
               End
               Begin VB.Label Label1 
                  Caption         =   "Available Wires"
                  Height          =   252
                  Index           =   47
                  Left            =   2160
                  TabIndex        =   182
                  Top             =   360
                  Width           =   1332
               End
               Begin VB.Label lblAmperage 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   252
                  Left            =   480
                  TabIndex        =   181
                  Top             =   2040
                  Width           =   852
               End
               Begin VB.Label lblWattsPerFoot 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   252
                  Left            =   480
                  TabIndex        =   180
                  Top             =   1320
                  Width           =   852
               End
               Begin VB.Label Label1 
                  Caption         =   "Amperage"
                  Height          =   252
                  Index           =   48
                  Left            =   480
                  TabIndex        =   179
                  Top             =   1800
                  Width           =   852
               End
               Begin VB.Label Label1 
                  Caption         =   "Watts Per Foot"
                  Height          =   252
                  Index           =   46
                  Left            =   480
                  TabIndex        =   178
                  Top             =   1080
                  Width           =   1212
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Wire Length"
               ClipControls    =   0   'False
               Height          =   2652
               Index           =   11
               Left            =   0
               TabIndex        =   160
               Top             =   120
               Width           =   4812
               Begin VB.Frame frmWirePasses 
                  BorderStyle     =   0  'None
                  Caption         =   "optWirePasses"
                  ClipControls    =   0   'False
                  Height          =   972
                  Left            =   2880
                  TabIndex        =   166
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1692
                  Begin VB.OptionButton optWirePasses 
                     Caption         =   "Double Pass"
                     Height          =   252
                     Index           =   1
                     Left            =   240
                     TabIndex        =   168
                     TabStop         =   0   'False
                     Top             =   600
                     Width           =   1332
                  End
                  Begin VB.OptionButton optWirePasses 
                     Caption         =   "Single Pass"
                     Height          =   252
                     Index           =   0
                     Left            =   240
                     TabIndex        =   167
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1332
                  End
               End
               Begin VB.Frame frmDoorStyle 
                  BorderStyle     =   0  'None
                  Caption         =   "optDoorStyle"
                  ClipControls    =   0   'False
                  Height          =   1212
                  Left            =   2880
                  TabIndex        =   163
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   1692
                  Begin VB.OptionButton optDoorStyle 
                     Caption         =   "3-Sided (Double Pass)"
                     Height          =   492
                     Index           =   1
                     Left            =   240
                     TabIndex        =   165
                     TabStop         =   0   'False
                     Top             =   720
                     Width           =   1332
                  End
                  Begin VB.OptionButton optDoorStyle 
                     Caption         =   "4-Sided (Single Pass)"
                     Height          =   372
                     Index           =   0
                     Left            =   240
                     TabIndex        =   164
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1332
                  End
               End
               Begin VB.OptionButton optLengthAlgorithm 
                  Caption         =   "Specify Door Dimensions"
                  Height          =   252
                  Index           =   1
                  Left            =   120
                  TabIndex        =   162
                  TabStop         =   0   'False
                  Top             =   1320
                  Value           =   -1  'True
                  Width           =   2532
               End
               Begin VB.OptionButton optLengthAlgorithm 
                  Caption         =   "Specify Overall Length"
                  Height          =   372
                  Index           =   0
                  Left            =   120
                  TabIndex        =   161
                  Top             =   360
                  Width           =   2052
               End
               Begin InchWorm.LengthEntry lenWireLength 
                  Height          =   288
                  Left            =   1200
                  TabIndex        =   169
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
                  TabIndex        =   170
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
                  TabIndex        =   171
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
                  Caption         =   "Length"
                  Height          =   252
                  Index           =   45
                  Left            =   480
                  TabIndex        =   174
                  Top             =   720
                  Width           =   612
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Width"
                  Height          =   252
                  Index           =   44
                  Left            =   480
                  TabIndex        =   173
                  Top             =   2040
                  Width           =   612
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Height"
                  Height          =   252
                  Index           =   43
                  Left            =   480
                  TabIndex        =   172
                  Top             =   1680
                  Width           =   612
               End
            End
         End
         Begin VB.Frame frmStock 
            Caption         =   "Finished Good Information"
            ClipControls    =   0   'False
            Height          =   1932
            Left            =   120
            TabIndex        =   145
            Top             =   3600
            Visible         =   0   'False
            Width           =   8952
            Begin VB.CommandButton cmdResearchPO 
               Caption         =   "Purchase Orders..."
               Height          =   312
               Index           =   1
               Left            =   4320
               TabIndex        =   150
               Top             =   1080
               Width           =   2052
            End
            Begin VB.CommandButton cmdVendorDetails 
               Caption         =   "Vendor Details..."
               Height          =   312
               Index           =   2
               Left            =   4320
               TabIndex        =   149
               Top             =   1440
               Width           =   2052
            End
            Begin VB.CommandButton cmdCopyModel 
               Caption         =   "Copy From Previous Item"
               Height          =   312
               Index           =   2
               Left            =   4320
               TabIndex        =   148
               Top             =   360
               Width           =   2052
            End
            Begin VB.CommandButton cmdViewCat 
               Caption         =   "View Catalog (pg 26)..."
               Height          =   312
               Index           =   0
               Left            =   4320
               TabIndex        =   147
               Top             =   720
               Width           =   2052
            End
            Begin VB.ComboBox cboMake 
               Height          =   315
               Index           =   2
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   146
               Top             =   360
               Width           =   3015
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtSerial 
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   151
               Top             =   1080
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   152
               Top             =   720
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               TabIndex        =   153
               Top             =   360
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewItem"
               Caption         =   "Item Remarks"
            End
            Begin VB.Label lblVendor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   158
               Top             =   1440
               Width           =   3015
            End
            Begin VB.Label Label1 
               Caption         =   "Vendor"
               Height          =   255
               Index           =   66
               Left            =   480
               TabIndex        =   157
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Serial #"
               Height          =   252
               Index           =   61
               Left            =   360
               TabIndex        =   156
               Top             =   1080
               Width           =   612
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Model"
               Height          =   255
               Index           =   89
               Left            =   480
               TabIndex        =   155
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Make"
               Height          =   255
               Index           =   95
               Left            =   480
               TabIndex        =   154
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Total Price"
            Height          =   252
            Index           =   58
            Left            =   3720
            TabIndex        =   235
            Top             =   4920
            Width           =   972
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpOrder 
         Height          =   5565
         Left            =   -99969
         TabIndex        =   236
         Top             =   360
         Visible         =   0   'False
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder2.frx":F858
         Begin VB.CommandButton cmdSpecialHandling 
            Caption         =   "Special Handling"
            Height          =   375
            Left            =   6196
            TabIndex        =   292
            Top             =   5160
            Width           =   1400
         End
         Begin VB.Frame Frame1 
            Caption         =   "General"
            ClipControls    =   0   'False
            Height          =   1515
            Index           =   5
            Left            =   120
            TabIndex        =   282
            Top             =   1200
            Width           =   4575
            Begin VB.ComboBox cboCSR 
               Height          =   315
               Left            =   1140
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   284
               Top             =   240
               Width           =   1935
            End
            Begin SOTACalendarControl.SOTACalendar calPromiseDate 
               Height          =   315
               Left            =   1140
               TabIndex        =   283
               Top             =   660
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskedText      =   "  /  /    "
               Text            =   "  /  /    "
            End
            Begin MMRemark.RemarkViewer rvOrder 
               Height          =   810
               Left            =   3360
               TabIndex        =   285
               Top             =   180
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   1429
               ContextID       =   "ViewOrder"
               Caption         =   "Order Remarks"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Created"
               Height          =   255
               Index           =   22
               Left            =   120
               TabIndex        =   291
               Top             =   1140
               Width           =   855
            End
            Begin VB.Label lblDate 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1140
               TabIndex        =   290
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "CSR"
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   289
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Last Updated"
               Height          =   255
               Index           =   17
               Left            =   2160
               TabIndex        =   288
               Top             =   1140
               Width           =   1020
            End
            Begin VB.Label lblLastUpdate 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3240
               TabIndex        =   287
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Promised"
               Height          =   255
               Index           =   85
               Left            =   120
               TabIndex        =   286
               Top             =   780
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Purchasing"
            ClipControls    =   0   'False
            Height          =   1812
            Index           =   6
            Left            =   4800
            TabIndex        =   276
            Top             =   60
            Width           =   4215
            Begin VB.ComboBox cboTerms 
               Height          =   315
               ItemData        =   "FOrder2.frx":F880
               Left            =   1560
               List            =   "FOrder2.frx":F882
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   279
               Top             =   960
               Width           =   1815
            End
            Begin VB.CheckBox chkReqPO 
               Caption         =   "PO Required"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1560
               TabIndex        =   278
               TabStop         =   0   'False
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox txtPO 
               Height          =   315
               Left            =   1560
               TabIndex        =   277
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Terms"
               Height          =   252
               Index           =   21
               Left            =   900
               TabIndex        =   281
               Top             =   1020
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Purchase Order"
               Height          =   252
               Index           =   20
               Left            =   180
               TabIndex        =   280
               Top             =   300
               Width           =   1212
            End
         End
         Begin VB.Frame frmCreditCard 
            Caption         =   "Credit Card Information"
            ClipControls    =   0   'False
            Height          =   2592
            Left            =   4800
            TabIndex        =   263
            Top             =   2040
            Width           =   4215
            Begin VB.Label Label1 
               Caption         =   "Name"
               Height          =   255
               Index           =   80
               Left            =   180
               TabIndex        =   275
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Expires"
               Height          =   255
               Index           =   74
               Left            =   180
               TabIndex        =   274
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "ZipCode"
               Height          =   255
               Index           =   82
               Left            =   180
               TabIndex        =   273
               Top             =   1800
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "Street"
               Height          =   255
               Index           =   81
               Left            =   180
               TabIndex        =   272
               Top             =   1440
               Width           =   675
            End
            Begin VB.Label lblExpireDate 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   271
               Top             =   660
               Width           =   1332
            End
            Begin VB.Label lblCCType 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   180
               TabIndex        =   270
               Top             =   300
               Width           =   972
            End
            Begin VB.Label lblCCNo 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   269
               Top             =   300
               Width           =   2292
            End
            Begin VB.Label lblHolderName 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   268
               Top             =   1020
               Width           =   2232
            End
            Begin VB.Label lblCCStreet 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   267
               Top             =   1380
               Width           =   2172
            End
            Begin VB.Label lblCCZipCode 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   266
               Top             =   1740
               Width           =   1212
            End
            Begin VB.Label Label3 
               Caption         =   "Status"
               Height          =   255
               Left            =   180
               TabIndex        =   265
               Top             =   2160
               Width           =   795
            End
            Begin VB.Label lblCCStatus 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Left            =   1200
               TabIndex        =   264
               Top             =   2100
               Width           =   1212
            End
         End
         Begin VB.CommandButton cmdCreateRMA 
            Caption         =   "Create RMA..."
            Height          =   375
            Left            =   4898
            TabIndex        =   262
            Top             =   5160
            Width           =   1215
         End
         Begin VB.CommandButton cmdLiteEdit 
            Caption         =   "Lite Editing"
            Height          =   375
            Left            =   7680
            TabIndex        =   261
            Top             =   5160
            Width           =   1335
         End
         Begin VB.CommandButton cmdCCEdit 
            Caption         =   "Edit Credit Card"
            Height          =   375
            Left            =   6180
            TabIndex        =   260
            Top             =   4680
            Visible         =   0   'False
            Width           =   1392
         End
         Begin VB.Frame Frame1 
            Caption         =   "Shipping"
            ClipControls    =   0   'False
            Height          =   2715
            Index           =   4
            Left            =   120
            TabIndex        =   244
            Top             =   2760
            Width           =   4575
            Begin VB.TextBox txtUPSAcct 
               Appearance      =   0  'Flat
               BackColor       =   &H80000009&
               Height          =   300
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   8
               MultiLine       =   -1  'True
               TabIndex        =   254
               Top             =   2055
               Width           =   1575
            End
            Begin VB.CheckBox chkBillRecipient 
               Caption         =   "Bill Recipient"
               Height          =   255
               Left            =   1080
               TabIndex        =   253
               Top             =   2355
               Width           =   1350
            End
            Begin VB.CommandButton cmdUPSUpdate 
               Caption         =   "&Change"
               Height          =   300
               Left            =   2760
               TabIndex        =   252
               Top             =   2055
               Width           =   995
            End
            Begin VB.CheckBox chkDefaultShipMeth 
               Caption         =   "Default Shipping Method"
               Enabled         =   0   'False
               Height          =   615
               Left            =   2880
               TabIndex        =   251
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox chkShipComplete 
               Caption         =   "Ship Complete"
               Height          =   195
               Left            =   1080
               TabIndex        =   250
               Top             =   180
               Width           =   1425
            End
            Begin VB.ComboBox cboWarehouse 
               Height          =   315
               Index           =   0
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   249
               Top             =   480
               Width           =   1095
            End
            Begin VB.ComboBox cboShipVia 
               Height          =   315
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   248
               Top             =   870
               Width           =   1680
            End
            Begin VB.CheckBox chkDropShip 
               Caption         =   "Drop Ship"
               Height          =   195
               Left            =   2880
               TabIndex        =   247
               Top             =   180
               Width           =   1215
            End
            Begin VB.TextBox txtShipToContact 
               Height          =   285
               Index           =   0
               Left            =   1080
               MaxLength       =   40
               TabIndex        =   246
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox txtShipToContact 
               Height          =   285
               Index           =   1
               Left            =   1080
               MaxLength       =   17
               TabIndex        =   245
               Top             =   1680
               Width           =   2655
            End
            Begin VB.Label lblUPSAcct 
               Alignment       =   1  'Right Justify
               Caption         =   "UPS Acct"
               Height          =   255
               Left            =   120
               TabIndex        =   259
               Top             =   2100
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ship From"
               Height          =   255
               Index           =   25
               Left            =   240
               TabIndex        =   258
               Top             =   555
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ship Via"
               Height          =   255
               Index           =   24
               Left            =   240
               TabIndex        =   257
               Top             =   945
               Width           =   735
            End
            Begin VB.Label lblPhone 
               Alignment       =   1  'Right Justify
               Caption         =   "Phone #"
               Height          =   255
               Left            =   240
               TabIndex        =   256
               Top             =   1725
               Width           =   735
            End
            Begin VB.Label lblShipTo 
               Alignment       =   1  'Right Justify
               Caption         =   "Ship To"
               Height          =   255
               Left            =   240
               TabIndex        =   255
               Top             =   1365
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ordered By"
            Height          =   1095
            Left            =   120
            TabIndex        =   237
            Top             =   60
            Width           =   4575
            Begin VB.ComboBox cboContact 
               Height          =   315
               Left            =   1140
               TabIndex        =   241
               Text            =   "cboContact"
               Top             =   660
               Width           =   1935
            End
            Begin VB.CommandButton cmdEditContact 
               Caption         =   "Edit"
               Height          =   315
               Left            =   3120
               TabIndex        =   240
               Top             =   660
               Width           =   495
            End
            Begin VB.TextBox txtInfo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1140
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   239
               Text            =   "FOrder2.frx":F884
               Top             =   240
               Width           =   2355
            End
            Begin VB.CommandButton cmdEmailQuote 
               Height          =   315
               Left            =   3720
               Picture         =   "FOrder2.frx":F88E
               Style           =   1  'Graphical
               TabIndex        =   238
               ToolTipText     =   "Email a quote to customer"
               Top             =   660
               Width           =   315
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Contact"
               Height          =   315
               Index           =   0
               Left            =   180
               TabIndex        =   243
               Top             =   720
               Width           =   795
            End
            Begin VB.Label lblNote 
               Alignment       =   1  'Right Justify
               Caption         =   "Info"
               Height          =   255
               Left            =   120
               TabIndex        =   242
               Top             =   300
               Width           =   855
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpCustomer 
         Height          =   5565
         Left            =   30
         TabIndex        =   293
         Top             =   360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   9816
         _Version        =   262144
         TabGuid         =   "FOrder2.frx":F969
         Begin VB.Frame Frame1 
            Caption         =   "Ship To"
            ClipControls    =   0   'False
            Height          =   3360
            Index           =   0
            Left            =   3000
            TabIndex        =   309
            Top             =   2055
            Width           =   6015
            Begin VB.CommandButton cmdEditAddr 
               Caption         =   "This Order Only"
               Height          =   315
               Index           =   0
               Left            =   1800
               TabIndex        =   314
               Top             =   2760
               Width           =   1335
            End
            Begin VB.CommandButton cmdEditAddr 
               Caption         =   "Change Address"
               Height          =   315
               Index           =   1
               Left            =   240
               TabIndex        =   313
               Top             =   2760
               Width           =   1335
            End
            Begin VB.CommandButton cmdContactMgr 
               Caption         =   "Contacts"
               Height          =   315
               Left            =   4440
               TabIndex        =   312
               Top             =   2760
               Width           =   1335
            End
            Begin VB.TextBox txtShipToNote 
               Height          =   312
               Left            =   240
               MaxLength       =   50
               TabIndex        =   311
               Top             =   1680
               Width           =   2655
            End
            Begin VB.CheckBox chkPricePackList 
               Caption         =   "Price Pack Slip"
               Height          =   255
               Left            =   240
               TabIndex        =   310
               Top             =   2160
               Width           =   1455
            End
            Begin VB.Label lblShipContact 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3720
               TabIndex        =   324
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label lblShipFax 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3720
               TabIndex        =   323
               Top             =   1080
               Width           =   2175
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Contact"
               Height          =   255
               Index           =   11
               Left            =   2880
               TabIndex        =   322
               Top             =   420
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Fax"
               Height          =   255
               Index           =   10
               Left            =   2880
               TabIndex        =   321
               Top             =   1140
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Phone"
               Height          =   255
               Index           =   9
               Left            =   2880
               TabIndex        =   320
               Top             =   780
               Width           =   735
            End
            Begin VB.Label lblShipAddr 
               BorderStyle     =   1  'Fixed Single
               Height          =   975
               Left            =   240
               TabIndex        =   319
               Top             =   360
               Width           =   2655
            End
            Begin VB.Label lblShipPhone 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3720
               TabIndex        =   318
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label lblCellPhone 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3720
               TabIndex        =   317
               Top             =   1440
               Width           =   2175
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Cell"
               Height          =   255
               Left            =   2880
               TabIndex        =   316
               Top             =   1500
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Note"
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   315
               Top             =   1440
               Width           =   495
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Bill To"
            ClipControls    =   0   'False
            Height          =   1755
            Index           =   8
            Left            =   120
            TabIndex        =   307
            Top             =   2040
            Width           =   2775
            Begin VB.Label lblBillAddr 
               BorderStyle     =   1  'Fixed Single
               Height          =   975
               Left            =   120
               TabIndex        =   308
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Customer Information"
            ClipControls    =   0   'False
            Height          =   1815
            Index           =   10
            Left            =   120
            TabIndex        =   294
            Top             =   120
            Width           =   8895
            Begin VB.TextBox txtCustName 
               Height          =   312
               Left            =   840
               TabIndex        =   299
               Top             =   360
               Width           =   3432
            End
            Begin VB.TextBox txtCustID 
               Height          =   312
               Left            =   5400
               TabIndex        =   298
               Top             =   360
               Visible         =   0   'False
               Width           =   1212
            End
            Begin VB.ComboBox cboCustType 
               Height          =   315
               ItemData        =   "FOrder2.frx":F991
               Left            =   840
               List            =   "FOrder2.frx":F99E
               Style           =   2  'Dropdown List
               TabIndex        =   297
               Top             =   780
               Width           =   1215
            End
            Begin VB.PictureBox picCustHold 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   300
               Picture         =   "FOrder2.frx":F9BE
               ScaleHeight     =   480
               ScaleWidth      =   480
               TabIndex        =   295
               Top             =   1260
               Visible         =   0   'False
               Width           =   480
            End
            Begin MMRemark.RemarkViewer rvCustomer 
               Height          =   810
               Left            =   7440
               TabIndex        =   296
               Top             =   360
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewCustomer"
               Caption         =   "Customer Remarks"
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
               TabIndex        =   306
               Top             =   1200
               Width           =   3492
            End
            Begin VB.Label lblCustName 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   840
               TabIndex        =   305
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Name"
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   304
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Type"
               Height          =   252
               Index           =   23
               Left            =   240
               TabIndex        =   303
               Top             =   780
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "ID"
               Height          =   255
               Index           =   6
               Left            =   4800
               TabIndex        =   302
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblCustID 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   0
               Left            =   5400
               TabIndex        =   301
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblCustType 
               BorderStyle     =   1  'Fixed Single
               Height          =   312
               Index           =   0
               Left            =   840
               TabIndex        =   300
               Top             =   780
               Width           =   1092
            End
         End
      End
      Begin MSComctlLib.ImageList imglStatus16 
         Left            =   5520
         Top             =   -120
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
               Picture         =   "FOrder2.frx":FAD1
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":FEB7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":10282
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":10672
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":10A2B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":10DF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":11148
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":1149A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":120EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrder2.frx":1243E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imglRemarks 
      Left            =   720
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
            Picture         =   "FOrder2.frx":12790
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":12BE2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglTypeAndStatus32 
      Left            =   720
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
            Picture         =   "FOrder2.frx":13034
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":135C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":13B64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":14149
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1456C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":14BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":150F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":15666
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":15C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":16128
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1657E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":16B1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":17771
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":183C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":18846
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":18CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1992E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglType64 
      Left            =   1200
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
            Picture         =   "FOrder2.frx":1A580
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1AFB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1BCA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1C525
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1CA50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1D7A5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDrop 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   65
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrder2.frx":1E23D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbOrderStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   325
      Top             =   6075
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "FOrder2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const m_skSource = "FOrder2"

Private m_lWindowID As Long

'This is non-standard. In support of creation of button at runtime.
Dim WithEvents cmdPriceHistory As VB.CommandButton
Attribute cmdPriceHistory.VB_VarHelpID = -1

Private Enum TabMainIndexes
    tmiCustomer = 1
    tmiOrder = 2
    tmiLines = 3
    tmiOrderStatus = 4
    tmiRmaLines = 5
    tmiOrderHistory = 6
End Enum

'Order Status tabs
Private Enum TabOrderStatus
    tosLineItem = 1
    tosShipment = 2
    tosInvoice = 3
End Enum

'indices for the button control array cmdSpecifyItem()
Private Const btnGasket = 0
Private Const btnShelf = 1
Private Const btnWire = 2
Private Const btnSPO = 3

'These objects wrap the GridEX controls to give a more convenient
'interface for getting the events we're interested in.
Private WithEvents m_gwItems As GridEXWrapper
Attribute m_gwItems.VB_VarHelpID = -1
Private WithEvents m_gwStatusLine As GridEXWrapper
Attribute m_gwStatusLine.VB_VarHelpID = -1
Private WithEvents m_gwShipments As GridEXWrapper
Attribute m_gwShipments.VB_VarHelpID = -1
Private WithEvents m_gwInvoice As GridEXWrapper
Attribute m_gwInvoice.VB_VarHelpID = -1
Private WithEvents m_gwOSLineItems As GridEXWrapper
Attribute m_gwOSLineItems.VB_VarHelpID = -1
Private WithEvents m_gwRMALine As GridEXWrapper
Attribute m_gwRMALine.VB_VarHelpID = -1

Private WithEvents m_oBrokenRules As BrokenRules
Attribute m_oBrokenRules.VB_VarHelpID = -1

Private m_oValidateItem As ValidateManual
Private m_oValidateDescr As ValidateManual

Private m_oOrder As Order
Attribute m_oOrder.VB_VarHelpID = -1
Private m_oCustomer As Customer

Private WithEvents m_oItems As Items
Attribute m_oItems.VB_VarHelpID = -1
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


Private m_bLoading As Boolean

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


Private m_ePreviousTab As TabMainIndexes


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


Public Property Get StatusCode() As String
    StatusCode = StatusCodeString(m_oOrder.StatusCode)
End Property


Public Property Get BrokenRules() As BrokenRules
    Set BrokenRules = m_oBrokenRules
End Property


Public Sub SetCaption(ByRef i_sTitle As String)
    Me.caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub



Public Sub Init()
    Form_Load
End Sub


Public Sub DoShowHelp()

End Sub


Public Function CancelButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean

End Function


'Called by MDIMain's DoExit function
Public Function ExitCheck() As Boolean

End Function


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Load()
    SetCaption "[New Order]"
    tabMain.Tabs(tmiCustomer).Selected = True
    gdxItems.ItemCount = 0
    
    Set m_oOrder = New Order
    Set m_oCustomer = m_oOrder.Customer
    Set m_oItems = m_oOrder.Items
    
    '11/10/04 LR
    'Note: since no userkey or id is specified this will load combo and select index=0
    '(first User in list). This calls cboCSR_click which writes selected User to Order.UserKey.
    
    'User.SetUpUsers cboCSR, g_rstUsers
    User.LoadActiveCSRs cboCSR
    
    Set m_oBrokenRules = New BrokenRules
    m_oBrokenRules.Form = Me
    LoadValidationRules

    'assign images
    LoadImageList imglTypeAndStatus32, gdxItems
    LoadImageList imglRemarks, gdxOSLineItems
    picItem.Picture = imglType64.ListImages(2).Picture
    cmdSpecifyItem(btnGasket).Picture = imglTypeAndStatus32.ListImages(3).Picture
    cmdSpecifyItem(btnShelf).Picture = imglTypeAndStatus32.ListImages(4).Picture
    cmdSpecifyItem(btnWire).Picture = imglTypeAndStatus32.ListImages(5).Picture
    cmdSpecifyItem(btnSPO).Picture = imglTypeAndStatus32.ListImages(6).Picture
    
    'Initialize grid wrappers
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
    Set m_gwInvoice = New GridEXWrapper
    m_gwInvoice.Grid = gdxOSInvoice

'!!! FOM?
'    If InStr(1, GetUserName, "WillCall", vbTextCompare) > 0 Then
'        LoadWCGridPrefs
'    End If

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

    cmdUpdateRMALine.Enabled = False

'HACK
    ' Add Price History button at runtime
    Set cmdPriceHistory = Controls.Add("VB.CommandButton", "cmdPriceHistory", frmPricing)
    With cmdPriceHistory
        .Move txtPrice.Left + txtPrice.width, txtPrice.Top, 250, txtPrice.Height
        .caption = "?"
        .TabIndex = txtPrice.TabIndex + 1
        .Visible = True
    End With
'***

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub LoadValidationRules()

End Sub


Private Sub Form_Resize()
    tabMain.width = Me.width - 225
    tabMain.Height = Me.Height - 645
End Sub


Private Sub tabMain_BeforeTabClick(ByVal NewTab As ActiveTabs.SSTab, ByVal Cancel As ActiveTabs.SSReturnBoolean)
    Static bInitialized As Boolean

    Select Case NewTab.Index

        Case tmiCustomer, tmiOrder, tmiOrderStatus, tmiRmaLines

        Case tmiLines

    End Select
End Sub


Private Sub tabMain_TabClick(ByVal NewTab As ActiveTabs.SSTab)

    Select Case NewTab.Index

        Case tmiCustomer

        Case tmiOrderHistory

        Case tmiOrder

        Case tmiLines

        Case tmiOrderStatus

        Case tmiRmaLines

    End Select

    With m_oBrokenRules
    End With
End Sub


Public Function DeleteButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean
End Function


Public Function PrintButton(ByVal i_bPrintOnly As Boolean, Optional ByVal i_bDoIt As Boolean = True) As Boolean
End Function


Public Function SaveButton(Optional ByVal i_bDoIt As Boolean = True, _
                           Optional ByVal i_bClose As Boolean = False) As Boolean
End Function


Public Function SplitOrderButton(Optional ByVal i_bDoIt As Boolean = True) As Boolean
    'TODO: The DoSplitOrder function in MDIMain should be moved here.
    '      The code works correctly as is, but violates the design
    '      concept of MDIMain not needing to know details of the forms
    '      that it manages.
End Function


Public Function Commit(Optional ByVal i_bDoIt As Boolean = True) As Boolean
End Function


'This public method is called by the CatalogRequest dialog box.

Public Sub AddMarketingItem(ByVal QtyBound As Integer, ByVal QtyNotebook As Integer, ByVal QtyPriceList As Integer)
End Sub


Public Sub AddGiftItem()
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


'******************************************************************
' Called by Form_Load

'Private Sub LoadWCGridPrefs()
'    Set m_colWCGridPrefs = New Collection
'    Select Case GetUserName
'        Case "LAWillCall", "LAWillCall2"
'            m_colWCGridPrefs.Add 1, "MPK"
'            m_colWCGridPrefs.Add 1, "MPK-Will Call"
'            m_colWCGridPrefs.Add 1, "Committed"
'            m_colWCGridPrefs.Add 1, "Ready to Commit"
'            m_colWCGridPrefs.Add 1, "Authorize"
'
'            gdxOrders.Groups.Clear
'            gdxOrders.GroupByBoxVisible = True
'            'Maximum number of groups that can be added is 4.
'            gdxOrders.Groups.Add gdxOrders.Columns("WhseKey").Index, jgexSortAscending
'            gdxOrders.Groups.Add gdxOrders.Columns("ShipMethID").Index, jgexSortDescending
'            gdxOrders.Groups.Add gdxOrders.Columns("StatusCode").Index, jgexSortDescending
'
'            gdxOrders.Columns("WhseKey").Visible = False
'            gdxOrders.Columns("ShipMethID").Visible = False
'            gdxOrders.Columns("StatusCode").Visible = False
'
'    End Select
'End Sub


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
    cmdPrint(0).Enabled = True

    If Order.Customer.HasAccount Then
        cmdContactMgr.Enabled = True
    Else
        cmdContactMgr.Enabled = False
    End If
    
    With tabMain
        m_ePreviousTab = .SelectedTab.Index
        .Tabs(tmiCustomer).Visible = True
        .Tabs(tmiCustomer).Selected = True
        .Tabs(tmiCustomer).Visible = True
        .Tabs(tmiOrder).Visible = True
        .Tabs(tmiOrderHistory).Visible = True
        .Tabs(tmiLines).Visible = True
        .Tabs(tmiOrderStatus).Visible = (m_oOrder.soKey > 0) And (m_oOrder.StatusCode <> iscDeleted)
        .Tabs(tmiRmaLines).Visible = (m_oOrder.StatusCode = ItemStatusCode.iscHasRMA)
    End With
    
End Sub

Public Sub UpdateOrderStatusBar()
    sbOrderStatus.Enabled = True
    AddStatusBarPanel
    LoadStatusBarPicture
    'If order's research stauts is valid, we will show valid research status. Otherwise,
    'Show the general 'Need Research' status
    'LoadStatusBar
    sbOrderStatus.Panels(1).text = StatusCode
    If m_oOrder.StatusCode > ItemStatusCode.iscReadyToCommit Then
        sbOrderStatus.Panels(2).text = "View"
    Else
        sbOrderStatus.Panels(2).text = "Edit"
    End If
    
    sbOrderStatus.Panels(3).text = "OP   " & m_oOrder.OPKey
    If m_oOrder.StatusCode = ItemStatusCode.iscHasRMA Then
        sbOrderStatus.Panels(5).text = "RMA   " & m_lRMAKey
    End If
    If m_oOrder.soKey > 0 Then
        sbOrderStatus.Panels(4).text = "SO  " & CStr(m_oOrder.TranNo)
    End If

    If m_oOrder.HasSpecialHandling Then
        sbOrderStatus.Panels(6).text = "Special Handling"
    End If
    
    If m_oOrder.IsDropShip Then
        sbOrderStatus.Panels(7).Picture = imgDrop.ListImages(1).Picture
    Else
        sbOrderStatus.Panels(7).Picture = Nothing
    End If
End Sub

Public Sub ResizeOrderMode()

    tabMain.width = Me.width - 255
    tabMain.Height = Me.Height - 750
    
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
'!!! no longer exists
'    Label1(69).Top = cmdAuthorizeAll.Top
'    Label1(62).Top = cmdAuthorizeAll.Top
    txtTotalTax.Top = cmdAuthorizeAll.Top
    txtTotalPrice.Top = cmdAuthorizeAll.Top
'    Label1(2).Top = cmdAuthorizeAll.Top
    cboWarehouse(1).Top = cmdAuthorizeAll.Top
    
    ResizeOS
    ResizeRMA
    ResizeOH
    
    DoEvents
    MDIMain.DoRefresh
End Sub


'Resizes Order Status tab for maximizing and minimizing

Private Sub ResizeOS()
    If tabMain.Tabs(tmiOrderStatus).Visible = True Then
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

    End If
End Sub


'Resizes RMA Line tab for maximizing and minimizing

Private Sub ResizeRMA()
    If tabMain.Tabs(tmiRmaLines).Visible = True Then
        tpRMA.width = tpCustomer.width
        tpRMA.Height = tpCustomer.Height
        
        tabRMADetail.width = tabMain.width - 250
        tabRMADetail.Height = tabMain.Height - 720
        
        gdxRMALine.width = tabRMADetail.width - 250
        gdxRMALineStatus.width = gdxRMALine.width
        gdxCallTag.width = gdxRMALine.width
    
        gdxRMALine.Height = (tabRMADetail.Height - 2025) / 2
        gdxRMALineStatus.Height = (tabRMADetail.Height - 2025) / 2
        gdxCallTag.Height = tabRMADetail.Height - 2520
    
        gdxRMALineStatus.Top = tabRMADetail.Height - gdxRMALineStatus.Height - 550
        lblRMALineStatus.Top = gdxRMALineStatus.Top - lblRMALineStatus.Height - 120
        gdxRMALineStatus.Visible = True
        
        cmdPrint(1).Top = gdxCallTag.Top + gdxCallTag.Height + 120
        cmdRefreshCallTag.Top = cmdPrint(1).Top
    End If
End Sub


'resize the Order History tab

Private Sub ResizeOH()
    If tabMain.Tabs(tmiOrderHistory).Visible = True Then
        gdxOrderEvent.width = tabMain.width - 300
        gdxOrderEvent.Height = tabMain.Height - 1000
        cmdPrint(0).Top = gdxOrderEvent.Top + gdxOrderEvent.Height + 60
        cmdPrint(0).Left = tabMain.width - 300 - cmdPrint(0).width
    End If
End Sub


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
                
        txtInfo.text = .Info
                
        'Initialize the Promise Date calendar control
        'this is an unconventional approach to handling the empty date problem
'        If .PromiseDate = Empty Then
'            If .SOKey > 0 Then
'                calPromiseDate.value = vbNullString
'            Else
'                'why do this?
'                dCurrentTime = Now()
'                If Weekday(dCurrentTime) = 6 Then
'                    .PromiseDate = DateAdd("d", 3, dCurrentTime)
'                ElseIf Weekday(dCurrentTime) = 7 Then
'                    .PromiseDate = DateAdd("d", 2, dCurrentTime)
'                Else
'                    .PromiseDate = DateAdd("d", 1, dCurrentTime)
'                End If
'
'                calPromiseDate.value = .PromiseDate
''                .PromiseDate = calPromiseDate.value    '??? removed 3/3/05 LR
'            End If
'        Else
'            calPromiseDate.value = .PromiseDate
''            .PromiseDate = calPromiseDate.value        '??? removed 3/3/05 LR
'        End If

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
    
    'cboShipVia can change if default shipping method is reset.
    'Set/Remove the Ship To Contact Info controls (visible property & required validation).
    EnableShipToContactCtrls

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
    txtTotalPrice.Amount = m_oItems.TotalPrice
    txtTotalTax.Amount = m_oItems.TotalTax(m_oOrder.SalesTax.TaxRate)
      
    m_oBrokenRules.Validate
    m_bLoading = bLoading
End Sub


Private Sub UpdateOrderCaption()
    If m_oOrder.StatusCode = ItemStatusCode.iscHasRMA Then
        SetCaption sCaption & "- RMA " & m_lRMAKey & "  " & cboWarehouse(1).text
    Else
        SetCaption sCaption & cboWarehouse(1).text
    End If
End Sub


Private Function sCaption() As String
    sCaption = m_oOrder.Customer.Name & "  OP " & m_oOrder.OPKey & " "
End Function


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
  
    'Only reset the UPS Account controls if the customer is not Bill Recipient.
    If Len(txtUPSAcct.text) = 0 Then
        chkBillRecipient.Enabled = False
        chkBillRecipient.value = vbUnchecked
            
        txtUPSAcct.text = ""
        cmdUPSUpdate.Enabled = False
    End If
End Sub


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

Private Sub SetOSTabsVisible()
    'Set all tabs' initial visible to be true
        
    Dim i As Integer
    
    For i = tosLineItem To tosInvoice
        SSAOrderDetails.Tabs(i).Visible = True
    Next
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
'***
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
'!!! no longer exist
'        Label1(12).Visible = False
'        Label1(13).Visible = False
        
        m_oBrokenRules.EnableClass ccShipToData, False
    End If
    
    m_oBrokenRules.Validate txtShipToContact(0)
    m_oBrokenRules.Validate txtShipToContact(1)
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


Private Sub AssignContactToOrder()
    UpdateShipContactInfo
    cmdEditContact.Enabled = True
End Sub

'move to utility module

Private Function FixNullField(field As Variant) As String
    FixNullField = IIf(IsNull(field), vbNullString, field)
End Function


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


Private Sub AddStatusBarPanel()
    Dim lIndex
    
    sbOrderStatus.Panels.Clear
    
    With sbOrderStatus.Panels
        For lIndex = 1 To 7
            .Add lIndex
            If lIndex = 1 Then
                .Item(lIndex).width = 2300
            ElseIf lIndex = 2 Then
                .Item(lIndex).width = 600
            ElseIf lIndex = 3 Then
                .Item(lIndex).width = 1100
            ElseIf lIndex = 4 Then
                .Item(lIndex).width = 1240
            ElseIf lIndex = 6 Then
                .Item(lIndex).width = 1440
            Else
                .Item(lIndex).width = 1240
            End If
        Next
    End With
End Sub


Private Sub LoadStatusBarPicture()
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
End Sub

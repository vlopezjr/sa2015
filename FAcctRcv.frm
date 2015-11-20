VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "mmremark.ocx"
Object = "{DAE11CD2-4384-11D7-9DBD-000102499D33}#1.0#0"; "currcontrol.ocx"
Begin VB.Form FAcctRcv 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   10230
   Begin ActiveTabs.SSActiveTabs tabMain 
      Height          =   7350
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   12965
      _Version        =   262144
      TabCount        =   8
      TagVariant      =   ""
      Tabs            =   "FAcctRcv.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   6960
         Left            =   30
         TabIndex        =   127
         Top             =   360
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   12277
         _Version        =   262144
         TabGuid         =   "FAcctRcv.frx":01D0
         Begin VB.Frame Frame11 
            Caption         =   "Credit Card Orders"
            Height          =   1935
            Left            =   6360
            TabIndex        =   148
            Top             =   240
            Width           =   3495
            Begin VB.CommandButton cmdOPFindCC 
               Caption         =   "View Credit Card"
               Height          =   315
               Left            =   900
               TabIndex        =   189
               Top             =   1380
               Width           =   1695
            End
            Begin VB.TextBox txtOPFindCC 
               Height          =   345
               Left            =   900
               TabIndex        =   150
               Top             =   480
               Width           =   1635
            End
            Begin VB.CommandButton cmdOPShowTrans 
               Caption         =   "View Transactions"
               Height          =   312
               Left            =   900
               TabIndex        =   149
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label lblOPFindCC 
               Caption         =   "OP#"
               Height          =   195
               Left            =   240
               TabIndex        =   151
               Top             =   540
               Width           =   495
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Statement Printing"
            Height          =   1935
            Left            =   6360
            TabIndex        =   145
            Top             =   240
            Visible         =   0   'False
            Width           =   3495
            Begin VB.CommandButton cmdStatementDisplay 
               Caption         =   "Display"
               Height          =   495
               Left            =   1800
               TabIndex        =   147
               Top             =   1200
               Width           =   1455
            End
            Begin VB.ComboBox cboStatementSetting 
               Height          =   315
               Left            =   360
               Style           =   2  'Dropdown List
               TabIndex        =   146
               Top             =   360
               Width           =   2415
            End
         End
         Begin GridEX20.GridEX gdxResearch 
            Height          =   4512
            Left            =   120
            TabIndex        =   136
            Top             =   2340
            Width           =   9792
            _ExtentX        =   17277
            _ExtentY        =   7964
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            Options         =   8
            RecordsetType   =   1
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   1
            ColumnHeaderHeight=   270
            FormatStylesCount=   6
            FormatStyle(1)  =   "FAcctRcv.frx":01F8
            FormatStyle(2)  =   "FAcctRcv.frx":02D8
            FormatStyle(3)  =   "FAcctRcv.frx":0410
            FormatStyle(4)  =   "FAcctRcv.frx":04C0
            FormatStyle(5)  =   "FAcctRcv.frx":0574
            FormatStyle(6)  =   "FAcctRcv.frx":064C
            ImageCount      =   0
            PrinterProperties=   "FAcctRcv.frx":0704
         End
         Begin VB.Frame Frame10 
            Height          =   1992
            Left            =   120
            TabIndex        =   128
            Top             =   180
            Width           =   6135
            Begin VB.OptionButton optShipment 
               Caption         =   "Shipment ID"
               Height          =   255
               Left            =   2100
               TabIndex        =   144
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton optOPID 
               Caption         =   "OP #"
               Height          =   255
               Left            =   3780
               TabIndex        =   143
               Top             =   780
               Width           =   1575
            End
            Begin VB.OptionButton optAcuity 
               Caption         =   "SO #"
               Height          =   255
               Left            =   3780
               TabIndex        =   142
               Top             =   1080
               Width           =   1575
            End
            Begin MSComCtl2.UpDown UpDown2 
               Height          =   315
               Left            =   5640
               TabIndex        =   137
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "lblMaxRecords"
               BuddyDispid     =   196633
               OrigLeft        =   5880
               OrigTop         =   300
               OrigRight       =   6120
               OrigBottom      =   612
               Max             =   50
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65537
               Enabled         =   -1  'True
            End
            Begin VB.OptionButton optCheck 
               Caption         =   "Check Number"
               Height          =   255
               Left            =   180
               TabIndex        =   134
               Top             =   780
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.TextBox txtInput 
               Height          =   285
               Left            =   180
               TabIndex        =   133
               Top             =   300
               Width           =   1695
            End
            Begin VB.CommandButton cmdGo 
               Caption         =   "Find"
               Height          =   312
               Left            =   2100
               TabIndex        =   132
               Top             =   300
               Width           =   972
            End
            Begin VB.OptionButton optInvoice 
               Caption         =   "Invoice"
               Height          =   255
               Left            =   180
               TabIndex        =   131
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton optPO 
               Caption         =   "Purchase Order"
               Height          =   255
               Left            =   180
               TabIndex        =   130
               Top             =   1440
               Width           =   1575
            End
            Begin VB.OptionButton optAmt 
               Caption         =   "Amount"
               Height          =   255
               Left            =   2100
               TabIndex        =   129
               Top             =   780
               Width           =   1575
            End
            Begin VB.Label Label42 
               Caption         =   "Max records to return"
               Height          =   195
               Left            =   3600
               TabIndex        =   138
               Top             =   360
               Width           =   1515
            End
            Begin VB.Label lblMaxRecords 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "25"
               Height          =   315
               Left            =   5280
               TabIndex        =   135
               Top             =   300
               Width           =   375
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   6960
         Left            =   30
         TabIndex        =   126
         Top             =   360
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   12277
         _Version        =   262144
         TabGuid         =   "FAcctRcv.frx":08DC
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6960
         Left            =   30
         TabIndex        =   94
         Top             =   360
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   12277
         _Version        =   262144
         TabGuid         =   "FAcctRcv.frx":0904
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   6540
            TabIndex        =   139
            Top             =   6360
            Width           =   975
         End
         Begin VB.TextBox txtBatch 
            Height          =   285
            Left            =   1440
            TabIndex        =   102
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Load"
            Height          =   375
            Left            =   3120
            TabIndex        =   103
            Top             =   6360
            Width           =   975
         End
         Begin VB.TextBox txtDisplay 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6720
            TabIndex        =   106
            Top             =   480
            Width           =   2652
         End
         Begin VB.CommandButton cmdReconcile 
            Caption         =   "Reconcile"
            Height          =   375
            Left            =   4260
            TabIndex        =   104
            Top             =   6360
            Width           =   975
         End
         Begin VB.CommandButton cmdPrintTape 
            Caption         =   "Print"
            Height          =   375
            Left            =   5400
            TabIndex        =   105
            Top             =   6360
            Width           =   975
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "0"
            Height          =   375
            Index           =   0
            Left            =   6720
            TabIndex        =   118
            Top             =   3180
            Width           =   1335
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "1"
            Height          =   375
            Index           =   1
            Left            =   6720
            TabIndex        =   115
            Top             =   2700
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "2"
            Height          =   375
            Index           =   2
            Left            =   7440
            TabIndex        =   116
            Top             =   2700
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "3"
            Height          =   375
            Index           =   3
            Left            =   8160
            TabIndex        =   117
            Top             =   2700
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "4"
            Height          =   375
            Index           =   4
            Left            =   6720
            TabIndex        =   112
            Top             =   2220
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "5"
            Height          =   375
            Index           =   5
            Left            =   7440
            TabIndex        =   113
            Top             =   2220
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "6"
            Height          =   375
            Index           =   6
            Left            =   8160
            TabIndex        =   114
            Top             =   2220
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "7"
            Height          =   375
            Index           =   7
            Left            =   6720
            TabIndex        =   109
            Top             =   1740
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "8"
            Height          =   375
            Index           =   8
            Left            =   7440
            TabIndex        =   110
            Top             =   1740
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "9"
            Height          =   375
            Index           =   9
            Left            =   8160
            TabIndex        =   111
            Top             =   1740
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "Backspace"
            Height          =   375
            Index           =   10
            Left            =   6720
            TabIndex        =   107
            Top             =   1260
            Width           =   1335
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "."
            Height          =   375
            Index           =   11
            Left            =   8160
            TabIndex        =   119
            Top             =   3180
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Caption         =   "Clear"
            Height          =   375
            Index           =   12
            Left            =   8160
            TabIndex        =   108
            Top             =   1260
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Caption         =   "+"
            Height          =   1095
            Index           =   13
            Left            =   8880
            TabIndex        =   121
            Top             =   2460
            Width           =   495
         End
         Begin VB.CommandButton cmdKey 
            Caption         =   "--"
            Height          =   1095
            Index           =   14
            Left            =   8880
            TabIndex        =   120
            Top             =   1260
            Width           =   495
         End
         Begin VB.Frame Frame8 
            Caption         =   "Options"
            Height          =   1575
            Left            =   6720
            TabIndex        =   122
            Top             =   3780
            Width           =   2655
            Begin VB.CheckBox ckbSubtraction 
               Caption         =   "Allow Subtraction"
               Height          =   255
               Left            =   480
               TabIndex        =   123
               Top             =   360
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox ckbEnter 
               Caption         =   "'Enter' = '+'"
               Height          =   255
               Left            =   480
               TabIndex        =   124
               Top             =   720
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox ckbDecimal 
               Caption         =   "Automatic Decimal"
               Height          =   255
               Left            =   480
               TabIndex        =   125
               Top             =   1080
               Value           =   1  'Checked
               Width           =   1935
            End
         End
         Begin MSComctlLib.ListView lvwBatch 
            Height          =   5055
            Left            =   180
            TabIndex        =   96
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   8916
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvwTape 
            Height          =   5055
            Left            =   4080
            TabIndex        =   99
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   8916
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label31 
            Caption         =   "Tape Detail"
            Height          =   255
            Left            =   4080
            TabIndex        =   98
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Batch Detail"
            Height          =   255
            Left            =   360
            TabIndex        =   95
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblTapeTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   100
            Top             =   5760
            Width           =   2295
         End
         Begin VB.Label lblBatchTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   97
            Top             =   5760
            Width           =   3615
         End
         Begin VB.Label Label32 
            Caption         =   "Batch Number"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   6360
            Width           =   1215
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6960
         Left            =   30
         TabIndex        =   70
         Top             =   360
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   12277
         _Version        =   262144
         TabGuid         =   "FAcctRcv.frx":092C
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            Height          =   315
            Left            =   8640
            TabIndex        =   90
            Top             =   6540
            Width           =   1215
         End
         Begin VB.Frame Frame6 
            Height          =   1572
            Left            =   120
            TabIndex        =   71
            Top             =   0
            Width           =   9735
            Begin VB.ComboBox cboTerritory 
               Height          =   315
               Left            =   2760
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   240
               Width           =   1815
            End
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "Refresh"
               Height          =   315
               Index           =   1
               Left            =   8400
               TabIndex        =   72
               Top             =   1140
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker dtpUpdateThreshold 
               Height          =   312
               Left            =   2760
               TabIndex        =   73
               Top             =   1080
               Width           =   1812
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   116916225
               CurrentDate     =   37069
            End
            Begin MSComCtl2.UpDown udCredit 
               Height          =   252
               Left            =   3120
               TabIndex        =   74
               Top             =   660
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   450
               _Version        =   393216
               Value           =   3
               BuddyControl    =   "lblMonthThreshold"
               BuddyDispid     =   196664
               OrigLeft        =   3180
               OrigTop         =   660
               OrigRight       =   3420
               OrigBottom      =   912
               Max             =   12
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65537
               Enabled         =   -1  'True
            End
            Begin MSComctlLib.Slider sldRecCount 
               Height          =   255
               Left            =   7200
               TabIndex        =   75
               Top             =   240
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   10
               SmallChange     =   5
               Max             =   100
               SelStart        =   50
               TickFrequency   =   10
               Value           =   50
            End
            Begin MSComCtl2.UpDown udNewCredit 
               Height          =   255
               Left            =   6480
               TabIndex        =   77
               Top             =   660
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   450
               _Version        =   393216
               Value           =   3
               BuddyControl    =   "lblMonthsIncrease"
               BuddyDispid     =   196670
               OrigLeft        =   4920
               OrigTop         =   720
               OrigRight       =   5160
               OrigBottom      =   975
               Max             =   12
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65537
               Enabled         =   -1  'True
            End
            Begin VB.Label Label29 
               Caption         =   "Show Customers with LESS than "
               Height          =   252
               Left            =   240
               TabIndex        =   86
               Top             =   660
               Width           =   2412
            End
            Begin VB.Label lblMonthThreshold 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Left            =   2760
               TabIndex        =   85
               Top             =   660
               Width           =   312
            End
            Begin VB.Label Label28 
               Caption         =   "months credit"
               Height          =   252
               Left            =   3480
               TabIndex        =   84
               Top             =   660
               Width           =   1092
            End
            Begin VB.Label Label27 
               Caption         =   "Select a Sales Territory"
               Height          =   255
               Left            =   240
               TabIndex        =   83
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label26 
               Caption         =   "Number of Customers to Return"
               Height          =   255
               Left            =   4800
               TabIndex        =   82
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label Label25 
               Caption         =   "Skip Customers that have been updated since"
               Height          =   492
               Left            =   240
               TabIndex        =   81
               Top             =   1020
               Width           =   2172
            End
            Begin VB.Label lblNewCredit 
               Caption         =   "Increase credit to"
               Height          =   252
               Left            =   4800
               TabIndex        =   80
               Top             =   660
               Width           =   1272
            End
            Begin VB.Label lblMonthsIncrease 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   252
               Left            =   6120
               TabIndex        =   79
               Top             =   660
               Width           =   312
            End
            Begin VB.Label Label23 
               Caption         =   "months"
               Height          =   255
               Left            =   6840
               TabIndex        =   78
               Top             =   720
               Width           =   615
            End
         End
         Begin GridEX20.GridEX gdxCredLimit 
            Height          =   2715
            Left            =   120
            TabIndex        =   87
            Top             =   1680
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   4789
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            CursorLocation  =   3
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            LockType        =   4
            Options         =   8
            RecordsetType   =   3
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            ColumnsCount    =   15
            Column(1)       =   "FAcctRcv.frx":0954
            Column(2)       =   "FAcctRcv.frx":0AF4
            Column(3)       =   "FAcctRcv.frx":0C6C
            Column(4)       =   "FAcctRcv.frx":0E0C
            Column(5)       =   "FAcctRcv.frx":0FB4
            Column(6)       =   "FAcctRcv.frx":1144
            Column(7)       =   "FAcctRcv.frx":1288
            Column(8)       =   "FAcctRcv.frx":13D0
            Column(9)       =   "FAcctRcv.frx":1554
            Column(10)      =   "FAcctRcv.frx":16E0
            Column(11)      =   "FAcctRcv.frx":186C
            Column(12)      =   "FAcctRcv.frx":19BC
            Column(13)      =   "FAcctRcv.frx":1B58
            Column(14)      =   "FAcctRcv.frx":1CF4
            Column(15)      =   "FAcctRcv.frx":1E9C
            FormatStylesCount=   6
            FormatStyle(1)  =   "FAcctRcv.frx":2044
            FormatStyle(2)  =   "FAcctRcv.frx":2124
            FormatStyle(3)  =   "FAcctRcv.frx":225C
            FormatStyle(4)  =   "FAcctRcv.frx":230C
            FormatStyle(5)  =   "FAcctRcv.frx":23C0
            FormatStyle(6)  =   "FAcctRcv.frx":2498
            ImageCount      =   0
            PrinterProperties=   "FAcctRcv.frx":2550
         End
         Begin GridEX20.GridEX gdxRemarks 
            Height          =   1800
            Left            =   120
            TabIndex        =   89
            Top             =   4680
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   3175
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   270
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   5
            Column(1)       =   "FAcctRcv.frx":2728
            Column(2)       =   "FAcctRcv.frx":28D4
            Column(3)       =   "FAcctRcv.frx":2A4C
            Column(4)       =   "FAcctRcv.frx":2B6C
            Column(5)       =   "FAcctRcv.frx":2C90
            FormatStylesCount=   6
            FormatStyle(1)  =   "FAcctRcv.frx":2DDC
            FormatStyle(2)  =   "FAcctRcv.frx":2EBC
            FormatStyle(3)  =   "FAcctRcv.frx":2FF4
            FormatStyle(4)  =   "FAcctRcv.frx":30A4
            FormatStyle(5)  =   "FAcctRcv.frx":3158
            FormatStyle(6)  =   "FAcctRcv.frx":3230
            ImageCount      =   0
            PrinterProperties=   "FAcctRcv.frx":32E8
         End
         Begin VB.Label Label30 
            Caption         =   " Remarks"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   4440
            Width           =   915
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpCollections 
         Height          =   6960
         Left            =   30
         TabIndex        =   41
         Top             =   360
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   12277
         _Version        =   262144
         TabGuid         =   "FAcctRcv.frx":34C0
      End
      Begin ActiveTabs.SSActiveTabPanel tpOOH 
         Height          =   6960
         Left            =   30
         TabIndex        =   21
         Top             =   360
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   12277
         _Version        =   262144
         TabGuid         =   "FAcctRcv.frx":34E8
         Begin VB.Frame frmAcctDetail 
            Height          =   2112
            Left            =   120
            TabIndex        =   30
            Top             =   4800
            Width           =   9855
            Begin VB.TextBox txtCustPmtTerms 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4380
               Locked          =   -1  'True
               TabIndex        =   140
               TabStop         =   0   'False
               Top             =   1140
               Width           =   975
            End
            Begin VB.CommandButton cmdReturnToCSR 
               Caption         =   "Return To CSR"
               Height          =   375
               Index           =   0
               Left            =   8100
               TabIndex        =   19
               Top             =   660
               Width           =   1335
            End
            Begin VB.TextBox txtLastPayDate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4380
               Locked          =   -1  'True
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtLastPayAmt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4380
               Locked          =   -1  'True
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   540
               Width           =   975
            End
            Begin VB.TextBox txtHighBalance 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4380
               Locked          =   -1  'True
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtOver90 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   1740
               Width           =   975
            End
            Begin VB.TextBox txtOver60 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox txtOver45 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   1140
               Width           =   975
            End
            Begin VB.TextBox txtOver30 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtCurBalance 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   540
               Width           =   975
            End
            Begin VB.TextBox txtTotalBalance 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdRelease 
               Caption         =   "Commit Order"
               Height          =   375
               Index           =   0
               Left            =   8100
               TabIndex        =   20
               Top             =   1080
               Width           =   1335
            End
            Begin VB.CommandButton cmdViewOrder 
               Caption         =   "View Order"
               Height          =   375
               Index           =   0
               Left            =   8100
               TabIndex        =   18
               Top             =   240
               Width           =   1335
            End
            Begin MMRemark.RemarkViewer rvROOH 
               Height          =   804
               Left            =   6720
               TabIndex        =   91
               Top             =   300
               Width           =   804
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ARCustLoad"
               Caption         =   "Customer Remarks"
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer PmtTerm"
               Height          =   255
               Left            =   2700
               TabIndex        =   141
               Top             =   1140
               Width           =   1455
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "Last Pmt Date"
               Height          =   252
               Left            =   2700
               TabIndex        =   39
               Top             =   840
               Width           =   1452
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "Last Pmt Amt"
               Height          =   252
               Left            =   2700
               TabIndex        =   38
               Top             =   540
               Width           =   1452
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Total Balance"
               Height          =   252
               Left            =   180
               TabIndex        =   37
               Top             =   240
               Width           =   1092
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Current"
               Height          =   252
               Left            =   180
               TabIndex        =   36
               Top             =   540
               Width           =   1092
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Over 30"
               Height          =   252
               Left            =   180
               TabIndex        =   35
               Top             =   840
               Width           =   1092
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "Highest Balance"
               Height          =   252
               Left            =   2940
               TabIndex        =   34
               Top             =   240
               Width           =   1212
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Over 90"
               Height          =   252
               Left            =   180
               TabIndex        =   33
               Top             =   1740
               Width           =   1092
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Over 60"
               Height          =   252
               Left            =   180
               TabIndex        =   32
               Top             =   1440
               Width           =   1092
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Over 45"
               Height          =   252
               Left            =   180
               TabIndex        =   31
               Top             =   1140
               Width           =   1092
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Grid Controls"
            Height          =   1455
            Left            =   3120
            TabIndex        =   23
            Top             =   120
            Width           =   6855
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "Refresh"
               Height          =   375
               Index           =   0
               Left            =   4560
               TabIndex        =   17
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox chkAutoRefresh 
               Caption         =   "Auto Refresh"
               Height          =   255
               Left            =   4560
               TabIndex        =   16
               Top             =   660
               Width           =   1935
            End
            Begin VB.CheckBox chkShowGroupBy 
               Caption         =   "Enable GroupBy"
               Height          =   255
               Left            =   360
               TabIndex        =   15
               Top             =   360
               Width           =   1815
            End
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   60000
               Left            =   6060
               Top             =   300
            End
            Begin VB.CommandButton cmdCollapseAll 
               Caption         =   "Collapse All"
               Enabled         =   0   'False
               Height          =   375
               Left            =   240
               TabIndex        =   25
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton cmdExpandAll 
               Caption         =   "Expand All"
               Enabled         =   0   'False
               Height          =   375
               Left            =   1560
               TabIndex        =   24
               Top             =   840
               Width           =   1095
            End
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   252
               Left            =   4800
               TabIndex        =   26
               Top             =   960
               Visible         =   0   'False
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   450
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "lblMinutes"
               BuddyDispid     =   196709
               OrigLeft        =   360
               OrigTop         =   4440
               OrigRight       =   600
               OrigBottom      =   4695
               Max             =   30
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65537
               Enabled         =   -1  'True
            End
            Begin VB.Label lblMinutes 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
               Height          =   252
               Left            =   4560
               TabIndex        =   28
               Top             =   960
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Label lblRefreshInterval 
               Caption         =   "Frequency (minutes)"
               Height          =   252
               Left            =   5160
               TabIndex        =   27
               Top             =   960
               Visible         =   0   'False
               Width           =   1452
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Filter By Collector"
            Height          =   1455
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   2895
            Begin VB.ComboBox cboCollector 
               Height          =   315
               Index           =   0
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton cmdMyCust 
               Caption         =   "My Customers"
               Height          =   375
               Left            =   240
               TabIndex        =   14
               Top             =   840
               Width           =   1215
            End
         End
         Begin GridEX20.GridEX gdxOrdersOnHold 
            Height          =   3135
            Left            =   120
            TabIndex        =   29
            Top             =   1680
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   5530
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
            ShowToolTips    =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            CursorLocation  =   3
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            Options         =   8
            RecordsetType   =   1
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   1
            ColumnHeaderHeight=   270
            ColumnsCount    =   8
            Column(1)       =   "FAcctRcv.frx":3510
            Column(2)       =   "FAcctRcv.frx":3624
            Column(3)       =   "FAcctRcv.frx":3730
            Column(4)       =   "FAcctRcv.frx":3864
            Column(5)       =   "FAcctRcv.frx":3994
            Column(6)       =   "FAcctRcv.frx":3AC4
            Column(7)       =   "FAcctRcv.frx":3C58
            Column(8)       =   "FAcctRcv.frx":3DD4
            FmtConditionsCount=   1
            FmtCondition(1) =   "FAcctRcv.frx":3F0C
            FormatStylesCount=   6
            FormatStyle(1)  =   "FAcctRcv.frx":3FD0
            FormatStyle(2)  =   "FAcctRcv.frx":40B0
            FormatStyle(3)  =   "FAcctRcv.frx":41E8
            FormatStyle(4)  =   "FAcctRcv.frx":4298
            FormatStyle(5)  =   "FAcctRcv.frx":434C
            FormatStyle(6)  =   "FAcctRcv.frx":4424
            ImageCount      =   0
            PrinterProperties=   "FAcctRcv.frx":44DC
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpManageCust 
         Height          =   6960
         Left            =   30
         TabIndex        =   40
         Top             =   360
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   12277
         _Version        =   262144
         TabGuid         =   "FAcctRcv.frx":46B4
         Begin VB.CommandButton cmdEditCC 
            Caption         =   "Edit Credit Cards"
            Height          =   312
            Left            =   1800
            TabIndex        =   194
            Top             =   5040
            Width           =   1452
         End
         Begin VB.CommandButton cmdEditContacts 
            Caption         =   "Edit Contacts"
            Height          =   312
            Left            =   1800
            TabIndex        =   193
            Top             =   4560
            Width           =   1452
         End
         Begin VB.TextBox txtExemptCert 
            Enabled         =   0   'False
            Height          =   288
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   190
            Top             =   5040
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Frame frmCustStatus 
            Caption         =   "Customer Detail"
            Height          =   3015
            Left            =   120
            TabIndex        =   45
            Top             =   1320
            Width           =   9615
            Begin CurrControl.CurrencyInput txtCreditLimit 
               Height          =   315
               Left            =   5100
               TabIndex        =   10
               Top             =   2460
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
            End
            Begin VB.ComboBox cboPmtTerms 
               Height          =   315
               Left            =   5100
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   2100
               Width           =   1455
            End
            Begin VB.TextBox txtHoldStatus 
               Appearance      =   0  'Flat
               Height          =   288
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   1740
               Width           =   972
            End
            Begin VB.TextBox txtAgingDate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   288
               Index           =   1
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   1380
               Width           =   975
            End
            Begin VB.TextBox txtCustID 
               Appearance      =   0  'Flat
               Height          =   288
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   300
               Width           =   1335
            End
            Begin VB.TextBox txtARStatus 
               Appearance      =   0  'Flat
               Height          =   288
               Left            =   5100
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   1740
               Width           =   975
            End
            Begin VB.ComboBox cboCustType 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   2100
               Width           =   1455
            End
            Begin VB.TextBox txtCSZ 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   900
               Width           =   2295
            End
            Begin VB.TextBox txtCustName 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   600
               Width           =   2295
            End
            Begin VB.ComboBox cboCollector 
               Height          =   315
               Index           =   2
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   2460
               Width           =   1455
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "Save"
               Enabled         =   0   'False
               Height          =   312
               Left            =   7680
               TabIndex        =   12
               Top             =   2460
               Width           =   1452
            End
            Begin VB.Frame frmARStatus 
               Caption         =   "AR Status Type"
               Height          =   1212
               Left            =   4200
               TabIndex        =   46
               Top             =   240
               Width           =   3132
               Begin VB.OptionButton opVIP 
                  Caption         =   "Never on Hold (VIP)"
                  Height          =   192
                  Left            =   120
                  TabIndex        =   4
                  ToolTipText     =   "Never on hold"
                  Top             =   300
                  Width           =   2292
               End
               Begin VB.OptionButton opManual 
                  Caption         =   "Always on Hold "
                  Height          =   192
                  Left            =   120
                  TabIndex        =   5
                  ToolTipText     =   "Always on hold"
                  Top             =   540
                  Width           =   2232
               End
               Begin VB.OptionButton opAuto 
                  Caption         =   "Hold determined by Status Update"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   6
                  ToolTipText     =   "On hold if over and/or late"
                  Top             =   780
                  Width           =   2832
               End
            End
            Begin VB.TextBox txtCreditLimit1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5100
               TabIndex        =   11
               Top             =   2460
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "Payment Terms"
               Height          =   252
               Left            =   3600
               TabIndex        =   93
               Top             =   2100
               Width           =   1272
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "AR Status"
               Height          =   312
               Left            =   3996
               TabIndex        =   68
               Top             =   1740
               Width           =   912
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "On Hold"
               Height          =   252
               Left            =   480
               TabIndex        =   66
               Top             =   1740
               Width           =   612
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               Caption         =   "Aging Date"
               Height          =   312
               Left            =   120
               TabIndex        =   65
               Top             =   1380
               Width           =   972
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "CustName"
               Height          =   252
               Left            =   240
               TabIndex        =   61
               Top             =   660
               Width           =   852
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "CustID"
               Height          =   255
               Left            =   360
               TabIndex        =   60
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "CustType"
               Height          =   252
               Left            =   240
               TabIndex        =   50
               Top             =   2100
               Width           =   852
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Credit Limit"
               Height          =   252
               Left            =   4020
               TabIndex        =   49
               Top             =   2460
               Width           =   852
            End
            Begin VB.Label lblCustType 
               Height          =   255
               Left            =   1200
               TabIndex        =   48
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "Collector"
               Height          =   252
               Left            =   240
               TabIndex        =   47
               Top             =   2460
               Width           =   852
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Find Existing Customer"
            Height          =   1035
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   9615
            Begin VB.ComboBox cboCustSearch 
               Height          =   315
               ItemData        =   "FAcctRcv.frx":46DC
               Left            =   5160
               List            =   "FAcctRcv.frx":46EF
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   480
               Width           =   1815
            End
            Begin VB.CommandButton cmdFindCust 
               Caption         =   "Find"
               Height          =   315
               Left            =   8110
               TabIndex        =   3
               Top             =   480
               Width           =   975
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtCustSearch 
               Height          =   285
               Left            =   1320
               TabIndex        =   1
               Top             =   480
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   503
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
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Search Type"
               Height          =   255
               Left            =   3840
               TabIndex        =   44
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Find"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   480
               Width           =   855
            End
         End
         Begin MMRemark.RemarkViewer rvCustomer 
            Height          =   810
            Left            =   480
            TabIndex        =   195
            Top             =   4560
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1429
            ContextID       =   "ARCustLoad"
            Caption         =   "Customer Remarks"
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "STax Exemption Certificate"
            Height          =   255
            Left            =   5520
            TabIndex        =   192
            Top             =   5040
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label45 
            Caption         =   "(max 15 characters)"
            Height          =   195
            Left            =   7800
            TabIndex        =   191
            Top             =   5400
            Visible         =   0   'False
            Width           =   1635
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   6960
         Left            =   30
         TabIndex        =   152
         Top             =   360
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   12277
         _Version        =   262144
         TabGuid         =   "FAcctRcv.frx":4728
         Begin VB.Frame Frame12 
            Height          =   5295
            Left            =   120
            TabIndex        =   155
            Top             =   1440
            Width           =   9735
            Begin VB.TextBox txtAuthCountry 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   1
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   175
               TabStop         =   0   'False
               Top             =   4800
               Width           =   2295
            End
            Begin VB.TextBox txtAuthZip 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   1
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   174
               TabStop         =   0   'False
               Top             =   4440
               Width           =   1575
            End
            Begin VB.TextBox txtAuthAddrName 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   1
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   171
               TabStop         =   0   'False
               Top             =   3000
               Width           =   2295
            End
            Begin VB.TextBox txtAuthAddrLine1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   1
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   170
               TabStop         =   0   'False
               Top             =   3360
               Width           =   2295
            End
            Begin VB.TextBox txtAuthAddrLine2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   1
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   169
               TabStop         =   0   'False
               Top             =   3720
               Width           =   2295
            End
            Begin VB.TextBox txtAuthCity 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   1
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   168
               TabStop         =   0   'False
               Top             =   4080
               Width           =   2295
            End
            Begin VB.TextBox txtAuthState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   1
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   167
               TabStop         =   0   'False
               Top             =   4440
               Width           =   615
            End
            Begin VB.TextBox txtAuthCountry 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   0
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   165
               TabStop         =   0   'False
               Top             =   4800
               Width           =   2295
            End
            Begin VB.TextBox txtAuthZip 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   0
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   164
               TabStop         =   0   'False
               Top             =   4440
               Width           =   1575
            End
            Begin VB.CommandButton cmdMovetoAuth0 
               Caption         =   "<"
               Height          =   375
               Left            =   4680
               TabIndex        =   187
               Top             =   1800
               Width           =   375
            End
            Begin VB.CommandButton cmdMovetoNonAuth91 
               Caption         =   ">"
               Height          =   375
               Left            =   4680
               TabIndex        =   186
               Top             =   1080
               Width           =   375
            End
            Begin VB.TextBox txtAuthState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   0
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   162
               TabStop         =   0   'False
               Top             =   4440
               Width           =   615
            End
            Begin VB.TextBox txtAuthCity 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   0
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   161
               TabStop         =   0   'False
               Top             =   4080
               Width           =   2295
            End
            Begin VB.ListBox lstNonAuthAddr 
               Height          =   2205
               Left            =   5400
               TabIndex        =   188
               Top             =   600
               Width           =   4095
            End
            Begin VB.ListBox lstAuthAddr 
               Height          =   2205
               ItemData        =   "FAcctRcv.frx":4750
               Left            =   240
               List            =   "FAcctRcv.frx":4752
               TabIndex        =   185
               Top             =   600
               Width           =   4095
            End
            Begin VB.TextBox txtAuthAddrLine2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   0
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   160
               TabStop         =   0   'False
               Top             =   3720
               Width           =   2295
            End
            Begin VB.TextBox txtAuthAddrLine1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   0
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   158
               TabStop         =   0   'False
               Top             =   3360
               Width           =   2295
            End
            Begin VB.TextBox txtAuthAddrName 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   288
               Index           =   0
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   156
               TabStop         =   0   'False
               Top             =   3000
               Width           =   2295
            End
            Begin VB.Label lblCustID 
               Height          =   255
               Left            =   2640
               TabIndex        =   182
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               Caption         =   "Non-Authorized Addresses"
               Height          =   255
               Left            =   6240
               TabIndex        =   181
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "Country"
               Height          =   255
               Index           =   7
               Left            =   5640
               TabIndex        =   180
               Top             =   4800
               Width           =   975
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "State / Zip Code"
               Height          =   255
               Index           =   6
               Left            =   5400
               TabIndex        =   179
               Top             =   4440
               Width           =   1215
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "Country"
               Height          =   255
               Index           =   5
               Left            =   360
               TabIndex        =   178
               Top             =   4800
               Width           =   975
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "State / Zip Code"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   177
               Top             =   4440
               Width           =   1215
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "City"
               Height          =   255
               Index           =   3
               Left            =   5760
               TabIndex        =   176
               Top             =   4080
               Width           =   855
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               Caption         =   " Name"
               Height          =   255
               Index           =   1
               Left            =   5760
               TabIndex        =   173
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "Address"
               Height          =   255
               Index           =   1
               Left            =   5760
               TabIndex        =   172
               Top             =   3360
               Width           =   855
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "City"
               Height          =   255
               Index           =   2
               Left            =   480
               TabIndex        =   166
               Top             =   4080
               Width           =   855
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               Caption         =   "Authorized Addresses for:"
               Height          =   255
               Left            =   480
               TabIndex        =   163
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "Address"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   159
               Top             =   3360
               Width           =   855
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               Caption         =   "Name"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   157
               Top             =   3000
               Width           =   855
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Enter Exact Customer ID"
            Height          =   1155
            Left            =   120
            TabIndex        =   153
            Top             =   120
            Width           =   9735
            Begin VB.CommandButton cmdLoadCust 
               Caption         =   "Load"
               Height          =   315
               Left            =   4200
               TabIndex        =   184
               Top             =   480
               Width           =   975
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtCustLoad 
               Height          =   285
               Left            =   1320
               TabIndex        =   183
               Top             =   480
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   503
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
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer ID"
               Height          =   255
               Left            =   120
               TabIndex        =   154
               Top             =   480
               Width           =   1095
            End
         End
      End
   End
End
Attribute VB_Name = "FAcctRcv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'validation rule classid
Private Const ccCustomer = 1

'AR Status Type Codes
'Warning: these must match those in ARStatus DLL (a separate project)
Private Const klARS_VIP = 1
Private Const klARS_GOOD = 2
Private Const klARS_MANUAL_HOLD = 3
Private Const klARS_AUTO_HOLD = 4

'*********************************************************************
'ActiveTab control tab stuff
'
'It's important to keep this stuff up to date.
'Warning: if you change this stuff, take a look at EnableTabs() as well

Private Const klNumTabs = 7

Private m_asTabRights(1 To klNumTabs) As String

Private Enum TabMainIndexes
    tmiOrdersOnHold = 1
    tmiCreditCards = 2  'todo: remove
    tmiManageCust = 3
    tmiCollections = 4  'todo: remove
    tmiCreditManager = 5
    tmiTurboTenKey = 6
    tmiResearch = 7
End Enum

'*********************************************************************

Private WithEvents m_oBrokenRules As BrokenRules
Attribute m_oBrokenRules.VB_VarHelpID = -1
Private m_oCustomer As Customer
Attribute m_oCustomer.VB_VarHelpID = -1

Private m_oPmtTerms As PaymentTerms

Private m_lWindowID As Long
Private m_lMinute As Long

Private m_oRstOrders As ADODB.Recordset
Private m_oRstOwed As ADODB.Recordset
Private m_oRstLate As ADODB.Recordset
Private m_oRstResearch As ADODB.Recordset

Private m_sARProfileID As String
Private m_iPercentOwed As Integer
Private m_iPercentLate As Integer
Private m_iOwedRSSize As Integer
Private m_iLateRSSize As Integer

Private m_sUserName As String
Private m_sOPID As String

Private m_lCustKey As Long          'NOTE: this is also used by code copied over from FBilling
Private m_bLoading As Boolean
Private m_bAcctListEmpty As Boolean
Private m_bStatusTypeChg As Boolean

' SALESTAX stuff copied over from FBilling
Private m_sCustID As String
Private m_sStateID As String
Private m_lAddrKey As Long
Private m_sExemptCert As String


Private m_oARStatus As ARStatus.CalcARStatus

Private WithEvents m_gwOrdersOnHold As GridEXWrapper
Attribute m_gwOrdersOnHold.VB_VarHelpID = -1

'Credit Manager variables
Private m_CLList As CreditLimitList
Private m_CLListTemp As CreditLimitList
Private m_CLUpdateList As CreditLimitList

'Turbo Ten Key variables
    
Private m_dBalance As Double
Private m_dCurEntry As Double
Private m_bNew As Boolean
Private m_sBatchCmnt As String

'To Pass mouse down location to HitTest
Private m_sngX As Single
Private m_sngY As Single

'This is the list item if the user has double
'clicked and wants to edit a tape entry
Private m_EditItem As ListItem

'Display Mask
'Private Const ksDisplayMask = "###,###,###.00"

Private Const klVScrollBarWidth = 350
Private Const klReconcileColWidth = 300

'Variables for Managing Customers backup
'used to catch changes for Event reporting
Private m_sBackupCollector As String

Private m_sBackupPmtTerms As String
Private m_sBackupARStatus As String
Private m_lBackupCreditLimit As Double

' Added for NA Auth Addr tab
Private Type AuthAddr
    lShipDays As Long
    lAddrKey As Long
    sAddrName As String
    sAddrLine1 As String
    sAddrLine2 As String
    sCity As String
    sStateID As String
    sCountryID As String
    sZipCode As String
End Type

Private m_utdAuthAddr() As AuthAddr
Private m_utdNonAuthAddr() As AuthAddr


'**********************************************************************************
'   The Form
'**********************************************************************************

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
    SetCaption "Accounts Receivable"
    LoadControls
    
    'Turbo 10-Key init
    SetUpTape
    SetUpBatch
    
    'Print Statements
    SetUpStatement
    
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Dim lTemp As Long
    
    tabMain.width = Me.width - 180
    tabMain.Height = Me.Height - 495
    
    gdxOrdersOnHold.width = tabMain.width - 315
    gdxOrdersOnHold.Height = tabMain.Height - 4300
    frmAcctDetail.Top = gdxOrdersOnHold.Top + gdxOrdersOnHold.Height + 30
    
    gdxResearch.width = tabMain.width - 315
    gdxResearch.Height = tabMain.Height - 2840
    
    gdxCredLimit.width = gdxOrdersOnHold.width
    
    lTemp = tabMain.Height - 2835
    gdxCredLimit.Height = lTemp * (2715 / 4515)
    gdxRemarks.Height = lTemp * (1800 / 4515)
    
    Label30.Top = gdxCredLimit.Top + gdxCredLimit.Height + 30
    gdxRemarks.width = gdxCredLimit.width
    gdxRemarks.Top = Label30.Top + Label30.Height + 30
    cmdUpdate.Top = gdxRemarks.Height + gdxRemarks.Top + 60
    cmdUpdate.Left = gdxRemarks.Left + gdxRemarks.width - cmdUpdate.width

'??? do we want to do this here?
    gdxOrdersOnHold.Refresh
    
'***DH 4/25/08 Len removed all controls from the CreditCardOrders tab.
'    gdxCCOrders.Refresh

    gdxResearch.Refresh
    gdxCredLimit.Refresh
    gdxRemarks.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '3/31/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwOrdersOnHold = Nothing

    m_oBrokenRules.Destroy
    Set m_oBrokenRules = Nothing
    Set m_oRstOrders = Nothing
    Set m_oRstOwed = Nothing
    Set m_oRstLate = Nothing

    'if we instantiated an ARStatus component, destroy it
    If Not m_oARStatus Is Nothing Then Set m_oARStatus = Nothing

    If cmdSave.Enabled = True Then
        If vbYes = MsgBox("Would you like to save your changes?", vbYesNo, "Save Changes") Then
            cmdSave_Click
        End If
    End If
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
    RefreshHoldList
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Public Sub DoShowHelp()
    ShowHelp "AcctRcv"
End Sub


Private Sub LoadControls()
    Dim rstCust As ADODB.Recordset

    m_bLoading = True

    EnableTabs
    
'Initialize the Orders On Hold tab

    Set m_gwOrdersOnHold = New GridEXWrapper
    m_gwOrdersOnHold.Grid = gdxOrdersOnHold
        
    '*** 9/4/03 LR
    'LoadCollectorCombo cboCollector(0), "spOPCollectors"
    LoadCombo cboCollector(0), g_rstCollectors, "UserID"
    cboCollector(0).AddItem "All", 0
    SetComboByText cboCollector(0), GetUserName, True

    If HasRight(k_sRightARReleaseOrder) Then
        cmdRelease(0).Enabled = True
        cmdReturnToCSR(0).Enabled = True
    Else
        cmdRelease(0).Enabled = False
        cmdReturnToCSR(0).Enabled = False
    End If

    If HasRight(k_sRightARViewCustomer) Then
        txtCreditLimit.ATMMode = g_bATMMode
    End If


'Initialize the Manage Customer tab

    Set m_oCustomer = New Customer

    Set rstCust = LoadDiscRst("SELECT CustClassID FROM tarCustClass WHERE companyid='cpc'")
    LoadCombo cboCustType, rstCust, "CustClassID"
    Set rstCust = Nothing

    Set m_oPmtTerms = New PaymentTerms

    LoadCombo cboCollector(2), g_rstCollectors, "UserID", , , True
    
    Call InitMgCustTab
    
    Set m_oBrokenRules = New BrokenRules
    m_oBrokenRules.Form = Me
    LoadValidationRules

    m_bStatusTypeChg = False
                
      
'Initialize Credit Manager

    Dim rst As ADODB.Recordset
    
    Set rst = LoadDiscRst("SELECT SalesTerritoryKey, SalesTerritoryID FROM tarSalesTerritory WHERE CompanyID = 'CPC' Order by salesTerritoryID")
    LoadCombo cboTerritory, rst, "SalesTerritoryID", "SalesTerritoryKey"
    Set rst = Nothing
    
    lblMonthsIncrease.Caption = udNewCredit.value
    lblMonthThreshold.Caption = udCredit.value
    
    'PRN#96 OK
    dtpUpdateThreshold.value = Now - 30

    m_bLoading = False

End Sub


Private Sub cmdStatementDisplay_Click()
    ' Research tab, Frame9 - set visible property to false - not used by A/R.
    ' this space is now used for finding all credit card transactions for a specific OP#.
    DisplayStatementCount
End Sub


Private Sub gdxOrdersOnHold_DblClick()
    rvROOH.Edit
End Sub


'**********************************************************************************
'   The Tab Control
'**********************************************************************************

'6/15/04 LR
'The Credit Card tab was removed 1/14/04.
'William's HACK below effectively disables, but does not delete the credit card tab.
'I've removed all of the unused credit card code

Private Sub EnableTabs()
    Dim TabIndex As Integer
    
    m_asTabRights(1) = k_sRightARViewOnHold
    m_asTabRights(2) = k_sRightARViewCreditCard
    m_asTabRights(3) = k_sRightARViewCustomer
    m_asTabRights(4) = k_sRightARViewCollections
    m_asTabRights(5) = k_sRightARViewCredit
    m_asTabRights(6) = k_sRightARViewTenKey
    m_asTabRights(7) = k_sRightARViewResearch

    'decide which tab will be selected
    For TabIndex = 1 To klNumTabs
        If HasRight(m_asTabRights(TabIndex)) Then
            
            tabMain.Tabs(TabIndex).Selected = True
            Exit For
        End If
    Next TabIndex
    
    For TabIndex = 1 To klNumTabs
        If Not tabMain.Tabs(TabIndex).Selected Then
            'Hack: this turns off the two tabs we've deleted (clean this up)
            If TabIndex <> tmiCreditCards And TabIndex <> tmiCollections Then
                tabMain.Tabs(TabIndex).Visible = HasRight(m_asTabRights(TabIndex))
            End If
        End If
    Next TabIndex
End Sub


Private Sub tabMain_TabClick(ByVal NewTab As ActiveTabs.SSTab)
    Select Case NewTab.Index
                                                                    
        Case tmiManageCust
            m_bLoading = True
            
            cmdFindCust.Default = True
            TryToSetFocus txtCustSearch
            
            m_bLoading = False
            
    End Select

    ' if we're changing tabs to something other than Maintain Customer and Maintain Customer is dirty
    ' then prompt to save or discard
    If NewTab.Index <> tmiManageCust And cmdSave.Enabled = True Then
        If vbYes = MsgBox("Would you like to save your changes?", vbYesNo, "Save Changes") Then
            cmdSave_Click
        End If
    End If
    
    With m_oBrokenRules
        .EnableClass ccCustomer, (NewTab.Index = tmiManageCust)
        .Validate
    End With
End Sub


'**********************************************************************************
'   Control Arrays
'**********************************************************************************

Private Sub cboCollector_Click(Index As Integer)
    Select Case Index
                                                                        
        'a selection is made from the Collector combo on the Orders On Hold tab
        Case 0
            FilterHoldList
            AttachGrid gdxOrdersOnHold, m_oRstOrders

        'a selection is made from the Collector combo on the Manage Customers tab
        Case 2
            If Not m_bLoading Then cmdSave.Enabled = True

    End Select
End Sub


Private Sub cmdRefresh_Click(Index As Integer)
    Select Case Index
        Case 0                  'Release Orders On Hold
            RefreshHoldList
            ResetTimer
            
        Case 1                  'Credit Manager
            SetWaitCursor True
    
            If gdxCredLimit.ItemCount > 0 Then
                gdxCredLimit.Refetch
            End If
            
            Set m_CLList = New CreditLimitList
            m_CLList.LoadData cboTerritory.ItemData(cboTerritory.ListIndex), _
                            dtpUpdateThreshold.value, udCredit.value, _
                            udNewCredit.value, sldRecCount.value, 0
                                            
            With gdxCredLimit
                .Row = -1
                .HoldSortSettings = True
                .HoldFields
                .ItemCount = m_CLList.Count
                .Refetch
                .Row = 1
                TryToSetFocus gdxCredLimit
            End With
            
            UpdateGridRemark
            Set m_CLUpdateList = New CreditLimitList
            SetWaitCursor False

    End Select
End Sub


Private Sub cmdRelease_Click(Index As Integer)
    
    Select Case Index
        Case 0:     'Orders On Hold
            If gdxOrdersOnHold.RowIndex(gdxOrdersOnHold.Row) <> 0 Then
                If OKForRelease(CLng(m_gwOrdersOnHold.value("OPKey")), ItemStatusCode.iscARHold) Then
                    ReleaseOrder CLng(m_gwOrdersOnHold.value("OPKey"))
                End If
                RefreshHoldList
            End If
            
    End Select

End Sub


'Might do some cleanup here since we no longer to CC Authorizations here. 10/9/04 LR

'called by cmdRelease_Click()
'for releasing both Orders on Hold and Credit Card Authorizations
'the value for oStatus differs in each case

Private Function OKForRelease(lOPKey As Long, oStatus As ItemStatusCode) As Boolean
    Dim rst As ADODB.Recordset
    
    Set rst = LoadDiscRst("SELECT tcpSO.StatusCode, isnull(tsoSalesOrder.Status, 0) as Status " _
                        & "FROM tcpSO left outer join tsoSalesOrder on tsoSalesOrder.SOKey = tcpSO.SOKey " _
                        & "WHERE tcpSO.OPKey = " & lOPKey)

    'if the OP order exists
    If Not rst.EOF Then
    
        '*** What are we trying to say here?
        'if Sage SO is not Cancelled (this matters only for CreditCardAuth)
        If rst.Fields("Status").value <> 3 Then
        
            'if OP Status > specified status return True (CAREFUL: this relies on a specific ordering)
            If (rst.Fields("StatusCode").value > oStatus) Then
                OKForRelease = False
                msg "This order has already been been committed.", vbCritical + vbOKOnly
            Else
                OKForRelease = True
            End If
        End If
    End If
End Function


'Called By:
'   cmdRelease_Click() release order on hold
    
'commit the order to Sage
Private Sub ReleaseOrder(lOPKey As Long)
    Dim oOrder As Order
    Dim oFrmChoosePmtTerms As FChoosePmtTerms
    
    Set oOrder = New Order
    Set oFrmChoosePmtTerms = New FChoosePmtTerms
    
    oOrder.Load lOPKey
    oFrmChoosePmtTerms.ReleaseAROrder oOrder
End Sub


Private Sub txtCreditLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not m_bLoading Then cmdSave.Enabled = True
End Sub


Private Sub txtCustSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtCustSearch.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            cmdFindCust_Click
        End If
    End If
End Sub


Private Sub txtCustSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub


Private Sub txtCustSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    cboCustSearch.ListIndex = GetSearchType(txtCustSearch.text)
    m_oBrokenRules.Validate txtCustSearch
End Sub


'Customer Lookup function for the Manage Customer Tab

Private Sub cmdFindCust_Click()

    m_bLoading = True
    Call InitMgCustTab
    m_bLoading = False

    If Len(txtCustSearch.text) = 0 Then
        Exit Sub      'no error on empty field
    End If

    m_oBrokenRules.Validate txtCustSearch
    
    m_lCustKey = Search.FindCustomer(txtCustSearch.text, cboCustSearch.ListIndex, m_oCustomer)

    If m_lCustKey <> 0 Then
        txtCustSearch.text = ""
        
        DisplayStatus
        
    Else
        txtCustSearch.SetFocus

    End If

End Sub


'**********************************************************************************
' tabMain(1)    Orders On Hold
'**********************************************************************************

Private Sub cmdMyCust_Click()
    SetComboByText cboCollector(0), GetUserName
End Sub

Private Sub chkShowGroupBy_Click()
    gdxOrdersOnHold.GroupByBoxVisible = (chkShowGroupBy.value = vbChecked)
End Sub

Private Sub cmdCollapseAll_Click()
    gdxOrdersOnHold.CollapseAll
End Sub

Private Sub cmdExpandAll_Click()
    gdxOrdersOnHold.ExpandAll
End Sub


Private Sub chkAutoRefresh_Click()
    If chkAutoRefresh.value = vbChecked Then
        Timer1.Enabled = True
        lblMinutes.Visible = True
        lblRefreshInterval.Visible = True
        UpDown1.Visible = True
    Else
        lblMinutes.Visible = False
        lblRefreshInterval.Visible = False
        UpDown1.Visible = False
    End If
End Sub


Private Sub Timer1_Timer()
    m_lMinute = m_lMinute + 1
    TestTimer
End Sub


Private Sub TestTimer()
    If m_lMinute >= UpDown1.value Then
        RefreshHoldList
        ResetTimer
    End If
End Sub


Private Sub UpDown1_LostFocus()
    ResetTimer
End Sub


Private Sub ResetTimer()
    Timer1.Enabled = False
    Timer1.Enabled = (chkAutoRefresh.value = vbChecked)
    m_lMinute = 0
End Sub


Private Sub gdxOrdersOnHold_AfterGroupChange()
    With gdxOrdersOnHold.Groups
        cmdExpandAll.Enabled = (.Count > 0)
        cmdCollapseAll.Enabled = (.Count > 0)
    End With
End Sub


Private Sub gdxOrdersOnHold_SelectionChange()
    DisplayAcctDetail m_gwOrdersOnHold.value("CustID")
End Sub


Private Sub m_gwOrdersOnHold_RowChosen()
    DisplayAcctDetail m_gwOrdersOnHold.value("CustID")
End Sub


Private Sub RefreshHoldList()
    ' cache the currently selected record
    m_sOPID = m_gwOrdersOnHold.value("OPKey")
 
'*** NOTE: an ADO issue
    'Set m_oRstOrders = CallSP("spOPOrdersOnHold")
    LoadList "spCPCarOrdersOnHold", m_oRstOrders
    
    FilterHoldList
    AttachGrid gdxOrdersOnHold, m_oRstOrders
    
    ' restore previously selected record (if still there)
    gdxOrdersOnHold.Find gdxOrdersOnHold.Columns("OPKey").Index, jgexEqual, m_sOPID
End Sub


' Release Orders On Hold tab
' Filter the list by (authorized user)

Private Sub FilterHoldList()
    ' if the list has not been loaded
    If m_oRstOrders Is Nothing Then Exit Sub
    
    If cboCollector(0).text = "All" Then
        m_oRstOrders.Filter = adFilterNone
    Else
        m_oRstOrders.Filter = "collector='" & cboCollector(0).text & "'"
    End If
    ' if the list is empty
    If m_oRstOrders.EOF Then ClearAcctDetail
End Sub


Private Sub DisplayAcctDetail(ByVal CustID As String)
    Dim ocmd As ADODB.Command
    Dim orst As ADODB.Recordset
    
    'Only assign attributes to Remark button if FAcctRcv is active.
    'This change is related with an automation error.
    If MDIMain.ActiveForm.Name = "FAcctRcv" Then
        'Use on error syntax to avoid can't show modally error here
        On Error Resume Next
        rvROOH.OwnerID = CustID
    End If
    
    ' This guard catches the case where this sub is invoked by
    ' gdxOrdersOnHold_SelectionChange when the Collapse/Expand All
    ' methods are invoked.
    If Len(CustID) > 0 Then
        Set ocmd = CreateCommandSP("spCPCarARSummary")
        With ocmd
            .Parameters("@CustID").value = Trim(CustID)
        End With
    
        Set orst = New ADODB.Recordset
        orst.Open ocmd, , , adLockReadOnly
        
        txtTotalBalance(1) = Format(orst.Fields("TotBal"), "#,###.00")
        txtCurBalance = Format(orst.Fields("CurBal"), "#,###.00")
        txtOver30 = Format(orst.Fields("Over30"), "#,###.00")
        txtOver45 = Format(orst.Fields("Over45"), "#,###.00")
        txtOver60 = Format(orst.Fields("Over60"), "#,###.00")
        txtOver90 = Format(orst.Fields("Over90"), "#,###.00")
        txtHighBalance = Format(orst.Fields("HighestBal"), "#,###.00")
        txtLastPayAmt = IIf(IsNull(orst.Fields("LastPmtAmt")), "", Format(orst.Fields("LastPmtAmt"), "#,###.00"))
        txtLastPayDate = Format(orst.Fields("LastPmtDate"))
        txtCustPmtTerms = Format(orst.Fields("PmtTermsID"))
        
        Set orst = Nothing
        Set ocmd = Nothing
    Else
        ClearAcctDetail
    End If
End Sub


Private Sub ClearAcctDetail()
    txtTotalBalance(1) = ""
    txtCurBalance = ""
    txtOver30 = ""
    txtOver45 = ""
    txtOver60 = ""
    txtOver90 = ""
    txtHighBalance = ""
    txtLastPayAmt = ""
    txtLastPayDate = ""
    txtCustPmtTerms = ""
End Sub


Private Sub cmdViewOrder_Click(Index As Integer)
    Dim oOrder As Order
    Dim frm As FViewOrder
    
    Select Case Index
        Case 0:     'Orders On Hold
            If gdxOrdersOnHold.RowIndex(gdxOrdersOnHold.Row) <> 0 Then
                
                Set oOrder = New Order
                oOrder.Load m_gwOrdersOnHold.value("OPKey")
                
                Set frm = New FViewOrder
                MDIMain.AddNewWindow frm
                frm.ShowOrder oOrder, False
            End If
            
    End Select
End Sub


Private Sub cmdReturnToCSR_Click(Index As Integer)
    Dim oOrder As Order
    Dim cmd As ADODB.Command
    Dim lTmpOPKey As Long
    Dim lTmpSOKey As Long
    Dim lTmpSOID As Long
    
    Select Case Index
                                                                
    Case 0:     'Orders On Hold
        If gdxOrdersOnHold.RowIndex(gdxOrdersOnHold.Row) <> 0 Then
            'restore order to Ready
            SetWaitCursor True
            lTmpOPKey = CLng(m_gwOrdersOnHold.value("OPKey"))
            
            Set oOrder = New Order
            oOrder.Load lTmpOPKey
            
            oOrder.Save _
                i_bForcePending:=False, _
                i_bSage:=False, _
                i_bCommitOrder:=False

            RefreshHoldList
            
            SendNotification "OP " & oOrder.OPKey & " for " & oOrder.Customer.ID _
                & " Returned by AR.", "As requested, OP " & oOrder.OPKey _
                & " for " & oOrder.Customer.ID & ": " & oOrder.Customer.Name _
                & ", has been returned by " & GetUserName & " to allow you to make changes." _
                & vbCrLf & vbCrLf & "Current database is " & g_DB.database, _
                Array(oOrder.UserID & "@caseparts.com")

            SetWaitCursor False
        End If

    End Select
    
    Set oOrder = Nothing
End Sub



'**********************************************************************************
' tabMain(3)    Manage Customers
'**********************************************************************************

Private Sub InitMgCustTab()
        
    'clear controls
    txtCustID.text = ""
    txtCustName.text = ""
    txtCSZ.text = ""
    txtAgingDate(1).text = ""
    txtHoldStatus.text = ""

    opVIP.value = False
    opManual.value = False
    opAuto.value = False
    
    txtARStatus.text = ""
    cboPmtTerms.ListIndex = -1
    txtCreditLimit.text = "$0.00"
    
    'initially disable controls on the Customer Detail frame
    frmARStatus.Enabled = False
    cmdSave.Enabled = False
    cboCustType.Enabled = False
    cboCollector(2).Enabled = False
    txtCreditLimit.Enabled = False
    cboPmtTerms.Enabled = False

    cmdEditCC.Enabled = False
    cmdEditContacts.Enabled = False

    txtExemptCert.text = ""
    m_sExemptCert = ""
     
    'PRN #531 - Context & OwnerID set the MM remarks
    'When this tab is loaded the ContextID is set to "ARCustLoad"
        'and the OwnerID is empty.  Clearing the OwnerID sets this
        'control back to it's inial state & the button is enabled, but
        'doesn't contain & will not show any remarks.
    rvCustomer.OwnerID = ""
    
    ' Moved from tabClick event.
    Set m_oARStatus = New ARStatus.CalcARStatus
    
    '*** 6/29/15 LR based on selectedserver in <username>.xml
    'm_oARStatus.UseProduction = Registry.GetRegNumberValue(HKEY_CURRENT_USER, "Software\Case Parts Company", "UseProductionDB", 1)
    'm_oARStatus.UseProduction = Not g_DB.IsDevelopment
End Sub


'initialize detail display
'Called by
'   cmdFindCust_Click

Private Sub DisplayStatus()
    Dim orst As ADODB.Recordset
    Dim oAddr As Address
    Dim ocmd As ADODB.Command
    Dim HoldStatusID As String

    m_bLoading = True
    
    'enable controls
    cboCustType.Enabled = True
    cboCollector(2).Enabled = True
    txtCreditLimit.Enabled = True
    cboPmtTerms.Enabled = True
    frmARStatus.Enabled = True
    cmdEditCC.Enabled = True
    cmdEditContacts.Enabled = True
    
'not dirty
    cmdSave.Enabled = False

    'get what I can from the Customer object
    With m_oCustomer
        .Load m_lCustKey
        rvCustomer.OwnerID = .ID
        txtCustID = .ID
        txtCustName = .Name
        Set oAddr = .BillAddr
        txtCSZ = Trim(oAddr.City) & ", " & Trim(oAddr.State) & " " & (oAddr.Zip)
        txtHoldStatus = IIf(.Hold, "Yes", "No")
        SetComboByText cboCustType, .CustType
    End With
   
    'get the remaining info
    Set ocmd = CreateCommandSP("spOParGetCustStatus")
    With ocmd
        .Parameters("@CustKey").value = m_lCustKey
    End With
    Set orst = New ADODB.Recordset
    orst.Open ocmd, , adOpenStatic, adLockReadOnly
    
    txtCreditLimit.Amount = orst.Fields("CreditLimit").value
    txtAgingDate(1).text = FormatDateTime(orst.Fields("AgingDate").value, vbShortDate)

    m_oPmtTerms.LoadComboBox cboPmtTerms
    SetComboByText cboPmtTerms, orst.Fields("PmtTermsID").value
    
    SetComboByText cboCollector(2), Format(orst.Fields("Collector").value)

    HoldStatusID = IIf(IsNull(orst.Fields("HoldStatusID")), "Good", Trim(orst.Fields("HoldStatusID")))
    
    Select Case HoldStatusID
        Case "VIP":
            opVIP.value = True
        Case "Manual_Hold":
            opManual.value = True
        Case "Good":
            opAuto.value = True
        Case "Auto_Hold":
            opAuto.value = True
    End Select
    
    txtARStatus.text = ""
    
    If m_oARStatus.OverLimit(orst) Then
        txtARStatus.text = "Over"
    End If
    
    If m_oARStatus.OverDue(orst) Then
        txtARStatus.text = txtARStatus.text & "Late"
    End If
    
    If Len(txtARStatus.text) = 0 Then
        txtARStatus.text = "Good"
    End If

    'moved from FBilling
    GetExemptCert (m_lCustKey)
    txtExemptCert.text = m_sExemptCert

    BackupManageCustomer HoldStatusID
    m_bLoading = False

End Sub


Private Sub BackupManageCustomer(ByVal sHoldStatusID As String)
    m_sBackupCollector = cboCollector(2).List(cboCollector(2).ListIndex)
    m_sBackupPmtTerms = cboPmtTerms.List(cboPmtTerms.ListIndex)
    m_lBackupCreditLimit = txtCreditLimit.Amount
    m_sBackupARStatus = sHoldStatusID
End Sub


Private Sub cmdEditCC_Click()
    Dim oFrm As FCreditCardEditor
    Set oFrm = New FCreditCardEditor
    Call oFrm.Init(m_oCustomer, Nothing, bShow:=True)
    'ignore any selection
    Unload oFrm
    Set oFrm = Nothing
End Sub


Private Sub cmdEditContacts_Click()
    If Not m_oCustomer.Contacts Is Nothing Then
        m_oCustomer.Contacts.Edit GetUserName
    End If
End Sub


Private Sub cboCustType_Click()
    If Not m_bLoading Then cmdSave.Enabled = True
End Sub


Private Sub cboPmtTerms_Click()
    If Not m_bLoading Then cmdSave.Enabled = True
End Sub


'NOTE: need to validate CreditLimit input

Private Sub opAuto_Click()
    If Not m_bLoading Then
        m_bStatusTypeChg = True
        cmdSave.Enabled = True
    End If
End Sub


Private Sub opManual_Click()
    If Not m_bLoading Then
        m_bStatusTypeChg = True
        cmdSave.Enabled = True
    End If
End Sub


Private Sub opVIP_Click()
    If Not m_bLoading Then
        m_bStatusTypeChg = True
        cmdSave.Enabled = True
    End If
End Sub



Private Sub cmdSave_Click()
    Dim ocmd As ADODB.Command

    SetWaitCursor True

    Set ocmd = CreateCommandSP("spOPUpdateARStatus2")
    With ocmd
        .Parameters("@CustKey") = m_lCustKey
        .Parameters("@Collector") = cboCollector(2).text
        .Parameters("@CreditLimit") = txtCreditLimit.Amount
        .Parameters("@CustType") = cboCustType.text
        .Parameters("@PmtTermsKey") = cboPmtTerms.ItemData(cboPmtTerms.ListIndex)
        .Execute
    End With
    Set ocmd = Nothing

'if the radio button selection has changed, CRUD the DB
'TODO: check the radio button click code to make sure this flag is
'set only if the selected button changes
'DH - Requires previous status to be loaded and cached.

    If m_bStatusTypeChg Then
        If opVIP.value Then
            m_oARStatus.Update m_lCustKey, klARS_VIP
        ElseIf opManual.value Then
            m_oARStatus.Update m_lCustKey, klARS_MANUAL_HOLD
        ElseIf opAuto.value Then
            m_oARStatus.Update m_lCustKey, klARS_AUTO_HOLD
        End If
    End If

    logManageCustomerEvent      'OA event log stuff
    
    m_bStatusTypeChg = False
    
'Corresponding controls are disabled
'    UpdateExemptCerts m_lCustKey, Trim(txtExemptCert.Text), m_sStateID

    SetWaitCursor False
    cmdSave.Enabled = False
End Sub


'**********************************************************************************
' tabMain(5)    Credit Manager
'**********************************************************************************

'Stuff for Teddy/Jon to Complete
'   0. Change 'Increase' updown from relative to absolute.
'   Impliment logic such that new level is greater than current.
'   1. Reformat grid column widths to appropriate dimensions
'   2. Add a checkbox to the first column of the grid
'   3. Add Memomeister functionality
'   4. Make sure the only columns editable in grid are Update and NewLimit
'   5. Batch Update
'   6. Use Memomeister to persist history data
'   DONE 7. Jon - rewrite views to hit memomeister not topCPCCreditLimitHist (deprecated)
'   DONE 8. Jon - rewrite spCPCUpdateCreditLimit to update memomeister


Private Sub cmdUpdate_Click()
    With gdxCredLimit
        .Refetch
        .Update
    End With
            
    If UpdateLimit Then
        cmdRefresh_Click (1)
    End If
End Sub


Private Function GetUpdateRecord() As CreditLimitRecord
    Dim m_CLRUpdateRecord As CreditLimitRecord
    Set m_CLRUpdateRecord = New CreditLimitRecord
    
    With gdxCredLimit
        m_CLRUpdateRecord.Update = .value(1)
        m_CLRUpdateRecord.NewLimit = .value(2)
        m_CLRUpdateRecord.CurLimit = .value(3)
        m_CLRUpdateRecord.YearRec = .value(4)
        m_CLRUpdateRecord.ColRems = .value(5)
        m_CLRUpdateRecord.CustID = .value(6)
        m_CLRUpdateRecord.CustName = .value(7)
        m_CLRUpdateRecord.CurAmount = .value(8)
        m_CLRUpdateRecord.l30DaysAmount = .value(9)
        m_CLRUpdateRecord.l45DaysAmount = .value(10)
        m_CLRUpdateRecord.Terms = .value(11)
        m_CLRUpdateRecord.LastChanged = .value(12)
        m_CLRUpdateRecord.AgingDate = .value(13)
        m_CLRUpdateRecord.TerritoryKey = .value(14)
        m_CLRUpdateRecord.CustKey = .value(15)
    End With
    
    Set GetUpdateRecord = m_CLRUpdateRecord
End Function


Private Function UpdateLimit() As Boolean
    Dim clrRecord As CreditLimitRecord
    
    If Not bUpdateChecked() Then
        msg "Sorry, please choose customers you want to update first!", vbOKOnly + vbExclamation, "Customer Credit Update"
        UpdateLimit = False
        Exit Function
    Else
        If vbYes = msg("Are you sure you want to change credit limit for these customer(s)", vbYesNo, "Change Credit Limit") Then
            SetWaitCursor True
            For Each clrRecord In m_CLUpdateList
                    CallSP "spCPCUpdateCreditLimit", _
                    "@_iCustKey", clrRecord.CustKey, _
                    "@_iOldLimit", clrRecord.CurLimit, _
                    "@_iNewLimit", clrRecord.NewLimit, _
                    "@_iUserID", GetUserID
                    LogOAEvent "Credit Manager", GetUserID, clrRecord.CustKey, , , "Change credit limit from $" & clrRecord.CurLimit & " to $" & clrRecord.NewLimit & " for " & clrRecord.CustID
            Next clrRecord
            SetWaitCursor False
            UpdateLimit = True
       Else
            UpdateLimit = False
        End If
    End If
End Function


Private Function bUpdateChecked() As Boolean
    Set m_CLUpdateList = New CreditLimitList

    Dim obj As CreditLimitRecord
    For Each obj In m_CLList
        If obj.Update Then
            m_CLUpdateList.Add obj
        End If
    Next

    If m_CLUpdateList.Count > 0 Then
        bUpdateChecked = True
    Else
        bUpdateChecked = False
    End If
End Function

Private Sub gdxCredLimit_Change()
    Dim temp As GridEX20.JSRowData
    Dim lIndex As Long
    
    With gdxCredLimit
        lIndex = .RowIndex(.Row)
    End With
End Sub


Private Sub gdxCredLimit_Click()
    UpdateGridRemark
End Sub


Private Sub gdxCredLimit_KeyUP(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        UpdateGridRemark
    End If
End Sub


Private Sub gdxCredLimit_LostFocus()
    gdxCredLimit.Update
End Sub

Private Sub gdxCredLimit_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxCredLimit.Update
End Sub


Private Sub gdxCredLimit_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_CLList Is Nothing Then Exit Sub
    If RowIndex > m_CLList.Count Then Exit Sub
    
    With m_CLList(RowIndex)
        Values(1) = .Update
        Values(2) = .NewLimit
        Values(3) = .CurLimit
        Values(4) = .YearRec
        Values(5) = .ColRems
        Values(6) = .CustID
        Values(7) = .CustName
        Values(8) = .CurAmount
        Values(9) = .l30DaysAmount
        Values(10) = .l45DaysAmount
        Values(11) = .Terms
        Values(12) = .LastChanged
        Values(13) = .AgingDate
        Values(14) = .TerritoryKey
        Values(15) = .CustKey
    End With
End Sub


Private Sub gdxCredLimit_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    With m_CLList(RowIndex)
        .Update = Values(1)
        .NewLimit = Values(2)
    End With
End Sub


Private Sub UpdateGridRemark()
    Dim sContextID As String
    Dim lEntityType As Long
    Dim rst As ADODB.Recordset
    
    sContextID = "ARCreditHist"
    lEntityType = 501
    
    If gdxCredLimit.value(15) = Empty Then
        Exit Sub
    Else
        SetWaitCursor True
        Set rst = CallSP("spCPCmmGetFilteredMemos", _
                         "@ContextID", sContextID, _
                         "@EntityType", lEntityType, _
                         "@OwnerKey", gdxCredLimit.value(15), _
                         "@Priority", 0)
    End If
    
    With gdxRemarks
        .HoldFields
        Set .ADORecordset = rst
    End With
    SetWaitCursor False
End Sub


Private Sub udCredit_UpClick()
    If udNewCredit.value < udCredit.value Then
        udNewCredit.value = udCredit.value
    End If
End Sub


Private Sub udNewCredit_DownClick()
    If udNewCredit.value < udCredit.value Then
        msg "Sorry. The new credit level should be greater than current level. ", vbOKOnly + vbExclamation, "New Credit Level"
        udNewCredit.value = udCredit.value
        Exit Sub
    End If
End Sub


'******************************************************************************
' UTILITY FUNCTIONS
'******************************************************************************

Private Sub LoadValidationRules()
    Dim oCtlWrapper As ControlWrapper

    With m_oBrokenRules
        'Customer search fields
        Set oCtlWrapper = .AddControl(txtCustSearch, k_sCustNameOrID, True, False)
        oCtlWrapper.AddRuleRequired "", ccCustomer, True, "Enter a value to search for a customer."
        .EnableClass ccCustomer, True

    End With
    
End Sub

' Load a recordset with an SP
' Should replace with a call to CallSP() - but there's an ADO problem here
Private Sub LoadList(i_sSPName As String, ByRef i_orst As ADODB.Recordset)
    Dim ocmd As ADODB.Command

    Set ocmd = CreateCommandSP(i_sSPName)
    Set i_orst = New ADODB.Recordset
    i_orst.Open ocmd, , adOpenStatic, adLockReadOnly
    Set ocmd = Nothing
End Sub


'********************************************************************************
' Tab(6)    Turbo Ten Key
'********************************************************************************

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Key preview is turned on.
    'This sub passes all keystrokes to the correct
    'element in the command button array for processing.
    
    If tabMain.SelectedTab.Index <> tmiTurboTenKey Then Exit Sub 'Ignore if not Turbo 10 Key
    If Me.ActiveControl.Name = "txtBatch" Then Exit Sub

    
    'Handle number keys
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        cmdKey_Click (KeyAscii - 48)
    End If
    
    'Handle other keys
    Select Case KeyAscii
        Case Is = 8     'Backspace
            cmdKey_Click (10)
        Case Is = 27    'Clear
            cmdKey_Click (12)
        Case Is = 43    'Plus
            cmdKey_Click (13)
        Case Is = 45    'Minus
            cmdKey_Click (14)
        Case Is = 46    'Decimal
            cmdKey_Click (11)
        Case Is = 13    'Enter
            If ckbEnter.value = vbChecked Then
                cmdKey_Click (13)   'Enter = Add
            End If
    End Select

    'kill the keystroke
    KeyAscii = 0
End Sub


Private Function FormatBatch(sInput As String, sMask As String) As String
    Dim sTemp As String
    
    sTemp = sMask

    sTemp = Left$(sTemp, Len(sTemp) - Len(sInput)) & sInput

    FormatBatch = sTemp
End Function


Private Sub txtBatch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtBatch.text)) > 0 And KeyCode = vbKeyReturn Then
        cmdLoad_Click
    End If
End Sub


Private Sub cmdLoad_Click()
    On Error GoTo ErrorHandler
    
    Dim sSQL As String
    Dim sBatchID As String
    Dim lstItm As ListItem
    Dim subItm As ListSubItem
    Dim dAmt As Double
    
    m_sBatchCmnt = ""
    lvwBatch.ListItems.Clear
    
    If Len(Trim(txtBatch.text)) = 0 Then
        TryToSetFocus txtBatch
        Exit Sub
    End If
    
    SetWaitCursor True
    
    Dim rst As ADODB.Recordset
    Set rst = CallSP("spcpcTurboTenARLoad", "@BatchID", txtBatch.text)
    
    With rst
        Do While Not .EOF
            m_sBatchCmnt = .Fields("BatchCmnt").value
            Set lstItm = lvwBatch.ListItems.Add(lvwBatch.ListItems.Count + 1, , Trim(.Fields("TranNo").value))
            Set subItm = lstItm.ListSubItems.Add(, , Format$(.Fields("TranAmt").value, g_DisplayMask))
            If .Fields("TranAmt").value < 0 Then
                subItm.ForeColor = vbRed
            End If
        
            .MoveNext
        Loop
    End With
    
    rst.Close
    Set rst = Nothing
    
    SetWaitCursor False
 
    If lvwBatch.ListItems.Count > 0 Then
        AdjustBatchTapeColumns
        CalcBatchTotal
    Else
        MsgBox "No items on this batch"
        txtBatch.SelStart = 0
        txtBatch.SelLength = Len(txtBatch.text)
        TryToSetFocus txtBatch
    End If
    Exit Sub

ErrorHandler:
    ClearWaitCursor
    msg Err.Description, vbOKOnly + vbCritical, Err.Source
End Sub


Private Function PrintTape() As Boolean
    Dim i As Integer
    Dim lstItm As ListItem
    Dim subItm As ListSubItem
    
    Printer.Print
    Printer.Print
    
    Printer.Print txtBatch.text
    Printer.Print m_sBatchCmnt
    Printer.Print
    
'   This code prints the tape
    For i = 1 To lvwTape.ListItems.Count
        Set lstItm = lvwTape.ListItems(i)
        Set subItm = lstItm.ListSubItems(1)
        Printer.CurrentX = 1200 - TextWidth(Format$(Val(Replace(subItm.text, ",", "")), g_DisplayMask))
        Printer.Print Format$(Val(Replace(subItm.text, ",", "")), g_DisplayMask) & vbTab & vbTab & lstItm.text
        Printer.CurrentX = 0
        If (i / 50) = Int(i / 50) Then
            Printer.NewPage
        End If
    Next

    Printer.Print
    Printer.Print "-----------------------"
    Printer.CurrentX = 1200 - TextWidth(Format$(Val(Replace(lblTapeTotal.Caption, ",", "")), g_DisplayMask))

    Printer.Print Format$(Val(Replace(lblTapeTotal.Caption, ",", "")), g_DisplayMask)
    PrintTape = True

End Function


Private Sub cmdReconcile_Click()
    
    Reconcile
    
    'Focus must be returned to txtdisplay so ENTER key functions properly
    TryToSetFocus txtDisplay
End Sub


Public Sub Reconcile()
    Dim iTape As Integer
    Dim iBatch As Integer
    
    'Clear all asterisks
    For iTape = 1 To lvwTape.ListItems.Count
        lvwTape.ListItems(iTape).SubItems(2) = ""
    Next
    
    For iBatch = 1 To lvwBatch.ListItems.Count
        lvwBatch.ListItems(iBatch).SubItems(2) = ""
    Next
    
    'Reconcile
    For iTape = 1 To lvwTape.ListItems.Count
        For iBatch = 1 To lvwBatch.ListItems.Count
            If lvwTape.ListItems(iTape).SubItems(1) = lvwBatch.ListItems(iBatch).SubItems(1) Then
                If lvwBatch.ListItems(iBatch).SubItems(2) <> "*" Then
                    lvwBatch.ListItems(iBatch).SubItems(2) = "*"
                    lvwTape.ListItems(iTape).SubItems(2) = "*"
                    Exit For
                End If
            End If
        Next
    Next

    'adjust the asterisk columns
    If lvwTape.ColumnHeaders(3).width = 0 Then
        'Make the 'Amount' Column narrower
        lvwTape.ColumnHeaders(2).width = lvwTape.ColumnHeaders(2).width - klReconcileColWidth
        lvwTape.ColumnHeaders(3).width = klReconcileColWidth
    End If
    
    If lvwBatch.ColumnHeaders(3).width = 0 Then
        'Make the 'Description' Column narrower
        lvwBatch.ColumnHeaders(1).width = lvwBatch.ColumnHeaders(1).width - klReconcileColWidth
        lvwBatch.ColumnHeaders(3).width = klReconcileColWidth
    End If
    
End Sub


Private Function FormatEntry(sEntry As String) As Double
    'If the entry does not contain a decimal and auto decimal is
    'checked, recalc
    If InStr(1, sEntry, ".") = 0 And ckbDecimal.value = vbChecked Then
        FormatEntry = CDbl(sEntry) / 100
    Else
        If IsNumeric(sEntry) Then
            FormatEntry = CDbl(sEntry)
        Else
            FormatEntry = 0
        End If
    End If

End Function


Private Sub lvwTape_DblClick()
    Dim lstItm As ListItem
    Dim subItm As ListSubItem
    
    Set lstItm = lvwTape.HitTest(m_sngX, m_sngY)
    If Not lstItm Is Nothing Then
        'if multiselected, clear all selections to reduce user ambiguity
        ClearSelections
        lstItm.Selected = True
        m_dCurEntry = Abs(FormatEntry(lstItm.SubItems(1)))
        txtDisplay = Format$(m_dCurEntry, g_DisplayMask)
        
        'Make the double clicked item stand out
        Set m_EditItem = lstItm
        Set subItm = m_EditItem.ListSubItems(1)
        subItm.Bold = True
        subItm.ForeColor = vbBlue
        TryToSetFocus txtDisplay
        txtDisplay.SelStart = Len(txtDisplay)
    Else
        Set m_EditItem = Nothing
        m_dCurEntry = 0
        txtDisplay = ""
        Beep
    End If
End Sub


Private Sub ClearSelections()
    Dim lstItm As ListItem
    For Each lstItm In lvwTape.ListItems
        lstItm.Selected = False
        lstItm.ListSubItems(1).Bold = False
    Next
End Sub

Private Sub lvwTape_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_sngX = X
    m_sngY = Y
End Sub

Private Sub cmdPrintTape_Click()
    'Focus must be returned to txtdisplay so ENTER key functions properly
    TryToSetFocus txtDisplay
    
    If PrintTape Then
        Printer.EndDoc
    Else
        Printer.KillDoc
    End If
    
End Sub

Private Sub cmdClear_Click()
    lvwTape.ListItems.Clear
End Sub

Private Sub AdjustBatchTapeColumns()
    'This adjusts the batch column widths based on the presence of
    'a vertical scroll bar
    With lvwBatch
        If .ListItems(1).Height * (.ListItems.Count + 1) > .Height Then
            .ColumnHeaders(1).width = .width - lvwBatch.ColumnHeaders(2).width - klVScrollBarWidth
        End If
    End With

End Sub

Private Sub CalcBatchTotal()
    Dim lstItm As ListItem
    Dim dTotal As Double
    
    dTotal = 0
    For Each lstItm In lvwBatch.ListItems
        dTotal = dTotal + CDbl(lstItm.SubItems(1))
    Next
    
    If dTotal < 0 Then
        lblBatchTotal.ForeColor = vbRed
    Else
        lblBatchTotal.ForeColor = vbBlack
    End If
    
    lblBatchTotal = Format$(dTotal, g_DisplayMask)
End Sub

Private Sub CalcTapeTotal()
    'Reads through the tape and displays/formats the total
    
    Dim lstItm As ListItem
    Dim dTotal As Double
    
    dTotal = 0
    For Each lstItm In lvwTape.ListItems
        dTotal = dTotal + CDbl(lstItm.SubItems(1))
    Next
    
    If dTotal < 0 Then
        lblTapeTotal.ForeColor = vbRed
    Else
        lblTapeTotal.ForeColor = vbBlack
    End If
    
    lblTapeTotal = Format$(dTotal, g_DisplayMask)
End Sub

Private Sub SetUpTape()
    lvwTape.LabelEdit = lvwManual
    lvwTape.View = lvwReport
    lvwTape.MultiSelect = True
    lvwTape.ColumnHeaders.Add , , , 350
    lvwTape.ColumnHeaders.Add , , "Amount", lvwTape.width - 400, lvwColumnRight
    lvwTape.ColumnHeaders.Add , , , 0, lvwColumnCenter
        
End Sub

Private Sub SetUpBatch()
    lvwBatch.LabelEdit = lvwManual
    lvwBatch.View = lvwReport
    lvwBatch.ColumnHeaders.Add , , "Description", 2600
    lvwBatch.ColumnHeaders.Add , , "Amount", lvwBatch.width - 2650, lvwColumnRight
    lvwBatch.ColumnHeaders.Add , , , 0, lvwColumnCenter
End Sub

Private Sub cmdKey_Click(Index As Integer)
    
    'Handle Number Buttons
    If Index >= 0 And Index <= 9 Then
        
        If m_bNew Then
            txtDisplay = ""
            m_bNew = False
        End If
        
        'only allow entry of 2 decimal places
        If InStr(1, txtDisplay.text, ".") > 0 Then
            If InStr(1, txtDisplay.text, ".") = (Len(Trim(txtDisplay.text)) - 2) Then
                Beep
                Exit Sub
            End If
        End If
        
        
        txtDisplay = txtDisplay & CStr(Index)
        txtDisplay.SelStart = Len(txtDisplay)
        txtDisplay.SelLength = 0
        Exit Sub
    End If
    
    'Handle Other Buttons
    Select Case Index
                                                                    
        Case Is = 10    'Backspace
            m_bNew = False
            If Len(txtDisplay) > 0 Then
                txtDisplay = Left(txtDisplay, Len(txtDisplay) - 1)
            Else
                Beep
            End If
            
         Case Is = 11    'Decimal
        
            If m_bNew Then
                txtDisplay = ""
                m_bNew = False
            End If
        
            If InStr(1, txtDisplay, ".") = 0 Then
                txtDisplay = txtDisplay & "."
            Else
                Beep
            End If
            
        Case Is = 12    'Clear
        
            txtDisplay = ""
            m_bNew = True
            m_dBalance = 0
            
        Case Is < 15    'Add/Subtract
            
            If Len(txtDisplay) = 0 Then
                m_dCurEntry = 0
            Else
            
                If Index = 13 Then  'Add
                    If Not m_bNew Then
                        m_dCurEntry = FormatEntry(txtDisplay)
                        txtDisplay = Format$(m_dCurEntry, g_DisplayMask)
                    Else
                        If m_dCurEntry < 0 Then m_dCurEntry = -m_dCurEntry
                    End If
                Else                'Subtract
                    If ckbSubtraction.value = vbChecked Then
                        If Not m_bNew Then
                            m_dCurEntry = -FormatEntry(txtDisplay)
                            txtDisplay = Format$(m_dCurEntry, g_DisplayMask)
                        Else
                            If m_dCurEntry > 0 Then m_dCurEntry = -m_dCurEntry
                        End If
                    End If
                End If
                
            End If
            
            'Disallow subtraction if appropriate
            If Index = 14 And ckbSubtraction.value <> vbChecked Then
                Beep
                Exit Sub
            Else
                If m_EditItem Is Nothing Then
                    AddTapeEntry m_dCurEntry
                Else
                    EditTapeEntry m_dCurEntry
                End If
                
                CalcTapeTotal
            End If

            m_bNew = True
            
    End Select
    
    txtDisplay.SelStart = Len(txtDisplay)
    txtDisplay.SelLength = 0
End Sub

Private Sub AddTapeEntry(dAmount As Double)
    Dim lstItm As ListItem
    Dim subItm As ListSubItem
        
    If dAmount < 0 Then
        Set lstItm = lvwTape.ListItems.Add(lvwTape.ListItems.Count + 1, , "-")
        lstItm.ForeColor = vbRed
    Else
        Set lstItm = lvwTape.ListItems.Add(lvwTape.ListItems.Count + 1, , "+")
    End If
    
    lstItm.Bold = True
    lstItm.EnsureVisible
    lstItm.Selected = False
    
    Set subItm = lstItm.ListSubItems.Add(, , Format$(dAmount, g_DisplayMask))
    subItm.ForeColor = lstItm.ForeColor
    lstItm.SubItems(2) = ""
    AdjustColumns
End Sub

Private Sub EditTapeEntry(dAmount As Double)
    Dim subItm As ListSubItem
        
    'Edit m_EditItem
    
    If m_EditItem Is Nothing Then
        Exit Sub 'something dreadful has happened
    End If
    
    If dAmount < 0 Then
        m_EditItem.text = "-"
        m_EditItem.ForeColor = vbRed
        Set subItm = m_EditItem.ListSubItems(1)
        subItm = Format$(dAmount, g_DisplayMask)
        subItm.ForeColor = vbRed
        subItm.Bold = False
    Else
        m_EditItem.text = "+"
        m_EditItem.ForeColor = vbBlack
        Set subItm = m_EditItem.ListSubItems(1)
        subItm = Format$(dAmount, g_DisplayMask)
        subItm.ForeColor = vbBlack
        subItm.Bold = False
    End If
    AdjustColumns
    ClearSelections
    Set m_EditItem = Nothing
    
End Sub

Private Sub AdjustColumns()
    'This mickey-mau routine adjusts the Tape Amount column width
    'depending on the presence of the vertical scroll bar
    If VScrollBar Then
        'If the reconcile column is visible...
        If lvwTape.ColumnHeaders(3).width = 0 Then
            lvwTape.ColumnHeaders(2).width = lvwTape.width - lvwTape.ColumnHeaders(1).width - klVScrollBarWidth
        Else
            lvwTape.ColumnHeaders(2).width = lvwTape.width - lvwTape.ColumnHeaders(1).width - klVScrollBarWidth - lvwTape.ColumnHeaders(3).width
        End If
    Else
        'If the reconcile column is visible...
        If lvwTape.ColumnHeaders(3).width = 0 Then
            lvwTape.ColumnHeaders(2).width = lvwTape.width - klVScrollBarWidth
        Else
            lvwTape.ColumnHeaders(2).width = lvwTape.width - klVScrollBarWidth - lvwTape.ColumnHeaders(3).width
        End If
    End If

End Sub

Private Function VScrollBar() As Boolean
    'Is there a veritcal scroll bar in the Tape?
    With lvwTape
        If .ListItems.Count > 0 Then
            If .ListItems(1).Height * (.ListItems.Count + 1) > .Height Then
                VScrollBar = True
            End If
        End If
    End With
End Function

Private Sub DeleteSelItems()
    Dim lstItm As ListItem
    Dim i As Integer
    
    'If at least something is selected...
    If Not (lvwTape.SelectedItem Is Nothing) Then
        i = 1
        Do Until i = lvwTape.ListItems.Count
            
            'If the item is selected, delete it
            If lvwTape.ListItems(i).Selected = True Then
                
                'If the EditItem is selected for deletion,
                'set the global to nothing
                If lvwTape.ListItems(i) Is m_EditItem Then
                    Set m_EditItem = Nothing
                End If
                
                'remove the item from the control
                lvwTape.ListItems.Remove lvwTape.ListItems(i).Index
                
            Else
                i = i + 1
            End If
        Loop
        
        'After you've fallen out of the loop, test the last item
        If lvwTape.ListItems(i).Selected = True Then
            'If the EditItem is selected for deletion,
            'set the global to nothing
            If lvwTape.ListItems(i) Is m_EditItem Then
                Set m_EditItem = Nothing
            End If
            lvwTape.ListItems.Remove lvwTape.ListItems(i).Index
        End If
        
        AdjustColumns
    End If
    
    CalcTapeTotal
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'If the user presses DELETE while the tape has focus,
    'delete any selected entries
    
    If Me.ActiveControl Is lvwTape Then
        If KeyCode = vbKeyDelete Then
            DeleteSelItems
        End If
    End If
End Sub

'******************************************************************************************
' Tab(7)    Research
'******************************************************************************************

Private Sub cmdGo_Click()
    
    Dim sSQL As String

    SetWaitCursor True
    
    If m_oRstResearch Is Nothing Then
        Set m_oRstResearch = New ADODB.Recordset
    Else
        'm_oRstResearch.Close
        CloseRst m_oRstResearch
        Set m_oRstResearch = New ADODB.Recordset
    End If
    
    If optCheck.value = True Then
    
        sSQL = "SELECT tarCustPmt.TranNo, tarCustPmt.TranAmt, " & _
            "tarCustPmt.TranDate, tarCustomer.CustID, tarCustomer.CustName " & _
            "FROM tarCustPmt INNER JOIN tarCustomer ON " & _
            "tarcustpmt.CustKey = tarCustomer.CustKey " & _
            "WHERE tarCustPmt.CompanyID = 'CPC' AND tarcustpmt.tranno LIKE '" & PrepSQLText(txtInput.text) & "%'"

    Else
    
'        If optPS = True Then
'            sSQL = "SELECT tarInvoice.TranCmnt, tarInvoice.TranDate, " & _
'                "tarCustomer.CustID, tarCustomer.CustName, " & _
'                "tarInvoice.CustPONo, tarInvoice.TranAmt, " & _
'                "tarInvoice.TranID FROM tarCustomer INNER JOIN " & _
'                "tarInvoice ON tarCustomer.CustKey = tarInvoice.CustKey " & _
'                "WHERE (tarInvoice.CompanyID = 'CPC') AND " & _
'                "tarInvoice.TranCmnt LIKE '%" & PrepSQLText(txtInput.text) & "%' Order By tarInvoice.TranCmnt"
'        Else
        
            If optInvoice.value = True Then
                sSQL = "SELECT tarInvoice.TranCmnt, tarInvoice.TranDate, " & _
                    "tarCustomer.CustID, tarCustomer.CustName, " & _
                    "tarInvoice.CustPONo, tarInvoice.TranAmt, " & _
                    "tarInvoice.TranID FROM tarCustomer INNER JOIN " & _
                    "tarInvoice ON tarCustomer.CustKey = tarInvoice.CustKey " & _
                    "WHERE (tarInvoice.CompanyID = 'CPC') AND " & _
                    "tarInvoice.TranNo LIKE '%" & PrepSQLText(txtInput.text) & "%' Order By tarInvoice.TranNo"
            Else
                If optPO Then
                    sSQL = "SELECT tarInvoice.TranCmnt, tarInvoice.TranDate, " & _
                        "tarCustomer.CustID, tarCustomer.CustName, " & _
                        "tarInvoice.CustPONo, tarInvoice.TranAmt, " & _
                        "tarInvoice.TranID FROM tarCustomer INNER JOIN " & _
                        "tarInvoice ON tarCustomer.CustKey = tarInvoice.CustKey " & _
                        "WHERE (tarInvoice.CompanyID = 'CPC') AND " & _
                        "tarInvoice.CustPONo LIKE '%" & PrepSQLText(txtInput.text) & "%' Order By tarInvoice.CustPONo"
                Else
                    If optAmt Then
                        sSQL = "SELECT tarInvoice.TranCmnt, tarInvoice.TranDate, " & _
                            "tarCustomer.CustID, tarCustomer.CustName, " & _
                            "tarInvoice.CustPONo, tarinvoice.TranAmt, " & _
                            "tarInvoice.TranID FROM tarCustomer INNER JOIN " & _
                            "tarInvoice ON tarCustomer.CustKey = tarInvoice.CustKey " & _
                            "WHERE (tarInvoice.CompanyID = 'CPC') AND " & _
                            "CONVERT(varchar(15), tarinvoice.TranAmt) LIKE '%" & PrepSQLText(txtInput.text) & "%' Order By tarInvoice.TranAmt"
                    Else
                        If optOPID Then
                            sSQL = "SELECT DISTINCT tcpSO.OPKey as OrderNo, tsoSalesOrder.TranNo as SalesOrder, tarInvoice.TranID as InvNo, tarInvoice.TranDate as InvDate, tarInvoice.TranAmt as InvAmt, " & _
                                "tarCustomer.CustID, tarCustomer.CustName, tarInvoice.CustPONo " & _
                                "FROM tarCustomer INNER JOIN tarInvoice ON tarCustomer.CustKey = tarInvoice.CustKey INNER JOIN tsoInvoiceShipment " & _
                                "ON tarInvoice.InvcKey = tsoInvoiceShipment.InvcKey INNER JOIN tsoShipLine ON tsoInvoiceShipment.ShipKey = tsoShipLine.ShipKey " & _
                                "INNER JOIN tsoSOLine ON tsoShipLine.SOLineKey = tsoSOLine.SOLineKey INNER JOIN tsoSalesOrder ON tsoSOLine.SOKey = tsoSalesOrder.SOKey " & _
                                "INNER JOIN tcpSO ON tsoSalesOrder.SOKey = tcpSO.SOKey " & _
                                "WHERE (tarInvoice.CompanyID = 'CPC') AND (CONVERT(varchar, tcpSO.OPKey) LIKE '%" & PrepSQLText(txtInput.text) & "%') ORDER BY tcpSO.OPKey"
                        Else
                            If optAcuity Then
                                sSQL = "SELECT DISTINCT tsoSalesOrder.TranNo as SalesOrder, tarInvoice.TranID as InvNo, tarInvoice.TranDate as InvDate, tarInvoice.TranAmt as InvAmt, " & _
                                    "tarCustomer.CustID, tarCustomer.CustName, tarInvoice.CustPONo " & _
                                    "FROM tarCustomer INNER JOIN tarInvoice ON tarCustomer.CustKey = tarInvoice.CustKey INNER JOIN tsoInvoiceShipment " & _
                                    "ON tarInvoice.InvcKey = tsoInvoiceShipment.InvcKey INNER JOIN tsoShipLine ON tsoInvoiceShipment.ShipKey = tsoShipLine.ShipKey " & _
                                    "INNER JOIN tsoSOLine ON tsoShipLine.SOLineKey = tsoSOLine.SOLineKey INNER JOIN tsoSalesOrder ON tsoSOLine.SOKey = tsoSalesOrder.SOKey " & _
                                    "WHERE (tarInvoice.CompanyID = 'CPC') AND (tsoSalesOrder.TranNo LIKE '%" & PrepSQLText(txtInput.text) & "%') ORDER BY tsoSalesOrder.TranNo"
                            Else 'optShipment
                                sSQL = "SELECT DISTINCT tcpSO.OPKey AS OrderNo, tsoSalesOrder.TranNo AS SalesOrder, tarInvoice.TranID AS InvNo, tarInvoice.TranDate AS InvDate, " & _
                                    "tarInvoice.TranAmt AS InvAmt, tarCustomer.CustID, tarCustomer.CustName, tarInvoice.CustPONo " & _
                                    "FROM tarCustomer INNER JOIN tarInvoice ON tarCustomer.CustKey = tarInvoice.CustKey INNER JOIN tsoInvoiceShipment ON tarInvoice.InvcKey = tsoInvoiceShipment.InvcKey " & _
                                    "INNER JOIN tsoShipLine ON tsoInvoiceShipment.ShipKey = tsoShipLine.ShipKey INNER JOIN tsoSOLine ON tsoShipLine.SOLineKey = tsoSOLine.SOLineKey INNER JOIN " & _
                                    "tsoSalesOrder ON tsoSOLine.SOKey = tsoSalesOrder.SOKey INNER JOIN tcpSO ON tsoSalesOrder.SOKey = tcpSO.SOKey INNER JOIN tsoShipment ON tsoShipLine.ShipKey = tsoShipment.ShipKey " & _
                                    "WHERE (tarInvoice.CompanyID = 'CPC') AND (tsoShipment.TranNo LIKE '%" & PrepSQLText(txtInput.text) & "%') ORDER BY tcpSO.OPKey"
                            End If
                        End If
                    End If
                End If
            End If
        'End If
    End If
    
    m_oRstResearch.MaxRecords = CInt(lblMaxRecords.Caption)
    m_oRstResearch.Open sSQL, g_DB.Connection
    
    SetWaitCursor False
    
    If m_oRstResearch.EOF Then
        msg "No matches were found."
    Else
        'AttachGrid gdxResearch, m_oRstResearch
        Set gdxResearch.ADORecordset = m_oRstResearch
    End If

End Sub


Private Sub cmdOPFindCC_Click()
    Dim sSQL As String
    Dim liCounter As Integer

    If m_oRstResearch Is Nothing Then
        Set m_oRstResearch = New ADODB.Recordset
    Else
        CloseRst m_oRstResearch
        Set m_oRstResearch = New ADODB.Recordset
    End If
    
    If IsNumeric(txtOPFindCC.text) = True Then
        sSQL = "select CrCardTypeName, CrCardNo, CrCardExp, CardHolderName, CrCardStreetNbrZip, CrCardZipCode " & _
        "from tcpso inner join tcpcreditcard on tcpso.cckey = tcpcreditcard.cckey " & _
        "inner join tcpcreditcardtype on tcpcreditcard.CrCardTypeKey = tcpcreditcardtype.CrCardTypeKey " & _
        "Where tcpso.OPKey = " & Trim$(txtOPFindCC.text)

        m_oRstResearch.Open sSQL, g_DB.Connection
        
        If m_oRstResearch.EOF Then
            msg "No matches were found."
        Else
            Set gdxResearch.ADORecordset = m_oRstResearch
            gdxResearch.Columns(1).Caption = "Card Type"
            gdxResearch.Columns(2).Caption = "Card No"
            gdxResearch.Columns(3).Caption = "Expire Date"
            gdxResearch.Columns(4).Caption = "Card Holder"
            gdxResearch.Columns(5).Caption = "Street"
            gdxResearch.Columns(6).Caption = "Zip"
            
            For liCounter = 1 To gdxResearch.Columns.Count
                gdxResearch.Columns(liCounter).AutoSize
            Next
        End If
    Else
        msg "Please enter an OP#", vbInformation
    End If
End Sub


Private Sub cmdOPShowTrans_Click()
    Dim sSQL As String
    Dim liCounter As Integer

    If m_oRstResearch Is Nothing Then
        Set m_oRstResearch = New ADODB.Recordset
    Else
        CloseRst m_oRstResearch
        Set m_oRstResearch = New ADODB.Recordset
    End If
    
    If IsNumeric(txtOPFindCC.text) = True Then
        sSQL = "select CreateDate, UserID, PNREF, Amount, TranType, RespMsg "
        sSQL = sSQL & "from tcpCCTransaction Where OPKey = " & Trim$(txtOPFindCC.text)
        
        m_oRstResearch.Open sSQL, g_DB.Connection
        
        If m_oRstResearch.EOF Then
            msg "No matches were found."
        Else
            Set gdxResearch.ADORecordset = m_oRstResearch
            gdxResearch.Columns(1).Caption = "Create Date"
            gdxResearch.Columns(2).Caption = "User ID"
            gdxResearch.Columns(5).Caption = "Tran Type"
            gdxResearch.Columns(6).Caption = "Response Message"
            
            For liCounter = 1 To gdxResearch.Columns.Count
                gdxResearch.Columns(liCounter).AutoSize
            Next
        End If
    Else
        msg "Please enter an Order Pad Number.", vbInformation
    End If
End Sub



'***************************************************************************
'Event Logging
'***************************************************************************

Private Function sEventCustomerMM(ByVal CustID As String) As String
    If Trim(rvROOH.OwnerID) = Trim(CustID) Then
        If rvROOH.RemarkContext.RemarkList.Dirty Then
            sEventCustomerMM = "Customer Remarks for " & CustID & " were changed"
        End If
    End If
End Function


'Persist changes to DB

Private Function sEventManageCustomer() As String
    If Trim(m_oCustomer.CustType) <> Trim(cboCustType.text) Then
        sEventManageCustomer = ", CustType changed from '" & Trim(m_oCustomer.CustType) & "' to '" & Trim(cboCustType.text) & "'"
    End If
    
    If Trim(m_sBackupCollector) <> Trim(cboCollector(2).text) Then
        sEventManageCustomer = sEventManageCustomer & _
                        ", Collector changed from '" & Trim(m_sBackupCollector) & "' to '" & Trim(cboCollector(2).text) & "'"
    End If

    If Trim(m_sBackupPmtTerms) <> Trim(cboPmtTerms.text) Then
        sEventManageCustomer = sEventManageCustomer & _
                        ", Pmt Terms changed from '" & Trim(m_sBackupPmtTerms) & "' to '" & Trim(cboPmtTerms.text) & "'"
    End If
    
    If Trim(m_lBackupCreditLimit) <> txtCreditLimit.Amount Then
        sEventManageCustomer = sEventManageCustomer & _
                        ", Credit Limit changed from '" & m_lBackupCreditLimit & "' to '" & txtCreditLimit.Amount & "'"
    End If
    
    If m_bStatusTypeChg Then
        If opVIP.value = True And Trim(m_sBackupARStatus) <> "VIP" Then
            sEventManageCustomer = sEventManageCustomer & _
                        ", AR Status changed from '" & Trim(m_sBackupARStatus) & "' to 'VIP'"
        ElseIf opManual.value = True And Trim(m_sBackupARStatus) <> "Manual_Hold" Then
            sEventManageCustomer = sEventManageCustomer & _
                        ", AR Status changed from '" & Trim(m_sBackupARStatus) & "' to 'Manual Hold'"
        ElseIf opAuto.value = True And Trim(m_sBackupARStatus) <> "Good" Or Trim(m_sBackupARStatus) <> "Auto_Hold" Then
            sEventManageCustomer = sEventManageCustomer & _
                        ", AR Status changed from '" & Trim(m_sBackupARStatus) & "' to 'Auto Hold'"
        End If
    End If
End Function


Private Sub logManageCustomerEvent()
    Dim sTemp As String
    
    sTemp = sEventManageCustomer
    
    If Trim(sTemp) <> "" Then
         sTemp = Right(sTemp, Len(sTemp) - 2)   'what are we trimming off?
         LogOAEvent "Manage Customer", GetUserID, m_oCustomer.Key, , , "Events for Customer " & m_oCustomer.ID & ": " & sTemp
    End If
End Sub

Private Sub SetUpStatement()
    'Populate the ComboBox
    cboStatementSetting.AddItem ("Over45 (1-C)")
    cboStatementSetting.AddItem ("Over45 (D-E)")
    cboStatementSetting.AddItem ("Over45 (F-H)")
    cboStatementSetting.AddItem ("Over45 (I-K)")
    cboStatementSetting.AddItem ("Over45 (L-0)")
    cboStatementSetting.AddItem ("Over45 (P-Q)")
    cboStatementSetting.AddItem ("Over45 (R-T)")
    cboStatementSetting.AddItem ("Over45 (U-Z)")
    cboStatementSetting.AddItem ("Over45 (1-C)")
    cboStatementSetting.ListIndex = 0
    
End Sub

Private Sub DisplayStatementCount()
    Dim sStartCust As String
    Dim sEndCust As String
    
    Select Case cboStatementSetting.text
        Case "Over45 (1-C)"
            sStartCust = "1"
            sEndCust = "CZZZZZZ"
        Case "Over45 (D-E)"
            sStartCust = "D"
            sEndCust = "EZZZZZZ"
        Case "Over45 (F-H)"
            sStartCust = "F"
            sEndCust = "HZZZZZZ"
        Case "Over45 (I-K)"
            sStartCust = "I"
            sEndCust = "KZZZZZZ"
        Case "Over45 (L-0)"
            sStartCust = "L"
            sEndCust = "LZZZZZZ"
        Case "Over45 (P-Q)"
            sStartCust = "P"
            sEndCust = "QZZZZZZ"
        Case "Over45 (R-T)"
            sStartCust = "R"
            sEndCust = "TZZZZZZ"
        Case "Over45 (U-Z)"
            sStartCust = "U"
            sEndCust = "ZZZZZZZ"
    End Select
    
    'PRN#70
    If m_oRstResearch Is Nothing Then
        Set m_oRstResearch = New ADODB.Recordset
    Else
        CloseRst m_oRstResearch
        Set m_oRstResearch = New ADODB.Recordset
    End If
    
    'LoadRst  m_oRstResearch
    SetWaitCursor True
    Set m_oRstResearch = CallSP("spcpcGetStatementPageCount", "@i_StartCustName", sStartCust, "@i_EndCustName", sEndCust)
    SetWaitCursor False

    m_oRstResearch.Filter = "StmtPageCount>1"
    If m_oRstResearch.EOF Then
        msg "No matches were found."
    End If
    'AttachGrid gdxResearch, m_oRstResearch
    Set gdxResearch.ADORecordset = m_oRstResearch
        
End Sub


'**********************************************************************************
' tabMain(8)    National Accounts Tool
'
' ShipDays < 90 = Authorized Accounts
' ShipDays > 89 = Non-Authorized Accounts
'**********************************************************************************

Private Sub cmdLoadCust_Click()
    If Len(txtCustLoad.text) = 0 Or InStr(1, txtCustLoad, "'") > 0 Then
        MsgBox "Please enter a valid Customer ID."
        txtCustLoad.text = ""
        lblCustID.Caption = ""
        txtCustLoad.SetFocus
    Else
        lblCustID.Caption = txtCustLoad.text
        txtCustLoad.text = ""
        Call ProcessAuthAddr
        Call ProcessNonAuthAddr
    End If
End Sub

Private Sub ClearlstAuthAddr()
    lstAuthAddr.Clear
    txtAuthAddrName(0).text = ""
    txtAuthAddrLine1(0).text = ""
    txtAuthAddrLine2(0).text = ""
    txtAuthCity(0).text = ""
    txtAuthState(0).text = ""
    txtAuthZip(0).text = ""
    txtAuthCountry(0).text = ""
End Sub

Private Sub ClearlstNonAuthAddr()
    lstNonAuthAddr.Clear
    txtAuthAddrName(1).text = ""
    txtAuthAddrLine1(1).text = ""
    txtAuthAddrLine2(1).text = ""
    txtAuthCity(1).text = ""
    txtAuthState(1).text = ""
    txtAuthZip(1).text = ""
    txtAuthCountry(1).text = ""
End Sub

Private Sub ProcessAuthAddr()
    Dim sSQL As String
    Dim m_rst As ADODB.Recordset
    Dim liCounter As Long
        
    ClearlstAuthAddr
    Set m_rst = New ADODB.Recordset
    sSQL = "SELECT * FROM tarCustAddr INNER JOIN tciAddress ON tarCustAddr.AddrKey = dbo.tciAddress.AddrKey "
    sSQL = sSQL & "INNER JOIN tarCustomer ON tarCustAddr.CustKey = tarCustomer.CustKey WHERE "
    sSQL = sSQL & "(dbo.tarCustAddr.ShipDays < 90) AND (dbo.tarCustomer.CustID = '" & lblCustID.Caption & "')"
    
    SetWaitCursor True
    m_rst.Open sSQL, g_DB.Connection
    SetWaitCursor False
    
    If m_rst.EOF And m_rst.BOF Then Exit Sub
    
    ReDim m_utdAuthAddr(m_rst.RecordCount - 1) As AuthAddr
    For liCounter = 0 To m_rst.RecordCount - 1
        With m_utdAuthAddr(liCounter)
            .lAddrKey = m_rst.Fields("AddrKey").value
            .lShipDays = m_rst.Fields("ShipDays").value
            .sAddrName = m_rst.Fields("AddrName").value
            .sAddrLine1 = m_rst.Fields("AddrLine1").value
            .sAddrLine2 = IIf(IsNull(m_rst.Fields("AddrLine2").value), "", m_rst.Fields("AddrLine2").value)
            .sCity = m_rst.Fields("City").value
            .sStateID = IIf(IsNull(m_rst.Fields("StateID").value), "", m_rst.Fields("StateID").value)
            .sCountryID = m_rst.Fields("CountryID").value
            .sZipCode = m_rst.Fields("PostalCode").value
            
            lstAuthAddr.AddItem m_rst.Fields("AddrName").value, liCounter
        End With
        m_rst.MoveNext
    Next liCounter
    If lstAuthAddr.ListCount > 0 Then lstAuthAddr.Selected(0) = True
End Sub

Private Sub ProcessNonAuthAddr()
    Dim sSQL As String
    Dim m_rst As ADODB.Recordset
    Dim liCounter As Long
        
    ClearlstNonAuthAddr
    Set m_rst = New ADODB.Recordset
    sSQL = "SELECT * FROM tarCustAddr INNER JOIN tciAddress ON tarCustAddr.AddrKey = dbo.tciAddress.AddrKey "
    sSQL = sSQL & "INNER JOIN tarCustomer ON tarCustAddr.CustKey = tarCustomer.CustKey WHERE "
    sSQL = sSQL & "(dbo.tarCustAddr.ShipDays > 89) AND (dbo.tarCustomer.CustID = '" & lblCustID.Caption & "')"
    
    SetWaitCursor True
    m_rst.Open sSQL, g_DB.Connection
    SetWaitCursor False

    If m_rst.EOF And m_rst.BOF Then Exit Sub
    
    ReDim m_utdNonAuthAddr(m_rst.RecordCount - 1) As AuthAddr
    For liCounter = 0 To m_rst.RecordCount - 1
        With m_utdNonAuthAddr(liCounter)
            .lAddrKey = m_rst.Fields("AddrKey").value
            .lShipDays = m_rst.Fields("ShipDays").value
            .sAddrName = m_rst.Fields("AddrName").value
            .sAddrLine1 = m_rst.Fields("AddrLine1").value
            .sAddrLine2 = IIf(IsNull(m_rst.Fields("AddrLine2").value), "", m_rst.Fields("AddrLine2").value)
            .sCity = m_rst.Fields("City").value
            .sStateID = IIf(IsNull(m_rst.Fields("StateID").value), "", m_rst.Fields("StateID").value)
            .sCountryID = m_rst.Fields("CountryID").value
            .sZipCode = m_rst.Fields("PostalCode").value
            
            lstNonAuthAddr.AddItem m_rst.Fields("AddrName").value, liCounter
        End With
        m_rst.MoveNext
    Next liCounter
    If lstNonAuthAddr.ListCount > 0 Then lstNonAuthAddr.Selected(0) = True
End Sub

Private Sub cmdMovetoNonAuth91_Click()
    If lstAuthAddr.ListCount = 0 Then Exit Sub
    Dim sSQL As String
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    With m_utdAuthAddr(lstAuthAddr.ListIndex) 'selected row = lstAuthAddr.listindex
        sSQL = "Update tarcustaddr set shipdays = 91 where addrkey = " & .lAddrKey
        Set cmd = CreateCommandSP(sSQL, adCmdText)
        cmd.Execute
        Set cmd = Nothing
    End With

    Call ProcessAuthAddr
    Call ProcessNonAuthAddr
End Sub

Private Sub cmdMovetoAuth0_Click()
    If lstNonAuthAddr.ListCount = 0 Then Exit Sub
    Dim sSQL As String
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With m_utdNonAuthAddr(lstNonAuthAddr.ListIndex) 'selected row = lstNonAuthAddr.listindex
        sSQL = "Update tarcustaddr set shipdays = 0 where addrkey = " & .lAddrKey
        Set cmd = CreateCommandSP(sSQL, adCmdText)
        cmd.Execute
        Set cmd = Nothing
    End With

    Call ProcessAuthAddr
    Call ProcessNonAuthAddr
End Sub

Private Sub lstAuthAddr_Click()
    With m_utdAuthAddr(lstAuthAddr.ListIndex)
        txtAuthAddrName(0).text = .sAddrName
        txtAuthAddrLine1(0).text = .sAddrLine1
        txtAuthAddrLine2(0).text = .sAddrLine2
        txtAuthCity(0).text = .sCity
        txtAuthState(0).text = .sStateID
        txtAuthZip(0).text = .sZipCode
        txtAuthCountry(0).text = .sCountryID
    End With
End Sub

Private Sub lstNonAuthAddr_Click()
    With m_utdNonAuthAddr(lstNonAuthAddr.ListIndex)
        txtAuthAddrName(1).text = .sAddrName
        txtAuthAddrLine1(1).text = .sAddrLine1
        txtAuthAddrLine2(1).text = .sAddrLine2
        txtAuthCity(1).text = .sCity
        txtAuthState(1).text = .sStateID
        txtAuthZip(1).text = .sZipCode
        txtAuthCountry(1).text = .sCountryID
    End With
End Sub


'**********************************************************************************
'Sales Tax Exemption
'**********************************************************************************

'Called by DisplayStatus

'Sets these module-level variables
'   m_lAddrKey
'   m_sExemptCert
'   m_sStateID

Private Sub GetExemptCert(ByVal CustKey As Long)
    Dim ocmd As ADODB.Command
    Set ocmd = CreateCommandSP("spcpcBillingGetExemptCert")
    With ocmd
        .Parameters("@_iCustKey").value = CustKey
        .Execute
        If Not IsNull(.Parameters("@_oExmptNo").value) Then
            m_lAddrKey = .Parameters("@_oAddrKey").value
            m_sExemptCert = Trim(.Parameters("@_oExmptNo").value)
            m_sStateID = Trim(.Parameters("@_oStateID").value)
        End If
    End With
    Set ocmd = Nothing
End Sub


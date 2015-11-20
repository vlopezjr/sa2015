VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "MMRemark.ocx"
Object = "{0FA91D91-3062-44DB-B896-91406D28F92A}#54.0#0"; "SOTACalendar.ocx"
Begin VB.Form FPurchAssist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   11070
   Begin ActiveTabs.SSActiveTabs tabMain 
      Height          =   5472
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10812
      _ExtentX        =   19076
      _ExtentY        =   9657
      _Version        =   262144
      TabCount        =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHotTracking {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "FPurchAssist.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   176
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":02DC
         Begin VB.CommandButton cmdRefreshDropShips 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   9360
            TabIndex        =   179
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboDropShipBuyers 
            Height          =   315
            Left            =   120
            TabIndex        =   178
            Top             =   120
            Width           =   1215
         End
         Begin MSComctlLib.TreeView tvDSPO 
            Height          =   4095
            Left            =   120
            TabIndex        =   177
            Top             =   840
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   7223
            _Version        =   393217
            Style           =   5
            Appearance      =   1
         End
         Begin VB.Label Label48 
            Caption         =   " (* requires fixing)"
            Height          =   255
            Left            =   120
            TabIndex        =   180
            Top             =   600
            Width           =   1695
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel13 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   152
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":0304
         Begin VB.CommandButton cmdPrint_SalesAnalysis 
            Caption         =   "&Print"
            Height          =   315
            Left            =   9720
            TabIndex        =   157
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdFind_SalesAnalysis 
            Caption         =   "&Find"
            Height          =   315
            Left            =   8640
            TabIndex        =   156
            Top             =   120
            Width           =   855
         End
         Begin VB.ComboBox cboWhse_SalesAnalysis 
            Height          =   315
            Left            =   3060
            Style           =   2  'Dropdown List
            TabIndex        =   154
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtItemID_SalesAnalysis 
            Height          =   315
            Left            =   510
            MaxLength       =   30
            TabIndex        =   153
            Top             =   120
            Width           =   1215
         End
         Begin VB.ComboBox cboVend_SalesAnalysis 
            Height          =   315
            Left            =   5430
            Style           =   2  'Dropdown List
            TabIndex        =   155
            Top             =   120
            Width           =   3075
         End
         Begin GridEX20.GridEX gdxSalesAnalysis 
            Height          =   4395
            Left            =   120
            TabIndex        =   158
            Top             =   600
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   7752
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            DefaultGroupMode=   1
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            SelectionStyle  =   1
            Options         =   8
            RecordsetType   =   1
            GroupByBoxVisible=   0   'False
            DataMode        =   1
            ColumnHeaderHeight=   285
            ColumnsCount    =   11
            Column(1)       =   "FPurchAssist.frx":032C
            Column(2)       =   "FPurchAssist.frx":0474
            Column(3)       =   "FPurchAssist.frx":05C0
            Column(4)       =   "FPurchAssist.frx":07A0
            Column(5)       =   "FPurchAssist.frx":0950
            Column(6)       =   "FPurchAssist.frx":0AA0
            Column(7)       =   "FPurchAssist.frx":0C84
            Column(8)       =   "FPurchAssist.frx":0E04
            Column(9)       =   "FPurchAssist.frx":0F94
            Column(10)      =   "FPurchAssist.frx":1124
            Column(11)      =   "FPurchAssist.frx":1264
            SortKeysCount   =   2
            SortKey(1)      =   "FPurchAssist.frx":13A4
            SortKey(2)      =   "FPurchAssist.frx":140C
            GroupConditionCountTitle=   ""
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPurchAssist.frx":1474
            FormatStyle(2)  =   "FPurchAssist.frx":1554
            FormatStyle(3)  =   "FPurchAssist.frx":168C
            FormatStyle(4)  =   "FPurchAssist.frx":173C
            FormatStyle(5)  =   "FPurchAssist.frx":17F0
            FormatStyle(6)  =   "FPurchAssist.frx":18C8
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":1980
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Warehouse"
            Height          =   195
            Index           =   3
            Left            =   2190
            TabIndex        =   161
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Item"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   160
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vendor"
            Height          =   195
            Index           =   1
            Left            =   4845
            TabIndex        =   159
            Top             =   180
            Width           =   510
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel12 
         Height          =   5085
         Left            =   30
         TabIndex        =   146
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":1B58
         Begin VB.Frame fraBranchTransfer 
            Caption         =   "Branch Transfer "
            Height          =   855
            Left            =   120
            TabIndex        =   163
            Top             =   4200
            Width           =   10575
            Begin VB.CommandButton cmdOpenTransRpt 
               Caption         =   "All Open POs"
               Height          =   315
               Left            =   120
               TabIndex        =   165
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton cmdTransSumRpt 
               Caption         =   "Summary by Whse"
               Default         =   -1  'True
               Height          =   315
               Left            =   1560
               TabIndex        =   164
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.ComboBox cboPOWhse 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   162
            Top             =   120
            Width           =   735
         End
         Begin VB.ComboBox cboBuyers 
            Height          =   315
            Index           =   2
            ItemData        =   "FPurchAssist.frx":1B80
            Left            =   120
            List            =   "FPurchAssist.frx":1B82
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh Grid"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   2160
            TabIndex        =   149
            Top             =   120
            Width           =   1095
         End
         Begin VB.ComboBox cboVendors 
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   120
            Width           =   2415
         End
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Use this Vendor"
            Height          =   315
            Index           =   2
            Left            =   6000
            TabIndex        =   147
            Top             =   120
            Width           =   1335
         End
         Begin GridEX20.GridEX gdxVendorsNew 
            Height          =   3495
            Left            =   120
            TabIndex        =   151
            Top             =   600
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   6165
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            Options         =   8
            RecordsetType   =   1
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   1
            ColumnHeaderHeight=   285
            ColumnsCount    =   5
            Column(1)       =   "FPurchAssist.frx":1B84
            Column(2)       =   "FPurchAssist.frx":1D28
            Column(3)       =   "FPurchAssist.frx":1E74
            Column(4)       =   "FPurchAssist.frx":204C
            Column(5)       =   "FPurchAssist.frx":2230
            SortKeysCount   =   3
            SortKey(1)      =   "FPurchAssist.frx":240C
            SortKey(2)      =   "FPurchAssist.frx":2474
            SortKey(3)      =   "FPurchAssist.frx":24DC
            FmtConditionsCount=   3
            FmtCondition(1) =   "FPurchAssist.frx":2544
            FmtCondition(2) =   "FPurchAssist.frx":26BC
            FmtCondition(3) =   "FPurchAssist.frx":2818
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPurchAssist.frx":296C
            FormatStyle(2)  =   "FPurchAssist.frx":2A4C
            FormatStyle(3)  =   "FPurchAssist.frx":2B84
            FormatStyle(4)  =   "FPurchAssist.frx":2C34
            FormatStyle(5)  =   "FPurchAssist.frx":2CE8
            FormatStyle(6)  =   "FPurchAssist.frx":2DC0
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":2E78
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel11 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   128
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":3050
         Begin GridEX20.GridEX gdxPOOrder 
            Height          =   2955
            Left            =   120
            TabIndex        =   175
            Top             =   1740
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   5212
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowCardSizing =   0   'False
            AllowColumnDrag =   0   'False
            AutomaticArrange=   0   'False
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            ColumnsCount    =   4
            Column(1)       =   "FPurchAssist.frx":3078
            Column(2)       =   "FPurchAssist.frx":31B0
            Column(3)       =   "FPurchAssist.frx":32D4
            Column(4)       =   "FPurchAssist.frx":3404
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPurchAssist.frx":3510
            FormatStyle(2)  =   "FPurchAssist.frx":35F0
            FormatStyle(3)  =   "FPurchAssist.frx":3728
            FormatStyle(4)  =   "FPurchAssist.frx":37D8
            FormatStyle(5)  =   "FPurchAssist.frx":388C
            FormatStyle(6)  =   "FPurchAssist.frx":3964
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":3A1C
         End
         Begin VB.Frame Frame9 
            Caption         =   "Find Drop Ship Customer from PO Number"
            Height          =   675
            Left            =   120
            TabIndex        =   171
            Top             =   960
            Width           =   10455
            Begin VB.CommandButton cmdFindAssociatedOrder 
               Caption         =   "Find Customer"
               Height          =   315
               Left            =   3000
               TabIndex        =   174
               Top             =   240
               Width           =   1275
            End
            Begin VB.TextBox txtLookupPONo 
               Height          =   315
               Left            =   840
               TabIndex        =   172
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PO #"
               Height          =   195
               Left            =   240
               TabIndex        =   173
               Top             =   300
               Width           =   375
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Find Unissued or Open Purchase Orders"
            Height          =   795
            Left            =   120
            TabIndex        =   130
            Top             =   120
            Width           =   10455
            Begin VB.CommandButton cmdLoadPO 
               Caption         =   "Load PO"
               Height          =   315
               Left            =   8640
               TabIndex        =   136
               Top             =   300
               Width           =   1275
            End
            Begin VB.CommandButton cmdFindPO 
               Caption         =   "Find PO"
               Height          =   315
               Left            =   7200
               TabIndex        =   135
               Top             =   300
               Width           =   1275
            End
            Begin VB.ComboBox cboPOStatus 
               Height          =   315
               ItemData        =   "FPurchAssist.frx":3BF4
               Left            =   3960
               List            =   "FPurchAssist.frx":3C01
               Style           =   2  'Dropdown List
               TabIndex        =   134
               Top             =   300
               Width           =   1335
            End
            Begin VB.ComboBox cboPOBuyer 
               Height          =   315
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   131
               Top             =   300
               Width           =   1815
            End
            Begin VB.Label lblPOCount 
               Caption         =   "[Count]"
               Height          =   255
               Left            =   5760
               TabIndex        =   170
               Top             =   300
               Width           =   975
            End
            Begin VB.Label Label41 
               Caption         =   "PO Status"
               Height          =   255
               Left            =   3000
               TabIndex        =   133
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label40 
               Caption         =   "Buyer"
               Height          =   255
               Left            =   240
               TabIndex        =   132
               Top             =   300
               Width           =   495
            End
         End
         Begin GridEX20.GridEX gdxPOs 
            Height          =   3255
            Left            =   120
            TabIndex        =   129
            Top             =   1740
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   5741
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            ColumnsCount    =   10
            Column(1)       =   "FPurchAssist.frx":3C1C
            Column(2)       =   "FPurchAssist.frx":3D5C
            Column(3)       =   "FPurchAssist.frx":3EC8
            Column(4)       =   "FPurchAssist.frx":3FEC
            Column(5)       =   "FPurchAssist.frx":416C
            Column(6)       =   "FPurchAssist.frx":43DC
            Column(7)       =   "FPurchAssist.frx":44FC
            Column(8)       =   "FPurchAssist.frx":4648
            Column(9)       =   "FPurchAssist.frx":47CC
            Column(10)      =   "FPurchAssist.frx":48F8
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPurchAssist.frx":4A3C
            FormatStyle(2)  =   "FPurchAssist.frx":4B1C
            FormatStyle(3)  =   "FPurchAssist.frx":4C54
            FormatStyle(4)  =   "FPurchAssist.frx":4D04
            FormatStyle(5)  =   "FPurchAssist.frx":4DB8
            FormatStyle(6)  =   "FPurchAssist.frx":4E90
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":4F48
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel10 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   104
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":5120
         Begin VB.Frame Frame1 
            Caption         =   "Gasket Orders Status Viewer"
            Height          =   2295
            Index           =   2
            Left            =   6960
            TabIndex        =   144
            Top             =   2400
            Width           =   3135
            Begin VB.ComboBox cboWhse_GaskOrdStatus 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   125
               Top             =   270
               Width           =   1515
            End
            Begin VB.CommandButton cmdViewGaskOrdStatus 
               Caption         =   "View"
               Height          =   375
               Left            =   1080
               TabIndex        =   126
               Top             =   660
               Width           =   1515
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "Warehouse:"
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   145
               Top             =   360
               Width           =   870
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Gasket Manufacturing Viewer"
            Height          =   1575
            Index           =   1
            Left            =   6960
            TabIndex        =   142
            Top             =   240
            Width           =   3135
            Begin VB.OptionButton optSO 
               Caption         =   "SO#"
               Height          =   255
               Left            =   1800
               TabIndex        =   122
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton optOP 
               Caption         =   "OP#"
               Height          =   255
               Left            =   1080
               TabIndex        =   121
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.CommandButton cmdView 
               Caption         =   "View"
               Height          =   375
               Left            =   720
               TabIndex        =   124
               Top             =   960
               Width           =   1755
            End
            Begin VB.TextBox txtOPKey 
               Height          =   315
               Left            =   720
               MaxLength       =   9
               TabIndex        =   123
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "Use"
               Height          =   195
               Index           =   6
               Left            =   720
               TabIndex        =   143
               Top             =   360
               Width           =   285
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Gasket Production Summary"
            Height          =   2352
            Left            =   120
            TabIndex        =   109
            Top             =   2400
            Width           =   6552
            Begin VB.CommandButton cmdGet 
               Caption         =   "Get"
               Height          =   375
               Left            =   4920
               TabIndex        =   120
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txtTrimAmt 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   312
               Left            =   1140
               TabIndex        =   119
               Top             =   1680
               Width           =   732
            End
            Begin VB.TextBox txtMoldAmt 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   312
               Left            =   1140
               TabIndex        =   118
               Top             =   1320
               Width           =   732
            End
            Begin VB.TextBox txtCutAmt 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   312
               Left            =   1140
               TabIndex        =   117
               Top             =   960
               Width           =   732
            End
            Begin MSComCtl2.DTPicker dtpEndDate 
               Height          =   312
               Left            =   3360
               TabIndex        =   113
               Top             =   480
               Width           =   1212
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               Format          =   53149697
               CurrentDate     =   37621
            End
            Begin MSComCtl2.DTPicker dtpStartDate 
               Height          =   312
               Left            =   1080
               TabIndex        =   112
               Top             =   480
               Width           =   1212
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               Format          =   53149697
               CurrentDate     =   37621
            End
            Begin VB.Label Label39 
               Caption         =   "#Trimmed"
               Height          =   252
               Index           =   5
               Left            =   240
               TabIndex        =   116
               Top             =   1740
               Width           =   792
            End
            Begin VB.Label Label39 
               Caption         =   "# Molded"
               Height          =   252
               Index           =   4
               Left            =   240
               TabIndex        =   115
               Top             =   1380
               Width           =   792
            End
            Begin VB.Label Label39 
               Caption         =   "# Cut"
               Height          =   252
               Index           =   3
               Left            =   240
               TabIndex        =   114
               Top             =   1020
               Width           =   552
            End
            Begin VB.Label Label39 
               Caption         =   "End Date"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   111
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label39 
               Caption         =   "Start Date"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   110
               Top             =   480
               Width           =   855
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Print Gasket Status Report"
            Height          =   1572
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   6615
            Begin VB.ComboBox cboWhse 
               Height          =   315
               Left            =   2760
               Style           =   2  'Dropdown List
               TabIndex        =   127
               Top             =   480
               Width           =   1392
            End
            Begin VB.CheckBox chkShip 
               Caption         =   "Ship Orders"
               Height          =   375
               Left            =   240
               TabIndex        =   107
               Top             =   420
               Value           =   1  'Checked
               Width           =   1512
            End
            Begin VB.CheckBox chkWillCall 
               Caption         =   "Will Call Orders"
               Height          =   375
               Left            =   240
               TabIndex        =   106
               Top             =   840
               Value           =   1  'Checked
               Width           =   1452
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "Print Report"
               Height          =   375
               Left            =   4800
               TabIndex        =   108
               Top             =   420
               Width           =   1392
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel9 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   88
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":5148
         Begin VB.CommandButton cmdUpdateXLS 
            Caption         =   "MinMax from Files"
            Height          =   375
            Left            =   2160
            TabIndex        =   97
            Top             =   4560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton cmdSpreadSheet 
            Caption         =   "Import Robert's Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   103
            Top             =   4560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton cmdMMUpdate 
            Caption         =   "&Update"
            Height          =   375
            Left            =   9360
            TabIndex        =   98
            Top             =   4560
            Width           =   1335
         End
         Begin GridEX20.GridEX gdxMMVendItem 
            Height          =   3255
            Left            =   120
            TabIndex        =   90
            Top             =   1200
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   5741
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
            ColumnsCount    =   11
            Column(1)       =   "FPurchAssist.frx":5170
            Column(2)       =   "FPurchAssist.frx":529C
            Column(3)       =   "FPurchAssist.frx":53C8
            Column(4)       =   "FPurchAssist.frx":550C
            Column(5)       =   "FPurchAssist.frx":5650
            Column(6)       =   "FPurchAssist.frx":5794
            Column(7)       =   "FPurchAssist.frx":58CC
            Column(8)       =   "FPurchAssist.frx":5A10
            Column(9)       =   "FPurchAssist.frx":5B60
            Column(10)      =   "FPurchAssist.frx":5CA0
            Column(11)      =   "FPurchAssist.frx":5DE0
            FormatStylesCount=   7
            FormatStyle(1)  =   "FPurchAssist.frx":5F20
            FormatStyle(2)  =   "FPurchAssist.frx":6058
            FormatStyle(3)  =   "FPurchAssist.frx":6108
            FormatStyle(4)  =   "FPurchAssist.frx":61BC
            FormatStyle(5)  =   "FPurchAssist.frx":6294
            FormatStyle(6)  =   "FPurchAssist.frx":634C
            FormatStyle(7)  =   "FPurchAssist.frx":642C
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":6478
         End
         Begin VB.Frame Frame5 
            Caption         =   "Search Existed Vendor Items"
            Height          =   975
            Left            =   120
            TabIndex        =   89
            Top             =   120
            Width           =   10575
            Begin VB.CommandButton cmdMMRefresh 
               Caption         =   "&Find"
               Height          =   375
               Left            =   8160
               TabIndex        =   95
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboMMWarehouse 
               Height          =   315
               Left            =   5280
               Style           =   2  'Dropdown List
               TabIndex        =   94
               Top             =   420
               Width           =   1815
            End
            Begin VB.TextBox txtVendID 
               Height          =   315
               Left            =   2640
               MaxLength       =   50
               TabIndex        =   92
               Top             =   420
               Width           =   1455
            End
            Begin VB.Label Label38 
               Caption         =   "Warehouse"
               Height          =   255
               Left            =   4320
               TabIndex        =   93
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label37 
               Caption         =   " PartNbr, VendID, or VendName"
               Height          =   255
               Left            =   240
               TabIndex        =   91
               Top             =   480
               Width           =   2415
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel8 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   69
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":6650
         Begin VB.CheckBox chkVIUnselectAll 
            Caption         =   "Unselect All"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1260
            TabIndex        =   77
            Top             =   4080
            Width           =   1215
         End
         Begin VB.CheckBox chkVISelectAll 
            Caption         =   "Select All"
            Enabled         =   0   'False
            Height          =   255
            Left            =   180
            TabIndex        =   76
            ToolTipText     =   "Select for update all qualified items"
            Top             =   4080
            Width           =   975
         End
         Begin VB.CommandButton cmdVICreate 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   315
            Left            =   7080
            TabIndex        =   80
            Top             =   4080
            Width           =   1095
         End
         Begin VB.ComboBox cboVIVendor 
            Height          =   288
            Index           =   1
            Left            =   4380
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   4080
            Width           =   2175
         End
         Begin VB.Frame Frame2 
            Caption         =   "Select Existing Vendor Items"
            Height          =   975
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   10455
            Begin VB.CommandButton cmdClear 
               Caption         =   "Clear"
               Height          =   312
               Left            =   9120
               TabIndex        =   84
               Top             =   360
               Width           =   1092
            End
            Begin VB.TextBox txtVIPartNumber 
               Height          =   315
               Left            =   5280
               MaxLength       =   30
               TabIndex        =   74
               Top             =   360
               Width           =   1455
            End
            Begin VB.CommandButton cmdVendItemFind 
               Caption         =   "Find"
               Height          =   312
               Left            =   7800
               TabIndex        =   75
               Top             =   360
               Width           =   1092
            End
            Begin VB.ComboBox cboVIVendor 
               Height          =   315
               Index           =   0
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label Label21 
               Caption         =   "Part Number"
               Height          =   255
               Left            =   3840
               TabIndex        =   73
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label20 
               Caption         =   "Vendor"
               Height          =   252
               Left            =   180
               TabIndex        =   72
               Top             =   360
               Width           =   612
            End
         End
         Begin GridEX20.GridEX gdxVendItem 
            Height          =   2535
            Left            =   120
            TabIndex        =   87
            Top             =   1440
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   4471
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            ItemCount       =   0
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   6
            Column(1)       =   "FPurchAssist.frx":6678
            Column(2)       =   "FPurchAssist.frx":67FC
            Column(3)       =   "FPurchAssist.frx":6968
            Column(4)       =   "FPurchAssist.frx":6AB4
            Column(5)       =   "FPurchAssist.frx":6C00
            Column(6)       =   "FPurchAssist.frx":6D5C
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPurchAssist.frx":6EA8
            FormatStyle(2)  =   "FPurchAssist.frx":6FE0
            FormatStyle(3)  =   "FPurchAssist.frx":7090
            FormatStyle(4)  =   "FPurchAssist.frx":7144
            FormatStyle(5)  =   "FPurchAssist.frx":721C
            FormatStyle(6)  =   "FPurchAssist.frx":72D4
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":73B4
         End
         Begin VB.Label Label34 
            Caption         =   "Note"
            Height          =   252
            Left            =   180
            TabIndex        =   86
            Top             =   4500
            Width           =   432
         End
         Begin VB.Label Label33 
            Caption         =   "Add New Vendor Items"
            Height          =   252
            Left            =   240
            TabIndex        =   85
            Top             =   1200
            Width           =   1812
         End
         Begin VB.Label Label31 
            Caption         =   "2. You cannot add a new Vendor/Item relationship if the item already exists for the vendor."
            Height          =   195
            Left            =   720
            TabIndex        =   83
            Top             =   4740
            Width           =   6375
         End
         Begin VB.Label Label32 
            Caption         =   "1. If there is more than one record for the same part number, you can only select one."
            Height          =   255
            Left            =   720
            TabIndex        =   82
            Top             =   4500
            Width           =   6075
         End
         Begin VB.Label Label22 
            Caption         =   "Vendor"
            Height          =   252
            Left            =   3600
            TabIndex        =   78
            Top             =   4080
            Width           =   612
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   55
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":758C
         Begin VB.TextBox txtVendCostItemID 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7140
            MaxLength       =   30
            TabIndex        =   67
            Top             =   4260
            Visible         =   0   'False
            Width           =   3312
         End
         Begin VB.CommandButton cmdAddRobertFile 
            Caption         =   "Import Robert's Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5220
            TabIndex        =   100
            Top             =   4620
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmdVendUpdate 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   315
            Left            =   6660
            TabIndex        =   68
            Top             =   1860
            Width           =   1212
         End
         Begin VB.TextBox txtVendComment 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7140
            MaxLength       =   225
            TabIndex        =   66
            Top             =   3756
            Visible         =   0   'False
            Width           =   3312
         End
         Begin VB.ComboBox cboVendID 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   1860
            Width           =   3312
         End
         Begin VB.CheckBox chkObsolete 
            Caption         =   "Obsolete"
            Enabled         =   0   'False
            Height          =   252
            Left            =   7140
            TabIndex        =   62
            Top             =   4740
            Visible         =   0   'False
            Width           =   1512
         End
         Begin VB.Frame frmItemSearch 
            Height          =   975
            Left            =   180
            TabIndex        =   56
            Top             =   660
            Width           =   9015
            Begin VB.CommandButton cmdVendFind 
               Caption         =   "Find"
               Height          =   315
               Left            =   6480
               TabIndex        =   61
               Top             =   360
               Width           =   1215
            End
            Begin VB.ComboBox cboVendWhse 
               Height          =   315
               Left            =   3960
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtItemID 
               Height          =   315
               Left            =   1080
               MaxLength       =   30
               TabIndex        =   58
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lblVendWhse 
               Alignment       =   1  'Right Justify
               Caption         =   "Warehouse"
               Height          =   252
               Left            =   2820
               TabIndex        =   59
               Top             =   360
               Width           =   972
            End
            Begin VB.Label lblVendItem 
               Alignment       =   1  'Right Justify
               Caption         =   "Item"
               Height          =   252
               Left            =   540
               TabIndex        =   57
               Top             =   360
               Width           =   432
            End
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "VendCostItemID"
            Enabled         =   0   'False
            Height          =   252
            Left            =   5640
            TabIndex        =   137
            Top             =   4260
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Label39 
            Caption         =   "This tab is used to manage the Warehouse/Vendor/Item relationship in OfficeAssistant."
            Height          =   372
            Index           =   0
            Left            =   240
            TabIndex        =   99
            Top             =   180
            Width           =   6552
         End
         Begin VB.Label lblVendName 
            Height          =   252
            Left            =   6420
            TabIndex        =   81
            Top             =   1920
            Width           =   2772
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Comments"
            Enabled         =   0   'False
            Height          =   252
            Left            =   6000
            TabIndex        =   65
            Top             =   3756
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Vendor"
            Height          =   192
            Index           =   0
            Left            =   480
            TabIndex        =   63
            Top             =   1920
            Width           =   672
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   40
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":75B4
         Begin VB.CommandButton cmdManageShipments 
            Caption         =   "Manage Shipments"
            Height          =   495
            Left            =   8700
            TabIndex        =   181
            Top             =   1860
            Width           =   1755
         End
         Begin VB.TextBox txtExPOShipMth 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   1320
            Width           =   2475
         End
         Begin VB.TextBox txtExPOStatus 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1395
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   312
            Left            =   7200
            TabIndex        =   5
            Top             =   240
            Width           =   852
         End
         Begin VB.TextBox txtVendFax 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   960
            Width           =   1392
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox txtVendPhone 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   960
            Width           =   1392
         End
         Begin MMRemark.RemarkViewer rvPO 
            Height          =   810
            Left            =   9840
            TabIndex        =   3
            Top             =   120
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1429
            ContextID       =   "ViewPO"
            Caption         =   "PO Remarks"
         End
         Begin VB.TextBox txtTranDate 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   600
            Width           =   1212
         End
         Begin VB.TextBox txtBuyerName 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   3300
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   600
            Width           =   1392
         End
         Begin VB.TextBox txtVendName 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   960
            Width           =   3552
         End
         Begin SOTACalendarControl.SOTACalendar calExpectedDate 
            Height          =   288
            Left            =   5700
            TabIndex        =   4
            Top             =   240
            Width           =   1392
            _ExtentX        =   2461
            _ExtentY        =   503
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
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Load"
            Height          =   312
            Index           =   1
            Left            =   2520
            TabIndex        =   1
            Top             =   180
            Width           =   672
         End
         Begin VB.TextBox txtPOID 
            Height          =   312
            Index           =   1
            Left            =   1140
            TabIndex        =   0
            Top             =   180
            Width           =   1212
         End
         Begin GridEX20.GridEX gdxPOLines 
            Height          =   2415
            Left            =   120
            TabIndex        =   2
            Top             =   2520
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   4260
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   270
            ColumnsCount    =   10
            Column(1)       =   "FPurchAssist.frx":75DC
            Column(2)       =   "FPurchAssist.frx":776C
            Column(3)       =   "FPurchAssist.frx":78BC
            Column(4)       =   "FPurchAssist.frx":7A28
            Column(5)       =   "FPurchAssist.frx":7B6C
            Column(6)       =   "FPurchAssist.frx":7CC8
            Column(7)       =   "FPurchAssist.frx":7E6C
            Column(8)       =   "FPurchAssist.frx":8004
            Column(9)       =   "FPurchAssist.frx":81A0
            Column(10)      =   "FPurchAssist.frx":82DC
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPurchAssist.frx":83F4
            FormatStyle(2)  =   "FPurchAssist.frx":84D4
            FormatStyle(3)  =   "FPurchAssist.frx":860C
            FormatStyle(4)  =   "FPurchAssist.frx":86BC
            FormatStyle(5)  =   "FPurchAssist.frx":8770
            FormatStyle(6)  =   "FPurchAssist.frx":8848
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":8900
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Caption         =   "Ship Mthd"
            Height          =   255
            Left            =   4800
            TabIndex        =   169
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "Status"
            Height          =   255
            Left            =   4800
            TabIndex        =   167
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label txtShipAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   1140
            TabIndex        =   141
            Top             =   1320
            Width           =   3552
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            Caption         =   "Address"
            Height          =   255
            Left            =   297
            TabIndex        =   140
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "Shipping"
            Height          =   255
            Left            =   297
            TabIndex        =   139
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact"
            Height          =   255
            Left            =   4800
            TabIndex        =   52
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax"
            Height          =   255
            Left            =   7080
            TabIndex        =   51
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Phone"
            Height          =   255
            Left            =   4800
            TabIndex        =   50
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Date"
            Height          =   192
            Left            =   480
            TabIndex        =   45
            Top             =   620
            Width           =   552
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Vendor"
            Height          =   252
            Left            =   240
            TabIndex        =   44
            Top             =   940
            Width           =   792
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Buyer"
            Height          =   192
            Left            =   2460
            TabIndex        =   43
            Top             =   600
            Width           =   672
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Request Date"
            Height          =   195
            Index           =   0
            Left            =   3960
            TabIndex        =   42
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "PO Number"
            Height          =   252
            Left            =   180
            TabIndex        =   41
            Top             =   240
            Width           =   852
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   29
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":8AD8
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Load"
            Height          =   312
            Index           =   0
            Left            =   3960
            TabIndex        =   39
            Top             =   600
            Width           =   1032
         End
         Begin VB.CommandButton cmdFreeze 
            Caption         =   "Freeze"
            Height          =   312
            Left            =   3960
            TabIndex        =   36
            Top             =   2220
            Width           =   1032
         End
         Begin VB.ListBox lstPO 
            Height          =   255
            ItemData        =   "FPurchAssist.frx":8B00
            Left            =   5340
            List            =   "FPurchAssist.frx":8B02
            TabIndex        =   35
            Top             =   1680
            Width           =   3192
         End
         Begin VB.ListBox lstSO 
            Height          =   255
            ItemData        =   "FPurchAssist.frx":8B04
            Left            =   360
            List            =   "FPurchAssist.frx":8B06
            TabIndex        =   34
            Top             =   1680
            Width           =   3192
         End
         Begin VB.TextBox txtPOID 
            Height          =   288
            Index           =   0
            Left            =   2700
            TabIndex        =   31
            Top             =   600
            Width           =   972
         End
         Begin VB.TextBox txtSOID 
            Height          =   288
            Left            =   2700
            TabIndex        =   30
            Top             =   240
            Width           =   972
         End
         Begin VB.Label lblVendorName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5400
            TabIndex        =   102
            Top             =   600
            Width           =   3732
         End
         Begin VB.Label lblCustName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5400
            TabIndex        =   101
            Top             =   240
            Width           =   3732
         End
         Begin VB.Label lblWarning 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   972
            Left            =   360
            TabIndex        =   96
            Top             =   3420
            Width           =   8172
         End
         Begin VB.Label Label17 
            Caption         =   "Purchase Order Line Items"
            Height          =   252
            Left            =   5340
            TabIndex        =   38
            Top             =   1380
            Width           =   1992
         End
         Begin VB.Label Label16 
            Caption         =   "Sales Order Line Items"
            Height          =   252
            Left            =   360
            TabIndex        =   37
            Top             =   1380
            Width           =   1872
         End
         Begin VB.Label Label15 
            Caption         =   "Acuity Purchase Order Number"
            Height          =   312
            Left            =   360
            TabIndex        =   33
            Top             =   660
            Width           =   2292
         End
         Begin VB.Label Label14 
            Caption         =   "Acuity Sales Order Number"
            Height          =   312
            Left            =   360
            TabIndex        =   32
            Top             =   300
            Width           =   2292
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   20
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":8B08
         Begin VB.CommandButton cmdPrintSPO 
            Caption         =   "Print"
            Height          =   312
            Left            =   9480
            TabIndex        =   28
            Top             =   720
            Width           =   972
         End
         Begin VB.ComboBox cboBuyers 
            Height          =   315
            Index           =   1
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   138
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optAll 
            Caption         =   "Show all SPO items from"
            Height          =   252
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   2052
         End
         Begin VB.OptionButton optNotOrdered 
            Caption         =   "Show SPO items not ordered yet (excludes DropShips)"
            Height          =   312
            Left            =   240
            TabIndex        =   24
            Top             =   180
            Value           =   -1  'True
            Width           =   4512
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   312
            Index           =   1
            Left            =   9480
            TabIndex        =   22
            Top             =   240
            Width           =   972
         End
         Begin SOTACalendarControl.SOTACalendar calStart 
            Height          =   288
            Left            =   2280
            TabIndex        =   26
            Top             =   540
            Width           =   1632
            _ExtentX        =   2884
            _ExtentY        =   503
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
         Begin SOTACalendarControl.SOTACalendar calEnd 
            Height          =   288
            Left            =   4320
            TabIndex        =   27
            Top             =   540
            Width           =   1632
            _ExtentX        =   2884
            _ExtentY        =   503
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
         Begin GridEX20.GridEX gdxSPOStatus 
            Height          =   3855
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   6800
            Version         =   "2.0"
            AutomaticSort   =   -1  'True
            ShowToolTips    =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            ColumnsCount    =   12
            Column(1)       =   "FPurchAssist.frx":8B30
            Column(2)       =   "FPurchAssist.frx":8E50
            Column(3)       =   "FPurchAssist.frx":8F64
            Column(4)       =   "FPurchAssist.frx":90D4
            Column(5)       =   "FPurchAssist.frx":9228
            Column(6)       =   "FPurchAssist.frx":933C
            Column(7)       =   "FPurchAssist.frx":94B0
            Column(8)       =   "FPurchAssist.frx":95D4
            Column(9)       =   "FPurchAssist.frx":9704
            Column(10)      =   "FPurchAssist.frx":9810
            Column(11)      =   "FPurchAssist.frx":9938
            Column(12)      =   "FPurchAssist.frx":9A5C
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPurchAssist.frx":9B80
            FormatStyle(2)  =   "FPurchAssist.frx":9C98
            FormatStyle(3)  =   "FPurchAssist.frx":9DD0
            FormatStyle(4)  =   "FPurchAssist.frx":9E80
            FormatStyle(5)  =   "FPurchAssist.frx":9F34
            FormatStyle(6)  =   "FPurchAssist.frx":A00C
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":A0C4
         End
         Begin VB.Label Label5 
            Caption         =   "to"
            Height          =   252
            Left            =   4020
            TabIndex        =   21
            Top             =   600
            Width           =   252
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5085
         Left            =   -99969
         TabIndex        =   7
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8969
         _Version        =   262144
         TabGuid         =   "FPurchAssist.frx":A29C
         Begin VB.CommandButton cmdThaw 
            Caption         =   "Thaw"
            Height          =   315
            Left            =   6000
            TabIndex        =   19
            Top             =   4560
            Width           =   1035
         End
         Begin VB.Frame Frame1 
            Height          =   855
            Index           =   0
            Left            =   180
            TabIndex        =   15
            Top             =   60
            Width           =   7035
            Begin VB.TextBox txtPONo 
               Height          =   315
               Left            =   1200
               TabIndex        =   17
               Top             =   300
               Width           =   1455
            End
            Begin VB.CommandButton cmdFind 
               Caption         =   "Find"
               Height          =   315
               Left            =   5820
               TabIndex        =   16
               Top             =   300
               Width           =   1035
            End
            Begin VB.Label Label1 
               Caption         =   "PO Number"
               Height          =   255
               Left            =   180
               TabIndex        =   18
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.TextBox txtVendor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtAgent 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1515
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1215
         End
         Begin GridEX20.GridEX gdxMatches 
            Height          =   2535
            Left            =   180
            TabIndex        =   8
            Top             =   1920
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   3
            Column(1)       =   "FPurchAssist.frx":A2C4
            Column(2)       =   "FPurchAssist.frx":A3FC
            Column(3)       =   "FPurchAssist.frx":A520
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPurchAssist.frx":A65C
            FormatStyle(2)  =   "FPurchAssist.frx":A73C
            FormatStyle(3)  =   "FPurchAssist.frx":A874
            FormatStyle(4)  =   "FPurchAssist.frx":A924
            FormatStyle(5)  =   "FPurchAssist.frx":A9D8
            FormatStyle(6)  =   "FPurchAssist.frx":AAB0
            ImageCount      =   0
            PrinterProperties=   "FPurchAssist.frx":AB68
         End
         Begin VB.Label Label2 
            Caption         =   "Vendor"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   1140
            Width           =   675
         End
         Begin VB.Label Label3 
            Caption         =   "Agent"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label4 
            Caption         =   "Date"
            Height          =   255
            Left            =   5280
            TabIndex        =   12
            Top             =   1140
            Width           =   675
         End
      End
   End
   Begin MSComctlLib.ImageList imglRemarks 
      Left            =   9660
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
            Picture         =   "FPurchAssist.frx":AD40
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPurchAssist.frx":B192
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnugdxVendorsPopUp 
      Caption         =   "gdxVendorsPopUp"
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuAutoFit 
         Caption         =   "AutoFit"
      End
      Begin VB.Menu mnuSaveLayout 
         Caption         =   "Save Layout"
      End
   End
End
Attribute VB_Name = "FPurchAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'activetab control tab indices
Private Enum TabMainIndexes
    tmiCreatePONew = 1
    tmiPOOrders = 2
    tmiSPOStatus = 3
    tmiDropShips = 4
    tmiThawPO = 5
    tmiManualFreeze = 6
    tmiExpeditePO = 7
    tmiVendSetup = 8
    tmiVendItems = 9
    tmiVendMaxMin = 10
    tmiGsktStatus = 11
    tmiAnnualSalesAnalysis = 12
End Enum

'increment this number when changes are made to gdxVendorsNew
'Private Const k_iPAVendorGridRev = 1

Private m_lWindowID As Long

Private m_rstVendors As ADODB.Recordset
Private WithEvents m_gwVendors As GridEXWrapper
Attribute m_gwVendors.VB_VarHelpID = -1

Private m_oRstFrozenItems As ADODB.Recordset
Private WithEvents m_gwMatches As GridEXWrapper
Attribute m_gwMatches.VB_VarHelpID = -1

Private m_sPONo As String
Private m_lPOKey As Long

Private WithEvents m_gwSPOStatus As GridEXWrapper
Attribute m_gwSPOStatus.VB_VarHelpID = -1
Private WithEvents m_gwPOOrder As GridEXWrapper
Attribute m_gwPOOrder.VB_VarHelpID = -1
Private WithEvents m_gwMMVendItem As GridEXWrapper
Attribute m_gwMMVendItem.VB_VarHelpID = -1

Private m_bAddToLA As Boolean

' cached info from item table
Private m_lItemKey As Long
Private m_dStdUnitCost As Double
Private m_dRplcmntUnitCost As Double
Private m_varPPLKey As Variant          'allow Null
Private m_lVItemKey As Long
Private m_lVWhseKey As Long
Private m_lMMWhseKey As Long

Private m_bLoadRobertData As Boolean

Private m_bLoad As Boolean
Private m_bDirty As Boolean

Private m_sOldVendComment As String

' cached info from inventory table
Private m_lWhseKey As Long
Private m_APList As AcctPayList
Private m_APListUpdate As AcctPayList

'*** Added 3/9/2011 Len
'*** timVendItem fix
Private m_defaultItemVendKey As Long
'****

'Removed 2/5/04 LR
''This is temporary to support transition. 8/21/03 LR
''Used by MDIMain to communicate which tab to display for Create PO.
'Private m_iVersion As Integer
'
'Public Property Let Version(ByVal lNewValue As Integer)
'    m_iVersion = lNewValue
'End Property


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


Private Sub cmdManageShipments_Click()
    Dim oFrm As New FProvisionalShipment
    oFrm.POKey = m_lPOKey
    oFrm.CreateShipment Me
End Sub




Private Sub cmdTransSumRpt_Click()
    SetWaitCursor True
    
    If cboPOWhse.ItemData(cboPOWhse.ListIndex) = 0 Then
        MsgBox "A warehouse must be selected for this report.", vbInformation, "Report"
    Else
        Dim oFrm As FViewer
        Set oFrm = New FViewer
        Call oFrm.ParamAdd(1, "ShippedFromLocation", cboPOWhse.text)
        Call oFrm.ViewReportByType("Transfer Status Summary")
        Set oFrm = Nothing
    End If
        
    SetWaitCursor False
End Sub

Private Sub Form_Load()
    SetCaption "Purchasing Assistant"
    
    LoadControls
    
    Set m_APList = New AcctPayList
    Set m_APListUpdate = New AcctPayList
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '3/31/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwVendors = Nothing
    Set m_gwMatches = Nothing
    Set m_gwSPOStatus = Nothing
    Set m_gwPOOrder = Nothing
    Set m_gwMMVendItem = Nothing
    
    MDIMain.FormUnregister Me
    
    Set m_APList = Nothing
    Set m_APListUpdate = Nothing
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    tabMain.Height = Me.Height - 540
    tabMain.width = Me.width - 240
    
    'New Create PO tab 8/21/03 LR
    'PRN #589 - resize logic for new Branch Transfer frame
    gdxVendorsNew.width = tabMain.width - 225
    'gdxVendorsNew.Height = tabMain.Height - 900
    gdxVendorsNew.Height = tabMain.Height - 1900
    gdxVendorsNew.Refresh
    
    'PRN #589 - resize logic for new Branch Transfer frame
    fraBranchTransfer.width = tabMain.width - 225
    fraBranchTransfer.Top = gdxVendorsNew.Height + 600
    
    gdxSPOStatus.width = tabMain.width - 240
    gdxSPOStatus.Height = tabMain.Height - 1860
    
    tvDSPO.width = gdxSPOStatus.width
    tvDSPO.Height = gdxSPOStatus.Height

    gdxPOLines.width = tabMain.width - 240
    gdxPOLines.Height = tabMain.Height - 3100
'    rvPO.Top = gdxPOLines.Top + gdxPOLines.Height + 120
'    Label25.Top = rvPO.Top
'    calExpectedDate.Top = Label25.Top + Label25.Height + 60
'    cmdSave.Top = calExpectedDate.Top
    
    gdxMatches.Height = tabMain.Height - 2940
    cmdThaw.Top = gdxMatches.Top + gdxMatches.Height + 120
    
    gdxVendItem.width = tabMain.width - 240
    gdxVendItem.Height = tabMain.Height - 2940
    
    cmdVICreate.Top = gdxVendItem.Top + gdxVendItem.Height + 120
    cboVIVendor(1).Top = cmdVICreate.Top
    Label22.Top = cmdVICreate.Top
    chkVISelectAll.Top = cmdVICreate.Top
    chkVIUnselectAll.Top = cmdVICreate.Top
    
    gdxMMVendItem.width = tabMain.width - 240
    gdxMMVendItem.Height = tabMain.Height - 2220
    cmdMMUpdate.Top = gdxMMVendItem.Top + gdxMMVendItem.Height + 120
    
'    cmdVICreate.Left = gdxVendItem.Left + gdxVendItem.Width - cmdVICreate.Width
'    cboVIVendor(1).Left = cmdVICreate.Left - 120 - cboVIVendor(1).Width
'    Label22.Left = cboVIVendor(1).Left - 120 - Label22.Width
    
    tvDSPO.Refresh
    gdxPOLines.Refresh
    gdxSPOStatus.Refresh
    gdxVendItem.Refetch
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Public Sub DoShowHelp()
    ShowHelp "FPurchAssist"
End Sub


Private Sub LoadPOBuyerCbo()
    Dim lBuyerKey As Long
        
    cboPOWhse.Clear
    If cboBuyers(2).ItemData(cboBuyers(2).ListIndex) = 0 Then
        cmdRefresh(2).Enabled = False
        cboPOWhse.Enabled = False
        cmdTransSumRpt.Enabled = False
    Else
        lBuyerKey = cboBuyers(2).ItemData(cboBuyers(2).ListIndex)
        Dim orst As ADODB.Recordset
        Set orst = CallSP("spcpcPOGetBuyerWhses", "@i_BuyerKey", lBuyerKey)

        'load whse cbo box
        While Not orst.EOF
            cboPOWhse.AddItem WhseKeyToID(orst.Fields("WhseKey").Value)
            cboPOWhse.ItemData(cboPOWhse.NewIndex) = orst.Fields("WhseKey").Value
            orst.MoveNext
        Wend


        Select Case orst.RecordCount
            Case 0
                'MsgBox "Please set up a warehouse for buyer " & lBuyerKey & " (" & cboBuyers(2).Text & ").", vbInformation, "Buyer"
                MsgBox cboBuyers(2).text & " has no vendors set up in SageAssistant.", vbInformation, "Buyer"
                cboBuyers(2).text = "<none>"
                cmdRefresh(2).Enabled = False
                cboPOWhse.Enabled = False
                cmdTransSumRpt.Enabled = False
                'SMR - clear cboVendors & gdxVendorsNew grid (Columns.Clear, research)
                cboVendors.Clear
                Exit Sub
            Case 1
                cboPOWhse.Enabled = False
                cboPOWhse.text = GetUserWhseID(GetUserKey(cboBuyers(2).text))
                cmdRefresh(2).Enabled = True
                cmdTransSumRpt.Enabled = True
            Case Is > 1
                cboPOWhse.Enabled = True
                cboPOWhse.text = GetUserWhseID(GetUserKey(cboBuyers(2).text))
                cmdRefresh(2).Enabled = True
                cmdTransSumRpt.Enabled = True
        End Select
        Set orst = Nothing
        'cboPOWhse.ListIndex = 0
    End If

End Sub

Private Sub LoadControls()

    m_bLoad = True
    
    LoadCombo cboBuyers(1), g_rstBuyers, "BuyerID", "BuyerKey"
    SetComboByKey cboBuyers(1), UserNameToBuyerKey(GetUserName)
        
    LoadCombo cboBuyers(2), g_rstBuyers, "BuyerID", "BuyerKey", , 1
    SetComboByKey cboBuyers(2), UserNameToBuyerKey(GetUserName)
    
    LoadCombo cboDropShipBuyers, g_rstBuyers, "BuyerID", "BuyerKey"
    SetComboByKey cboDropShipBuyers, UserNameToBuyerKey(GetUserName)
    
    LoadPOBuyerCbo
    
    Set m_gwVendors = New GridEXWrapper
    m_gwVendors.Grid = gdxVendorsNew

    m_gwVendors.InitGridLayout GetUserKey, g_PAVendorGridRev
    
    LoadVendorCombo
    LoadCombo cboVIVendor(1), g_rstVendors, "VendName", "VendKey", , True
    
    Set m_gwMatches = New GridEXWrapper
    m_gwMatches.Grid = gdxMatches

    Set m_gwSPOStatus = New GridEXWrapper
    m_gwSPOStatus.Grid = gdxSPOStatus
    
    Set m_gwPOOrder = New GridEXWrapper
    m_gwPOOrder.Grid = gdxPOs
    
    'load icons into the grid
    LoadImageList imglRemarks, gdxSPOStatus
    LoadImageList imglRemarks, gdxPOs

    'SOTA controls on SPO Status tab
    calStart.Value = Date - 14
    calEnd.Value = Date
    
    'tab(tmiManualFreeze) is always visible
    If HasRight(k_sRightPurchasing) Then
        tabMain.Tabs(tmiCreatePONew).Visible = True
        tabMain.Tabs(tmiCreatePONew).Selected = True
        tabMain.Tabs(tmiThawPO).Visible = True
'Removed 2/5/04 LR
'        'temporary 8/21/03 LR
'        If m_iVersion = 1 Then
'            tabMain.Tabs(tmiCreatePONew).Visible = False
'        ElseIf m_iVersion = 2 Then
'            tabMain.Tabs(tmiCreatePONew).Visible = True
'            tabMain.Tabs(tmiCreatePONew).Selected = True
'        End If
    End If
    
    'We will set more exact tab rights for purchasing tab in the future
'    LoadCombo cboSPOWhse, g_rstWhses, "WhseID", "WhseKey", GetUserWhseKey
    
    '*** changed rst & added filter 8/26/03 LR
    g_rstWhses.Filter = "transit = 0"
    LoadCombo cboMMWarehouse, g_rstWhses, "WhseID", "WhseKey"
    cboMMWarehouse.ListIndex = 0
    LoadCombo cboWhse_GaskOrdStatus, g_rstWhses, "WhseID", "WhseKey"
    g_rstWhses.Filter = adFilterNone
    
    Set m_gwMMVendItem = New GridEXWrapper
    m_gwMMVendItem.Grid = gdxMMVendItem

    LoadCombo cboPOBuyer, g_rstBuyers, "BuyerID", "BuyerKey"
    cboPOBuyer.AddItem "<Any>", 0
    SetComboByKey cboPOBuyer, UserNameToBuyerKey(GetUserName)
    SetComboByText cboPOStatus, "<Any>"
    
    m_bDirty = False
    
    m_bLoad = False
End Sub


Private Sub gdxSalesAnalysis_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Dim colCurrent As JSColumn
    If gdxSalesAnalysis.EditMode = jgexEditModeOff And gdxSalesAnalysis.col <> 0 Then
        'enter edit mode
        gdxSalesAnalysis.EditMode = jgexEditModeOn
        Set colCurrent = gdxSalesAnalysis.Columns.ItemByPosition(gdxSalesAnalysis.col)
        
        'select all the text in the cell
        gdxSalesAnalysis.SelStart = 0
        gdxSalesAnalysis.SelLength = Len(gdxSalesAnalysis.Value(colCurrent.Index))
    End If
End Sub



'**********************************************************************************
'   The Main Tab Control
'**********************************************************************************

Private Sub tabMain_TabClick(ByVal NewTab As ActiveTabs.SSTab)
    Select Case NewTab.Index
                                                                           
        Case tmiManualFreeze
            TryToSetFocus txtSOID
            cmdLoad(0).Default = True
            cmdFreeze.Enabled = False
            
        Case tmiExpeditePO
            cmdLoad(1).Default = True
            TryToSetFocus txtPOID(1)
        
        Case tmiVendSetup
            TryToSetFocus txtItemID
            g_rstWhses.Filter = "transit = 0"
            LoadCombo cboVendWhse, g_rstWhses, "WhseID", "WhseKey"
            g_rstWhses.Filter = adFilterNone
            cmdVendUpdate.Enabled = False
            
        Case tmiGsktStatus
            g_rstWhses.Filter = "transit = 0"
            LoadCombo cboWhse, g_rstWhses, "WhseID", "WhseKey"
            g_rstWhses.Filter = adFilterNone
            txtCutAmt.text = ""
            txtMoldAmt.text = ""
            txtTrimAmt.text = ""
            dtpEndDate.Value = Date
            dtpStartDate.Value = DateAdd("d", -7, Date)
    
        Case tmiAnnualSalesAnalysis
            'PRN 520 LR
            'LoadCombo cboWhse_SalesAnalysis, g_rstWhses, "WhseID", "WhseKey"
            'filter out transit warehouses
            g_rstWhses.Filter = "transit = 0"
            LoadCombo cboWhse_SalesAnalysis, g_rstWhses, "WhseID", "WhseKey"
            'PRN 526 SMR - added 2 lines below to allow search by all warehouses
            cboWhse_SalesAnalysis.AddItem "<ALL>", 0
            cboWhse_SalesAnalysis.ListIndex = 0
            g_rstWhses.Filter = adFilterNone
            
            LoadCombo cboVend_SalesAnalysis, g_rstVendors, "VendName", "VendKey"
            cboVend_SalesAnalysis.AddItem "<ALL>", 0
            cboVend_SalesAnalysis.ListIndex = 0
    End Select
End Sub


'**********************************************************************************
'   Control Array Handlers
'**********************************************************************************
        
Private Sub cmdRefresh_Click(Index As Integer)
    Dim orst As ADODB.Recordset
    Dim lBuyerKey As Long
    Dim lWhseKey As Long
    Dim i As Integer
    
    Select Case Index
                                                                    
    Case 1  'SPO Status tab
       
        SetWaitCursor True
                        
        If optNotOrdered.Value = True Then
            Set orst = database.CallSP("spcpcGetSPOStatus", "@_iNotFrozen", 1, "@_iBuyerKey", cboBuyers(1).ItemData(cboBuyers(1).ListIndex))
        Else
            Set orst = database.CallSP("spCPCGetSPOPurchStatus", "@_iStartDate", calStart.Value, _
                                "@_iEndDate", calEnd.Value, "@_iBuyerKey", cboBuyers(1).ItemData(cboBuyers(1).ListIndex))
        End If
        
        With gdxSPOStatus
            .HoldFields
            Set .ADORecordset = orst
            
            For i = 2 To .Columns.Count
                .Columns(i).AutoSize
            Next
        End With
        
        Set orst = Nothing
        SetWaitCursor False
        
    Case 2  ' New Create PO tab 8/21/03 LR
        
        SetWaitCursor True
        
        On Error GoTo EH
        
        'get the key of the currently selected buyer
        lBuyerKey = cboBuyers(2).ItemData(cboBuyers(2).ListIndex)
        
        lWhseKey = cboPOWhse.ItemData(cboPOWhse.ListIndex)
    
        Set orst = database.CallSP("spcpcPOVendSummary", "@i_BuyerKey", lBuyerKey, "@i_WhseKey", lWhseKey)

        With gdxVendorsNew
            .HoldFields
            .HoldSortSettings = True
            Set .ADORecordset = orst
        End With

        'load the combo based on buyer and what's in the grid
        Set orst = database.CallSP("spcpcPOGetUnspecifiedVendors", "@i_BuyerKey", lBuyerKey, "@i_WhseKey", lWhseKey)
        LoadCombo cboVendors, orst, "VendName", "VendKey"
        Set orst = Nothing
                
        SetWaitCursor False
        
    End Select
    
    Exit Sub
    
EH:
    SetWaitCursor False

'10/29/15 LR
'misleading error reporting
' err.number = 0
' err.description =
'   CallSP spcpcPOVendSummaryTest
'   -2147217871 (Microsoft OLE DB Provider for SQL Server) Query timeout expired
'   Additional Information:
'   ErrMsg 0: Query timeout expired

    'If Err.Number = -2147217871 Then  'ODBC SQL Server Driver  timeout expired
        LogEvent "FPurchAssist", "CreatePO", "Query timed out"
        msg "Your query timed out." & vbCrLf & "Please try again.", vbExclamation, "Create PO Warning"
    '    msg "Your query timed out." & vbCrLf & "Please try again.", vbExclamation, "OrderPad Warning"
    'Else
    '    LogError "FPurchAssist", "CreatePO", "", Err.Source, Err.Number, Err.Description
    '    msg "FOrder.FindOrdersByCriteria" & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "Call the computer guys.", vbCritical, "OrderPad Error"
    'End If

    
End Sub


Private Sub cmdLoad_Click(Index As Integer)
    Dim oRstSO As ADODB.Recordset
    Dim oRstPO As ADODB.Recordset
    Dim oRstPOLine As ADODB.Recordset
    Dim i As Integer
    Dim sSQL As String
    Dim sWarning As String
    
    Select Case Index
                                                                    
        '*** Manual Freeze
        Case 0
    
            If Len(txtSOID.text) <= 0 Then
                msg "You must enter a Sage Sales Order (SO) Number."
                Exit Sub
            ElseIf Not IsNumeric(txtSOID.text) Then
                msg "Invalid SO Number", vbCritical
                Exit Sub
            End If
            If Len(txtPOID(0).text) <= 0 Then
                msg "You must enter a Sage Purchase Order (PO) Number."
                Exit Sub
            ElseIf Not IsNumeric(txtPOID(0).text) Then
                msg "Invalid PO Number", vbCritical
                Exit Sub
            End If
            
            LoadLists
            cmdFreeze.Enabled = True
            
      '*** Expediate PO
      Case 1
            If Len(txtPOID(1).text) <= 0 Then
                ClearScreen
                msg "You must enter a Sage Purchase Order (PO) Number."
                Exit Sub
            ElseIf Not IsNumeric(txtPOID(1).text) Then
                msg "Invalid PO Number", vbCritical
                ClearScreen
                Exit Sub
            End If
            
            Set oRstPO = CallSP("spCPCPOGetExpeditePO", "@i_TranKey", CLng(Trim(txtPOID(1))))

            If oRstPO.EOF Then
                ClearScreen
                msg ("PO " & Trim(txtPOID(1)) & " not found.")
            Else
                With oRstPO
                    txtVendName.text = "[" & Trim(.Fields("VendID").Value) & "] " & .Fields("VendName").Value
                    txtBuyerName.text = .Fields("BuyerName").Value
                    txtTranDate.text = Format(.Fields("TranDate").Value, "m/d/yy")
                    txtContact.text = .Fields("ContactName").Value
                    txtVendPhone.text = FormatPhoneNumber(.Fields("Phone").Value, .Fields("PhoneExt").Value)
                    txtVendFax.text = FormatPhoneNumber(.Fields("Fax").Value, .Fields("FaxExt").Value)
                    txtShipAddress.Caption = CompAddr(.Fields("AddrName").Value, .Fields("AddrLine1").Value, .Fields("AddrLine2").Value, _
                                    .Fields("City").Value, .Fields("StateID").Value, .Fields("PostalCode").Value, .Fields("CountryID").Value)
                                    
                    '*** SMR Added PO Status and Ship Method to the ExpeditePO tab...
                    Select Case oRstPO.Fields("Status").Value
                        Case 0: txtExPOStatus.text = "Unissued"
                        Case 1: txtExPOStatus.text = "Open"
                        Case 2: txtExPOStatus.text = "Inactive"
                        Case 3: txtExPOStatus.text = "Canceled"
                        Case 4: txtExPOStatus.text = "Closed"
                        Case 5: txtExPOStatus.text = "Incomplete"
                        Case 6: txtExPOStatus.text = "Pending Approval"
                    End Select
                    txtExPOShipMth.text = .Fields("ShipMethod").Value
                                    
                    rvPO.Visible = True
                    rvPO.OwnerID = Trim(txtPOID(1).text)
                    m_lPOKey = .Fields("POKey")
                    calExpectedDate.Value = .Fields("RequestDate").Value
                    cmdSave.Enabled = True
                    
                    Set oRstPOLine = CallSP("spCPCPOGetPOOrderLine", "@i_POKey", .Fields("POKey").Value)
                End With
                
                With gdxPOLines
                    .HoldFields
                    .HoldSortSettings = True
                    Set .ADORecordset = oRstPOLine
                    
                    For i = 1 To .Columns.Count
                        .Columns(i).AutoSize
                    Next
                End With
                
'                AttachGrid gdxPOLines, m_oRstPOLines
            End If
 
        '*** New Create PO
        'Use this Vendor on the new Create PO tab 8/21/03 LR
        'Create a POWiz for a vendor who's not on the Red/Orange/Yellow list
        Case 2
            Dim lVendKey As Long
            Dim lBuyerKey As Long
            Dim lWhseKey As Long
            Dim sUserID As String
            Dim oFrm As FPOWiz
        
            'check to make sure that the combo has been loaded
            If cboVendors.ListCount > 0 Then
                sUserID = BuyerKeyToUserID(cboBuyers(2).ItemData(cboBuyers(2).ListIndex))
                lVendKey = cboVendors.ItemData(cboVendors.ListIndex)
                lWhseKey = cboPOWhse.ItemData(cboPOWhse.ListIndex)
        
                Set oFrm = New FPOWiz
                'Call ofrm.LoadVendor(lVendKey, sUserID, lWhseKey, True)
                'the last parameter is set to true, because we are loading POWiz by Vendor
                Call oFrm.Init(lVendKey, sUserID, lWhseKey, True)
            Else
                msg "You first need to refresh your vendor list."
            End If

    End Select
End Sub


Private Sub cboBuyers_Click(Index As Integer)
    If Index = 2 Then
        LoadPOBuyerCbo
    End If
End Sub


Private Sub cboBuyers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 2 Then
        If KeyCode = vbKeyReturn And cmdRefresh(2).Enabled = True Then
            cmdRefresh_Click (2)
        End If
    End If
End Sub



'*************************************************************************
' Tab(1) Create POs
'*************************************************************************

Private Sub CheckForSavedWork(sUserID As String)
    Dim orst As ADODB.Recordset

    'Test to see if entries already exist in tcpPOProformaLn
'SQL:
    Set orst = LoadDiscRst("SELECT CreateDate FROM tcpPOProformaLn WHERE vendkey=" & CStr(m_gwVendors.Value("VendKey")) & " AND UserID = '" & sUserID & "'")
    
    If Not orst.EOF Then
        'If entries exist, determine if user wants to start over for this vendor/user
        If vbYes = msg("Do you want to erase your work from " & orst.Fields("CreateDate").Value & " on this vendor?", vbYesNo) Then
            LoadLineItems
            'ToDo: If there are any open POMaker windows for this, they need to be destroyed
        End If
    Else
        LoadLineItems
    End If
    Set orst = Nothing
End Sub


'*************************************************************************************************
' Tab(2)  New CreatePOs tab
'*************************************************************************************************

Private Function OKtoEdit(sUserID As String) As Boolean
    Dim orst As ADODB.Recordset
    Dim sMsg As String
    Dim iResponse As Integer

    'Test to see if entries already exist in tcpPOProformaLn
'SQL:
    'PRN #541 smr - 2/22/05
    'Set oRst = LoadDiscRst("SELECT CreateDate FROM tcpPOProformaLn WHERE vendkey=" & CStr(m_gwVendors.value("VendKey")) & " AND UserID = '" & sUserID & "'")
    Set orst = LoadDiscRst("SELECT CreateDate FROM tcpPOProformaLn WHERE vendkey=" & CStr(m_gwVendors.Value("VendKey")) & " AND UserID = '" & sUserID & "' And DirtyData = 1")
    
''    'SMR - I was going to just get the count, but I need the create date for the msg box.
''    Set oRst = LoadDiscRst("SELECT count(*) as Count FROM tcpPOProformaLn WHERE vendkey=" & CStr(m_gwVendors.value("VendKey")) & " AND UserID = '" & sUserID & "' And DirtyData = 1")
''    If oRst.Fields("Count") > 0 Then
    
    If Not orst.EOF Then
        'If entries exist, determine if user wants to start over for this vendor/user
            '(only if find we have saved data)
        sMsg = "You have saved work from " & Format(orst.Fields("CreateDate").Value, "Long Date") & "." & vbCrLf _
            & "Click YES to use it," & vbCrLf & "NO to erase your work and start over," & vbCrLf & "CANCEL to bail out."
        
        iResponse = msg(sMsg, vbYesNoCancel, "Create PO")
        
        Select Case iResponse
            Case vbYes:
                OKtoEdit = True
            Case vbNo:
                LoadLineItems
                OKtoEdit = True
                'ToDo: If there are any open POWizard windows for this, they need to be destroyed
            Case vbCancel:
                OKtoEdit = False
        End Select
    Else
        LoadLineItems
        OKtoEdit = True
    End If
    Set orst = Nothing
End Function


Private Sub gdxVendorsNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnugdxVendorsPopUp
    End If
End Sub


Private Sub mnuFont_click()
    ChangeGridFont gdxVendorsNew
End Sub


Private Sub mnuAutoFit_Click()
'12/2004 - smr
    m_gwVendors.GridAutoFit
End Sub


Private Sub mnuSaveLayout_Click()
'12/2004 - smr
    SetWaitCursor True
    Call m_gwVendors.GridSaveLayout(GetUserKey)
    SetWaitCursor False
End Sub


Private Sub cmdOpenTransRpt_Click()
    SetWaitCursor True
    
    'smr - 01/05 - New logic to call FViewer - uses: "Crystal Reports 8.5 ActiveX Designer Run Time Library"
    Dim oFrm As FViewer
    Set oFrm = New FViewer
    Call oFrm.ViewReportByType("Transfer Status")
    Set oFrm = Nothing
    
    SetWaitCursor False
End Sub


'**********************************************************************************
' Used by both Create PO tabs
'**********************************************************************************

'Called by
'   CheckForSavedWork()     Version 1
'   OKtoEdit()              Version 2
'
'The effect of this routine is to update tcpPOProformaLn(Test)

Private Sub LoadLineItems()
    Dim oCmd As ADODB.Command
    Dim bLoad As Boolean

    SetWaitCursor True
    bLoad = m_bLoad
    m_bLoad = True

    Set oCmd = CreateCommandSP("spcpcPOProforma")
    With oCmd
        .Parameters("@i_BuyerKey") = cboBuyers(2).ItemData(cboBuyers(2).ListIndex)
        .Parameters("@i_VendKey") = m_gwVendors.Value("VendKey")
        .Parameters("@i_UserID") = Trim(BuyerKeyToUserID(cboBuyers(2).ItemData(cboBuyers(2).ListIndex)))
        .Execute
    End With
    
    SetWaitCursor False
    m_bLoad = bLoad
End Sub


Private Sub m_gwVendors_RowChosen()
    Dim oFrm As Form
    Dim sUserID As String
    
    On Error GoTo ErrorHandler
    
    SetWaitCursor True
    
    sUserID = BuyerKeyToUserID(cboBuyers(2).ItemData(cboBuyers(2).ListIndex))
    
    If OKtoEdit(sUserID) Then
        Set oFrm = New FPOWiz
        Call oFrm.Init(m_gwVendors.Value("VendKey"), sUserID, cboPOWhse.ItemData(cboPOWhse.ListIndex), False)
    End If

    SetWaitCursor False
    Exit Sub
    
ErrorHandler:
    ClearWaitCursor
    msg "Can't show POWizard. The error is " & Err.Number & " - " & Err.Description, vbOKOnly + vbExclamation, Err.Source
End Sub


'**********************************************************************************
' Tab(3)  PO Orders
'**********************************************************************************

Private Sub cmdLoadPO_Click()
    If IsEmpty(m_gwPOOrder.Value("PONumber")) Then Exit Sub
    
    If tabMain.Tabs(tmiExpeditePO).Visible = False Then
        msg "Sorry. You don't have right to load this PO in Expedite PO tab." & vbCrLf & _
            "Please contact IT department for more details.", vbOKOnly + vbExclamation, "Can't Load PO"
    Else
        tabMain.Tabs(tmiExpeditePO).Selected = True
        txtPOID(1).text = m_gwPOOrder.Value("PONumber")
        cmdLoad_Click 1
    End If
    
End Sub


Private Sub cboPOBuyer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFindPO_Click
    End If
End Sub


Private Sub cboPOStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFindPO_Click
    End If
End Sub


Private Sub cmdFindAssociatedOrder_Click()
    Dim orst As ADODB.Recordset

    SetWaitCursor True
                
    If Len(Trim$(txtLookupPONo.text)) = 0 Or Not IsNumeric(Trim$(txtLookupPONo.text)) Then
        MsgBox "Requires a numeric PO number", vbExclamation + vbOKOnly, "Incorrect Parameter"
        Exit Sub
    End If
    
    Set orst = CallSP("spCPCGetPOCustomer", "@PONbr", Trim$(txtLookupPONo.text))
    
    gdxPOs.Visible = False
    gdxPOOrder.Visible = True
    
    SetWaitCursor False
                    
    With gdxPOOrder
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
    End With
    
    Set orst = Nothing
End Sub


Private Sub cmdFindPO_Click()
    Dim cmd As ADODB.Command
    Dim orst As ADODB.Recordset

    SetWaitCursor True
    
    Set cmd = CreateCommandSP("spcpcPOGetPOOrders")
    
    With cmd
        If cboPOBuyer.text <> "<Any>" Then
            .Parameters("@_iBuyerKey") = cboPOBuyer.ItemData(cboPOBuyer.ListIndex)
        End If
        
        If cboPOStatus.text <> "<Any>" Then
            .Parameters("@_iStatus") = cboPOStatus.ItemData(cboPOStatus.ListIndex)
        End If
    End With
    
    Set orst = New ADODB.Recordset
    orst.Open cmd, , , adLockReadOnly
    
    gdxPOOrder.Visible = False
    gdxPOs.Visible = True
    
    SetWaitCursor False
    
    With gdxPOs
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
    End With
    
    lblPOCount.Caption = orst.RecordCount
    
    Set orst = Nothing
    Set cmd = Nothing
End Sub


Private Sub gdxPOs_DblClick()
    EditPORemarks m_gwPOOrder.Value("TranID")
End Sub


Private Sub gdxPOs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        EditPORemarks m_gwPOOrder.Value("TranID")
    End If
End Sub


Private Sub EditPORemarks(ByVal POKey As String)
    If POKey = "" Then Exit Sub
    
    Dim oRC As RemarkContext
    
    Set oRC = New RemarkContext
    oRC.Edit "ViewPO", POKey
End Sub


'*************************************************************************
' Tab(4) SPO Status
'*************************************************************************

Private Sub optAll_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRefresh_Click 1
    End If
End Sub


Private Sub optDropShips_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRefresh_Click 1
    End If
End Sub


Private Sub optNotOrdered_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRefresh_Click 1
    End If
End Sub


Private Sub calEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRefresh_Click 1
    End If
End Sub


Private Sub calStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRefresh_Click 1
    End If
End Sub


Private Sub gdxSPOStatus_DblClick()
    EditRemarks m_gwSPOStatus.Value("OPLineKey")
End Sub


Private Sub gdxSPOStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        EditRemarks m_gwSPOStatus.Value("OPLineKey")
    End If
End Sub


Private Sub EditRemarks(ByVal OPLineKey As Long)
    Dim oRC As RemarkContext
    
    Set oRC = New RemarkContext
    oRC.Edit "ViewOrderLine", OPLineKey
End Sub


Private Sub cmdPrintSPO_Click()
    'PRN#64
    gdxSPOStatus.PrintGrid True
End Sub

'*************************************************************************
' Tab(x) Unissued Drop Ships
'*************************************************************************

Private Sub cmdRefreshDropShips_Click()
    Dim orst As ADODB.Recordset
    SetWaitCursor True

    Set orst = CallSP("spcpcGetDropShipPOUnissued", "@_iBuyerKey", cboDropShipBuyers.ItemData(cboDropShipBuyers.ListIndex))
    LoadTree orst

    Set orst = Nothing
    SetWaitCursor False
End Sub


Private Sub LoadTree(orst As ADODB.Recordset)
    Dim sHoldNode As String
    Dim sPOKeyVal As String
    Dim sOPKeyVal As String
    Dim sTextVal As String
    Dim oNode As Node
    Dim sVendorInfo As String

    tvDSPO.Style = tvwPictureText
    tvDSPO.ImageList = imglRemarks
    sHoldNode = ""
    With tvDSPO.Nodes
        .Clear
        .Add , , "root", "Unissued DropShip POs"
        
        Do While Not orst.EOF
            If orst.Fields("POID").Value <> sHoldNode Then
                sHoldNode = orst.Fields("POID").Value
                sPOKeyVal = "PO-" & StripLeadingZeros(sHoldNode)
                
                'if buyerkey = 9 ("unassigned") then PO has not been fixed. mark it.
                sTextVal = IIf(orst.Fields("BuyerKey").Value = 9, sPOKeyVal & "*", sPOKeyVal)
                
                Set oNode = .Add("root", tvwChild, sPOKeyVal, sTextVal)
                oNode.EnsureVisible
            End If
            
            sOPKeyVal = "OP-" & orst.Fields("OPLineKey").Value
            
            'sVendorInfo = Trim(orst.Fields("VendID").Value) & " " & Trim(orst.Fields("VendName")) & "][" & FormatCurrency(orst.Fields("UnitCost"))
            sVendorInfo = Trim(orst.Fields("VendName")) & "][" & FormatCurrency(orst.Fields("UnitCost"))
            Set oNode = .Add(sPOKeyVal, tvwChild, sOPKeyVal, "[OP" & StripLeadingZeros(orst.Fields("OPID").Value) & "/SO" & StripLeadingZeros(orst.Fields("SOID").Value) & "][" & orst.Fields("Description").Value & "][" & sVendorInfo & "][" & orst.Fields("ShipMethID") & "]", IIf(orst.Fields("Remarks") = 0, 1, 2))
            oNode.EnsureVisible
            orst.MoveNext
        Loop
    
    End With

End Sub


'prevent double-click from collapsing the expanded tree
Private Sub tvDSPO_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Expanded = True
End Sub


Private Sub tvDSPO_DblClick()
    Dim orst As ADODB.Recordset
    
    'if a PO node with an asterisk is double-clicked, run fixer
    
    If InStr(1, tvDSPO.SelectedItem.Key, "PO") And tvDSPO.SelectedItem.Key <> "root" Then
        If InStr(1, tvDSPO.SelectedItem.text, "*") Then
            Dim frmFixPO As FFixPO
            Dim sPOID As String
            sPOID = tvDSPO.SelectedItem.Key
            
            Set frmFixPO = New FFixPO
            frmFixPO.FixPO Mid$(sPOID, InStr(1, sPOID, "PO-") + 3)    'strip off the alpha prefix required by the TreeView
            If frmFixPO.Fixed Then
                'refresh the TreeView to clear the asterisk
                SetWaitCursor True
                Set orst = CallSP("spcpcGetDropShipPOUnissued", "@_iBuyerKey", cboDropShipBuyers.ItemData(cboDropShipBuyers.ListIndex))
                LoadTree orst
                SetWaitCursor False
            End If
            Set frmFixPO = Nothing
        End If

    'if a Line Item node is double-clicked, run MM editor
    ElseIf InStr(1, tvDSPO.SelectedItem.Key, "OP") Then
        Dim sOPLineKey As String
        sOPLineKey = tvDSPO.SelectedItem.Key
        EditRemarks Mid$(sOPLineKey, InStr(1, sOPLineKey, "OP-") + 3)    'strip off the alpha prefix required by the TreeView
    End If

End Sub

'*************************************************************************
' Tab(5) Thaw PO
'*************************************************************************

Private Sub cmdFind_Click()
    If Not IsNumeric(Trim(txtPONo.text)) Then
        msg "Invalid PO Number", vbCritical
        Exit Sub
    End If
    m_sPONo = Trim(txtPONo.text)
    LoadFrozenList
End Sub

Private Sub cmdThaw_Click()
    Dim oCmd As ADODB.Command
    Dim orst As ADODB.Recordset
    
    If IsEmpty(m_gwMatches.Value("POLineKey")) Then Exit Sub
'SQL:
    Set orst = LoadDiscRst("select tcpPOFreeze.*, tsoSOLine.Description as SODescr, tpoPOLine.Description as PODescr, tapVendor.VendName " _
                        & "From tcpPOFreeze inner join tsoSOLine on tsoSOLine.SOLineKey = tcpPOFreeze.SOLineKey " _
                        & "inner join tpoPOLine on tcpPOFreeze.POLineKey = tpoPOLine.POLIneKey inner join " _
                        & "tpoPurchOrder on tpoPurchOrder.POKey = tpoPOLine.POKey inner join tapVendor " _
                        & "on tapVendor.VendKey = tpoPurchOrder.VendKey where tpoPOLine.POLineKey = " & m_gwMatches.Value("POLineKey"))
    ' remove item from Freeze table
    Set oCmd = CreateCommandSP("DELETE tcpPOFreeze WHERE POLineKey=" & m_gwMatches.Value("POLineKey"), adCmdText)
    oCmd.Execute
    Set oCmd = Nothing
    
    
    While Not orst.EOF

        LogDB.LogOAEvent "Thaw PO", GetUserID, orst.Fields("POLineKey").Value, orst.Fields("SOLineKey").Value, , _
                    "Thaw item - " & Trim(orst.Fields("SODescr").Value) & " from PO " & Trim(orst.Fields("PODescr").Value) & ". " _
                    & "The vendor is " & orst.Fields("VendName").Value
        
        'LogDB.LogActivity "SA",
                            
        orst.MoveNext
    Wend
    
    RefreshFrozenList
End Sub


Private Sub LoadFrozenList()
    Dim oCmd As ADODB.Command
    Dim orst As ADODB.Recordset
    
    SetWaitCursor True

    'check to ensure the PO exists and that it's not a dropship
'SQL:
    Set orst = LoadDiscRst("SELECT POKey, DfltDropShip FROM tpoPurchOrder WHERE TranNo like '%" & Trim(m_sPONo) & "'")
    If orst.EOF Then
        SetWaitCursor False
        msg "PO " & m_sPONo & " does not exist.", vbCritical, "Purchasing"
        Exit Sub
    ElseIf orst.Fields("DfltDropShip") = 1 Then
        SetWaitCursor False
        msg "PO " & m_sPONo & " is a dropship. It has no frozen line items.", vbInformation, "Purchasing"
        Exit Sub
    End If

    Set orst = CallSP("spOPPOGetFrozenItems", "@PONo", "%" & m_sPONo)
    
    With orst
        If .EOF Then
            SetWaitCursor False
            msg "Couldn't find any Sales Order line items frozen to PO " & txtPONo.text, vbExclamation, "Purchasing"
            Exit Sub
        Else
            txtVendor.text = .Fields("VendName").Value
            txtAgent.text = .Fields("BuyerID").Value
            txtDate.text = .Fields("TranDate").Value
        End If
    End With
    
    With gdxMatches
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
    End With

    SetWaitCursor False

End Sub


Private Sub RefreshFrozenList()
    Dim oCmd As ADODB.Command
    Dim orst As ADODB.Recordset
    
    SetWaitCursor True

    Set orst = CallSP("spOPPOGetFrozenItems", "@PONo", "%" & m_sPONo)
       
    With gdxMatches
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
    End With

    SetWaitCursor False
End Sub


'*************************************************************************
' Tab(6) Manual Freeze
'*************************************************************************

Private Sub cmdFreeze_Click()
    Dim lSOLineKey As Long
    Dim lPOLineKey As Long
    Dim orst As ADODB.Recordset
    Dim oCmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo EH
    
    lSOLineKey = lstSO.ItemData(lstSO.ListIndex)
    lPOLineKey = lstPO.ItemData(lstPO.ListIndex)

    'check to see if they are frozen yet
'SQL:
    Set orst = LoadDiscRst("SELECT TranNo FROM tpoPOLine INNER JOIN tcpPOFreeze ON tpoPOLine.POLineKey = tcpPOFreeze.POLineKey " & _
        "INNER JOIN tpoPurchOrder ON tpoPOLine.POKey = tpoPurchOrder.POKey " & _
        "WHERE SOLineKey=" & lSOLineKey)
    
    
    If Not orst.EOF Then
    
        'Msg "The Sales Order Line Item has already been frozen to an item on PO " & StripLeadingZeros(oRst.Fields("TranNo"))
        lblWarning = "SO line item " & lstSO.ListIndex + 1 & " has already been frozen to an item on PO " & StripLeadingZeros(orst.Fields("TranNo"))
    Else
        'freeze it
        Set oCmd = CreateCommandSP("spCPCInsertPOFreeze")
        oCmd.Parameters("@_iSOLineKey").Value = lSOLineKey
        oCmd.Parameters("@_iPOLineKey").Value = lPOLineKey
        'PRN#96
        oCmd.Execute
        Set oCmd = Nothing
'SQL:
        Set rst = LoadDiscRst("select tapVendor.VendName " _
               & "From tpoPOLine inner join tpoPurchOrder on tpoPurchOrder.POKey = tpoPOLine.POKey inner join tapVendor " _
               & "on tapVendor.VendKey = tpoPurchOrder.VendKey where tpoPOLine.POLineKey = " & lPOLineKey)
               
        LogDB.LogOAEvent _
            "Manual Freeze", _
            GetUserID, lPOLineKey, lSOLineKey, , _
            "Freeze item - " & lstSO.list(lstSO.ListIndex) & " on PO line " & lstPO.ItemData(lstPO.ListIndex) & " - " & lstPO.list(lstPO.ListIndex) & ". The vendor is " & rst.Fields("VendName").Value
            
        'LogDB.LogActivity "SA",
        
    End If

    LoadLists
    
    Set orst = Nothing
    Exit Sub

EH:
    If Err.Number = 381 Then     'no list item selected
        msg "To freeze an SO line to a PO line, select a matching item from each list."
    Else
        msg Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description
    End If
   
End Sub


Private Sub txtItemID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdVendFind_Click
    End If
End Sub


Private Sub txtPOID_GotFocus(Index As Integer)
    If Index = 0 Then
        txtPOID(0).SelStart = 0
        txtPOID(0).SelLength = Len(txtPOID(0).text)
    End If
End Sub


Private Sub txtPOID_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdLoad_Click Index
    End If
End Sub

Private Sub txtPONo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFind_Click
    End If
End Sub


Private Sub txtSOID_GotFocus()
    txtSOID.SelStart = 0
    txtSOID.SelLength = Len(txtSOID.text)
End Sub


Private Sub LoadLists()
    Dim oRstSO As ADODB.Recordset
    Dim oRstPO As ADODB.Recordset
    Dim sSQL As String
    Dim sWarning As String
    
    lblWarning.Caption = vbNullString

'    sSQL = "SELECT tsoSOLine.Description, tsoSOLine.SOLineKey " & _
'    "FROM  tsoSalesOrder INNER JOIN tsoSOLine ON tsoSalesOrder.SOKey = tsoSOLine.SOKey " & _
'    "LEFT OUTER JOIN tcpPOFreeze ON tsoSOLine.SOLineKey = tcpPOFreeze.SOLineKey " & _
'    "WHERE (tsoSalesOrder.TranNo LIKE '%" & Trim(txtSOID) & "') AND (tcpPOFreeze.SOLineKey IS NULL)"

    'get the unfrozen SO line items (include Customer name for reference)
'SQL:
    sSQL = "SELECT dbo.tsoSOLine.Description, dbo.tsoSOLine.SOLineKey, dbo.tarCustomer.CustName " _
        & "FROM  dbo.tsoSalesOrder INNER JOIN dbo.tsoSOLine ON dbo.tsoSalesOrder.SOKey = dbo.tsoSOLine.SOKey " _
        & "INNER JOIN dbo.tarCustomer ON dbo.tsoSalesOrder.CustKey = dbo.tarCustomer.CustKey LEFT OUTER JOIN " _
        & "dbo.tcpPOFreeze ON dbo.tsoSOLine.SOLineKey = dbo.tcpPOFreeze.SOLineKey " _
        & "WHERE (dbo.tsoSalesOrder.TranNo LIKE '%" & Trim(txtSOID) & "') AND (tcpPOFreeze.SOLineKey IS NULL)"

    Set oRstSO = LoadDiscRst(sSQL)

    If oRstSO.RecordCount = 0 Then
        sWarning = sWarning & "There are no unfrozen items on this SO. "
    Else
        lblCustName.Caption = oRstSO.Fields("CustName").Value
    End If

'    sSQL = "SELECT tpoPOLine.Description, tpoPOLine.POLineKey " & _
'    "FROM tpoPurchOrder INNER JOIN tpoPOLine ON tpoPurchOrder.POKey = tpoPOLine.POKey " & _
'    " LEFT OUTER JOIN tcpPOFreeze ON tpoPOLine.POLineKey = tcpPOFreeze.POLineKey " & _
'    "WHERE (tpoPurchOrder.TranNo LIKE '%" & Trim(txtPOID(0)) & "') AND (tcpPOFreeze.POLineKey IS NULL)"

    'get the unfrozen PO line items (include Vendor name for reference)
'SQL:
    sSQL = "SELECT dbo.tpoPOLine.Description, dbo.tpoPOLine.POLineKey, dbo.tapVendor.VendName " _
        & "FROM dbo.tpoPurchOrder INNER JOIN dbo.tpoPOLine ON dbo.tpoPurchOrder.POKey = dbo.tpoPOLine.POKey " _
        & "INNER JOIN dbo.tapVendor ON dbo.tpoPurchOrder.VendKey = dbo.tapVendor.VendKey LEFT OUTER JOIN " _
        & "dbo.tcpPOFreeze ON dbo.tpoPOLine.POLineKey = dbo.tcpPOFreeze.POLineKey " _
        & "WHERE (dbo.tpoPurchOrder.TranNo LIKE '%" & Trim(txtPOID(0)) & "') AND (tcpPOFreeze.POLineKey IS NULL)"

    Set oRstPO = LoadDiscRst(sSQL)

    If oRstPO.RecordCount = 0 Then
        sWarning = sWarning & "There are no unfrozen items on this PO."
    Else
        lblVendorName.Caption = oRstPO.Fields("VendName").Value
    End If
    
    lstSO.Clear
    Do While Not oRstSO.EOF
        lstSO.AddItem oRstSO.Fields("Description").Value
        lstSO.ItemData(lstSO.NewIndex) = oRstSO.Fields("SOLineKey").Value
        oRstSO.MoveNext
    Loop
    
    lstPO.Clear
    Do While Not oRstPO.EOF
        lstPO.AddItem oRstPO.Fields("Description").Value
        lstPO.ItemData(lstPO.NewIndex) = oRstPO.Fields("POLineKey").Value
        oRstPO.MoveNext
    Loop
    
    lblWarning.Caption = sWarning
    Set oRstSO = Nothing
    Set oRstPO = Nothing
End Sub



'*************************************************************************
' Tab(8) Expedite PO
'*************************************************************************

Private Sub ClearScreen()
    txtVendName.text = vbNullString
    txtBuyerName.text = vbNullString
    txtTranDate.text = vbNullString
    txtContact.text = vbNullString
    txtVendPhone.text = vbNullString
    txtVendFax.text = vbNullString
    txtShipAddress.Caption = vbNullString
    rvPO.Visible = False
    calExpectedDate.Value = Null
    cmdSave.Enabled = False
    gdxPOLines.HoldFields
    Set gdxPOLines.ADORecordset = Nothing
    txtPOID(1).SetFocus
    m_lPOKey = 0
End Sub

'See cmdLoad_Click() above.

'Private Sub gdxPOLines_AfterColUpdate(ByVal ColIndex As Integer)
'    m_bDirty = True
'    cmdSave.Enabled = True
'End Sub

Private Sub cmdSave_Click()
    Dim oCmd As ADODB.Command
    
    SetWaitCursor True
    
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = g_DB.Connection
    oCmd.CommandType = adCmdText
    If calExpectedDate.Value = "" Then
'SQL:
        oCmd.CommandText = "UPDATE tpoPurchOrder SET DfltRequestDate = null WHERE POKey = " & m_lPOKey
    Else
'SQL:
        oCmd.CommandText = "UPDATE tpoPurchOrder SET DfltRequestDate = '" & calExpectedDate.Value & "' WHERE POKey = " & m_lPOKey
    End If
    oCmd.Execute
    
    SetWaitCursor False
    Set oCmd = Nothing
End Sub


'*************************************************************************
' Tab(9) Vendor Setup
'*************************************************************************

Private Sub cboVendID_Click()
    If m_bDirty Then Exit Sub
    
    cmdVendUpdate.Enabled = (cboVendID.list(cboVendID.ListIndex) <> "<none>")
    If cboVendID.ListIndex > 0 Then
        g_rstVendors.Filter = "vendkey=" & cboVendID.ItemData(cboVendID.ListIndex)
        g_rstVendors.Filter = adFilterNone
    End If
End Sub


Private Sub cboVendWhse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdVendFind_Click
    End If
End Sub


Private Sub cmdVendFind_Click()
    If Len(txtItemID.text) = 0 Then
        'restore focus to textbox and get out of here
        TryToSetFocus txtItemID
        Exit Sub
    End If
    
    Dim rstItemVend As ADODB.Recordset

    'look for an inventory record for a specified item/warehouse and if found return the itemkey
    Set rstItemVend = CallSP("spCPCGetItemVend", "@_iItemID", Trim(txtItemID.text), "@_iWhseKey", cboVendWhse.ItemData(cboVendWhse.ListIndex))
    
    If rstItemVend.EOF Then
        msg "Sorry. The item doesn't exist in inventory. Please check the item ID.", vbExclamation + vbOKOnly, "Searching Result"
        
        'clear/reset controls
        Dim bDirty As Boolean
        bDirty = m_bDirty
        m_bDirty = True     'is 'dirty' the appropriate descr (really 'altering'?)
        
        cmdVendUpdate.Enabled = False
        cboVendID.Clear
        txtVendComment.text = ""
        
        'give textbox the focus with contents selected
        txtItemID.SelStart = 0
        txtItemID.SelLength = Len(txtItemID.text)
        TryToSetFocus txtItemID
        
        m_bDirty = bDirty
    Else
        'we found the item in the specified warehouse inventory
        'cache these keys
        m_lVItemKey = rstItemVend.Fields("ItemKey").Value
        m_lVWhseKey = cboVendWhse.ItemData(cboVendWhse.ListIndex)
        
        'in effect all we're using in the rst is itemkey
        LoadVendInfo rstItemVend.Fields("ItemKey").Value, rstItemVend.Fields("WhseKey").Value
    End If

    Set rstItemVend = Nothing
End Sub


'uses
'   cmdVendUpdate.Caption
'   m_lVWhseKey,
'   m_lVItemKey,
'   cboVendID,
'   chkObsolete,
'   txtVendComment

'cmdVendUpdate.Caption is initially = 'add'

Private Sub cmdVendUpdate_Click()
    Dim bDirty As Boolean

'*** Added 3/9/2011 Len
'*** timVendItem fix
'*** make sure there is a timVendItem record for the delected vendor
    
    Dim VendOK As Boolean
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spCPCCheckItemVendRelationship")
    With cmd
        .Parameters("@_iVendKey").Value = cboVendID.ItemData(cboVendID.ListIndex)
        .Parameters("@_iItemKey").Value = m_lVItemKey
        .Execute
        VendOK = IIf(IsNull(.Parameters("@_oVendKey").Value), False, True)
    End With

    If Not VendOK Then
        MsgBox "Sage is not setup to buy this item from this vendor", vbOKOnly, "Item Vendor Setup"
        'reset combobox selection
        SetComboByKey cboVendID, m_defaultItemVendKey
        Exit Sub
    End If
    
'***
    
    bDirty = m_bDirty
    m_bDirty = True
    
    SetWaitCursor True

'NOTE: check the length of the comment field

    Select Case cmdVendUpdate.Caption
        Case "Add"
            CallSP "spCPCAddItemVend", "@_iWhseKey", m_lVWhseKey, _
                    "@_iItemKey", m_lVItemKey, "@_iVendKey", cboVendID.ItemData(cboVendID.ListIndex), _
                    "@_iComment", Trim(txtVendComment.text), "@_iVendCostItemID", Trim(txtVendCostItemID.text)
                    '"@_iObsolete", chkObsolete.Value,

            cmdVendUpdate.Enabled = False
            
        Case "Update"
            CallSP "spCPCUpdateItemVend", "@_iWhseKey", m_lVWhseKey, _
                    "@_iItemKey", m_lVItemKey, "@_iVendKey", cboVendID.ItemData(cboVendID.ListIndex), _
                    "@_iComment", Trim(txtVendComment.text), "@_iVendCostItemID", Trim(txtVendCostItemID.text)

            cmdVendUpdate.Enabled = False

    End Select
    
    cmdVendUpdate.Caption = "Update"
    
    SetWaitCursor False
    m_bDirty = bDirty
End Sub


'Called by:
'   cmdVendFind_Click()

Private Sub LoadVendInfo(ByVal lItemKey As Long, ByVal lWhseKey As Long)
    Dim rstItemVendInfo As Recordset
    Dim bDirty As Boolean
    
    Set rstItemVendInfo = CallSP("spCPCGetItemVendInfo", "@_iWhseKey", lWhseKey, "@_iItemKey", lItemKey)
    
    bDirty = m_bDirty
    m_bDirty = True
    
'***464
'not sure why we changed the sort order
'removed this at Bob's request 2/21/06 LR
'    g_rstVendors.Sort = "VendID"      'change sort order
    LoadCombo cboVendID, g_rstVendors, "VendName", "VendKey", , True
'    g_rstVendors.Sort = "VendName"    'restore
    
    If rstItemVendInfo.EOF Then
        txtVendComment.text = ""
        txtVendCostItemID.text = ""
        cmdVendUpdate.Caption = "Add"
        cmdVendUpdate.Enabled = False
    Else
        SetComboByKey cboVendID, rstItemVendInfo.Fields("VendKey").Value
        
'*** Added 3/9/2011 Len
'*** timVendItem fix
'*** cache the initial vendkey for the item
        m_defaultItemVendKey = rstItemVendInfo.Fields("VendKey").Value
'***

        If cboVendID.ListIndex > 0 Then
            g_rstVendors.Filter = "vendkey=" & cboVendID.ItemData(cboVendID.ListIndex)
            g_rstVendors.Filter = adFilterNone
        End If
        txtVendComment.text = rstItemVendInfo.Fields("Comment").Value
        m_sOldVendComment = Trim(rstItemVendInfo.Fields("Comment").Value)
        txtVendCostItemID.text = rstItemVendInfo.Fields("VendCostItemID").Value
        cmdVendUpdate.Caption = "Update"
        cmdVendUpdate.Enabled = True
    End If
    
    m_bDirty = bDirty
End Sub


'*************************************************************************
' Tab(10) Vendor/Items tab
'*************************************************************************


Private Sub LoadVendorCombo()
    Dim sSQL As String
    Dim orst As ADODB.Recordset
'SQL:
    sSQL = "SELECT distinct tapVendor.VendKey, tapVendor.VendID, tapVendor.VendName, tapVendor.DfltPurchAcctKey " & _
        "FROM timItem INNER JOIN timVendItem ON " & _
        "timItem.ItemKey = timVendItem.ItemKey INNER JOIN tapVendor ON " & _
        "timVendItem.VendKey = tapVendor.VendKey INNER JOIN timItemDescription ON " & _
        "timItem.ItemKey = timItemDescription.ItemKey " & _
        "WHERE (tapVendor.CompanyID = 'CPC') AND tapVendor.DfltPurchAcctKey = 3088 " & _
        "Order by tapVendor.VendName "
                
    Set orst = LoadDiscRst(sSQL)
    
    LoadCombo cboVIVendor(0), orst, "VendName", "VendKey", , True
    Set orst = Nothing
End Sub


Private Sub txtSOID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdLoad_Click 0
    End If
End Sub


Private Sub txtVIPartNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtVIPartNumber.text)) > 0 And KeyCode = vbKeyReturn Then
        cmdVendItemFind_Click
    End If
End Sub


Private Sub txtVIPartNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub

Private Sub cmdClear_Click()
    txtVIPartNumber.text = ""
    cboVIVendor(0).ListIndex = 0
    TryToSetFocus cboVIVendor(0)
End Sub


Private Sub cboVIVendor_Click(Index As Integer)
    If m_bLoad Then Exit Sub

    Dim oAPay As AcctPay
    Dim oAPay1 As AcctPay
    Dim bTemp As Boolean

    SetWaitCursor True
    m_bLoad = True
    If Index = 1 Then
        cmdVICreate.Enabled = False
        If m_APList.Count > 0 And cboVIVendor(1).ItemData(cboVIVendor(1).ListIndex) > 0 Then
            'Msg "Only those records that the part number from vendor " & cboVIVendor(1).List(cboVIVendor(1).ListIndex) & " doesn't exist will be selected.", vbOKOnly + vbExclamation, "Update Vendor"
            For Each oAPay In m_APList
                If oAPay.bUpdate = True Then
                    If VendorExists(oAPay.ItemKey, cboVIVendor(1).ItemData(cboVIVendor(1).ListIndex)) Then
                        oAPay.bUpdate = False
                    Else
                        bTemp = True
                    End If
                End If
            Next
            gdxVendItem.Refetch
            cmdVICreate.Enabled = True
        End If
        If (Not bTemp) And chkVISelectAll.Value = vbChecked Then
            chkVISelectAll.Value = vbUnchecked
        End If
    End If
    m_bLoad = False
    SetWaitCursor False
End Sub


Private Sub chkVISelectAll_Click()
    Dim bLoad As Boolean
    Dim lCount As Long
    If m_bLoad Then Exit Sub
    If m_APList Is Nothing Then Exit Sub
    If m_APList.Count = 0 Then Exit Sub
    
    bLoad = m_bLoad
    m_bLoad = True
    
    SetWaitCursor True
    If chkVISelectAll.Value = vbChecked Then
    
        Dim oAPay As AcctPay
        Dim lIndex As Long
        chkVIUnselectAll.Value = vbUnchecked
        
        If cboVIVendor(1).ItemData(cboVIVendor(1).ListIndex) > 0 Then
            msg "Only part numbers which are not already assigned to " & cboVIVendor(1).list(cboVIVendor(1).ListIndex) & " can be selected.", vbOKOnly + vbExclamation, "Vendor Items"
            For lIndex = 1 To m_APList.Count
                If Not VendorExists(m_APList(lIndex).ItemKey, cboVIVendor(1).ItemData(cboVIVendor(1).ListIndex)) Then
                    If lIndex > 1 Then
                        If m_APList(lIndex).ItemKey = m_APList(lIndex - 1).ItemKey And m_APList(lIndex - 1).bUpdate = True Then
                            m_APList(lIndex).bUpdate = False
                        Else
                            m_APList(lIndex).bUpdate = True
                            lCount = lCount + 1
                        End If
                    Else
                        m_APList(lIndex).bUpdate = True
                        lCount = lCount + 1
                    End If
                Else
                    m_APList(lIndex).bUpdate = False
                End If
            Next
        Else
            For lIndex = 1 To m_APList.Count
                If lIndex > 1 Then
                    If m_APList(lIndex).ItemKey = m_APList(lIndex - 1).ItemKey And m_APList(lIndex - 1).bUpdate = True Then
                        m_APList(lIndex).bUpdate = False
                    Else
                        m_APList(lIndex).bUpdate = True
                        lCount = lCount + 1
                    End If
                Else
                    m_APList(lIndex).bUpdate = True
                    lCount = lCount + 1
                End If
            Next
        End If
        
        If lCount = 0 Then
            msg "Sorry. no items are qualified for creating Vend/Item relationship."
            chkVISelectAll.Value = vbUnchecked
        End If
        gdxVendItem.Refetch
    End If
    m_bLoad = bLoad
    SetWaitCursor False
End Sub


Private Sub chkVIUnselectAll_Click()
    Dim bLoad As Boolean
    If m_bLoad Then Exit Sub
    If m_APList Is Nothing Then Exit Sub
    If m_APList.Count = 0 Then Exit Sub
    
    bLoad = m_bLoad
    m_bLoad = True
    
    If chkVIUnselectAll.Value = vbChecked Then
        Dim oAPay As AcctPay
        
        chkVISelectAll.Value = vbUnchecked
        For Each oAPay In m_APList
            oAPay.bUpdate = False
        Next
        
        gdxVendItem.Refetch
    End If
    m_bLoad = False
End Sub


Private Sub cmdVendItemFind_Click()
    Dim bLoad As Boolean
    Dim bTemp As Boolean
    
    If Trim(cboVIVendor(0).list(cboVIVendor(0).ListIndex)) = "<none>" And Trim(txtVIPartNumber) = "" Then
        TryToSetFocus cboVIVendor(0)
        Exit Sub
    End If
    
    Set m_APList = New AcctPayList
    
    bTemp = m_APList.LoadVendItem(PrepSQLText(Trim(txtVIPartNumber)), cboVIVendor(0).ItemData(cboVIVendor(0).ListIndex), cboVIVendor(1).ItemData(cboVIVendor(1).ListIndex))
    
    bLoad = m_bLoad
    m_bLoad = True
    If m_APList.Count = 0 Then
        MsgBox "Sorry, No records satisfy this request."
        If cboVIVendor(0).ListIndex = 0 Then
            txtVIPartNumber.SelLength = Len(txtVIPartNumber.text)
            txtVIPartNumber.SelStart = 0
            TryToSetFocus txtVIPartNumber
        Else
            TryToSetFocus cboVIVendor(0)
        End If
        chkVISelectAll.Enabled = False
        chkVIUnselectAll.Enabled = False
        chkVISelectAll.Value = vbUnchecked
        chkVIUnselectAll.Value = vbUnchecked
        cmdVICreate.Enabled = False
    Else
        chkVISelectAll.Enabled = True
        chkVIUnselectAll.Enabled = True
        If bTemp Then
            chkVISelectAll.Value = vbChecked
        Else
            chkVISelectAll.Value = vbUnchecked
        End If
        chkVIUnselectAll.Value = vbUnchecked
    End If
    
    m_bLoad = bLoad
    LoadVIGrid
End Sub


Private Sub LoadVIGrid()
    With gdxVendItem
        .Row = -1
        .HoldSortSettings = True
        .ColumnAutoResize = True
        .HoldFields
        .ItemCount = m_APList.Count
        .Refetch
        .Row = 1
        TryToSetFocus gdxVendItem
    End With
    
    If cboVIVendor(1).ItemData(cboVIVendor(1).ListIndex) > 0 Then
        cmdVICreate.Enabled = True
    End If
End Sub


Private Function checkUpdateChecked() As Boolean
    Set m_APListUpdate = New AcctPayList    'NOTE: This is done on Form_Load

    Dim AP As AcctPay
    For Each AP In m_APList
        If AP.bUpdate Then
            m_APListUpdate.Add AP
        End If
    Next
    
    checkUpdateChecked = (m_APListUpdate.Count > 0)
End Function


Private Function VendorExists(ByVal ItemKey As Long, ByVal VendKey As Long) As Boolean
    Dim orst As ADODB.Recordset
    
    Set orst = CallSP("spCPCExistVendItem", "@_iItemKey", ItemKey, "@_iVendKey", VendKey)
    
    If Not orst.EOF Then VendorExists = True
    
    Set orst = Nothing
End Function


Private Sub cmdVICreate_Click()
    Dim oAcctPay As AcctPay
    
    If m_APListUpdate Is Nothing Then Exit Sub
    
    If checkUpdateChecked Then
        If vbYes = msg("Are you sure that your want to create these Vend/Item(s)?", vbYesNo + vbExclamation, "Create Vend/Item?") Then
            AddVendItem
            cmdVendItemFind_Click
        End If
    Else
        msg "Please select part numbers you want to add first.", vbOKOnly, "Vendor Items"
    End If
    
    cmdVICreate.Enabled = False
End Sub


'NOTE: why the transaction? All the AP objects or nothing?

Private Sub AddVendItem()
    
    On Error GoTo ErrorHandler
    
    Dim oCmd As ADODB.Command
    Dim AP As AcctPay
    Dim lVendKey As Long
    
    m_bLoad = True
    g_DB.Connection.BeginTrans
    
    lVendKey = cboVIVendor(1).ItemData(cboVIVendor(1).ListIndex)
    For Each AP In m_APListUpdate
        Set oCmd = CreateCommandSP("spcpcInsertVendItem")
        With oCmd
            .Parameters("@_iVendKey").Value = lVendKey
            .Parameters("@_iItemKey").Value = AP.ItemKey
            .Parameters("@_iBreakType").Value = AP.BreakType
            .Parameters("@_iDiscMeth").Value = AP.DiscMeth
            .Parameters("@_iItemAliasKey").Value = AP.ItemAliasKey
            .Parameters("@_iListPrice").Value = AP.ListPrice
            .Parameters("@_iOrigCountry").Value = Trim(AP.OrigCountry)
            .Parameters("@_iPurchUnitMeasKey").Value = AP.PurchUnitMeasKey
            .Parameters("@_iRplcmntUnitCost").Value = AP.NewReplCost
            .Parameters("@_iVendFamilyKey").Value = Null
            .Parameters("@_iVendItemID").Value = Trim(AP.VendPartNbr)
            .Parameters("@_iSubjToVendFamDisc").Value = AP.SubjToVendFamDisc
            .Execute
        End With
    Next
    cboVIVendor(1).ListIndex = 0
    m_bLoad = False
    Set oCmd = Nothing
    g_DB.Connection.CommitTrans
    Exit Sub
    
    
ErrorHandler:
    DisplayWarning "Unexpected Internal Error in creating VendItem"
    g_DB.Connection.RollbackTrans
    m_bLoad = False
    Set oCmd = Nothing
End Sub


Private Sub gdxVendItem_LostFocus()
    gdxVendItem.Update
End Sub

Private Sub gdxVendItem_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    
    If m_APList Is Nothing Then Exit Sub
    
    If RowIndex > m_APList.Count Then Exit Sub
    
    With m_APList(RowIndex)
        Values(1) = .bUpdate
        Values(2) = .NewReplCost
        Values(3) = .VendPartNbr
        Values(4) = .CPCPartNbr
        Values(5) = .Descr
        Values(6) = .Vendor
    End With
End Sub


Private Sub gdxVendItem_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim oAPay As AcctPay
    Dim lIndex As Long
    Dim bMulRecords As Boolean
    
    On Error GoTo ErrorHandler
    
    If Values(1) = True Then
        If VendorExists(m_APList(RowIndex).ItemKey, cboVIVendor(1).ItemData(cboVIVendor(1).ListIndex)) Then
            msg "This item already exists for the vendor. You cannot add a new Vendor/Item relationship for it."
            Exit Sub
        End If
        
        For lIndex = 1 To m_APList.Count
            If lIndex <> RowIndex And m_APList(RowIndex).ItemKey = m_APList(lIndex).ItemKey And m_APList(lIndex).bUpdate = True Then
                msg "There is already one " & m_APList(RowIndex).CPCPartNbr & " selected. Please select only one entry for a particular part number.", vbExclamation + vbOKOnly, "Vendor Items"
                bMulRecords = True
                Exit For
            End If
        Next
    End If
    
    If Not bMulRecords Then
        If Not IsNumeric(Values(2)) Then
            msg "The value in the cost field is not valid.", vbOKOnly + vbExclamation, "Vendor Item"
        Else
            With m_APList(RowIndex)
                .NewReplCost = Values(2)
                .VendPartNbr = Values(3)
                .bUpdate = Values(1)
            End With
        End If
    End If
    Exit Sub
    
ErrorHandler:
    msg "The value in the Cost or PartNbr field is not valid. Please enter effective Cost or PartNbr.", vbOKOnly + vbExclamation, "Vendor Item"
End Sub


'*********************************************************************
'Tab(11): VendItem MaxQty and MinQty tab
'*********************************************************************

Private Sub cmdMMRefresh_Click()
    Dim bTemp As Boolean

    If Trim(txtVendID.text = "") Then
        TryToSetFocus txtVendID
    Else
        Set m_APList = New AcctPayList
        
        bTemp = m_APList.LoadMMVendItem(PrepSQLText(Trim(txtVendID.text)), cboMMWarehouse.ItemData(cboMMWarehouse.ListIndex))
        
        If Not bTemp Then
            txtVendID.SelStart = 0
            txtVendID.SelLength = Len(txtVendID.text)
            TryToSetFocus txtVendID
        End If
        LoadMMVIGrid
    End If
    
    m_bLoadRobertData = False
    m_lMMWhseKey = cboMMWarehouse.ItemData(cboMMWarehouse.ListIndex)
End Sub


Private Sub cboMMWarehouse_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtVendID.text)) > 0 And KeyCode = vbKeyReturn Then
        cmdMMRefresh_Click
    End If
End Sub


Private Sub LoadMMVIGrid()
    With gdxMMVendItem
        .Row = -1
        .HoldSortSettings = True
        .ColumnAutoResize = True
        .HoldFields
        .ItemCount = m_APList.Count
        .Refetch
        .Row = 1
        TryToSetFocus gdxMMVendItem
    End With
End Sub


Private Sub gdxMMVendItem_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_APList Is Nothing Then Exit Sub
    
    If RowIndex > m_APList.Count Then Exit Sub
    
    With m_APList(RowIndex)
        Values(1) = .NewMax
        Values(2) = .NewMin
        Values(3) = .MaxQty
        Values(4) = .MinQty
        Values(5) = .VendPartNbr
        Values(6) = .Descr
        Values(7) = .VendID
        Values(8) = .Vendor
        Values(9) = .bUpdate
        Values(10) = .ItemKey
        Values(11) = .VendKey
    End With
End Sub


Private Sub gdxMMVendItem_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo ErrorHandler
    
    If Not IsNumeric(Values(1)) Then
        msg "The value in New Max Qty field is not valid.", vbOKOnly + vbExclamation, "Vend Items MaxQty"
    ElseIf Not IsNumeric(Values(2)) Then
        msg "The value in New Min Qty field is not valid.", vbOKOnly + vbExclamation, "Vend Items MinQty"
    Else
        With m_APList(RowIndex)
            .bUpdate = (CLng(Values(1)) <> CLng(Values(3)) Or CLng(Values(2)) <> CLng(Values(4)))
            .NewMax = CLng(Values(1))
            .NewMin = CLng(Values(2))
        End With
    End If
    Exit Sub
    
ErrorHandler:
    msg "The value in New Max or Min Stock field is not valid. Please enter effective New Max or Min Stock.", vbOKOnly + vbExclamation, "Vendor Item Max&Min Stock"

End Sub


Private Sub gdxMMVendItem_LostFocus()
    gdxMMVendItem.Update
End Sub


Private Sub cmdMMUpdate_Click()
    SetWaitCursor True
    If checkUpdateChecked Then
        If vbYes = msg("Are you sure that your want to update Max&Min stock for those vend items?", vbYesNo + vbExclamation, "Update VendItem Max&Min Stock?") Then
            UpdateVendItemMM
            If m_bLoadRobertData Then
                cmdSpreadSheet_Click
            Else
                cmdMMRefresh_Click
            End If
        End If
    Else
        If m_bLoadRobertData Then
            msg "All Min&Max stock values from Robert's file are the same as in Inventory table"
        Else
            msg "Please edit New Max Stock or Min Stock of Vendor items that you want to edit first.", vbOKOnly, "Vend Items Max&Min Stocks"
        End If
    End If
    SetWaitCursor False
End Sub


Private Sub UpdateSpreadSheetMM(ByRef oFrm As FProgress)
    On Error GoTo ErrorHandler
    
    Dim oCmd As ADODB.Command
    Dim AP As AcctPay
    Dim lAmount As Long

    g_DB.Connection.BeginTrans
    
    For Each AP In m_APList
        lAmount = lAmount + 1
    
        oFrm.SecondStepProgress lAmount / m_APList.Count
        Set oCmd = CreateCommandSP("spCPCUpdateVendItemMM")
        With oCmd
            .Parameters("@_iWhseKey").Value = m_lMMWhseKey
            .Parameters("@_iItemKey").Value = AP.ItemKey
            .Parameters("@_iMaxStockQty").Value = AP.NewMax
            .Parameters("@_iMinStockQty").Value = AP.NewMin
            .Execute
        End With
        WriteResultToLogFile "Vendor - " & AP.VendID & ", Vend PartNbr - " & AP.VendPartNbr & ", New Max Qty - " & AP.NewMax & ", New Min Qty - " & AP.NewMin
    Next
    Set oCmd = Nothing
    g_DB.Connection.CommitTrans
    Exit Sub
    
ErrorHandler:
    DisplayWarning "Unexpected Internal Error in updating Vend Items Max&Min stocks"
    g_DB.Connection.RollbackTrans
    Set oCmd = Nothing
End Sub




Private Sub UpdateVendItemMM()
    On Error GoTo ErrorHandler
    
    Dim oCmd As ADODB.Command
    Dim AP As AcctPay
    
    g_DB.Connection.BeginTrans
    
    For Each AP In m_APListUpdate
        Set oCmd = CreateCommandSP("spCPCUpdateVendItemMM")
        With oCmd
            .Parameters("@_iWhseKey").Value = m_lMMWhseKey
            .Parameters("@_iItemKey").Value = AP.ItemKey
            .Parameters("@_iMaxStockQty").Value = AP.NewMax
            .Parameters("@_iMinStockQty").Value = AP.NewMin
            .Execute
        End With
    Next
    Set oCmd = Nothing
    g_DB.Connection.CommitTrans
    Exit Sub
    
ErrorHandler:
    DisplayWarning "Unexpected Internal Error in updating Vend Items Max&Min stocks"
    g_DB.Connection.RollbackTrans
    Set oCmd = Nothing
End Sub


Private Sub txtVendID_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtVendID.text)) > 0 And KeyCode = vbKeyReturn Then
        cmdMMRefresh_Click
    End If
End Sub


Private Sub txtVendID_KeyPress(KeyAscii As Integer)
     If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub


'Update inventory MinMax value from spreadsheets

Private Sub cmdUpdateXLS_Click()
    'The first step is prompting users to choose folder where
    'those spreadsheet files located
    Dim oShell As Shell
    Dim oFolder As Folder
    Dim oFolderFile As FolderItem
    Dim oFrm As FProgress
    Dim sFileName As String
    Dim lAmount As Long
    
    On Error Resume Next
    
    SetWaitCursor True
    Set oFrm = New FProgress
    Set m_APList = New AcctPayList
    Set oShell = New Shell
    Set oFolder = oShell.BrowseForFolder(Me.hWnd, "Please select the folder where spreadsheet located", 111)
    oFrm.Show
    For Each oFolderFile In oFolder.Items
        If Right(oFolderFile.Name, 4) = ".xls" Then
            'The first step is to load data from spread sheets
            lAmount = lAmount + 1
            oFrm.FirstStepProgress lAmount / oFolder.Items.Count
            m_APList.LoadSpreadSheet oFolderFile.path
            'The second step is to update data from spread sheets.
        End If
    Next
    
    If lAmount = 0 Then
        oFrm.HideProgress
        msg "No file available, Please check if you choose the correct folder"
    ElseIf m_APList.Count = 0 Then
        oFrm.HideProgress
        msg "No record is qualifed for updating Min/Max values in inventory table." & vbCrLf & _
            "Please check spread sheet files in the folder."
    Else
        m_lMMWhseKey = 24
        UpdateSpreadSheetMM oFrm
        oFrm.HideProgress
        'Also log the updating result to
    End If

    Set m_APList = Nothing
    Set oFrm = Nothing
    SetWaitCursor False
End Sub


'Called by:
'   UpdateSpreadSheetMM()
'

Private Sub WriteResultToLogFile(ByVal sErrMsg As String)
    Dim oFSO As FileSystemObject
    Dim oTS As TextStream
    
    Set oFSO = New FileSystemObject
    Set oTS = oFSO.OpenTextFile(App.path & "\SpreadSheetMM.txt", ForAppending, True)
    
    If Not oTS Is Nothing Then
        oTS.Write sErrMsg & vbCrLf
        oTS.Close
    End If
    
    Set oTS = Nothing
    Set oFSO = Nothing
End Sub


Private Sub cmdSpreadSheet_Click()
    Dim bTemp As Boolean

    Set m_APList = New AcctPayList
    
    bTemp = m_APList.LoadSpreadSheet(App.path & "\MaxMinSea2001ForRobert.xls")
    If Not bTemp Then
        msg "No results are retrieved from Robert's spreadsheet. Please contact TechSupport for details."
    End If
    LoadMMVIGrid
    
    m_bLoadRobertData = True
    m_lMMWhseKey = 24
End Sub


Private Sub cmdAddRobertFile_Click()
    Dim oConn As ADODB.Connection
    Dim orst As ADODB.Recordset
    Dim orstItem As ADODB.Recordset
    Dim rstItemVendInfo As ADODB.Recordset
    Dim lItemKey As Long
    Dim lVendorKey As Long
    Dim sVendID As String
    Dim i As Long
    
    SetWaitCursor True
    
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & App.path & "\FixInventory1.xls;" & _
               "Extended Properties=""Excel 8.0;HDR=yes;"""
    Debug.Print oConn.ConnectionString
'SQL:
    Set orst = LoadDiscRst("Select * from [VendItem$] where [Update] is not null and rtrim([Update]) <> '' ", oConn)
    If orst.EOF Then
        msg "Worksheet is empty"
    Else
        Debug.Print "The total rows are " & orst.RecordCount
        While Not orst.EOF
            i = 0
'            Debug.Print RTrim(oRst.Fields("ItemID").Value) & " " & RTrim(oRst.Fields("VendID").Value) & " " & RTrim(oRst.Fields("Update").Value) & " "; RTrim(oRst.Fields("NewVendID").Value)
'SQL:
             Set orstItem = LoadDiscRst("Select ItemKey from timItem where ItemID = '" & RTrim(orst.Fields("ItemID").Value) & "'")
             If Not orstItem.EOF Then
                lItemKey = orstItem.Fields("ItemKey").Value
                lVendorKey = 0
                sVendID = ""
                If RTrim(orst.Fields("Update").Value) = "Y" Then
                    sVendID = RTrim(orst.Fields("VendID").Value)
                     g_rstVendors.Filter = "VendID='" & RTrim(orst.Fields("VendID").Value) & "'"
                        If Not g_rstVendors.EOF Then
                            lVendorKey = g_rstVendors.Fields("VendKey").Value
                        End If
                    g_rstVendors.Filter = adFilterNone
                ElseIf RTrim(orst.Fields("Update").Value) = "N" Then
                    sVendID = RTrim(orst.Fields("NewVendID").Value)
                    g_rstVendors.Filter = "VendID='" & RTrim(orst.Fields("NewVendID").Value) & "'"
                        If Not g_rstVendors.EOF Then
                            lVendorKey = g_rstVendors.Fields("VendKey").Value
                        End If
                    g_rstVendors.Filter = adFilterNone
                End If

             
                 If lItemKey <> 0 And lVendorKey <> 0 Then
                 '10/09/02 TeddyX
                 'Do we have to check if the entries existed in tcpWhseItemVend?
                    Set rstItemVendInfo = CallSP("spCPCGetItemVendInfo", "@_iWhseKey", g_MPKWhseKey, "@_iItemKey", lItemKey)
                    If rstItemVendInfo.EOF Then
                        CallSP "spCPCAddItemVend", "@_iWhseKey", g_MPKWhseKey, _
                                "@_iItemKey", lItemKey, "@_iVendKey", lVendorKey, _
                            "@_iObsolete", 0, "@_iComment", ""
                    Else
                        rstItemVendInfo.MoveFirst

                        If Not IsNull(rstItemVendInfo.Fields("VendKey").Value) Then
                            g_rstVendors.Filter = "VendKey = " & rstItemVendInfo.Fields("VendKey").Value
                            If Not g_rstVendors.EOF Then
                                Debug.Print RTrim(orst.Fields("ItemID").Value) & " " & RTrim(g_rstVendors.Fields("VendID").Value) & " " & sVendID
                            End If
                            
                            g_rstVendors.Filter = adFilterNone
                        Else
                            Debug.Print RTrim(orst.Fields("ItemID").Value) & " NULL" & " " & sVendID
                        End If
                        CallSP "spCPCUpdateRobertItemVend", "@_iWhseKey", g_MPKWhseKey, _
                            "@_iItemKey", lItemKey, "@_iVendKey", lVendorKey
                        'Debug.Print RTrim(oRst.Fields("ItemID").Value) & " " & RTrim(g_rstVendors.Fields("VendID").Value) & " " & sVendID
                       ' g_rstVendors.Filter = adFilterNone
                    End If
                    i = i + 1
                End If
            End If

            orst.MoveNext
        Wend
    End If
    
    Debug.Print "The total inserted rows are " & i
    SetWaitCursor False
    Set orst = Nothing
    oConn.Close
    Set oConn = Nothing
End Sub


'*********************************************************************
'Tab(12): Gasket Status tab
'*********************************************************************

Private Sub cmdViewGaskOrdStatus_Click()
    'PRN#130
    Dim frm As FXSLViewer
    Dim sGasketStatus As String
    On Error GoTo EH
    
        sGasketStatus = GetGasketOrdersStatusXML(cboWhse_GaskOrdStatus.ItemData(cboWhse_GaskOrdStatus.ListIndex))
        If sGasketStatus <> "" Then
            Set frm = New FXSLViewer
            MDIMain.AddNewWindow frm
            frm.ShowViewer "Gasket Orders Status Viewer (" & cboWhse_GaskOrdStatus.text & ")", g_XsltPath & "GasketOrdersStatus.xsl", sGasketStatus
        Else
            msg "There is no open gasket order", vbOKOnly, "Gasket Orders Status Viewer"
            cboWhse_GaskOrdStatus.SetFocus
        End If
    
    Exit Sub
EH:
    MsgBox "Gasket Orders Status viewer can't be loaded due to error " & Err.Number & " '" & Err.Description & "'", vbInformation
End Sub


Private Sub cmdView_Click()
    'PRN#21
    Dim frm As FXSLViewer
    Dim sGasketStatus As String
    On Error GoTo EH
    
    txtOPKey.text = Trim(txtOPKey.text)
    If txtOPKey.text <> "" Then
        sGasketStatus = GetGasketStatusXML(txtOPKey.text, optOP.Value = True)
        If sGasketStatus <> "" Then
            Set frm = New FXSLViewer
            MDIMain.AddNewWindow frm
            If optOP.Value = True Then
                frm.ShowViewer "Gasket Manufacturing Viewer - OP " & txtOPKey.text, g_XsltPath & "GasketReport.xsl", sGasketStatus
            Else
                frm.ShowViewer "Gasket Manufacturing Viewer - SO " & txtOPKey.text, g_XsltPath & "GasketReport.xsl", sGasketStatus
            End If
        Else
            msg "No record exists. Please enter a different OP#/SO#", vbOKOnly, "Gasket Viewer"
            txtOPKey.SetFocus
        End If
    Else
        MsgBox "Please enter an OP#/SO#.", vbInformation
        txtOPKey.SetFocus
    End If
    
    Exit Sub
EH:
    MsgBox "Gasket viewer can't be loaded due to error " & Err.Number & " '" & Err.Description & "'", vbInformation
End Sub


'Called by:
'   cmdView_Click()

Private Function GetGasketStatusXML(ByVal Key As Long, ByVal IsByOPKey As Boolean) As String
    Dim sTemp As String
    On Error GoTo EH
    
    SetWaitCursor True
    
    Dim orst As ADODB.Recordset
    If IsByOPKey Then
        Set orst = CallSP("spcpcGetGasketStatus_ByOPKey", "@_iOPKey", Key)
    Else
        Set orst = CallSP("spcpcGetGasketStatus_BySONum", "@_iTranKey", Key)
    End If
    
    With orst
        If Not orst.EOF Then
            Do Until .EOF
                sTemp = sTemp & "<Order>" & vbCrLf
                sTemp = sTemp & "<OPID>" & .Fields("OPKey") & "</OPID>" & vbCrLf
                sTemp = sTemp & "<SOID>" & .Fields("TranKey") & "</SOID>" & vbCrLf
                sTemp = sTemp & "<SODate>" & .Fields("SODate") & "</SODate>" & vbCrLf
                sTemp = sTemp & "<OrderStatus>" & .Fields("OrderStatus") & "</OrderStatus>" & vbCrLf
                sTemp = sTemp & "<CustName>" & CleanXML(.Fields("CustName")) & "</CustName>" & vbCrLf
                sTemp = sTemp & "<WhseKey>" & .Fields("WhseKey") & "</WhseKey>" & vbCrLf
                sTemp = sTemp & "<Descr>" & CleanXML(.Fields("Descr")) & "</Descr>" & vbCrLf
                sTemp = sTemp & "<LineStatus>" & .Fields("LineStatus") & "</LineStatus>" & vbCrLf
                sTemp = sTemp & "<PickNum>" & .Fields("PickListNo") & "</PickNum>" & vbCrLf
                sTemp = sTemp & "<PickDate>" & .Fields("PickDate") & "</PickDate>" & vbCrLf
                sTemp = sTemp & "<BeginUser>" & .Fields("BeginUser") & "</BeginUser>" & vbCrLf
                sTemp = sTemp & "<BeginTime>" & .Fields("BeginTime") & "</BeginTime>" & vbCrLf
                sTemp = sTemp & "<CutUser>" & .Fields("CutUser") & "</CutUser>" & vbCrLf
                sTemp = sTemp & "<CutTime>" & .Fields("CutTime") & "</CutTime>" & vbCrLf
                sTemp = sTemp & "<MoldUser>" & .Fields("MoldUser") & "</MoldUser>" & vbCrLf
                sTemp = sTemp & "<MoldTime>" & .Fields("MoldTime") & "</MoldTime>" & vbCrLf
                sTemp = sTemp & "<TrimUser>" & .Fields("TrimUser") & "</TrimUser>" & vbCrLf
                sTemp = sTemp & "<TrimTime>" & .Fields("TrimTime") & "</TrimTime>" & vbCrLf
                sTemp = sTemp & "</Order>"
                .MoveNext
            Loop
            GetGasketStatusXML = "<Orders>" & sTemp & "</Orders>"
        End If
    End With
    Set orst = Nothing
    
    SetWaitCursor False
    Exit Function
    
EH:
    ClearWaitCursor
    Set orst = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'Called by:
'   cmdViewGaskOrdStatus_Click()

Private Function GetGasketOrdersStatusXML(ByVal WhseKey As Long) As String
    Dim sTemp As String
    Dim sTempSummary As String
    Dim lPrevOPKey As Long
    Dim lMyRecNum As Long
    Dim lInDept As Long
    
    On Error GoTo EH
    
    SetWaitCursor True
    
    Dim orst As ADODB.Recordset
    Set orst = CallSP("spcpcGetGasketOrdersSummary", "@_iWhseKey", WhseKey)
    With orst
        Do Until .EOF
            Select Case .Fields("StateKey")
                Case 1 'In Dept
                    sTempSummary = sTempSummary & "<InDept>" & .Fields("Total") & "</InDept>" & vbCrLf
                    lInDept = .Fields("Total")
                Case 2 'Cut
                    sTempSummary = sTempSummary & "<NeedCut>" & lInDept - .Fields("Total") & "</NeedCut>" & vbCrLf
                Case 3 'Mold
                    sTempSummary = sTempSummary & "<NeedMold>" & lInDept - .Fields("Total") & "</NeedMold>" & vbCrLf
            End Select
            .MoveNext
        Loop
        If .State <> adStateClosed Then
            .Close
        End If
    End With
    Set orst = Nothing
    
    Set orst = CallSP("spcpcGetGasketOrdersStatus", "@_iWhseKey", WhseKey)
    With orst
        If Not orst.EOF Then
            Do Until .EOF
                If lPrevOPKey <> .Fields("OPKey") Then
                    lPrevOPKey = .Fields("OPKey")
                    lMyRecNum = lMyRecNum + 1
                    If sTemp <> "" Then
                        sTemp = sTemp & Space(2) & "</Lines>" & vbCrLf
                        sTemp = sTemp & "</Order>" & vbCrLf
                    End If
                    sTemp = sTemp & "<Order>" & vbCrLf
                    sTemp = sTemp & Space(2) & "<MyRecNum>" & lMyRecNum & "</MyRecNum>" & vbCrLf
                    sTemp = sTemp & Space(2) & "<OPID>" & .Fields("OPKey") & "</OPID>" & vbCrLf
                    sTemp = sTemp & Space(2) & "<SODate>" & .Fields("SODate") & "</SODate>" & vbCrLf
                    sTemp = sTemp & Space(2) & "<CustName>" & CleanXML(.Fields("CustName")) & "</CustName>" & vbCrLf
                    sTemp = sTemp & Space(2) & "<Lines>" & vbCrLf
                End If
                sTemp = sTemp & Space(4) & "<Line>" & vbCrLf
                sTemp = sTemp & Space(6) & "<LineStatus>" & .Fields("LineStatus") & "</LineStatus>" & vbCrLf
                sTemp = sTemp & Space(6) & "<Descr>" & CleanXML(.Fields("Descr")) & "</Descr>" & vbCrLf
                sTemp = sTemp & Space(6) & "<BeginUser>" & .Fields("BeginUser") & "</BeginUser>" & vbCrLf
                sTemp = sTemp & Space(6) & "<BeginTime>" & .Fields("BeginTime") & "</BeginTime>" & vbCrLf
                sTemp = sTemp & Space(6) & "<CutUser>" & .Fields("CutUser") & "</CutUser>" & vbCrLf
                sTemp = sTemp & Space(6) & "<CutTime>" & .Fields("CutTime") & "</CutTime>" & vbCrLf
                sTemp = sTemp & Space(6) & "<MoldUser>" & .Fields("MoldUser") & "</MoldUser>" & vbCrLf
                sTemp = sTemp & Space(6) & "<MoldTime>" & .Fields("MoldTime") & "</MoldTime>" & vbCrLf
                sTemp = sTemp & Space(6) & "<TrimUser>" & .Fields("TrimUser") & "</TrimUser>" & vbCrLf
                sTemp = sTemp & Space(6) & "<TrimTime>" & .Fields("TrimTime") & "</TrimTime>" & vbCrLf
                sTemp = sTemp & Space(4) & "</Line>" & vbCrLf
                .MoveNext
            Loop
            sTemp = sTemp & Space(2) & "</Lines>" & vbCrLf
            sTemp = sTemp & "</Order>" & vbCrLf
        End If
    End With
    Set orst = Nothing
    
    GetGasketOrdersStatusXML = "<Orders><WhseKey>" & WhseKey & "</WhseKey>" & vbCrLf & sTempSummary & sTemp & "</Orders>"
    
    SetWaitCursor False
    Exit Function
    
EH:
    ClearWaitCursor
    Set orst = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'Called by:
'   GetGasketStatusXML()
'   GetGasketOrdersStatusXML()

Private Function CleanXML(ByVal sXML As String) As String
    Dim sTemp As String
    sTemp = sXML
    sTemp = Replace(sTemp, "&", "&amp;")
    sTemp = Replace(sTemp, "<", "&lt;")
    sTemp = Replace(sTemp, ">", "&gt;")
    CleanXML = sTemp
End Function

Private Sub cmdPrint_Click()
    On Error GoTo ErrorHandler
    
    If Not (chkWillCall.Value = vbChecked Or chkShip.Value = vbChecked) Then
        msg "Please select the report(s) you want to print", vbExclamation + vbOKOnly, "Gasket Status"
        Exit Sub
    End If

    SetWaitCursor True
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spcpcGetGasketStatus")
    cmd.Parameters("@_iWhseKey").Value = cboWhse.ItemData(cboWhse.ListIndex)
    cmd.Execute
    
    'smr - 12/04 - New logic to call FViewer - uses: "Crystal Reports 8.5 ActiveX Designer Run Time Library"
    Dim oFrm As FViewer
    
    If chkWillCall.Value = vbChecked Then
        Set oFrm = New FViewer
        Call oFrm.ViewReportByType("Gasket Report Will Call")
    End If
    
    If chkShip.Value = vbChecked Then
        Set oFrm = New FViewer
        Call oFrm.ViewReportByType("Gasket Report Ship")
    End If
    
    Set oFrm = Nothing
    
    Set cmd = Nothing
    SetWaitCursor False
    Exit Sub
    
ErrorHandler:
    ClearWaitCursor
    Set cmd = Nothing
    msg Err.Number & " - " & Err.Description, vbOKOnly + vbExclamation, Err.Source
End Sub


''Private Sub PrintGasketReport(ByVal sReportFileName As String)
''    crGsktStatus.Destination = 0
''    crGsktStatus.ReportFileName = sReportFileName
''    crGsktStatus.WindowShowPrintBtn = True
''    crGsktStatus.WindowShowRefreshBtn = True
''    crGsktStatus.WindowState = crptMaximized
''    crGsktStatus.PrintReport
''End Sub


Private Sub cmdGet_Click()
    On Error GoTo ErrorHandler
    
    Dim orst As ADODB.Recordset
    
    SetWaitCursor True
    txtCutAmt.text = ""
    txtMoldAmt.text = ""
    txtTrimAmt.text = ""
        
    Set orst = CallSP("spCPCGsktSummary", "@_iStartDate", dtpStartDate.Value, "@_iEndDate", dtpEndDate.Value)
    
    If orst.EOF Then
        msg "No record exists. Please reset the start and end date", vbOKOnly, "Gasket Summary"
        
    Else
        orst.MoveFirst
        While Not orst.EOF
            Select Case Trim(orst.Fields("StateID").Value)
            Case "Cut"
                txtCutAmt.text = orst.Fields("Amt").Value
            Case "Mold"
                txtMoldAmt.text = orst.Fields("Amt").Value
            Case "Trim"
                txtTrimAmt.text = orst.Fields("Amt").Value
            End Select
            orst.MoveNext
        Wend
    End If
    
    Set orst = Nothing
    SetWaitCursor False
    Exit Sub
    
ErrorHandler:
    Set orst = Nothing
    ClearWaitCursor
    msg Err.Number & " - " & Err.Description, vbOKOnly + vbExclamation, Err.Source
End Sub


'*********************************************************************
'Tab(13): Sales History tab
'*********************************************************************

Private Sub cmdFind_SalesAnalysis_Click()
    Dim orst As ADODB.Recordset
    On Error GoTo EH
    
    txtItemID_SalesAnalysis.text = Trim$(txtItemID_SalesAnalysis.text)
    If txtItemID_SalesAnalysis.text = "" And cboVend_SalesAnalysis.text = "<ALL>" Then
        MsgBox "Please enter a part number and/or select a vendor.", vbInformation
        txtItemID_SalesAnalysis.SetFocus
        Exit Sub
    End If
    
    SetWaitCursor True
    
    'smr - 01-24-2005 - if Whse = <All>, the @WhseKey will equal 0
    If cboVend_SalesAnalysis.ListIndex = 0 Then
        Set orst = CallSP("cpoaGetSalesByPartAndVendor", _
            "@WhseKey", cboWhse_SalesAnalysis.ItemData(cboWhse_SalesAnalysis.ListIndex), _
            "@ItemID", txtItemID_SalesAnalysis.text)
    'smr - 01-21-2005 - to show all part#s based on Vendor
    ElseIf txtItemID_SalesAnalysis.text = "" Then
        Set orst = CallSP("cpoaGetSalesByPartAndVendor", _
            "@WhseKey", cboWhse_SalesAnalysis.ItemData(cboWhse_SalesAnalysis.ListIndex), _
            "@VendKey", cboVend_SalesAnalysis.ItemData(cboVend_SalesAnalysis.ListIndex))
    Else
        Set orst = CallSP("cpoaGetSalesByPartAndVendor", _
            "@WhseKey", cboWhse_SalesAnalysis.ItemData(cboWhse_SalesAnalysis.ListIndex), _
            "@ItemID", txtItemID_SalesAnalysis.text, _
            "@VendKey", cboVend_SalesAnalysis.ItemData(cboVend_SalesAnalysis.ListIndex))
    End If
        
    SetWaitCursor False
    
    'cboWhse_SalesAnalysis.Tag = cboWhse_SalesAnalysis.ItemData(cboWhse_SalesAnalysis.ListIndex)
    
'    'smr
'    Dim col As JSColumn
'    Dim fmtCon As JSFmtCondition
'    Set col = gdxSalesAnalysis.Columns("Min")
'    Set fmtCon = gdxSalesAnalysis.FmtConditions.Add(col.Index, jgexLessThan, 1)
'    fmtCon.FormatStyle.ForeColor = vbRed
    
    
    If orst.EOF Then
        msg "No item found.", vbInformation, "Sales History"
    Else
        With gdxSalesAnalysis
            .HoldFields
            .HoldSortSettings = True
            Set .ADORecordset = orst
            
            'smr - 01/24/2005 - only show the col whse if the search is by All warehouses
            If cboWhse_SalesAnalysis.ListIndex = 0 Then
                .Columns(11).Visible = True
            Else
                .Columns(11).Visible = False
            End If
            
        End With
    End If
    Exit Sub
EH:
    MsgBox "Failed to retrieve Sales History due to error " & Err.Number & " " & Err.Description, vbInformation
End Sub


Private Sub cmdPrint_SalesAnalysis_Click()
    gdxSalesAnalysis.PrinterProperties.Orientation = jgexPPLandscape
    gdxSalesAnalysis.PrintGrid True
End Sub


Private Sub gdxSalesAnalysis_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    gdxSalesAnalysis.PrinterProperties.FooterString(jgexHFRight) = "Page " & PageNumber & " of " & nPages
End Sub


Private Sub gdxSalesAnalysis_DblClick()
    If Not IsEmpty(gdxSalesAnalysis.Value(10)) Then
        Dim oFrm As Form
        Set oFrm = New FInventoryHistory
        If oFrm.ShowHistory(cboWhse_SalesAnalysis.ItemData(cboWhse_SalesAnalysis.ListIndex), gdxSalesAnalysis.Value(10)) Then
            MDIMain.AddNewWindow oFrm
            oFrm.SetCaption "Inventory History (" & cboWhse_SalesAnalysis.text & "): " & gdxSalesAnalysis.Value(1)  ' ItemID
        Else
            Unload oFrm
            Set oFrm = Nothing
        End If
    End If
End Sub


Private Sub gdxSalesAnalysis_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX20.JSRetBoolean)
    'PRN #523 - smr - 2/4/2005 - added logic to be able to change min/max values in the grid
    'BeforeColUpdate contains error handling to determine if the cell should be updated
    With gdxSalesAnalysis
        '***.Value(8) = Min column in gdxSalesAnalysis grid ***
        '***.Value(9) = Max column in gdxSalesAnalysis grid ***
        If .Value(8) = "" Or .Value(9) = "" Then
            Cancel = True
            MsgBox "Please enter a numeric value.", vbInformation, "Min/Max Value"
            Exit Sub
        ElseIf Not IsNumeric(.Value(8)) Or Not IsNumeric(.Value(9)) Then
            Cancel = True
            MsgBox "Please enter a numeric value.", vbInformation, "Min/Max Value"
            Exit Sub
        ElseIf CLng(.Value(8)) > CLng(.Value(9)) Then
            Cancel = True
            MsgBox "'Min' quantity must be less than 'Max' quantity.", vbInformation, "Min/Max Value"
            Exit Sub
        End If
    End With
End Sub


Private Sub gdxSalesAnalysis_AfterColUpdate(ByVal ColIndex As Integer)
    'PRN #523 - smr - 2/4/2005 - added logic to be able to change min/max values in the grid
    'If cancel is set to true in BeforeColUpdate, grid will cancel change & not call AfterColUpdate
    
    Dim lWhseKey As Long
    On Error GoTo EH
                    
    SetWaitCursor True
    
    lWhseKey = cboWhse_SalesAnalysis.ItemData(cboWhse_SalesAnalysis.ListIndex)
        
    '***.Value(8) = Min column in gdxSalesAnalysis grid ***
    '***.Value(9) = Max column in gdxSalesAnalysis grid ***
    With gdxSalesAnalysis
        CallSP "spcpcPOUpdateMinMax", _
                "@_iWhseKey", lWhseKey, _
                "@_iItemKey", .Value(10), _
                "@_iMinQty", CLng(.Value(8)), _
                "@_iMaxQty", CLng(.Value(9))
    End With

    'SMR - this will not be left as is, I need to research!!!
        'this will not produce an error, but take to much time & will be changed!
    cmdFind_SalesAnalysis_Click
    'gdxSalesAnalysis.DataChanged = False

    SetWaitCursor False
    Exit Sub
EH:
    SetWaitCursor False
    msg Err.Number & " " & Err.Description, vbCritical, "gdxSalesAnalysis_AfterColUpdate"
End Sub



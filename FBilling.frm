VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "mmremark.ocx"
Begin VB.Form FBilling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   9855
   Begin ActiveTabs.SSActiveTabs tabMain 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11880
      _Version        =   262144
      TabCount        =   8
      Tabs            =   "FBilling.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel9 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   1
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FBilling.frx":01CC
         Begin VB.ComboBox cboRMACreditSortBy 
            Height          =   315
            ItemData        =   "FBilling.frx":01F4
            Left            =   7920
            List            =   "FBilling.frx":01FE
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   5460
            Width           =   1335
         End
         Begin VB.CommandButton cmdRMAPrint 
            Caption         =   "Print Rep&ort"
            Height          =   375
            Left            =   1080
            TabIndex        =   5
            Top             =   5880
            Width           =   1095
         End
         Begin VB.CommandButton cmdRMAUpdate 
            Caption         =   "&Save"
            Height          =   375
            Index           =   1
            Left            =   6840
            TabIndex        =   4
            Top             =   5880
            Width           =   1095
         End
         Begin VB.CommandButton cmdRMARefresh 
            Caption         =   "&Refresh"
            Height          =   375
            Index           =   1
            Left            =   8100
            TabIndex        =   3
            Top             =   5880
            Width           =   1095
         End
         Begin VB.ComboBox cboWhse 
            Height          =   315
            Index           =   1
            ItemData        =   "FBilling.frx":0212
            Left            =   5820
            List            =   "FBilling.frx":0214
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   5460
            Width           =   1335
         End
         Begin MMRemark.RemarkViewer rvRMACredit 
            Height          =   804
            Left            =   120
            TabIndex        =   7
            Top             =   5484
            Width           =   804
            _ExtentX        =   1429
            _ExtentY        =   1402
            ContextID       =   "ViewRMA"
            Caption         =   "RMA Remarks"
         End
         Begin GridEX20.GridEX gdxRMACred 
            Height          =   5235
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   9234
            Version         =   "2.0"
            ScrollToolTips  =   -1  'True
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            ColumnsCount    =   24
            Column(1)       =   "FBilling.frx":0216
            Column(2)       =   "FBilling.frx":0376
            Column(3)       =   "FBilling.frx":04E2
            Column(4)       =   "FBilling.frx":0636
            Column(5)       =   "FBilling.frx":0762
            Column(6)       =   "FBilling.frx":08DE
            Column(7)       =   "FBilling.frx":0A9E
            Column(8)       =   "FBilling.frx":0BE2
            Column(9)       =   "FBilling.frx":0D42
            Column(10)      =   "FBilling.frx":0EA2
            Column(11)      =   "FBilling.frx":103E
            Column(12)      =   "FBilling.frx":1152
            Column(13)      =   "FBilling.frx":12FA
            Column(14)      =   "FBilling.frx":1486
            Column(15)      =   "FBilling.frx":15E6
            Column(16)      =   "FBilling.frx":1752
            Column(17)      =   "FBilling.frx":186E
            Column(18)      =   "FBilling.frx":1982
            Column(19)      =   "FBilling.frx":1B26
            Column(20)      =   "FBilling.frx":1CB2
            Column(21)      =   "FBilling.frx":1E0A
            Column(22)      =   "FBilling.frx":1F42
            Column(23)      =   "FBilling.frx":20CE
            Column(24)      =   "FBilling.frx":21FE
            GroupCount      =   1
            Group(1)        =   "FBilling.frx":239A
            FormatStylesCount=   6
            FormatStyle(1)  =   "FBilling.frx":2402
            FormatStyle(2)  =   "FBilling.frx":253A
            FormatStyle(3)  =   "FBilling.frx":25EA
            FormatStyle(4)  =   "FBilling.frx":269E
            FormatStyle(5)  =   "FBilling.frx":2776
            FormatStyle(6)  =   "FBilling.frx":282E
            ImageCount      =   0
            PrinterProperties=   "FBilling.frx":290E
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sort By"
            Height          =   195
            Index           =   2
            Left            =   7365
            TabIndex        =   10
            Top             =   5520
            Width           =   510
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Warehouse"
            Height          =   255
            Index           =   1
            Left            =   4860
            TabIndex        =   9
            Top             =   5520
            Width           =   915
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   11
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FBilling.frx":2AE6
         Begin VB.Frame Frame13 
            Caption         =   "Lookup Sage SO#"
            Height          =   912
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   8592
            Begin VB.CommandButton cmdFindSOID 
               Caption         =   "Find"
               Height          =   312
               Left            =   2400
               TabIndex        =   38
               Top             =   360
               Width           =   912
            End
            Begin VB.TextBox txtOPID 
               Height          =   312
               Left            =   780
               TabIndex        =   37
               Top             =   360
               Width           =   1452
            End
            Begin VB.Label Label48 
               Caption         =   "SOID"
               Height          =   252
               Left            =   3540
               TabIndex        =   41
               Top             =   420
               Width           =   432
            End
            Begin VB.Label lblSOID 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   312
               Left            =   4080
               TabIndex        =   40
               Top             =   360
               Width           =   1452
            End
            Begin VB.Label Label46 
               Caption         =   "OPID"
               Height          =   192
               Left            =   240
               TabIndex        =   39
               Top             =   420
               Width           =   552
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Add Customer PO to Order"
            Height          =   1395
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   8652
            Begin VB.CommandButton cmdUpdateSO 
               Caption         =   "Update"
               Height          =   312
               Left            =   2760
               TabIndex        =   32
               Top             =   780
               Width           =   912
            End
            Begin VB.CommandButton cmdFindSO 
               Caption         =   "Find"
               Height          =   312
               Left            =   2760
               TabIndex        =   31
               Top             =   360
               Width           =   912
            End
            Begin VB.TextBox txtPO 
               Height          =   312
               Left            =   1560
               TabIndex        =   30
               Top             =   780
               Width           =   912
            End
            Begin VB.TextBox txtSO 
               Height          =   312
               Left            =   1560
               TabIndex        =   29
               Top             =   360
               Width           =   912
            End
            Begin VB.Label lblSOInfo 
               BackColor       =   &H80000005&
               Height          =   792
               Left            =   4200
               TabIndex        =   35
               Top             =   240
               Width           =   4212
            End
            Begin VB.Label Label39 
               Caption         =   "Customer PO#"
               Height          =   315
               Left            =   300
               TabIndex        =   34
               Top             =   780
               Width           =   1155
            End
            Begin VB.Label Label40 
               Caption         =   "Sales Order"
               Height          =   255
               Left            =   300
               TabIndex        =   33
               Top             =   420
               Width           =   1035
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Price Pack Slip"
            Height          =   1815
            Left            =   120
            TabIndex        =   12
            Top             =   2760
            Width           =   8655
            Begin VB.TextBox txtAcuitySO 
               Height          =   285
               Left            =   2040
               TabIndex        =   16
               Top             =   240
               Width           =   1695
            End
            Begin VB.CommandButton cmdLookUp 
               Caption         =   "&Look Up"
               Enabled         =   0   'False
               Height          =   375
               Left            =   7440
               TabIndex        =   15
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox chkPricePacking 
               Caption         =   "Price Packing Slip"
               Height          =   255
               Left            =   5040
               TabIndex        =   14
               Top             =   1320
               Width           =   1695
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "&Update"
               Enabled         =   0   'False
               Height          =   375
               Left            =   7440
               TabIndex        =   13
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblAcuitySO 
               Alignment       =   1  'Right Justify
               Caption         =   "Sage SO#"
               Height          =   375
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblAssCustName 
               Alignment       =   1  'Right Justify
               Caption         =   "Cust Name"
               Height          =   255
               Left            =   480
               TabIndex        =   26
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lblAssistCustName 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1560
               TabIndex        =   25
               Top             =   600
               Width           =   2175
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "CSR"
               Height          =   255
               Left            =   480
               TabIndex        =   24
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lblAssistCSR 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2040
               TabIndex        =   23
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label11 
               Caption         =   "Ship To City"
               Height          =   252
               Left            =   3960
               TabIndex        =   22
               Top             =   600
               Width           =   972
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "Order Date"
               Height          =   375
               Left            =   3960
               TabIndex        =   21
               Top             =   960
               Width           =   855
            End
            Begin VB.Label lblShipToCity 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5040
               TabIndex        =   20
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label lblOrderDate 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5040
               TabIndex        =   19
               Top             =   960
               Width           =   1935
            End
            Begin VB.Label lblAssistCustID 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2040
               TabIndex        =   18
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "Cust ID"
               Height          =   255
               Left            =   480
               TabIndex        =   17
               Top             =   960
               Width           =   975
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel8 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   42
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FBilling.frx":2B0E
         Begin VB.CommandButton cmdFindWCOrder 
            Caption         =   "Find"
            Height          =   315
            Left            =   4200
            TabIndex        =   135
            Top             =   120
            Width           =   795
         End
         Begin VB.TextBox txtFindWCOrder 
            Height          =   345
            Left            =   2880
            TabIndex        =   134
            Top             =   120
            Width           =   1155
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   312
            Left            =   8220
            TabIndex        =   43
            Top             =   120
            Width           =   1092
         End
         Begin GridEX20.GridEX gdxWillCallOrders 
            Height          =   5655
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   9975
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
            ColumnsCount    =   7
            Column(1)       =   "FBilling.frx":2B36
            Column(2)       =   "FBilling.frx":2C66
            Column(3)       =   "FBilling.frx":2D72
            Column(4)       =   "FBilling.frx":2E9A
            Column(5)       =   "FBilling.frx":2FBE
            Column(6)       =   "FBilling.frx":3102
            Column(7)       =   "FBilling.frx":323A
            FmtConditionsCount=   1
            FmtCondition(1) =   "FBilling.frx":34C6
            FormatStylesCount=   6
            FormatStyle(1)  =   "FBilling.frx":35F6
            FormatStyle(2)  =   "FBilling.frx":36D6
            FormatStyle(3)  =   "FBilling.frx":380E
            FormatStyle(4)  =   "FBilling.frx":38BE
            FormatStyle(5)  =   "FBilling.frx":3972
            FormatStyle(6)  =   "FBilling.frx":3A4A
            ImageCount      =   0
            PrinterProperties=   "FBilling.frx":3B02
         End
         Begin VB.Label Label9 
            Caption         =   "OP#"
            Height          =   255
            Left            =   2400
            TabIndex        =   137
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblOpenWCOrders 
            Caption         =   "open Will Call orders"
            Height          =   255
            Left            =   180
            TabIndex        =   136
            Top             =   180
            Width           =   2115
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   45
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FBilling.frx":3CDA
         Begin VB.CommandButton cmdBillingSummary 
            Caption         =   "&Look Up"
            Height          =   300
            Left            =   4620
            TabIndex        =   47
            Top             =   360
            Width           =   972
         End
         Begin VB.TextBox txtBatchNumber 
            Height          =   285
            Left            =   2340
            TabIndex        =   46
            Top             =   360
            Width           =   1935
         End
         Begin GridEX20.GridEX gdxCCRD 
            Height          =   1368
            Left            =   240
            TabIndex        =   48
            Top             =   4920
            Visible         =   0   'False
            Width           =   8952
            _ExtentX        =   15796
            _ExtentY        =   2408
            Version         =   "2.0"
            ScrollToolTips  =   -1  'True
            ShowToolTips    =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   10
            Column(1)       =   "FBilling.frx":3D02
            Column(2)       =   "FBilling.frx":3F76
            Column(3)       =   "FBilling.frx":409A
            Column(4)       =   "FBilling.frx":41BE
            Column(5)       =   "FBilling.frx":42EE
            Column(6)       =   "FBilling.frx":448E
            Column(7)       =   "FBilling.frx":45D6
            Column(8)       =   "FBilling.frx":470A
            Column(9)       =   "FBilling.frx":484A
            Column(10)      =   "FBilling.frx":4996
            FormatStylesCount=   6
            FormatStyle(1)  =   "FBilling.frx":4ADE
            FormatStyle(2)  =   "FBilling.frx":4C16
            FormatStyle(3)  =   "FBilling.frx":4CC6
            FormatStyle(4)  =   "FBilling.frx":4D7A
            FormatStyle(5)  =   "FBilling.frx":4E52
            FormatStyle(6)  =   "FBilling.frx":4F0A
            ImageCount      =   0
            PrinterProperties=   "FBilling.frx":4FEA
         End
         Begin GridEX20.GridEX gdxDI 
            Height          =   1428
            Left            =   240
            TabIndex        =   49
            Top             =   3120
            Width           =   8952
            _ExtentX        =   15796
            _ExtentY        =   2514
            Version         =   "2.0"
            ScrollToolTips  =   -1  'True
            ShowToolTips    =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   4
            Column(1)       =   "FBilling.frx":51C2
            Column(2)       =   "FBilling.frx":530A
            Column(3)       =   "FBilling.frx":542E
            Column(4)       =   "FBilling.frx":555E
            FormatStylesCount=   6
            FormatStyle(1)  =   "FBilling.frx":56FE
            FormatStyle(2)  =   "FBilling.frx":5836
            FormatStyle(3)  =   "FBilling.frx":58E6
            FormatStyle(4)  =   "FBilling.frx":599A
            FormatStyle(5)  =   "FBilling.frx":5A72
            FormatStyle(6)  =   "FBilling.frx":5B2A
            ImageCount      =   0
            PrinterProperties=   "FBilling.frx":5C0A
         End
         Begin GridEX20.GridEX gdxCOD 
            Height          =   1428
            Left            =   240
            TabIndex        =   50
            Top             =   1260
            Width           =   8952
            _ExtentX        =   15796
            _ExtentY        =   2514
            Version         =   "2.0"
            ScrollToolTips  =   -1  'True
            ShowToolTips    =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   4
            Column(1)       =   "FBilling.frx":5DE2
            Column(2)       =   "FBilling.frx":5F2A
            Column(3)       =   "FBilling.frx":604E
            Column(4)       =   "FBilling.frx":617E
            FormatStylesCount=   6
            FormatStyle(1)  =   "FBilling.frx":631E
            FormatStyle(2)  =   "FBilling.frx":6456
            FormatStyle(3)  =   "FBilling.frx":6506
            FormatStyle(4)  =   "FBilling.frx":65BA
            FormatStyle(5)  =   "FBilling.frx":6692
            FormatStyle(6)  =   "FBilling.frx":674A
            ImageCount      =   0
            PrinterProperties=   "FBilling.frx":682A
         End
         Begin VB.Label Label37 
            Caption         =   "Credit Card Billing Summary"
            Height          =   252
            Left            =   300
            TabIndex        =   54
            Top             =   4680
            Visible         =   0   'False
            Width           =   2052
         End
         Begin VB.Label Label38 
            Caption         =   "Duplicate Invoice Billing Summary"
            Height          =   252
            Left            =   300
            TabIndex        =   53
            Top             =   2820
            Width           =   2532
         End
         Begin VB.Label Label35 
            Caption         =   "Invoice Batch Number"
            Height          =   252
            Left            =   300
            TabIndex        =   52
            Top             =   420
            Width           =   1752
         End
         Begin VB.Label Label47 
            Caption         =   "Confirming Copy Only Billing"
            Height          =   252
            Left            =   300
            TabIndex        =   51
            Top             =   960
            Width           =   2532
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   55
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FBilling.frx":6A02
         Begin VB.ComboBox cboWhse 
            Height          =   315
            Index           =   0
            ItemData        =   "FBilling.frx":6A2A
            Left            =   7920
            List            =   "FBilling.frx":6A2C
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   5520
            Width           =   1335
         End
         Begin VB.CommandButton cmdRMARefresh 
            Caption         =   "&Refresh"
            Height          =   375
            Index           =   0
            Left            =   8160
            TabIndex        =   57
            Top             =   5940
            Width           =   1095
         End
         Begin VB.CommandButton cmdRMAUpdate 
            Caption         =   "&Save"
            Height          =   375
            Index           =   0
            Left            =   6960
            TabIndex        =   56
            Top             =   5940
            Width           =   1095
         End
         Begin MMRemark.RemarkViewer rvARRemarks 
            Height          =   810
            Left            =   1080
            TabIndex        =   59
            Top             =   5460
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1429
            ContextID       =   "RMACustLoad"
            Caption         =   "A/R Remarks"
         End
         Begin MMRemark.RemarkViewer rvRMAApprove 
            Height          =   810
            Left            =   180
            TabIndex        =   60
            Top             =   5460
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1429
            ContextID       =   "ViewRMA"
            Caption         =   "RMA Remarks"
         End
         Begin GridEX20.GridEX gdxRMAApproval 
            Height          =   5295
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   9340
            Version         =   "2.0"
            ScrollToolTips  =   -1  'True
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            ColumnsCount    =   21
            Column(1)       =   "FBilling.frx":6A2E
            Column(2)       =   "FBilling.frx":6B8E
            Column(3)       =   "FBilling.frx":6CFA
            Column(4)       =   "FBilling.frx":6E4E
            Column(5)       =   "FBilling.frx":6FC2
            Column(6)       =   "FBilling.frx":7106
            Column(7)       =   "FBilling.frx":7266
            Column(8)       =   "FBilling.frx":73C6
            Column(9)       =   "FBilling.frx":7562
            Column(10)      =   "FBilling.frx":768E
            Column(11)      =   "FBilling.frx":7836
            Column(12)      =   "FBilling.frx":79C2
            Column(13)      =   "FBilling.frx":7B22
            Column(14)      =   "FBilling.frx":7C8E
            Column(15)      =   "FBilling.frx":7DB2
            Column(16)      =   "FBilling.frx":7ECE
            Column(17)      =   "FBilling.frx":7FE2
            Column(18)      =   "FBilling.frx":8186
            Column(19)      =   "FBilling.frx":8312
            Column(20)      =   "FBilling.frx":846A
            Column(21)      =   "FBilling.frx":85A2
            GroupCount      =   1
            Group(1)        =   "FBilling.frx":8722
            FormatStylesCount=   6
            FormatStyle(1)  =   "FBilling.frx":878A
            FormatStyle(2)  =   "FBilling.frx":88C2
            FormatStyle(3)  =   "FBilling.frx":8972
            FormatStyle(4)  =   "FBilling.frx":8A26
            FormatStyle(5)  =   "FBilling.frx":8AFE
            FormatStyle(6)  =   "FBilling.frx":8BB6
            ImageCount      =   0
            PrinterProperties=   "FBilling.frx":8C96
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Warehouse"
            Height          =   252
            Index           =   0
            Left            =   6960
            TabIndex        =   62
            Top             =   5520
            Width           =   912
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   63
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FBilling.frx":8E6E
         Begin VB.Frame Frame12 
            Caption         =   "Pre-Post Shipment Check"
            Height          =   6135
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   9015
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   4320
               TabIndex        =   131
               Top             =   2880
               Width           =   3135
               Begin VB.OptionButton Option2 
                  Caption         =   "Exceptions"
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   133
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "All"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   132
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   975
               End
            End
            Begin VB.CommandButton cmdPrintExceptionReport 
               Caption         =   "Print"
               Height          =   375
               Left            =   7680
               TabIndex        =   130
               Top             =   2880
               Width           =   1095
            End
            Begin VB.ComboBox cboWhse 
               Height          =   315
               Index           =   2
               ItemData        =   "FBilling.frx":8E96
               Left            =   6240
               List            =   "FBilling.frx":8E98
               Style           =   2  'Dropdown List
               TabIndex        =   128
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton cmdGetBatches 
               Caption         =   "Refresh"
               Height          =   375
               Left            =   7680
               TabIndex        =   67
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtShipWarnings 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2535
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   66
               Top             =   3360
               Width           =   8532
            End
            Begin VB.CommandButton cmdBalanceFreight 
               Caption         =   "Balance Frgt"
               Height          =   375
               Left            =   7680
               TabIndex        =   65
               Top             =   840
               Width           =   1095
            End
            Begin GridEX20.GridEX gdxShipmentBatch 
               Height          =   2055
               Left            =   240
               TabIndex        =   68
               Top             =   840
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   3625
               Version         =   "2.0"
               HoldSortSettings=   -1  'True
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               MethodHoldFields=   -1  'True
               AllowEdit       =   0   'False
               GroupByBoxVisible=   0   'False
               ColumnHeaderHeight=   285
               IntProp1        =   0
               ColumnsCount    =   6
               Column(1)       =   "FBilling.frx":8E9A
               Column(2)       =   "FBilling.frx":904E
               Column(3)       =   "FBilling.frx":91C2
               Column(4)       =   "FBilling.frx":9356
               Column(5)       =   "FBilling.frx":951A
               Column(6)       =   "FBilling.frx":9692
               FormatStylesCount=   6
               FormatStyle(1)  =   "FBilling.frx":97F2
               FormatStyle(2)  =   "FBilling.frx":992A
               FormatStyle(3)  =   "FBilling.frx":99DA
               FormatStyle(4)  =   "FBilling.frx":9A8E
               FormatStyle(5)  =   "FBilling.frx":9B66
               FormatStyle(6)  =   "FBilling.frx":9C1E
               ImageCount      =   0
               PrinterProperties=   "FBilling.frx":9CFE
            End
            Begin VB.Label Label5 
               Caption         =   "Warehouse"
               Height          =   255
               Left            =   5160
               TabIndex        =   129
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label43 
               Caption         =   "Select a Shipment Batch"
               Height          =   255
               Left            =   240
               TabIndex        =   70
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label42 
               Caption         =   "Pre-Post Exception Report"
               Height          =   255
               Left            =   240
               TabIndex        =   69
               Top             =   3120
               Width           =   2055
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6345
         Left            =   30
         TabIndex        =   71
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FBilling.frx":9ED6
         Begin VB.Frame Frame11 
            Caption         =   "Search Existing Customer"
            Height          =   1875
            Left            =   5640
            TabIndex        =   110
            Top             =   120
            Width           =   3735
            Begin VB.CommandButton cmdFindCustomer 
               Caption         =   "Find"
               Height          =   375
               Left            =   2400
               TabIndex        =   112
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox txtZipCode 
               Height          =   285
               Left            =   1560
               TabIndex        =   111
               Top             =   720
               Width           =   1935
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtPhoneNumber 
               Height          =   285
               Left            =   1560
               TabIndex        =   113
               Top             =   360
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
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
               mask            =   "(###)###-####"
               text            =   "          "
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               Caption         =   "Zip Code"
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               Caption         =   "Phone Number"
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Create Sage Account"
            Height          =   6135
            Left            =   120
            TabIndex        =   77
            Top             =   120
            Width           =   5295
            Begin VB.TextBox txtCoName 
               Height          =   285
               Left            =   1680
               TabIndex        =   91
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox txtZip 
               Height          =   285
               Left            =   1680
               TabIndex        =   90
               Top             =   720
               Width           =   1935
            End
            Begin VB.OptionButton optEndUser 
               Caption         =   "End User"
               Height          =   255
               Left            =   240
               TabIndex        =   89
               Top             =   2520
               Width           =   1095
            End
            Begin VB.OptionButton optDealer 
               Caption         =   "Dealer"
               Height          =   255
               Left            =   1440
               TabIndex        =   88
               Top             =   2520
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optWholesaler 
               Caption         =   "Wholesaler"
               Height          =   255
               Left            =   2400
               TabIndex        =   87
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtStreet 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   86
               Top             =   1080
               Width           =   1935
            End
            Begin VB.TextBox txtCity 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   85
               Top             =   1440
               Width           =   1935
            End
            Begin VB.TextBox txtState 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   84
               Top             =   1800
               Width           =   1935
            End
            Begin VB.TextBox txtCountry 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   83
               Top             =   2160
               Width           =   1935
            End
            Begin VB.TextBox txtTaxSchdKey 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   82
               Top             =   5520
               Width           =   735
            End
            Begin VB.TextBox txtTaxRate 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   81
               Top             =   5136
               Width           =   615
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "Print"
               Height          =   375
               Left            =   3960
               TabIndex        =   80
               Top             =   5160
               Width           =   1095
            End
            Begin VB.CommandButton cmdClear 
               Caption         =   "Clear"
               Height          =   375
               Left            =   3960
               TabIndex        =   79
               Top             =   5640
               Width           =   1095
            End
            Begin VB.CheckBox chkGovt 
               Caption         =   "Government"
               Height          =   375
               Left            =   3840
               TabIndex        =   78
               Top             =   2460
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Company Name"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   109
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Billing Zip"
               Height          =   255
               Left            =   480
               TabIndex        =   108
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "Street"
               Height          =   255
               Left            =   240
               TabIndex        =   107
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               Caption         =   "City"
               Height          =   255
               Left            =   360
               TabIndex        =   106
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "State"
               Height          =   255
               Left            =   480
               TabIndex        =   105
               Top             =   1800
               Width           =   855
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "Country"
               Height          =   255
               Left            =   600
               TabIndex        =   104
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label lblTaxSchdKey 
               Caption         =   "Tax Schedule Key"
               Height          =   255
               Left            =   60
               TabIndex        =   103
               Top             =   5520
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "Tax Rate"
               Height          =   255
               Left            =   660
               TabIndex        =   102
               Top             =   5160
               Width           =   735
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "Cust ID"
               Height          =   255
               Left            =   240
               TabIndex        =   101
               Top             =   3600
               Width           =   1155
            End
            Begin VB.Label lblCustName 
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   3240
               Width           =   3495
            End
            Begin VB.Label lblWarehouse 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1680
               TabIndex        =   99
               Top             =   4752
               Width           =   1935
            End
            Begin VB.Label lblTerritory 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1680
               TabIndex        =   98
               Top             =   4368
               Width           =   1935
            End
            Begin VB.Label lblTax 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1680
               TabIndex        =   97
               Top             =   3984
               Width           =   1935
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Warehouse"
               Height          =   255
               Left            =   240
               TabIndex        =   96
               Top             =   4740
               Width           =   1155
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Sales Territory"
               Height          =   255
               Left            =   240
               TabIndex        =   95
               Top             =   4380
               Width           =   1155
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Tax Schedule"
               Height          =   255
               Left            =   240
               TabIndex        =   94
               Top             =   3960
               Width           =   1155
            End
            Begin VB.Label lblAcuity 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1680
               TabIndex        =   93
               Top             =   3600
               Width           =   1935
            End
            Begin VB.Label Label4 
               Caption         =   "Sage Account Info"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   92
               Top             =   2880
               Width           =   3975
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Manage Sales Tax Exemption Certificates"
            Height          =   4095
            Left            =   5640
            TabIndex        =   72
            Top             =   2160
            Width           =   3735
            Begin VB.TextBox txtCustID 
               Height          =   285
               Left            =   240
               TabIndex        =   74
               Top             =   600
               Width           =   1815
            End
            Begin VB.CommandButton cmdFindExemption 
               Caption         =   "Find"
               Height          =   375
               Left            =   2400
               TabIndex        =   73
               Top             =   600
               Width           =   1095
            End
            Begin GridEX20.GridEX gdxSTaxState 
               Height          =   255
               Left            =   360
               TabIndex        =   75
               Top             =   3600
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               Version         =   "2.0"
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               MethodHoldFields=   -1  'True
               Options         =   8
               RecordsetType   =   1
               DataMode        =   1
               ColumnHeaderHeight=   285
               ColumnsCount    =   2
               Column(1)       =   "FBilling.frx":9EFE
               Column(2)       =   "FBilling.frx":9FC6
               FormatStylesCount=   6
               FormatStyle(1)  =   "FBilling.frx":A06A
               FormatStyle(2)  =   "FBilling.frx":A14A
               FormatStyle(3)  =   "FBilling.frx":A282
               FormatStyle(4)  =   "FBilling.frx":A332
               FormatStyle(5)  =   "FBilling.frx":A3E6
               FormatStyle(6)  =   "FBilling.frx":A4BE
               ImageCount      =   0
               PrinterProperties=   "FBilling.frx":A576
            End
            Begin GridEX20.GridEX gdxSTaxCert 
               Height          =   1815
               Left            =   240
               TabIndex        =   76
               Top             =   1560
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   3201
               Version         =   "2.0"
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               Options         =   8
               RecordsetType   =   1
               DataMode        =   1
               ColumnHeaderHeight=   285
               ColumnsCount    =   2
               Column(1)       =   "FBilling.frx":A74E
               Column(2)       =   "FBilling.frx":A816
               FormatStylesCount=   6
               FormatStyle(1)  =   "FBilling.frx":A8BA
               FormatStyle(2)  =   "FBilling.frx":A99A
               FormatStyle(3)  =   "FBilling.frx":AAD2
               FormatStyle(4)  =   "FBilling.frx":AB82
               FormatStyle(5)  =   "FBilling.frx":AC36
               FormatStyle(6)  =   "FBilling.frx":AD0E
               ImageCount      =   0
               PrinterProperties=   "FBilling.frx":ADC6
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel10 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   116
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FBilling.frx":AF9E
         Begin VB.CommandButton cmdCreateReport 
            Caption         =   "Create Report"
            Height          =   372
            Left            =   6720
            TabIndex        =   120
            Top             =   240
            Width           =   1272
         End
         Begin VB.CommandButton cmdChargeCC 
            Caption         =   "Charge Credit Card Now"
            Height          =   372
            Left            =   4620
            TabIndex        =   119
            Top             =   240
            Width           =   1932
         End
         Begin VB.TextBox txtBatchNumberCC 
            Height          =   312
            Left            =   1920
            TabIndex        =   118
            Top             =   300
            Width           =   1032
         End
         Begin VB.CommandButton cmdGetBatchCC 
            Caption         =   "Get Batch"
            Height          =   372
            Left            =   3240
            TabIndex        =   117
            Top             =   240
            Width           =   1215
         End
         Begin GridEX20.GridEX gdxCCCharged 
            Height          =   2052
            Left            =   120
            TabIndex        =   121
            Top             =   3960
            Width           =   9192
            _ExtentX        =   16219
            _ExtentY        =   3625
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   5
            Column(1)       =   "FBilling.frx":AFC6
            Column(2)       =   "FBilling.frx":B132
            Column(3)       =   "FBilling.frx":B276
            Column(4)       =   "FBilling.frx":B3C6
            Column(5)       =   "FBilling.frx":B566
            FormatStylesCount=   6
            FormatStyle(1)  =   "FBilling.frx":B6AE
            FormatStyle(2)  =   "FBilling.frx":B78E
            FormatStyle(3)  =   "FBilling.frx":B8C6
            FormatStyle(4)  =   "FBilling.frx":B976
            FormatStyle(5)  =   "FBilling.frx":BA2A
            FormatStyle(6)  =   "FBilling.frx":BB02
            ImageCount      =   0
            PrinterProperties=   "FBilling.frx":BBBA
         End
         Begin GridEX20.GridEX gdxCCOrders 
            Height          =   2292
            Left            =   120
            TabIndex        =   122
            Top             =   1200
            Width           =   9192
            _ExtentX        =   16219
            _ExtentY        =   4048
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
            ColumnsCount    =   8
            Column(1)       =   "FBilling.frx":BD92
            Column(2)       =   "FBilling.frx":BF2A
            Column(3)       =   "FBilling.frx":C072
            Column(4)       =   "FBilling.frx":C1C2
            Column(5)       =   "FBilling.frx":C362
            Column(6)       =   "FBilling.frx":C4AE
            Column(7)       =   "FBilling.frx":C612
            Column(8)       =   "FBilling.frx":C772
            FormatStylesCount=   6
            FormatStyle(1)  =   "FBilling.frx":C89A
            FormatStyle(2)  =   "FBilling.frx":C97A
            FormatStyle(3)  =   "FBilling.frx":CAB2
            FormatStyle(4)  =   "FBilling.frx":CB62
            FormatStyle(5)  =   "FBilling.frx":CC16
            FormatStyle(6)  =   "FBilling.frx":CCEE
            ImageCount      =   0
            PrinterProperties=   "FBilling.frx":CDA6
         End
         Begin VB.Label Label51 
            Caption         =   "Credit Card Orders Charged"
            Height          =   252
            Left            =   600
            TabIndex        =   127
            Top             =   3660
            Width           =   2172
         End
         Begin VB.Label Label50 
            Caption         =   "Credit Card Orders Not Yet Charged"
            Height          =   192
            Left            =   540
            TabIndex        =   126
            Top             =   900
            Width           =   2712
         End
         Begin VB.Label Label49 
            Caption         =   "Invoice Batch Number"
            Height          =   252
            Left            =   180
            TabIndex        =   125
            Top             =   360
            Width           =   1752
         End
         Begin VB.Label lblCardCount 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   180
            TabIndex        =   124
            Top             =   840
            Width           =   312
         End
         Begin VB.Label lblChargeCount 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   180
            TabIndex        =   123
            Top             =   3600
            Width           =   312
         End
      End
   End
   Begin MSComctlLib.ImageList imglRemarks 
      Left            =   9000
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
            Picture         =   "FBilling.frx":CF7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FBilling.frx":D3D0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************
'ActiveTab control tab stuff
'Warning: if you change this stuff, check EnableTabs() as well
'*******************************************************************
Private Const klNumTabs = 8
Private m_asTabRights(1 To klNumTabs) As String

Private Enum SSTBillingIndexes
    tmiAccount = 1
    tmiAssist = 2
    tmiShipment = 3
    tmiRMAApprovalMgr = 4
    tmiRMACredMgr = 5
    tmiWillCall = 6
    tmiCreditCard = 7
    tmiSummary = 8
End Enum

Private Const k_lSearchByCustID = 0
Private Const k_lSearchByRMAID = 1
Private Const k_lSearchByOPID = 2

Private m_lWindowID  As Long
Private m_bLoading As Boolean

Private m_lAddPOOPKey As Long
Private m_lAddPOSOKey As Long

Private m_lOPKey As Long
Private m_lRMAKey As Long
Private m_lRMACustKey As Long
Private m_lSOKey As Long

Private m_orstWillCallOrders As ADODB.Recordset
Private m_oRstSO As ADODB.Recordset
Private m_oRstPO As ADODB.Recordset

Private m_oRMAApprovalList As RMAList
Private m_oRMACreditList As RMAList

Private m_colReason As Collection
Private m_colDisposition As Collection

Private WithEvents m_gwShipment As GridEXWrapper
Attribute m_gwShipment.VB_VarHelpID = -1
Private WithEvents m_gwWillCallOrders As GridEXWrapper
Attribute m_gwWillCallOrders.VB_VarHelpID = -1

Private m_arrayCCOrders() As Variant

Private m_sZip As String
Private m_sCountry As String
Private m_sState As String
Private m_sCity As String

'*** Sales Tax Exemption Stuff ***
Const iRH = 300
Dim m_lCustKey As Long


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



Private Sub cboWhse_Click(Index As Integer)
    If (Index = 2) Then
        SetWaitCursor True
        GetShipmentBatches
        SetWaitCursor False
    End If
End Sub

Private Sub chkGovt_Click()
    Validate
End Sub


Private Sub cmdPrintExceptionReport_Click()
    Dim sType As String
    Dim iType As Integer
    Dim sCreateDate As String
    Dim sPackStation As String
    Dim sTemp As String
    Dim bExceptionsOnly As Boolean
    
    sCreateDate = Trim(CStr(m_gwShipment.value("CreateDate")))
    sType = Trim(CStr(m_gwShipment.value("TypeId")))
    iType = CInt(m_gwShipment.value("TypeKey"))

    If Option1.value = True Then
        bExceptionsOnly = False
    ElseIf Option2.value = True Then
        bExceptionsOnly = True
    End If
    
    
    If Not IsNull(m_gwShipment.value("PackStation")) Then
        sPackStation = m_gwShipment.value("PackStation")
    End If
    
    If sPackStation <> "" Then
        sTemp = "Shipment Details Whse: " & cboWhse(2).text & " Type: " & sType & " Station: " & sPackStation & " DATE: " & sCreateDate & vbCrLf
    Else
        sTemp = "Shipment Details Whse: " & cboWhse(2).text & " Type: " & sType & " DATE: " & sCreateDate & vbCrLf
    End If
   
    txtShipWarnings.text = sTemp & GetShipWarnings(sCreateDate, cboWhse(2).ItemData(cboWhse(2).ListIndex), iType, sPackStation, bExceptionsOnly)
     
    SetWaitCursor True

    Printer.Font = "Courier New"
    Printer.FontSize = 14
    Printer.Print "Pre-post Exception Report " & Date
    Printer.Print
    Printer.FontSize = 12
    
    If sPackStation <> "" Then
        Printer.Print "Shipment Details Whse:" & cboWhse(2).text & " Type: " & sType & " - " & sPackStation & " DATE: " & sCreateDate
    Else
        Printer.Print "Shipment Details Whse:" & cboWhse(2).text & " Type: " & sType & " DATE: " & sCreateDate
    End If
    
    Printer.FontSize = 10
    Printer.Print GetShipWarnings(sCreateDate, cboWhse(2).ItemData(cboWhse(2).ListIndex), iType, sPackStation, bExceptionsOnly)
    Printer.Print
    
    Printer.EndDoc
    
    SetWaitCursor False
    msg "Report printed on " & Printer.DeviceName
End Sub



Private Sub PrintExceptionsAndShipments()

End Sub
'***************************************************************************
'Form events
'***************************************************************************

Private Sub Form_Load()

    SetCaption "Billing Management"
    
    m_bLoading = True
    
    EnableTabs
    
    LoadImageList imglRemarks, gdxWillCallOrders
    Set m_gwWillCallOrders = New GridEXWrapper
    m_gwWillCallOrders.Grid = gdxWillCallOrders
    
    If tabMain.Tabs(tmiShipment).Visible Then
        Set m_gwShipment = New GridEXWrapper
        m_gwShipment.Grid = gdxShipmentBatch
        
        g_rstWhses.Filter = "transit = 0"
        LoadCombo cboWhse(2), g_rstWhses, "WhseID", "WhseKey"
        g_rstWhses.Filter = adFilterNone
        cboWhse(2).text = GetUserWhseID
        
                
        
        'SetComboByKey cboWhse(2), GetUserWhseKey
    
    
    End If
    
    If tabMain.Tabs(tmiRMAApprovalMgr).Visible Or tabMain.Tabs(tmiRMACredMgr).Visible Then
        
        Set m_colReason = Billing.GetRMAReasons()
        Set m_colDisposition = Billing.GetRMADispositions()
        
        If tabMain.Tabs(tmiRMAApprovalMgr).Visible Then
            'filter out transit warehouses
            g_rstWhses.Filter = "transit = 0"
            LoadCombo cboWhse(0), g_rstWhses, "WhseID", "WhseKey"
            g_rstWhses.Filter = adFilterNone
            
            cboWhse(0).text = GetUserWhseID
        End If
        
        If tabMain.Tabs(tmiRMACredMgr).Visible Then
            g_rstWhses.Filter = "transit = 0"
            LoadCombo cboWhse(1), g_rstWhses, "WhseID", "WhseKey"
            g_rstWhses.Filter = adFilterNone
            cboWhse(1).AddItem "All", 0
            cboWhse(1).ListIndex = 0
        End If

' Can't be true due to surrounding condition - LR 12/16/14
'        If tabMain.Tabs(tmiDropShip).Visible Then
'            g_rstWhses.Filter = "transit = 0"
'            LoadCombo cboWhse(2), g_rstWhses, "WhseID", "WhseKey"
'            g_rstWhses.Filter = adFilterNone
'            cboWhse(2).AddItem "All", 0
'            cboWhse(2).ListIndex = 0
'        End If
    End If
       
    cboRMACreditSortBy.ListIndex = 0
       
'*** Sales Tax Exemption Stuff ***
    
    Dim colTemp As JSColumn
    
    Set gdxSTaxState.ADORecordset = Billing.GetNexusStates()

    With gdxSTaxState
        .ColumnAutoResize = True
        .RecordsetType = jgexRSADOStatic
        .CursorLocation = jgexUseClient
        .LockType = jgexLockReadOnly
        .Options = adCmdText
        .ColumnHeaders = False
        .width = 690
        .Height = (iRH * .RowCount)
        .ActAsDropDown = True
        'Required when setting the ActAsDropDown property = True
        'Also, the Parent control must a column with EditType = jgexCombo
        .AllowAddNew = False
        .AllowDelete = False
        .AllowEdit = False
        .ContinuousScroll = True
        .DetectRowDrag = False
        .GroupByBoxVisible = False
        .HideSelection = jgexHighLightNormal
        .MultiSelect = False
        .View = jgexTable
    End With
    
    With gdxSTaxCert
        .RecordsetType = jgexRSADOKeyset
        .CursorLocation = jgexUseClient
        .LockType = jgexLockOptimistic
        .Options = adCmdText
        .GroupByBoxVisible = False
        .ColumnAutoResize = True
        .NewRowPos = jgexBottom
        .RowHeaders = True
        .SelectionStyle = jgexSingleCell
        .AllowDelete = True
        .AllowEdit = True
        .width = 3225
        .Columns.Clear
        Set colTemp = .Columns.Add("CustKey", jgexText, , "CustKey")
        Set colTemp = .Columns.Add("State", jgexText, , "State")
        Set colTemp = .Columns.Add("ExemptNo", jgexText, , "ExemptNo")
    End With
    FormatCertColumns
    
    m_bLoading = False
End Sub


Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    tabMain.width = Me.width - 270
    tabMain.Height = Me.Height - 540
    
    gdxRMAApproval.width = tabMain.width - 240
    gdxRMAApproval.Height = tabMain.Height - 1500
    gdxRMACred.width = tabMain.width - 240
    gdxRMACred.Height = tabMain.Height - 1500
    
    rvRMAApprove.Top = gdxRMAApproval.Top + gdxRMAApproval.Height + 120
    rvARRemarks.Top = gdxRMAApproval.Top + gdxRMAApproval.Height + 120
    rvRMACredit.Top = rvRMAApprove.Top
    
    cmdRMAPrint.Top = rvRMACredit.Top + rvRMACredit.Height - cmdRMAPrint.Height
    Label34(0).Top = rvRMAApprove.Top
    cboWhse(0).Top = rvRMAApprove.Top
    cmdRMAUpdate(0).Top = cmdRMAPrint.Top
    cmdRMARefresh(0).Top = cmdRMAUpdate(0).Top
   
    Label34(1).Top = Label34(0).Top
    cboWhse(1).Top = cboWhse(0).Top
    Label34(2).Top = Label34(0).Top
    cboRMACreditSortBy.Top = Label34(0).Top
    
    cmdRMAUpdate(1).Top = cmdRMAPrint.Top
    cmdRMARefresh(1).Top = cmdRMAPrint.Top

    gdxCOD.width = tabMain.width - 450
    gdxDI.width = tabMain.width - 450

    gdxCOD.Height = (tabMain.Height - 2880) / 3
    gdxDI.Height = (tabMain.Height - 2880) / 3
    
    Label38.Top = gdxCOD.Top + gdxCOD.Height + 90
    gdxDI.Top = Label38.Top + Label38.Height
    
    gdxWillCallOrders.Height = tabMain.Height - 1200
    
    DoEvents

'*** Is this necessary?

    gdxRMAApproval.Refresh
    gdxDI.Refresh
    gdxCOD.Refresh
    gdxWillCallOrders.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '3/31/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwWillCallOrders = Nothing
    Set m_gwShipment = Nothing
    
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Public Sub DoShowHelp()
    ShowHelp "FBilling", True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub EnableTabs()
    Dim TabIndex As Integer
    
    m_asTabRights(1) = k_sRightBillingAccount
    m_asTabRights(2) = k_sRightBillingAssist
    m_asTabRights(3) = k_sRightBillingTemp
    m_asTabRights(4) = k_sRightBillingRMAApprovalMgr
    m_asTabRights(5) = k_sRightBillingRMACredMgr
    m_asTabRights(6) = k_sRightBillingWillCall
    m_asTabRights(7) = k_sRightBillingCreditCard
    m_asTabRights(8) = k_sRightBillingSummary
    
    'decide which tab will be selected
    For TabIndex = 1 To tabMain.Tabs.Count
        If HasRight(m_asTabRights(TabIndex)) Then
            tabMain.Tabs(TabIndex).Selected = True
            Exit For
        End If
    Next TabIndex
    
    For TabIndex = 1 To tabMain.Tabs.Count
        If Not tabMain.Tabs(TabIndex).Selected Then
            tabMain.Tabs(TabIndex).Visible = HasRight(m_asTabRights(TabIndex))
        End If
    Next TabIndex
End Sub


'****************************************************************************
'Control arrays
'****************************************************************************

Private Sub cmdPrint_Click()
    If PrintData Then
        Printer.EndDoc
    Else
        Printer.KillDoc
    End If
End Sub


'*********************************************************************************
'Create Account tab
'*********************************************************************************

Private Sub CreateAcct()
    Dim sTemp As String
    Dim rst As ADODB.Recordset
    Dim iDashPos As Integer
    Dim sZip As String
    Dim taxClass As SalesTax

    SetWaitCursor True
    
    Set rst = New ADODB.Recordset

    sTemp = CreateAcctNbr(Trim(txtCoName.text), m_sZip)
    
    rst.Open "SELECT CustID, CustName FROM tarCustomer WHERE CustID LIKE '" & sTemp & "%' Order by CustID Desc", g_DB.Connection, adOpenDynamic, adLockReadOnly
    
    If Not rst.EOF Then
        iDashPos = InStr(1, rst.Fields("CustID").value, "-", vbTextCompare)
        If iDashPos = 0 Then
            lblCustName.caption = rst.Fields("CustName").value
            sTemp = sTemp & "-1"
        Else
            lblCustName.caption = rst.Fields("CustName").value & " (Plus Others)"
            sTemp = Left(rst.Fields("CustID").value, iDashPos) & CStr(Val(Mid$(rst.Fields("CustID").value, iDashPos + 1)) + 1)
        End If
    Else
        lblCustName.caption = ""
    End If
    Set rst = Nothing
    
    lblAcuity.caption = sTemp
    
    txtCity = m_sCity
    txtState.text = m_sState
    txtCountry.text = m_sCountry

    'lblCollector.Caption = GetCollector
    lblWarehouse.caption = GetWhseID(m_sState, m_sCountry)
    lblTerritory.caption = Billing.GetTerritory(Trim$(lblWarehouse.caption))

    'Look up Tax Schedule information
    If m_sCountry = "USA" Then
        If chkGovt.value = vbChecked Then
            lblTax.caption = "Government"
            txtTaxRate.text = g_dGovtDfltTaxRate
            txtTaxSchdKey.text = g_lGovtDfltSchdKey
        Else
            Set taxClass = New SalesTax
            
            On Error GoTo EH
            taxClass.SetTax m_sZip, m_sState, txtStreet.text, txtCity.text
            On Error GoTo 0
            
            lblTax.caption = taxClass.STaxSchdID
            txtTaxRate.text = taxClass.TaxRate
            txtTaxSchdKey.text = taxClass.STaxSchdKey
            Set taxClass = Nothing
        End If
    Else
        lblTax.caption = "International"
        txtTaxRate.text = g_dIntlDfltTaxRate
        txtTaxSchdKey.text = g_lIntlDfltSchdKey
    End If
    
    SetWaitCursor False
    Exit Sub
    
EH:
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description
    
    'Clean up
    Set taxClass = Nothing
    txtCity.text = ""
    txtState.text = ""
    txtCountry.text = ""
    txtTaxRate.text = ""
    txtTaxSchdKey.text = ""
    lblCustName.caption = ""
    lblAcuity.caption = ""
    'lblCollector.Caption = ""
    lblTax.caption = ""
    lblTerritory.caption = ""
    lblWarehouse.caption = ""
    SetWaitCursor False
    
    Exit Sub
End Sub


Private Sub cmdFindCustomer_Click()
    Dim sPhoneNumber As String
    Dim sZipCode As String
    Dim oCustomer As Customer
    
    sPhoneNumber = Trim(txtPhoneNumber.text)
    sZipCode = Trim(txtZipCode.text)
    
    If sPhoneNumber = "" And sZipCode = "" Then Exit Sub
    
    Set oCustomer = New Customer
    If sZipCode = "" Then
        If Search.FindCustomer(sPhoneNumber, 3, oCustomer, True) = 0 Then
            txtPhoneNumber.SetSel 0, Len(txtPhoneNumber.text)
            txtPhoneNumber.SetFocus
        End If
    ElseIf sPhoneNumber = "" Then
        If Search.FindCustomer(sZipCode, 2, oCustomer, True) = 0 Then
            txtZipCode.SelStart = 0
            txtZipCode.SelLength = Len(txtZipCode.text)
            txtZipCode.SetFocus
        End If
    End If
    
    Set oCustomer = Nothing
End Sub


Private Sub Option1_Click()
    RefreshExceptionReport
End Sub

Private Sub Option2_Click()
    RefreshExceptionReport
End Sub


Private Sub txtPhoneNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtPhoneNumber.text) <> "" And KeyCode = vbKeyReturn Then
        cmdFindCustomer_Click
    End If
End Sub


Private Sub txtPhoneNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(txtPhoneNumber.text) <> "" Then
        txtZipCode.text = ""
    End If
End Sub

Private Sub txtZipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtZipCode.text) <> "" And KeyCode = vbKeyReturn Then
        cmdFindCustomer_Click
    End If
End Sub


Private Sub txtZipCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(txtZipCode.text) <> "" Then
        txtPhoneNumber.text = ""
    End If
End Sub


Private Sub cmdClear_Click()
    txtCoName.text = ""
    txtZip.text = ""
    m_sZip = ""
    m_sCountry = ""
    m_sCity = ""
    m_sState = ""
    ClearForm
    TryToSetFocus txtCoName
End Sub


Private Sub ClearForm()
    txtCity.text = ""
    txtState.text = ""
    txtStreet.text = ""
    txtCountry.text = ""
    
'***For testing
    txtTaxRate.text = ""
    txtTaxSchdKey.text = ""
'***

    lblCustName.caption = ""
    lblAcuity.caption = ""
    lblTax.caption = ""
    lblTerritory.caption = ""
    lblWarehouse.caption = ""
End Sub


'Change the way collectors are assigned.
' 1. Collectors have been changing frequently, so we need to specify collectors in the database, not in code.
' 2. Use a random number to select a collector from the list rather than assigning zip code ranges.

'Private Function GetCollector() As String
'
'    g_rstCollectors.AbsolutePosition = Int(g_rstCollectors.RecordCount * Rnd + 1)
'    GetCollector = g_rstCollectors.Fields("UserID").value
'
'    cmdPrint.Enabled = Len(Trim(GetCollector)) > 1
'End Function


Private Function AcctValid() As Boolean
    If m_sState = "WA" Then
        AcctValid = Len(Trim(txtCoName.text)) > 2 And Len(Trim(txtZip.text)) > 4 _
                And Len(m_sCity) > 0 And Len(Trim(txtStreet.text)) > 0
    Else
        AcctValid = Len(Trim(txtCoName.text)) > 2 And Len(Trim(txtZip.text)) > 3
    End If
End Function


Private Sub Validate()
    If AcctValid Then
        CreateAcct
    Else
        ClearForm
    End If
End Sub

Private Sub tabMain_BeforeTabClick(ByVal NewTab As ActiveTabs.SSTab, ByVal Cancel As ActiveTabs.SSReturnBoolean)
    
    SetWaitCursor True

    Select Case NewTab.Index
        Case tmiWillCall
            RefreshWillCallList
        Case tmiRMAApprovalMgr
            LoadRMAReceived cboWhse(0).ItemData(cboWhse(0).ListIndex)
            UpdateRMAApproveRemarks
        Case tmiRMACredMgr
            LoadRMAApproved cboWhse(1).ItemData(cboWhse(1).ListIndex)
            UpdateRMACreditRemarks
    End Select
    
    'I'm not sure if the Default control is in any way affected by the tab control
    If tabMain.SelectedTab.Index = tmiCreditCard Then
        cmdGetBatchCC.Default = False
        cmdChargeCC.Default = False
        cmdCreateReport.Default = False
    End If
    
    SetWaitCursor False
    
End Sub


Private Sub tabMain_TabClick(ByVal NewTab As ActiveTabs.SSTab)
        If NewTab.Index = tmiCreditCard Then
            txtBatchNumberCC.SetFocus
            cmdGetBatchCC.Default = True
        End If
End Sub


Private Sub txtAcuitySO_KeyDown(KeyCode As Integer, Shift As Integer)
    If cmdLookUp.Enabled And KeyCode = vbKeyReturn Then
        cmdLookUp_Click
    End If
End Sub


Private Sub txtAcuitySO_KeyUp(KeyCode As Integer, Shift As Integer)
    cmdLookUp.Enabled = (Len(Trim(txtAcuitySO.text)) > 0)
End Sub


'Private Sub txtBatchNbr_KeyDown(KeyCode As Integer, Shift As Integer)
'    If cmdGetBatch.Enabled And KeyCode = vbKeyReturn Then
'        cmdGetBatch_Click
'    End If
'End Sub
'
'
'Private Sub txtBatchNbr_KeyUp(KeyCode As Integer, Shift As Integer)
'    cmdGetBatch.Enabled = (Len(Trim(txtBatchNbr.text)) > 0)
'End Sub


Private Sub txtBatchNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtBatchNumber.text)) > 0 _
        And KeyCode = vbKeyReturn Then
        cmdBillingSummary_Click
    End If
End Sub

'***DH 9/17/08 Added
Private Sub txtCoName_LostFocus()
    Validate
End Sub

Private Sub txtCoName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCoName_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub

 
Private Sub txtZip_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub


'***DH 9/9/08 Added
Private Sub txtZip_LostFocus()
    m_sZip = Trim$(Left(txtZip.text, 5))
    
    If m_sZip = "" Then Exit Sub
    
    GetCountyData

    If m_sState = "<>" Then
        msg "Warning - Zip Code not found.", vbCritical + vbOKOnly, "Not Found"
        
'        '***For testing
'        txtTaxRate.Text = ""
'        txtTaxSchdKey.Text = ""
'        '***

        txtZip.SetFocus
        ClearForm
        Exit Sub
    Else
        If m_sState = "WA" Then
            txtStreet.Enabled = True
            txtStreet.BackColor = &H80000005 'White
            txtStreet.SetFocus
        Else
            txtStreet.Enabled = False
            txtStreet.BackColor = &H8000000F 'Grey
            txtStreet.text = ""
       End If
    End If

    Validate
    
End Sub


Private Sub txtStreet_LostFocus()
    Validate
End Sub


Private Sub GetCountyData()
    Dim rst As ADODB.Recordset
    Dim sSQL As String

    m_sState = "<>"
    m_sCity = "<none>"
    
    sSQL = "SELECT * FROM tsmPostalCode WHERE PostalCode like  '" & m_sZip & "%'"
    
    Set rst = LoadDiscRst(sSQL)
    
    If Not rst.EOF Then
        If Trim(rst.Fields("StateID").value) <> "NF" Then
            If Trim(rst.Fields("CountryID").value) = "USA" And Len(m_sZip) <> 5 Then
                'm_sState = ""
            Else
                m_sCountry = Trim(rst.Fields("CountryID").value)
                m_sState = Trim(rst.Fields("StateID").value)
                m_sCity = Trim(rst.Fields("City").value)
            End If
       End If
    End If
    Set rst = Nothing
End Sub


Private Function PrintData() As Boolean
    Printer.Print "Company Name: " & vbTab & txtCoName.text & vbCrLf
    Printer.Print "Billing Zip : " & vbTab & txtZip.text & vbCrLf

    If m_sState = "WA" Then
        Printer.Print "Street      : " & vbTab & txtStreet.text & vbCrLf
    End If
    Printer.Print "City        : " & vbTab & m_sCity & vbCrLf
    Printer.Print "State       : " & vbTab & m_sState & vbCrLf
    Printer.Print "Country     : " & vbTab & m_sCountry & vbCrLf
    Printer.Print "--------------" & vbCrLf
    Printer.Print "Sage Acct : " & vbTab & lblAcuity.caption & vbCrLf

    If lblTax.caption = "Interstate" Then
        Printer.Print "Sales Tax   : " & vbTab & lblTax.caption & vbCrLf
    Else
        Printer.Print "Sales Tax   : " & vbTab & lblTax.caption & " OR Resale " & vbCrLf
    End If
    Printer.Print "Territory  : " & vbTab & lblTerritory.caption & vbCrLf
    Printer.Print "Warehouse   : " & vbTab & lblWarehouse.caption & vbCrLf
    Printer.Print "--------------" & vbCrLf
    Printer.Print "(ABA Number)"

    PrintData = True
End Function


Private Function CreateAcctNbr(sCustomerName As String, sBillingZip As String) As String
    Dim sTemp As String
    'Handle Baskin-Robbins
    If UCase(Left$(sCustomerName, 6)) = "BASKIN" Then
        If ContainsStoreNbr(sCustomerName) Then
            sTemp = BuildBR(sCustomerName)
            CreateAcctNbr = Left$(sTemp, 12)
            Exit Function
        End If
    End If
    
    'Remove "The "
    sTemp = Replace(sCustomerName, "THE ", "", 1, , vbTextCompare)
    sTemp = UCase(AlphanumericOnly(sTemp))
    
    'Pad and reduce to 5 chars
    sTemp = Left$(sTemp & "     ", 5)

    'Maximum Sage CustID Length = 12...
'***DH 9/10/08 The zip has already been trimmed to 5.
    'CreateAcctNbr = Left$(sTemp & Left$(sBillingZip, 5), 12)
    CreateAcctNbr = Left$(sTemp & sBillingZip, 12)
    
    'Remove empty space when customer name < 5 characters
    CreateAcctNbr = Replace(CreateAcctNbr, " ", "")
End Function


Private Function BuildBR(sCustName As String) As String
    Dim lInputLength As Long
    Dim ichar As Integer
    Dim sTemp As String
    Dim i As Integer
    
    sTemp = ""
    lInputLength = Len(sCustName)
    For i = 1 To lInputLength
        ichar = Asc(Mid$(sCustName, i, 1))
        If IsDigit(ichar) Then
              sTemp = sTemp & Mid$(sCustName, i, 1)
        End If
    Next
    
    If Len(sTemp) = 0 Then
        BuildBR = "BASKI"
    Else
        BuildBR = "BR:" & sTemp
    End If

End Function


Private Function ContainsStoreNbr(sBRName As String) As Boolean
    Dim i As Integer
    Dim ichar As Integer
    
    ContainsStoreNbr = False
    
    For i = 1 To Len(sBRName)
        ichar = Asc(Mid$(sBRName, i, 1))
        If IsDigit(ichar) Then
              ContainsStoreNbr = True
              Exit Function
        End If
    Next
End Function


'*** Sales Tax Exemption Stuff ***


Private Function LookupState(sState As String) As Boolean
    Dim vBookmark As Variant
    
    'Cache the bookmark
    vBookmark = gdxSTaxCert.RowBookmark(gdxSTaxCert.RowIndex(gdxSTaxCert.Row))
    
    With gdxSTaxCert.ADORecordset
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If InStr(1, Trim$(sState), Trim$(.Fields("State").value), vbTextCompare) > 0 Then
                    If IIf(IsNull(vBookmark), 0, vBookmark) <> .AbsolutePosition Then
                        LookupState = True
                    End If
                End If
                .MoveNext
            Loop
            'This is testing for new records. In some cases the grid will return a Null Bookmark,
            'and in others it will return Empty as the Bookmark. I don't understand why. I thought
            'it had something to do with the .AllowAddNew property of the grid, but that wasn't the case.
            If Not IsEmpty(vBookmark) Then
                If Not IsNull(vBookmark) Then
                    'set the bookmark back
                    .Bookmark = vBookmark
                End If
            End If
        
        End If
    End With
End Function

Private Sub FormatCertColumns()
    With gdxSTaxCert
        .Columns("CustKey").Visible = False
        .Columns("State").width = 730
        .Columns("ExemptNo").width = 2495
        .Columns("State").TextAlignment = jgexAlignLeft
        'Required for ActAsDropDown grid to work.
        .Columns("State").EditType = jgexEditCombo
        Set .Columns("State").DropDownControl = gdxSTaxState
    End With
    
    gdxSTaxState.BoundColumnIndex = "State"
    gdxSTaxState.ReplaceColumnIndex = "State"
End Sub

Private Function CanAddState() As Boolean
    CanAddState = Not (gdxSTaxCert.RowCount = gdxSTaxState.RowCount)
End Function



'******

'*** Events ***
Private Sub cmdFindExemption_Click()
    'Dim orst As ADODB.Recordset
    Dim sSQL As String
    
    If txtCustID.text = "" Then Exit Sub
    
    m_lCustKey = Billing.GetCustKey(Trim$(txtCustID.text))
    
    If m_lCustKey > 0 Then
        sSQL = "SELECT * FROM tcpSTaxExempt WHERE CustKey = " & m_lCustKey & _
                " ORDER BY State"

        gdxSTaxCert.DatabaseName = g_DB.ConnectionString
        gdxSTaxCert.RecordSource = sSQL
        gdxSTaxCert.ClearFields
        gdxSTaxCert.Rebind
        FormatCertColumns
        gdxSTaxCert.AllowAddNew = CanAddState
    End If
End Sub

Private Sub gdxSTaxCert_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    If Trim$(gdxSTaxCert.value(gdxSTaxCert.Columns("State").ColPosition)) = "" Or _
       Trim$(gdxSTaxCert.value(gdxSTaxCert.Columns("ExemptNo").ColPosition)) = "" Then
        MsgBox "All fields are required.", vbOKOnly, "Validation Error"
        Cancel = True
        gdxSTaxCert.SetFocus
    ElseIf LookupState(gdxSTaxCert.value(gdxSTaxCert.Columns("State").ColPosition)) = True Then
        'Check to see if if an entry already exists in the grid for this State.
        MsgBox "An Exemption Certificate already exists for this state.", vbOKOnly, "Validation Error"
        Cancel = True
        gdxSTaxCert.SetFocus
    Else
        gdxSTaxCert.value(1) = m_lCustKey
    End If
End Sub

Private Sub gdxSTaxCert_AfterUpdate()
    'The order of the following three statements matters.
    'Something is going on inside the grid that I don't understand.
    gdxSTaxCert.AllowAddNew = CanAddState
    gdxSTaxCert.Refetch
    FormatCertColumns
End Sub

Private Sub gdxSTaxCert_AfterDelete()
    'The order of the following three statements matters.
    'Something is going on inside the grid that I don't understand.
    gdxSTaxCert.AllowAddNew = CanAddState
    gdxSTaxCert.Refetch
    FormatCertColumns
End Sub
'******



'***************************************************************************
'Assist tab
'***************************************************************************

Private Sub txtOPID_Change()
    cmdFindSOID.Default = True
End Sub


Private Sub cmdFindSOID_Click()
    Dim cmd As ADODB.Command
    Dim lRetVal As Long
    
    If Len(txtOPID.text) > 0 Then
        lRetVal = Billing.GetSoIdByOpId(txtOPID.text)
        If lRetVal = 0 Then
            lblSOID.caption = "none"
        Else
            lblSOID.caption = lRetVal
        End If
    End If
    Set cmd = Nothing
    txtOPID.SelStart = 0
    txtOPID.SelLength = Len(txtOPID.text)
    txtOPID.SetFocus
End Sub


Private Sub cmdFindSO_Click()
   
    On Error GoTo ErrorHandler
    
    Dim orst As ADODB.Recordset
    If Trim(txtSO.text) = "" Then
        Exit Sub
    Else
        If Not IsNumeric(Trim(txtSO.text)) Then
            msg "The value in the Sales Order textbox is not valid. Please enter effective SOID.", vbOKOnly + vbExclamation, "Add Customer PO"
            txtSO.SelStart = 0
            txtSO.SelLength = Len(txtSO.text)
            TryToSetFocus txtSO
            Exit Sub
        ElseIf CLng(Trim(txtSO.text)) = 0 Then
            Exit Sub
        End If
    End If
    
    Set orst = LoadDiscRst("Select tcpSO.*, tarCustomer.CustID, tarCustomer.CustName " & _
                "FROM tcpSO inner join tarCustomer on tarCustomer.CustKey = tcpSO.CustKey " & _
                "WHERE TranKey = " & CLng(Trim(txtSO.text)))
    
    If orst.EOF Then
        txtSO.SelStart = 0
        txtSO.SelLength = Len(txtSO.text)
        TryToSetFocus txtSO
        lblSOInfo.caption = ""
        m_lAddPOOPKey = 0
        m_lAddPOSOKey = 0
    Else
        lblSOInfo.caption = orst.Fields("CustID").value & " - " & orst.Fields("CustName").value & vbCrLf & _
                            "Order Date: " & orst.Fields("CreateDate").value & vbCrLf & _
                            "CSR: " & orst.Fields("UserID").value & vbCrLf
        If Trim(orst.Fields("PurchOrd")) <> "" Then
            lblSOInfo.caption = lblSOInfo.caption & "PO: " & orst.Fields("PurchOrd").value
        End If
        m_lAddPOOPKey = orst.Fields("OPKey").value
        m_lAddPOSOKey = orst.Fields("SOKey").value
    End If
    Exit Sub
    
ErrorHandler:
    msg "The value in the Sales Order textbox is not valid. Please enter effective SOID.", vbOKOnly + vbExclamation, "Add Customer PO"

End Sub


Private Sub cmdUpdateSO_Click()
    If Trim(txtPO.text) = "" Then Exit Sub
    If m_lAddPOOPKey = 0 Or m_lAddPOSOKey = 0 Then Exit Sub
    
    Dim oCmd As ADODB.Command
    
    Set oCmd = CreateCommandSP("spCPCInsertPO")
    With oCmd
        .Parameters("@_iOPKey") = m_lAddPOOPKey
        .Parameters("@_iSOKey") = m_lAddPOSOKey
        .Parameters("@_iPurchOrd") = Left(Trim(txtPO.text), 15)  'Left() addded 8/28/02 LR
        .Execute
    End With
    
    Set oCmd = Nothing
    m_lAddPOOPKey = 0
    m_lAddPOSOKey = 0
    
    cmdFindSO_Click
End Sub



Private Sub BalanceFreight(iType As Integer, sCreateDate As String, sPackStation As String)
    Dim dFreightAmt As Double

    Dim rst As ADODB.Recordset
    Dim sSQL As String
    Dim sShipMethWhere As String
    Dim sPaymentTermsWhere As String
        
    
    
    
    On Error GoTo ErrorHandler
    
    SetWaitCursor True
    
    If iType = 1 Then 'REG SHIPMENTS
        sShipMethWhere = " COALESCE((SELECT ShipMethKey FROM tciShipMethod t WHERE ShipMethKey not in (27,37,32,40,44,52,69,70,72,78,80,85) and t.ShipMethKey = tsoSalesOrder.DfltShipMethKey), -1) "
        sPaymentTermsWhere = " COALESCE((SELECT PmtTermsKey FROM tciPaymentTerms t WHERE t.PmtTermsKey not in (48) and t.PmtTermsKey = tsoSalesOrder.PmtTermsKey), -1)"
    ElseIf iType = 2 Then 'WILL CALL NET
        sShipMethWhere = " COALESCE((SELECT ShipMethKey FROM tciShipMethod t WHERE ShipMethKey in (27,37,32) and t.ShipMethKey = tsoSalesOrder.DfltShipMethKey), -1)"
        sPaymentTermsWhere = " COALESCE((SELECT PmtTermsKey FROM tciPaymentTerms t WHERE PmtTermsKey in (22,29,30,31,32,47) and t.PmtTermsKey = tsoSalesOrder.PmtTermsKey), -1)"
    ElseIf iType = 3 Then 'WILL CALL CHASH
        sShipMethWhere = " COALESCE((SELECT ShipMethKey FROM tciShipMethod t WHERE ShipMethKey in (27,37,32) and t.ShipMethKey = tsoSalesOrder.DfltShipMethKey), -1)"
        sPaymentTermsWhere = "COALESCE((SELECT PmtTermsKey FROM tciPaymentTerms t WHERE PmtTermsKey in (36,37,40,41,44) and t.PmtTermsKey = tsoSalesOrder.PmtTermsKey), -1)"
    ElseIf iType = 4 Then 'USPS
        sShipMethWhere = " COALESCE((SELECT ShipMethKey FROM tciShipMethod t WHERE ShipMethKey in (40,44,52,69,70,72,78,80,85,77,66,73,51,23,42,67,71,64,82,81,83,84,43,33,39,65,68,68,28,59) and t.ShipMethKey = tsoSalesOrder.DfltShipMethKey), -1)"
        sPaymentTermsWhere = " COALESCE((SELECT PmtTermsKey FROM tciPaymentTerms t WHERE t.PmtTermsKey not in (48) and t.PmtTermsKey = tsoSalesOrder.PmtTermsKey), -1)"
    ElseIf iType = 5 Then 'MO CRCARD
        sShipMethWhere = " COALESCE((SELECT ShipMethKey FROM tciShipMethod t WHERE t.ShipMethKey = tsoSalesOrder.DfltShipMethKey), -1)"
        sPaymentTermsWhere = " COALESCE((SELECT PmtTermsKey FROM tciPaymentTerms t WHERE PmtTermsKey in (48) and t.PmtTermsKey = tsoSalesOrder.PmtTermsKey), -1)"
    End If
    
    Set rst = New ADODB.Recordset

    sSQL = "SELECT DISTINCT tsoPendShipment.ShipKey, tsoPendShipment.FreightAmt, tsoPendShipment.WhseKey, tsoSalesOrder.SalesAmt, tsoSalesOrder.STaxAmt , " & vbCrLf & _
        "tciPaymentTerms.DueDayOrMonth, RTRIM(tsoPendShipment.UserFld1) UserFld1 " & vbCrLf & _
        "FROM tsoPendShipment INNER JOIN " & vbCrLf & _
        "tsoShipLine ON tsoPendShipment.ShipKey = tsoShipLine.ShipKey INNER JOIN " & vbCrLf & _
        "tsoSOLine ON tsoShipLine.SOLineKey = tsoSOLine.SOLineKey INNER JOIN " & vbCrLf & _
        "tsoSalesOrder ON tsoSOLine.SOKey = tsoSalesOrder.SOKey INNER JOIN " & vbCrLf & _
        "tciPaymentTerms ON tsoSalesOrder.PmtTermsKey = tciPaymentTerms.PmtTermsKey " & vbCrLf & _
        "Where convert(char, tsoPendShipment.CreateDate, 101) = '" & sCreateDate & "' " & vbCrLf & _
        "  AND tsoPendShipment.WhseKey = " & cboWhse(2).ItemData(cboWhse(2).ListIndex) & vbCrLf & _
        "  AND tsoSalesOrder.DfltShipMethKey = " & sShipMethWhere & vbCrLf & _
        "  AND tsoSalesOrder.PmtTermsKey = " & sPaymentTermsWhere

    If sPackStation <> "" Then
        sSQL = sSQL & " AND tsoPendShipment.CreateUserId = '" & sPackStation & "' " & vbCrLf
    End If

rst.Open sSQL, g_DB.Connection, adOpenDynamic, adLockReadOnly

    'Set rst = CallSP("spCPCGetShipmentToBalanceFreight", "@CreateDate", sCreateDate, "@WhseKey", cboWhse(2).ItemData(cboWhse(2).ListIndex), "@Type", iType, "@UserId", sPackStation)

    If rst.EOF Then Exit Sub
    
    Debug.Print "Balance Freight Record Count: " & rst.RecordCount
        
    With rst
        Do While Not .EOF
'            'PRN#266
'            If Trim$(.Fields("UserFld1").Value & "") <> "1" Then ' UserFld1 { 1 = AlreadyProcessed | else = Not yet }
'
                 dFreightAmt = .Fields("FreightAmt").value
'
'                 'PRN#266
'                 If .Fields("WhseKey") = 24 Then ' 24 = Seattle
'                     If .Fields("DueDayOrMonth") = 0 Then ' 0 = COD
'                         If .Fields("FreightAmt") > 0 And .Fields("SalesAmt") > 0 And .Fields("STaxAmt") > 0 Then
'                             nTaxRate = .Fields("STaxAmt") / .Fields("SalesAmt")
'                             dFreightAmt = FormatCurrency(dFreightAmt / (1 + nTaxRate), 2)
'                         End If
'                     End If
'                 End If
                 
                 UpdateFreightLine .Fields("ShipKey").value, dFreightAmt
'
'            End If
            .MoveNext
        Loop
    End With
        
    rst.Close
    SetWaitCursor False
    Set rst = Nothing
    Exit Sub
    
ErrorHandler:
    ClearWaitCursor
    msg Err.Description, vbOKOnly + vbExclamation, Err.Source
End Sub


Private Function UpdateFreightLine(lShipKey As Long, dFreightAmt As Double)
'    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim bIsFirstLine As Boolean
    
'    On Error GoTo EH
    
    Set rst = New ADODB.Recordset
    
'     'PRN#266
'     ' FreightAmt must be same in both tsoPendShipment and tsoShipLine, if they are not,
'     ' Sage would fail when posting shipment. To remedy this problem, we have to update
'     ' tsoPendShipment.FreightAmt to be same as tsoShipLine.FreightAmt. To prevent tsoPendShipment.FreightAmt
'     ' from being updated more than once, we use UserFld1 to flag whether we have already updated before.
'    ' We always want to flag the UserFld1, so that we don't have to process it twice
'    ' Update tsoPendShipment
'    Set cmd = New ADODB.Command
'    Set cmd = CreateCommandSP("update tsoPendShipment set FreightAmt = " & dFreightAmt & " , UserFld1 = 1 Where (tsoPendShipment.ShipKey =  " & lShipKey & ")", adCmdText)
'    cmd.Execute
'    Set cmd = Nothing
    
    ' Update tsoShipLineDist
    With rst
        'Get ALL the LineDists for this shipment
        .Open "SELECT tsoShipLineDist.FreightAmt FROM tsoShipLine INNER JOIN " & _
            "tsoShipLineDist ON tsoShipLine.ShipLineKey = tsoShipLineDist.ShipLineKey " & _
            "Where tsoShipLine.ShipKey = " & CStr(lShipKey), g_DB.Connection, adOpenDynamic, adLockOptimistic
        
        'Update the first LineDist with the entire freight amount; zero all others
        bIsFirstLine = True
        Do While Not .EOF
            If bIsFirstLine Then
                .Fields("FreightAmt").value = dFreightAmt
            Else
                .Fields("FreightAmt").value = 0
            End If
            .Update
            bIsFirstLine = False
            .MoveNext
        Loop
    
        .Close
    End With
    
    Set rst = Nothing
    
'    Exit Function
'EH:
'    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Sub cmdLookUp_Click()
    
    If Len(Trim(txtAcuitySO.text)) = 0 Then Exit Sub
    
    Dim sql As String
    Dim rst As ADODB.Recordset
    SetWaitCursor True
       
    sql = "SELECT tarCustomer.CustID, tarCustomer.CustName, " _
        & "tsoSalesOrder.CreateUserID as CSR, tsoSalesOrder.SOKey, " _
        & "tciAddress.City, tsoSalesOrder.UserFld1 AS PricePackSlip, " _
        & "tsoSalesOrder.CreateDate FROM tsoSalesOrder INNER JOIN " _
        & "tarCustomer ON tsoSalesOrder.CustKey = tarCustomer.CustKey " _
        & "INNER JOIN tciAddress ON " _
        & "tsoSalesOrder.DfltShipToAddrKey = tciAddress.AddrKey " _
        & "WHERE tsoSalesOrder.Status = 1 AND " _
        & "tsoSalesOrder.CompanyID = 'CPC' AND " _
        & "tsoSalesOrder.TranNo = '" & FormatBatchID(Trim(txtAcuitySO.text), "0000000000") & "'"
        
    Set rst = LoadDiscRst(sql)
    
    If rst.RecordCount = 0 Then
        cmdUpdate.Enabled = False
        lblAssistCustName.caption = ""
        lblAssistCustID.caption = ""
        lblAssistCSR.caption = ""
        lblShipToCity.caption = ""
        lblOrderDate.caption = ""
        chkPricePacking.value = vbUnchecked
        msg "Sorry, No records satisfy this request"
        TryToSetFocus txtAcuitySO
        txtAcuitySO.SelStart = 0
        txtAcuitySO.SelLength = Len(txtAcuitySO.text)
        m_lSOKey = 0
    ElseIf rst.RecordCount > 0 Then
        cmdUpdate.Enabled = True
        If Not IsNull(rst.Fields("CustName")) Then lblAssistCustName.caption = rst.Fields("CustName").value
        If Not IsNull(rst.Fields("CustID")) Then lblAssistCustID.caption = rst.Fields("CustID").value
        If Not IsNull(rst.Fields("CSR")) Then lblAssistCSR.caption = rst.Fields("CSR").value
        If Not IsNull(rst.Fields("City")) Then lblShipToCity.caption = rst.Fields("City").value
        If Not IsNull(rst.Fields("CreateDate")) Then lblOrderDate.caption = rst.Fields("CreateDate").value
        If IsNull(rst.Fields("PricePackSlip")) Then
            chkPricePacking.value = vbUnchecked
        Else
            chkPricePacking.value = rst.Fields("PricePackSlip").value
        End If
        m_lSOKey = rst.Fields("SOKey").value
    End If
    
    Set rst = Nothing
    SetWaitCursor False
End Sub


Private Sub cmdUpdate_Click()
    If m_lSOKey = 0 Then Exit Sub
    
    Dim sSQL As String
    Dim oCmd As ADODB.Command
    
    SetWaitCursor True
    sSQL = "Update tsoSalesOrder set UserFld1 = " & chkPricePacking.value _
            & "Where SOKey = " & m_lSOKey
    Set oCmd = CreateCommandSP(sSQL, adCmdText)
    oCmd.Execute
    
    Set oCmd = Nothing
    SetWaitCursor False
    If vbYes = msg("Update Complete! Would you like to set a new search? ", _
                    vbYesNo + vbExclamation, "New Price Pack Slip Search?") Then
        lblAssistCustName.caption = ""
        lblAssistCustID.caption = ""
        lblAssistCSR.caption = ""
        lblShipToCity.caption = ""
        lblOrderDate.caption = ""
        chkPricePacking.value = vbUnchecked
        txtAcuitySO.text = ""
        TryToSetFocus txtAcuitySO
    End If
End Sub


'**************************************************************************
'Shipment tab
'**************************************************************************

Private Sub gdxShipmentBatch_SelectionChange()
    RefreshExceptionReport
End Sub

Private Sub RefreshExceptionReport()
Dim sTemp As String
    Dim sType As String
    Dim iType As Integer
    Dim sCreateDate As String
    Dim sPackStation As String
    Dim bExceptionsOnly As Boolean
    
    SetWaitCursor True
    
    sCreateDate = Trim(CStr(m_gwShipment.value("CreateDate")))
    sType = Trim(CStr(m_gwShipment.value("TypeId")))
    iType = CInt(m_gwShipment.value("TypeKey"))
    
    If Option1.value = True Then
        bExceptionsOnly = False
    ElseIf Option2.value = True Then
        bExceptionsOnly = True
    End If
    
'    If (m_gwShipment.Value("PackStation") <> Null) Then
'        sPackStation = Trim(CStr(m_gwShipment.Value("PackStation")))
'    End If
    
    If Not IsNull(m_gwShipment.value("PackStation")) Then
        sPackStation = m_gwShipment.value("PackStation")
    End If
    
    If sPackStation <> "" Then
        sTemp = "Shipment Details Whse: " & cboWhse(2).text & " Type: " & sType & " Station: " & sPackStation & " DATE: " & sCreateDate & vbCrLf
    Else
        sTemp = "Shipment Details Whse: " & cboWhse(2).text & " Type: " & sType & " DATE: " & sCreateDate & vbCrLf
    End If
    
    
    txtShipWarnings.text = sTemp & GetShipWarnings(sCreateDate, cboWhse(2).ItemData(cboWhse(2).ListIndex), iType, sPackStation, bExceptionsOnly)
       
    SetWaitCursor False
End Sub

Private Sub cmdGetBatches_Click()
    GetShipmentBatches
End Sub

Private Sub GetShipmentBatches()
    Dim sSQL As String
    Dim rst As ADODB.Recordset
    
    Set rst = CallSP("spcpcGetAccountingBatches", "@WhseKey", cboWhse(2).ItemData(cboWhse(2).ListIndex))
    
    Dim iCount As Integer
    iCount = rst.RecordCount
    
    With gdxShipmentBatch
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = rst
    End With
End Sub

Private Sub cmdBalanceFreight_Click()
    Dim cmd As ADODB.Command
    Dim iType As Integer
    Dim sCreateDate As String
    Dim sPackStation As String
    
    SetWaitCursor False
    
    iType = CInt(m_gwShipment.value("TypeKey"))
    sCreateDate = Trim(CStr(m_gwShipment.value("CreateDate")))
    
    If Not IsNull(m_gwShipment.value("PackStation")) Then
        sPackStation = m_gwShipment.value("PackStation")
    End If

    BalanceFreight iType, sCreateDate, sPackStation
        
    SetWaitCursor False
    
    msg "Update Complete"
End Sub




Private Function GetShipWarnings(sCreateDate As String, iWhseKey As Integer, iType As Integer, sPackStation As String, bExceptionsOnly) As String
    Dim rst As ADODB.Recordset
    Dim sShipmentID As String
    Dim OPKey As Long
    Dim SONbr As String
    Dim sWarnings As String
    Dim sBuffer As String
    Dim iCounter As Integer
    Dim iExceptionCount As Integer
    
    Set rst = CallSP("spcpcShipmentCheck2", "@CreateDate", sCreateDate, "@WhseKey", iWhseKey, "@Type", iType, "@UserId", sPackStation)

    If rst.EOF Then Exit Function

    With rst
        Do While Not .EOF
            
            
            OPKey = CLng((.Fields("OPNbr").value))
            
            'Debug.Print ("OPKEY: " & OPKey)
            
            'Free Parts
            If .Fields("PartsNoCharge").value = 1 Then
                sWarnings = sWarnings & Space(4) & "Order is marked Parts No Charge" & vbCrLf
                sWarnings = sWarnings & GetSpecialHandlingRemarks(Space(8), OPKey, "Order.PartsNoCharge")
                If .Fields("OrderAmt").value > 0 Then
                    sWarnings = sWarnings & Space(4) & "Error - Order is marked No Charge but shipment has value" & vbCrLf
                End If
            End If
                  
            'No Charge
            If .Fields("PartsNoCharge").value = 0 And .Fields("OrderAmt").value = 0 Then
                sWarnings = sWarnings & Space(4) & "Warning - Order is not marked No Charge but shipment has no value" & vbCrLf
            End If
            
            'International
            If Trim(.Fields("ShiptoCountryId").value) <> "USA" Then
                sWarnings = sWarnings & Space(4) & "Warning - Order is shipped to " & Trim(.Fields("ShiptoCountryId").value) _
                & "." & vbCrLf
            End If
            
            'Sales Tax
            If .Fields("TaxError").value = 1 Then
                sWarnings = sWarnings & Space(4) & "Error - OP and SO Sales Tax values don't match" & vbCrLf
            End If

            'Interstate
            If (.Fields("ShipAddrSTaxSchdKey").value = g_lIStateDfltSchdKey) _
                And (.Fields("STaxAmt").value > 0) Then
                sWarnings = sWarnings & Space(4) & "Error - Customer is Interstate and is being charged Sales Tax" & vbCrLf
            End If
            
            'Government
            If (.Fields("ShipAddrSTaxSchdKey").value = g_lGovtDfltSchdKey) _
                And (.Fields("STaxAmt").value > 0) Then
                sWarnings = sWarnings & Space(4) & "Error - Customer is Government and is being charged Sales Tax" & vbCrLf
            End If
            
            'International
            If (.Fields("ShipAddrSTaxSchdKey").value = g_lIntlDfltSchdKey) _
                And (.Fields("STaxAmt").value > 0) Then
                sWarnings = sWarnings & Space(4) & "Error - Customer is International and is being charged Sales Tax" & vbCrLf
            End If
            
            'Resale
            If (Len(Trim(.Fields("ShipAddrSTaxExemptNo").value)) > 0) And (.Fields("STaxAmt").value > 0) Then
                sWarnings = sWarnings & Space(4) & "Warning - Customer has a Resale Certificate and is being charged Sales Tax" _
                & vbCrLf
            End If
            
            'Debug.Print "OPKey: " & OPKey & " Free Freight: " & .Fields("FreeFreight").Value
            'Free Freight
            If .Fields("FreeFreight").value = 1 Then
                sWarnings = sWarnings & Space(4) & "Order is marked Free Freight" & vbCrLf
                sWarnings = sWarnings & GetSpecialHandlingRemarks(Space(8), OPKey, "Order.FreeFreight")
                If .Fields("FreightAmt").value > 0 And InStr(1, .Fields("ShipMethod").value, "Call") = 0 Then
                    sWarnings = sWarnings & Space(4) & "Error - Order is marked Free Freight but shipment has freight" & vbCrLf
                End If
            End If
            
            'Consider looking up CSWReportDatabase data to compare with Freight value in shipment
            'cpsqlpro currently takes 8s to execute this query
            
            'High Freight
            If .Fields("FreightAmt").value > 75 Then
                sWarnings = sWarnings & Space(4) & "Warning - Order has high freight - " & _
                Format$(.Fields("FreightAmt").value, "$###,###.##") & vbCrLf
            End If

            'Zero Freight and Not Free Freight or WillCall
            If .Fields("FreeFreight").value = 0 And .Fields("FreightAmt").value = 0 And InStr(1, .Fields("ShipMethod").value, "Call") = 0 Then
                sWarnings = sWarnings & Space(4) & "Warning - The Shipment has no freight but the order is not WillCall or marked Free Freight" & vbCrLf
            End If
                        
            'Ship Complete
            If .Fields("ShipComplete").value = 1 And .Fields("BackOrders").value = 1 Then
                sWarnings = sWarnings & Space(4) & "Error - Order is marked Ship Complete but shipment has backorders" & vbCrLf
            End If
            
            'Debug.Print "OPKey: " & OPKey & " Inbound Freight: " & .Fields("InboundFreight").Value
            'Inbound Freight
            If .Fields("InboundFreight").value = 1 Then
                sWarnings = sWarnings & Space(4) & "Order is marked Inbound Freight" & vbCrLf
                sWarnings = sWarnings & GetSpecialHandlingRemarks(Space(8), OPKey, "Order.InboundFreight")
            End If
            
            'BillDifferentRate
            If .Fields("BillDifferentRate").value = 1 Then
                If .Fields("BillMethKey").value > 0 Then
                    g_rstShipVia.Filter = "ShipMethKey = " & .Fields("BillMethKey").value
                    
                    sWarnings = sWarnings & Space(4) & "Warning - Order is marked Bill Different Rate (shipped " _
                    & .Fields("ShipMethod").value & ", billed " & g_rstShipVia.Fields("ShipMethID").value & ")" & vbCrLf
                    
                    g_rstShipVia.Filter = adFilterNone
                Else
                    sWarnings = sWarnings & Space(4) & "Warning - Order is marked Bill Different Rate" & vbCrLf
                End If
            End If
            
            'Debug.Print "OPKey: " & OPKey & " Reduced Freight: " & .Fields("ReducedFreight").Value
            'Reduced Freight
            If .Fields("ReducedFreight").value = 1 Then
                sWarnings = sWarnings & Space(4) & "Order is marked Reduced Freight" & vbCrLf
                sWarnings = sWarnings & GetSpecialHandlingRemarks(Space(8), OPKey, "Order.ReducedFreight")
            End If
            
            'Deposit
            If .Fields("Deposit").value = 1 Then
                sWarnings = sWarnings & Space(4) & "Order is marked Deposit" & vbCrLf
                sWarnings = sWarnings & GetSpecialHandlingRemarks(Space(8), OPKey, "Order.Deposit")
            End If
            
            'No Tracking Number
            If Trim(.Fields("ShipTrackNo").value) = "" And InStr(1, .Fields("ShipMethod").value, "Call") = 0 Then
                sWarnings = sWarnings & Space(4) & "Warning - Order is missing a tracking number" & vbCrLf
            End If
            
            'Bill Recipient - Freight more then handling charge.
            'Is marked Bill Recipient, is not WillCall,and the
            'freight does not contain only Handling charges.
            If Len(Trim(.Fields("UPSAcct").value)) > 0 And Not IsHandlingCharge(.Fields("FreightAmt")) Then
                sWarnings = sWarnings & Space(4) & "Error - Order is UPS Bill Recipient, " _
                            & "but has inappropriate handling. Freight: " & _
                            Format$(.Fields("FreightAmt").value, "$###,###.##") & vbCrLf
            End If
                        
            'Not Bill Recipient - Freight less then or egual to $2.00.
            If Len(Trim(.Fields("UPSAcct"))) = 0 And InStr(1, .Fields("ShipMethod").value, "Call") = 0 _
            And .Fields("FreightAmt") <= 2 And .Fields("FreightAmt") > 0 And .Fields("FreeFreight") = 0 Then
                sWarnings = sWarnings & Space(4) & "Warning - Freight is too low. Freight: " & _
                            Format$(.Fields("FreightAmt").value, "$###,###.##") & vbCrLf
            End If

            If .Fields("HasTrueCompressor") > 0 Then
                sWarnings = sWarnings & Space(4) & "Attention - Has a True Compressor" & vbCrLf
            End If
            
            If sShipmentID <> CStr(.Fields("Shipment").value) Then
                
                'get the data for the next one
                sShipmentID = CStr(.Fields("Shipment").value)
                OPKey = CLng((.Fields("OPNbr").value))
                SONbr = CStr(.Fields("SONbr").value)
                
                If Len(sWarnings) > 0 Then
                    iExceptionCount = iExceptionCount + 1
                End If
                                    
                If bExceptionsOnly Then
                    If Len(sWarnings) > 0 Then
                        iCounter = iCounter + 1
                        sBuffer = sBuffer & vbCrLf & iCounter & ".  Shipment " & sShipmentID & "(OP-" & OPKey & " / " & SONbr & ")" & vbCrLf & sWarnings
                    End If
                Else
                    iCounter = iCounter + 1
                    sBuffer = sBuffer & vbCrLf & iCounter & ".  Shipment " & sShipmentID & "(OP-" & OPKey & " / " & SONbr & ")" & vbCrLf & sWarnings
                End If
                
                sWarnings = vbNullString
            End If
            
            .MoveNext
        Loop
        
        
        If Len(sWarnings) > 0 Then
            sBuffer = sBuffer & vbCrLf & "  Shipment " & sShipmentID & vbCrLf & sWarnings & vbCrLf
            'sBuffer = sBuffer & vbCrLf & "  Shipment " & sShipmentID & " OPKey: " & OPKey & "    SO#: " & SONbr & vbCrLf & sWarnings
        End If
        
        GetShipWarnings = sBuffer & vbCrLf & iExceptionCount & " Exception(s)"
    End With
    
    CloseRst rst
End Function


Private Function GetSpecialHandlingRemarks(indent As String, OPKey As Long, RemarkType As String) As String
    Dim rst As ADODB.Recordset
    Dim sql As String
    sql = "select sender, effectivedate, memotext from tcimemo where addressee='" & RemarkType & "' and memoownerkey=" & OPKey
    
    Set rst = LoadDiscRst(sql)
    
    With rst
        If Not rst.EOF Then
            Do While Not .EOF
                GetSpecialHandlingRemarks = GetSpecialHandlingRemarks & indent & .Fields("sender") & " " & .Fields("effectivedate") & vbCrLf & _
                    indent & .Fields("memotext") & vbCrLf
                .MoveNext
            Loop
        Else
            GetSpecialHandlingRemarks = indent & "No remarks" & vbCrLf
        End If
    End With
    
End Function


Private Function IsHandlingCharge(Amount As Double) As Boolean
    'Handling charges for MPK, STL, and SEA repectively.
    If Amount = 1.5 Or Amount = 2 Or Amount = 1.75 Then
        IsHandlingCharge = True
    Else
        IsHandlingCharge = False
    End If
End Function


'************************************************************************************
'RMA Approval tab(6)
'************************************************************************************

'Called By
'   Form_Load  (probably don't want to do this)
'   cmdRMARefresh_Click
'   UpdateRMAApproval

Private Sub LoadRMAReceived(ByVal lWhseKey As Long)
    Set m_oRMAApprovalList = Billing.LoadReceivedRMAsToApprove(lWhseKey)
    
    Dim i As Integer
    With gdxRMAApproval
        .HoldFields
        .ItemCount = m_oRMAApprovalList.Count
        .Refetch    '??????
        .Row = 1
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
End Sub


Private Sub cmdRMARefresh_Click(Index As Integer)
    Select Case Index
        Case 0: 'RMA Approval tab
            SetWaitCursor True
            'load the Recieved items for approval
            LoadRMAReceived cboWhse(Index).ItemData(cboWhse(Index).ListIndex)
            UpdateRMAApproveRemarks
            SetWaitCursor False
        Case 1: 'RMA Credit tab
            SetWaitCursor True
            'load the Approved items for crediting
            LoadRMAApproved cboWhse(Index).ItemData(cboWhse(Index).ListIndex)
            UpdateRMACreditRemarks
            SetWaitCursor False
    End Select
End Sub


Private Sub UpdateRMAApproveRemarks()
'value(3) = OPKey
'value(12) = grouping caption

    'Why check for both values? It's unbound. We can establish a state on init.
    'Handling the case when we click on a grouping line?
    If gdxRMAApproval.value(3) = Empty Or IsNull(gdxRMAApproval.value(3)) Then
        rvRMAApprove.Visible = False
        rvARRemarks.Visible = False
    Else
        rvRMAApprove.Visible = True
        rvRMAApprove.OwnerID = ""   'Why are we doing this? Is it masking a problem in the control?
        rvRMAApprove.OwnerID = gdxRMAApproval.value(3)

'5/18/03 LR
'This is a Q&D fix to allow the RMA czar to attach AR remarks to the customer
'associated with the RMA.
'The CustID appears in a string called RMAInfo (column 12 in the grid) which
'is used to label the grid grouping line.
'ParseCustID parses out the ID so it can be used as the OwnerID in the rvARRemarks control.
        rvARRemarks.Visible = True
        rvARRemarks.OwnerID = ParseCustID(gdxRMAApproval.value(12))
        
    End If
End Sub


'Called by:
'   UpdateRMAApproveRemarks
'
'Parses out the ID so it can be used as the OwnerID in the rvARRemarks control

Private Function ParseCustID(RMAInfo As String) As String
    Dim i As Integer
    Dim s As String
    
    i = InStr(1, RMAInfo, "Customer")
    s = Mid$(RMAInfo, i + 10)
    i = InStr(1, s, " ")
    ParseCustID = Mid$(s, 1, i - 1)
End Function

'this handles mouse selection
Private Sub gdxRMAApproval_Click()
    UpdateRMAApproveRemarks
End Sub

'this handles the up/down arrow keys
Private Sub gdxRMAApproval_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        UpdateRMAApproveRemarks
    End If
End Sub


Private Sub gdxRMAApproval_LostFocus()
    gdxRMAApproval.Update
End Sub


Private Sub gdxRMAApproval_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxRMAApproval.Update
End Sub


Private Sub gdxRMAApproval_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    If (ColIndex = 18) And Not gdxRMAApproval.GetRowData(gdxRMAApproval.Row).value(4) Then
        Cancel = True
    End If
End Sub


Private Sub gdxRMAApproval_AfterColEdit(ByVal ColIndex As Integer)
    gdxRMAApproval.Update
End Sub


Private Sub gdxRMAApproval_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    
    If m_oRMAApprovalList Is Nothing Then Exit Sub
    
    If RowIndex > m_oRMAApprovalList.Count Then Exit Sub
    
    With m_oRMAApprovalList.Item(RowIndex)
        Values(1) = .RMAKey
        Values(2) = .SOLineKey
        Values(3) = .OPKey
        Values(4) = .Approved
        Values(5) = .ItemID
        Values(6) = .QtyAuthorized
        Values(7) = .AuthBy
        Values(8) = .AuthDate
        Values(9) = .QtyRcvd
        Values(10) = .Cost
        Values(11) = .Price
        Values(12) = .RMAInfo
        Values(13) = .OPLineKey
        Values(14) = m_colDisposition(CStr(.Disposition))
        Values(15) = .QtyPreCred
        Values(16) = m_colReason(CStr(.Reason))
        Values(17) = .Restock * 100
        Values(18) = .CreditFreight
        Values(19) = .RcvdWhseID
        Values(20) = .Descr
        Values(21) = .ReceiveDate
    End With
End Sub


Private Sub gdxRMAApproval_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    With m_oRMAApprovalList.Item(RowIndex)
        .Approved = IIf(Values(4), True, False)
        .CreditFreight = IIf(Values(18), True, False)
        
        '06/12/03 AVH PRN#103 if the receiver receives too many of an item on an RMA
        '( RMA for quantity of 2, 1 is being returned but the receiver receives 2
        'there is no way to edit the ones received in error.
        
        If Values(9) > .QtyRcvdOriginal Then
            msg "Sorry. QtyRcvd must be less than or equal to " & .QtyRcvdOriginal _
            & vbCrLf & " as only downward adjustment is allowed. Please re-enter.", vbExclamation + vbOKOnly, "Enter Qty Creidt"
        Else
            .QtyRcvd = Values(9)
        End If
        
    End With
End Sub


Private Sub cmdRMAPrint_Click()
    On Error GoTo EH
    
    Dim oFrm As FViewer
    Set oFrm = New FViewer
    Call oFrm.ParamAdd(1, "WhseKey", cboWhse(1).ItemData(cboWhse(1).ListIndex))
    Call oFrm.ViewReportByType("RMA Credit")
    Set oFrm = Nothing

    Exit Sub
EH:
    Exit Sub
End Sub


Private Sub cmdRMAUpdate_Click(Index As Integer)
    Select Case Index
        Case 0:
            UpdateRMAApproval
        Case 1:
            UpdateRMACredit
    End Select
End Sub


Private Sub UpdateRMAApproval()
    Dim m_oAdjs As InvAdjustments
    Dim lIndex As Long
    Dim lStatus As Long
    Dim cmd As ADODB.Command
    Dim sEventStrValue As String
    Dim lsMessage As String
    Dim lsInventoryMessage As String
    
    If m_oRMAApprovalList.Count = 0 Then Exit Sub
           
    
    On Error GoTo ErrorRollback
    
    g_DB.Connection.BeginTrans
    
    Set m_oAdjs = New InvAdjustments
    
    SetWaitCursor True
    With m_oRMAApprovalList
        For lIndex = 1 To .Count
            If .Item(lIndex).Approved = True Then
                
                Dim itemFound As Boolean
                itemFound = True
                
                If .Item(lIndex).OPItemType = itFinishedGood Or .Item(lIndex).OPItemType = itBTOKit Then
                    itemFound = Billing.IsItemInInventory(.Item(lIndex).ItemID, cboWhse(0).ItemData(cboWhse(0).ListIndex))
                End If
                
                If itemFound Then
                    
                    Billing.ApproveRMAItem .Item(lIndex).RmaLineKey, .Item(lIndex).Approved, .Item(lIndex).CreditFreight, .Item(lIndex).QtyRcvd
                    
                    If .Item(lIndex).Disposition = 2 Then  '2 = Return to Stock
                        If Not m_oAdjs.Add(.Item(lIndex).ItemID, .Item(lIndex).QtyRcvd, "Inv Adj - " & .Item(lIndex).ItemID & " Returned to Stock") Then
                            If Len(lsMessage) = 0 Then
                                lsMessage = "The following parts are not in inventory and have not been approved. " & _
                                "Please review disposition(s) for: " & chr(13) & chr(10) & chr(13) & chr(10)
                            End If
                            lsMessage = lsMessage & "RMA#: " & .Item(lIndex).RMAKey & ", Part#: " & .Item(lIndex).ItemID & ", Whse: " & _
                                        .Item(lIndex).RcvdWhseID & chr(13) & chr(10)
                        End If
                    End If
                Else
                    If Len(lsInventoryMessage) = 0 Then
                        lsInventoryMessage = "The following parts are not set up in inventory: " & chr(13) & chr(10) & chr(13) & chr(10)
                    End If
                    
                    lsInventoryMessage = lsInventoryMessage & "RMA#: " & .Item(lIndex).RMAKey & ", Part#: " & .Item(lIndex).ItemID & ", Whse: " & _
                                    .Item(lIndex).RcvdWhseID & chr(13) & chr(10)
                End If
                
            Else
                'Check if qty has been adjusted downward
                If .Item(lIndex).QtyRcvd < .Item(lIndex).QtyRcvdOriginal Then
                    Billing.ApproveRMAAdjustmentForQtyReceived .Item(lIndex).RmaLineKey, .Item(lIndex).QtyRcvd
                End If
            End If
        Next
    End With
    
    If m_oAdjs.Count > 0 Then
        m_oAdjs.whseid = cboWhse(0).ItemData(cboWhse(0).ListIndex)
        m_oAdjs.BatchDescr = "Process Inventory Transactions " & GetUserName & " " & Format(Now, "MM/DD/YYYY")
        Call m_oAdjs.ProcessBatch
    End If
    
    g_DB.Connection.CommitTrans
    
    If Len(lsMessage) > 0 Or Len(lsInventoryMessage) > 0 Then MsgBox lsMessage & lsInventoryMessage, vbInformation, "RMA Approval"
    
    LoadRMAReceived cboWhse(0).ItemData(cboWhse(0).ListIndex)
    UpdateRMAApproveRemarks
    SetWaitCursor False
Exit Sub

ErrorRollback:
    On Error Resume Next
    Set m_oAdjs = Nothing
    g_DB.Connection.RollbackTrans
    SetWaitCursor False
End Sub


'*********************************************************************************
'RMA Credit tab
'*********************************************************************************

Private Sub LoadRMAApproved(ByVal lWhseKey As Long)

    Set m_oRMACreditList = Billing.LoadRMACredit(lWhseKey)
    
    Dim i As Integer
    With gdxRMACred
        .HoldFields
        .ItemCount = m_oRMACreditList.Count
        .Refetch
        .Row = 1
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
End Sub


Private Sub gdxRMACred_Click()
    UpdateRMACreditRemarks
End Sub


Private Sub gdxRMACred_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        UpdateRMACreditRemarks
    End If
End Sub


Private Sub gdxRMACred_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    If (ColIndex = 4 Or ColIndex = 23) And Not gdxRMACred.GetRowData(gdxRMACred.Row).value(5) Then
        Cancel = True
    End If
End Sub


Private Sub gdxRMACred_AfterColEdit(ByVal ColIndex As Integer)
    gdxRMACred.Update
End Sub


Private Sub gdxRMACred_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_oRMACreditList Is Nothing Then Exit Sub
    
    If RowIndex > m_oRMACreditList.Count Then Exit Sub
    
    With m_oRMACreditList.Item(RowIndex)
        Values(1) = .RMAKey
        Values(2) = .SOLineKey
        Values(3) = .OPKey
        Values(4) = .QtyCred
        Values(5) = .Credited
        Values(6) = .ExtPrice
        Values(7) = .ItemID
        Values(8) = .QtyAuthorized
        Values(9) = .AuthBy
        Values(10) = .AuthDate
        Values(11) = .QtyRcvd
        Values(12) = .Cost
        Values(13) = .Price
        Values(14) = .RMAInfo
        Values(15) = .OPLineKey
        Values(16) = .QtyPreCred
        Values(17) = m_colReason(CStr(.Reason))
        Values(18) = .Restock * 100
        Values(19) = .CreditFreight
        Values(20) = .RcvdWhseID
        Values(21) = .Descr
        Values(22) = .ApproveDate
        Values(23) = .CMNbr
        '06/12/03 AVH PRN#5 Provide an option to sort the grid by RMA# or by CustID
        Values(24) = .RMAInfo_ByCustID
    End With
End Sub


Private Sub gdxRMACred_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    With m_oRMACreditList.Item(RowIndex)
        .Credited = IIf(Values(5), True, False)
        .CMNbr = Values(23)
        
        If Not IsNumeric(Values(4)) Then Exit Sub
        
        If CLng(Values(4) > CLng(Values(11)) - CLng(Values(16))) Then
            msg "Sorry. QtyCred must be less than or equal to " & (CLng(Values(11)) - CLng(Values(16))) _
            & vbCrLf & "Please re-enter.", vbExclamation + vbOKOnly, "Enter Qty Creidt"
        Else
            .QtyCred = CLng(Values(4))
        End If
    End With
End Sub


Private Sub gdxRMACred_LostFocus()
    gdxRMACred.Update
End Sub


Private Sub gdxRMACred_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxRMACred.Update
End Sub


Private Sub UpdateRMACreditRemarks()
    If gdxRMACred.value(3) = Empty Or IsNull(gdxRMACred.value(3)) Then
        rvRMACredit.Visible = False
    Else
        rvRMACredit.Visible = True
        rvRMACredit.OwnerID = ""
        rvRMACredit.OwnerID = gdxRMACred.value(3)
    End If
End Sub


Private Sub UpdateRMACredit()
    Dim lIndex As Long
    Dim lStatus As Long
    Dim cmd As ADODB.Command
    Dim oCurrentTime As Date
    Dim sEventStrValue As String
    
    'if there are no RMAs in the collection, exit
    If m_oRMACreditList.Count = 0 Then Exit Sub
       
    SetWaitCursor True
    With m_oRMACreditList
        For lIndex = 1 To .Count
            If .Item(lIndex).Credited And .Item(lIndex).QtyCred > 0 Then
                Billing.UpdateRMAItemCredit .Item(lIndex).RmaLineKey, .Item(lIndex).QtyCred + .Item(lIndex).QtyPreCred, .Item(lIndex).CMNbr
            End If
        Next
    End With
    
    LoadRMAApproved cboWhse(1).ItemData(cboWhse(1).ListIndex)
    UpdateRMAApproveRemarks
    SetWaitCursor False
End Sub


'*********************************************************************************
'Billing Summary tab (tmiSummary = 7)
'*********************************************************************************

Private Sub cmdBillingSummary_Click()
    Dim sBatchID As String
    Dim sSQL As String
    Dim rstCOD As ADODB.Recordset
    Dim rstDI As ADODB.Recordset
    
    If Len(Trim(txtBatchNumber.text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtBatchNumber.text) Then
        msg "Invalid Batch Number", vbCritical
        Exit Sub
    End If
    
    SetWaitCursor True
    
    Set rstCOD = CallSP("spCPCGetBillingSummaryCOD", "@_iBatchNo", CLng(Trim(txtBatchNumber.text)))
    Set rstDI = CallSP("spCPCGetBillingSummaryDI", "@_iBatchNo", CLng(Trim(txtBatchNumber.text)))
 
    If rstCOD.EOF And rstDI.EOF Then
        msg "Search returns no result."
        txtBatchNumber.SelStart = 0
        txtBatchNumber.SelLength = Len(txtBatchNumber.text)
        TryToSetFocus txtBatchNumber
    End If
    
    With gdxCOD
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = rstCOD
    End With
    FormatTranAmt gdxCOD
    
    With gdxDI
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = rstDI
    End With
    FormatTranAmt gdxDI
    
    SetWaitCursor False
End Sub


'************************************************************************
' Will Call Orders tab
'************************************************************************

'Private Sub gdxWillCallOrders_DblClick()
'    With m_gwWillCallOrders
'        LoadOrder .value("OPKey"), .value("CustID"), .value("Gasket")
'    End With
'End Sub

Private Sub m_gwWillCallOrders_ColumnChosen(columnName As String)
    Select Case columnName
        Case "OP#"
            MsgBox "Open Order " & m_gwWillCallOrders.value("OP")
'            With m_gwWillCallOrders
'                LoadOrder .value("OPKey"), .value("CustID"), .value("Gasket")
'            End With
        Case "Remarks"
            Dim oRC As RemarkContext
            Set oRC = New RemarkContext
            oRC.Edit "WillCallBilling", m_gwWillCallOrders.value("OP")
    
    End Select
End Sub




'Called by:
'   gdxWillCallOrders_DblClick

'********** Is there similar code in other forms? *****************

Private Sub LoadOrder(OPKey As Long, CustID As String, Gasket As Boolean)
    Dim oFrm As FOrder
    Dim bChgCust As Boolean
    Dim lBillingAddressKey As Long
    Dim lShippingAddressKey As Long
    Dim lMiscBillAddrKey As Long
    Dim lMiscShipAddrKey As Long
    Dim lDfltCntctKey As Long
    Dim lGsktCustKey As Long
    Dim oCustomer As Customer
    
    'if the order is for a MISC account AND the order contains a gasket then
    bChgCust = False
    If Right(Trim(CustID), 4) = "MISC" And Gasket Then
        If vbYes = MsgBox("Would you like to change this order from a MISC customer to a GSKT customer before proceeding?", vbYesNo, "Open WillCall Order") Then
            bChgCust = True
        End If
    End If

    LogEvent "FBilling", "LoadOrder", GetUserName & " instantiating FOrder from FBilling for OP " & OPKey
    
    SetWaitCursor True
    Set oFrm = New FOrder
    Set oCustomer = New Customer
        
    MDIMain.AddNewWindow oFrm
    With oFrm
        .Show
        .Order.Load OPKey
        If bChgCust Then
            lGsktCustKey = getGaskCustKey(Left(Trim(CustID), 4) & "GSKT", lBillingAddressKey, lShippingAddressKey) 'reload default customer info
            oCustomer.Load lGsktCustKey
            GetMiscAddrKey Trim(CustID), lMiscBillAddrKey, lMiscShipAddrKey
            lDfltCntctKey = getDefaultCntctKey(lGsktCustKey)
    
            If .Order.Customer.BillAddr.AddrKey = 0 Then
                oCustomer.BillAddr = .Order.Customer.BillAddr
            Else
                oCustomer.BillAddr.Load lBillingAddressKey
            End If
            
            If .Order.Customer.ShipAddr.AddrKey = 0 Then
                oCustomer.ShipAddr = .Order.Customer.ShipAddr
            Else
                oCustomer.ShipAddr.Load lShippingAddressKey
            End If
            
            .Order.Customer = oCustomer
            .Order.Save
        End If
        .Customer = .Order.Customer
        .Items = .Order.Items
        .lblCustName.Visible = True
        .lblCustType(0).Visible = True
        .txtCustName.Visible = False
        .cboCustType.Visible = False

        .TransitionTabs False
    End With
    SetWaitCursor False
End Sub


'Called by:
'   LoadOrder

Private Function getDefaultCntctKey(lCustKey As Long)
    Dim orst As ADODB.Recordset
    
    Set orst = LoadDiscRst("Select DfltCntctKey from tarCustAddr where CustKey = " & lCustKey)
    If Not orst.EOF Then
        getDefaultCntctKey = orst.Fields("DfltCntctKey").value
    End If
    Set orst = Nothing
End Function


'Called by:
'   LoadOrder

Private Sub GetMiscAddrKey(sCustID As String, ByRef lMiscBillAddrKey As Long, ByRef lMiscShipAddrKey As Long)
    Dim orst As ADODB.Recordset

    Set orst = LoadDiscRst("Select CustKey, DfltBillToAddrKey, DfltShipToAddrKey from tarCustomer Where CompanyID = 'CPC' and CustID = '" & sCustID & "'")
    If Not orst.EOF Then
        lMiscBillAddrKey = orst.Fields("DfltBillToAddrKey").value
        lMiscShipAddrKey = orst.Fields("DfltShipToAddrKey").value
    End If
    Set orst = Nothing
End Sub


'Called by:
'   LoadOrder

Private Function getGaskCustKey(sCustID As String, ByRef lBillAddrKey As Long, ByRef lShipAddrKey As Long) As Long
    Dim sSQL As String
    Dim rst As ADODB.Recordset
    
    sSQL = "Select CustKey, DfltBillToAddrKey, DfltShipToAddrKey  from tarCustomer where CompanyID = 'CPC' and CustID = '" & sCustID & "'"
    Set rst = LoadDiscRst(sSQL)
    
    If Not rst.EOF Then
        getGaskCustKey = rst.Fields("CustKey").value
        lBillAddrKey = rst.Fields("DfltBillToAddrKey").value
        lShipAddrKey = rst.Fields("DfltShipToAddrKey").value
    End If
    Set rst = Nothing
End Function


Private Sub cboRMACreditSortBy_Click()
    Const LISTINDEX_RMA_INFO = 0
    Const LISTINDEX_RMA_INFO_BY_CUSTID = 1
    Const COLINDEX_RMA_INFO = 14
    Const COLINDEX_RMA_INFO_BY_CUSTID = 24
    
    '06/12/03 AVH PRN#5 Provide an option to sort the grid by RMA# or by CustID
    If m_bLoading Then Exit Sub
    With gdxRMACred
        .ItemCount = m_oRMACreditList.Count
        'Sort by combo selection
        Select Case cboRMACreditSortBy.ListIndex
            Case LISTINDEX_RMA_INFO
                .Groups.Item(1).ColIndex = COLINDEX_RMA_INFO
            Case LISTINDEX_RMA_INFO_BY_CUSTID
                .Groups.Item(1).ColIndex = COLINDEX_RMA_INFO_BY_CUSTID
            Case Else
                MsgBox "Sort by " & cboRMACreditSortBy.text & " has not been implemented. Please inform the system administrator.", vbInformation
                Exit Sub
        End Select
        .Refetch
    End With
End Sub


'************************************************************************
' Utility Functions
'************************************************************************

Private Sub AttachGrid(ByRef i_oGrid As GridEX, ByRef i_orst As ADODB.Recordset)
    With i_oGrid
        Dim i As Long
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = i_orst
        For i = 1 To .Columns.Count
            If .Columns(i).Key <> "TrackingNo" Then
                .Columns(i).AutoSize
            End If
        Next
    End With
End Sub


Private Sub FormatTranAmt(ByRef oGrid As GridEX)
    Dim fmtcon As JSFmtCondition
    Dim col As JSColumn
    
    Set col = oGrid.Columns("TranAmt")
    Set fmtcon = oGrid.FmtConditions.Add(col.Index, jgexLessThan, 0)
    fmtcon.FormatStyle.BackColor = vbRed
End Sub


'********** Are there similar functions elsewhere?  How about ValidationRules? ***********

'*******************************************************************************
' VALIDATION FUNCTIONS
'*******************************************************************************

Private Function IsInteger(ByVal i_sText As String) As Boolean
    Dim i As Long
    Dim chr As Long

    For i = 1 To Len(i_sText)
        chr = Asc(Mid(i_sText, i, 1))
        If chr < Asc("0") Or chr > Asc("9") Then
            IsInteger = False
            Exit Function
        End If
    Next
    IsInteger = True
End Function


Private Function AlphanumericOnly(sInput As String) As String
    Dim lInputLength As Long
    Dim ichar As Integer
    Dim sTemp As String
    Dim i As Integer
    
    sTemp = ""
    lInputLength = Len(sInput)
    For i = 1 To lInputLength
        ichar = Asc(Mid$(sInput, i, 1))
        If (IsAlpha(ichar) Or IsDigit(ichar)) Then
              sTemp = sTemp & Mid$(sInput, i, 1)
        End If
    Next
    
    AlphanumericOnly = sTemp

End Function


Private Function IsAlpha(ByVal i_iChar As Integer) As Boolean
    If (i_iChar >= Asc("a") And i_iChar <= Asc("z")) _
    Or (i_iChar >= Asc("A") And i_iChar <= Asc("Z")) Then
        IsAlpha = True
    End If
End Function


Private Function IsDigit(ByVal i_iChar As Integer) As Boolean
    If (i_iChar >= Asc("0") And i_iChar <= Asc("9")) Then
        IsDigit = True
    End If
End Function


'***********************************************************************************
' BEGIN CREDIT CARD STUFF
'***********************************************************************************

'Is William mapping the recordset to an array and using a unbound grid simply to support
'the selection checkbox column?

'orstCCOrders (spcpcGetBillingSummaryCCRD2) now contains only
'    OPKey
'    ChargeCC = 1
'    InvcKey
'    TranID (Invoice)
'    TranAmt
'    BatchNo

Private Sub cmdGetBatchCC_Click()
    Dim orstCCOrders As ADODB.Recordset
    Dim oOrder As Order
    Dim orstCCCharged As ADODB.Recordset
    Dim RowIndex As Integer
    Dim i As Integer
    Dim NumRows As Integer

    On Error GoTo EH
    
    If Len(Trim(txtBatchNumberCC.text)) = 0 Then
        MsgBox "Please enter a valid batch number."
        Exit Sub
    End If
    If IsNumeric(txtBatchNumberCC.text) = False Then
        MsgBox "Please enter a valid batch number."
        Exit Sub
    End If
    
    SetWaitCursor True
    
    Set orstCCOrders = CallSP("spcpcGetBillingSummaryCCRD2", "@_iBatchNo", CLng(txtBatchNumberCC.text))
    Set orstCCCharged = CallSP("spcpGetCCOrderCharged", "@_iBatchNo", CLng(txtBatchNumberCC.text))
    
    If orstCCOrders.EOF Then
        cmdGetBatchCC.Default = False
        cmdCreateReport.Default = True
    Else
        cmdGetBatchCC.Default = False
        cmdChargeCC.Default = True
    End If
    
    If orstCCOrders.EOF = True And orstCCCharged.EOF = True Then
        MsgBox "There are no credit card orders for that batch number."
        'and continue on...
    End If

'Build the array we'll bind (through events) with the grid
'Will need to destroy the Order objects references when done

    NumRows = orstCCOrders.RecordCount
    Debug.Print "NumRows: " + CStr(NumRows)
    If Not orstCCOrders.EOF Then
        'transfer the contents of the recordset to an array
        
        '*** BE SURE TO change the column range if you add or remove columns ***
        ReDim m_arrayCCOrders(1 To 11, 1 To NumRows) As Variant
        RowIndex = 1
        orstCCOrders.MoveFirst
        
        Do Until orstCCOrders.EOF
        
            Set oOrder = New Order
            oOrder.Load orstCCOrders.Fields("OPKey").value
            
            '10/25/04 LR added to catch null CCKeys (a bug upstream)
            If oOrder.CreditCard Is Nothing Then
                'EMail.SendToList "01", GetUserName & "@caseparts.com", "Error loading Invoice Batch " & txtBatchNumberCC.text, "OP " & oOrder.OPKey & " has a Null CC Key", False 'TextFormat
                Debug.Print "Error loading Invoice Batch " & txtBatchNumberCC.text, "OP " & oOrder.OPKey & " has a Null CC Key"
                'LogEvent "FBilling", "cmdGetBatchCC_Click", "Error loading Invoice Batch " & txtBatchNumberCC.text & ". OP " & oOrder.OPKey & " has a Null CC Key."
                
                NumRows = NumRows - 1
                Debug.Print "NumRows: " + CStr(NumRows)
                If NumRows > 0 Then
                    ReDim Preserve m_arrayCCOrders(1 To 11, 1 To NumRows)
                End If
            Else
                m_arrayCCOrders(1, RowIndex) = orstCCOrders.Fields("ChargeCC").value
                m_arrayCCOrders(2, RowIndex) = orstCCOrders.Fields("TranID").value
                m_arrayCCOrders(3, RowIndex) = oOrder.Customer.Name
                m_arrayCCOrders(4, RowIndex) = orstCCOrders.Fields("TranAmt").value
                m_arrayCCOrders(5, RowIndex) = oOrder.CreditCard.MaskedCCNo
                m_arrayCCOrders(6, RowIndex) = oOrder.CreditCard.CardHolderName
                m_arrayCCOrders(7, RowIndex) = oOrder.CreditCard.TypeID
                m_arrayCCOrders(8, RowIndex) = orstCCOrders.Fields("InvcKey").value
                'I don't think this column is needed (in the array or the recordset)
                m_arrayCCOrders(9, RowIndex) = orstCCOrders.Fields("BatchNo").value
                Set m_arrayCCOrders(10, RowIndex) = oOrder
                m_arrayCCOrders(11, RowIndex) = orstCCOrders.Fields("RespMsg").value
                lblCardCount.caption = RowIndex
                DoEvents
                RowIndex = RowIndex + 1
            End If
            
            Set oOrder = Nothing
            orstCCOrders.MoveNext
        Loop
    End If

    'If the recordset is empty, NumRows = 0, so the upper grid will be empty
    lblCardCount.caption = NumRows
    With gdxCCOrders
        .HoldSortSettings = True
        .HoldFields
        .ItemCount = NumRows
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
        TryToSetFocus gdxCCOrders
    End With
       
    AttachGrid gdxCCCharged, orstCCCharged
    lblChargeCount.caption = orstCCCharged.RecordCount
    
    SetWaitCursor False
        
    Exit Sub

EH:
    
    msg "* Runtime error" & vbCrLf & _
            vbTab & "Err Num: " & vbTab & Err.Number & vbCrLf & _
            vbTab & "Err Source: " & vbTab & Err.Source & vbCrLf & _
            vbTab & "Err Desc: " & vbTab & Err.Description & vbCrLf & _
            vbCrLf & _
            "* Source: " & "SageAssistant.cmdGetBatchCC_Click()" & vbCrLf & _
            vbCrLf & _
            "*Timestamp: " & Now & " " & Timer
   
    SetWaitCursor False
    
End Sub


Private Sub cmdChargeCC_Click()
    Dim oOrder As Order
    'Dim oCCTran As CCTransaction
    Dim RowIndex As Integer
    Dim AuthID As String
    Dim RetCode As Long

    On Error GoTo EH

    If gdxCCOrders.ItemCount = 0 Then
        MsgBox "Charge Credit Card Grid is empty." & vbCrLf & "You cannot perform this operation until its filled."
        Exit Sub
    End If

    SetWaitCursor True

    'By default the Janus grid does not update until you move off the cell.
    'This will catch the last checkbox change made with the mouse.
    gdxCCOrders.Update
   
    'for each row in the array associated with the unbound grid
    For RowIndex = 1 To UBound(m_arrayCCOrders, 2)
    
        'if the checkbox is checked for this row
        If m_arrayCCOrders(1, RowIndex) = True Then
'            Set oCCTran = New CCTransaction
            
            'grab the order reference and destroy it in the array
            Set oOrder = m_arrayCCOrders(10, RowIndex)
            Set m_arrayCCOrders(10, RowIndex) = Nothing  '*** not sufficient

            'get the invoice key
'            oCCTran.InvcKey = CLng(m_arrayCCOrders(8, RowIndex))
            'get the invoice amount
'            oCCTran.Amount = CDbl(m_arrayCCOrders(4, RowIndex))
'            oCCTran.Charge oOrder, CStr(m_arrayCCOrders(2, RowIndex))   'Index 2 = Invoice TranID
'            oOrder.CCTransactions.Add oCCTran
'            If oCCTran.Result <> 0 Then

             If oOrder.CreditCard.Charge(CStr(m_arrayCCOrders(2, RowIndex)), CLng(m_arrayCCOrders(8, RowIndex)), CDbl(m_arrayCCOrders(4, RowIndex))) <> 0 Then
                MarkCCOrderAsBad oOrder 'CLng(m_arrayCCOrders(9, RowIndex))
            End If
            
'***
            lblChargeCount.caption = RowIndex
            DoEvents
        End If
    Next RowIndex

'Executing the button click event on this tab as a a grid refresh
'this will REDIM m_arrayCCOrders and any remaining orders
'what about the Order reference pointers?
    cmdGetBatchCC_Click

    SetWaitCursor False

    cmdChargeCC.Default = False
    cmdCreateReport.Default = True
    
    Exit Sub

EH:
    msg "* Runtime error" & vbCrLf & _
            vbTab & "Err Num: " & vbTab & Err.Number & vbCrLf & _
            vbTab & "Err Source: " & vbTab & Err.Source & vbCrLf & _
            vbTab & "Err Desc: " & vbTab & Err.Description & vbCrLf & _
            vbCrLf & _
            "* Source: " & "SageAssistant.cmdChargeCC_Click()"

    SetWaitCursor False

End Sub


Private Sub cmdCreateReport_Click()
    On Error GoTo ErrorHandler
    
    If Len(Trim(txtBatchNumberCC.text)) = 0 Then
        MsgBox "Please enter a valid batch number."
        Exit Sub
    End If
    If IsNumeric(txtBatchNumberCC.text) = False Then
        MsgBox "Please enter a valid batch number."
        Exit Sub
    End If
    
    Dim sURL As String
    Dim frm As FViewReport
    Set frm = New FViewReport
    MDIMain.AddNewWindow frm
    With frm
        .SetCaption ("CreditCard Report")
        sURL = g_CrCardReportURL & "?BatchNumber=" & Trim(CStr(txtBatchNumberCC.text)) & "&Server=" & g_DB.server
        .PopUp sURL, False
    End With
    
    Exit Sub
    
ErrorHandler:

    msg "* Runtime error" & vbCrLf & _
            vbTab & "Err Num: " & vbTab & Err.Number & vbCrLf & _
            vbTab & "Err Source: " & vbTab & Err.Source & vbCrLf & _
            vbTab & "Err Desc: " & vbTab & Err.Description & vbCrLf & _
            vbCrLf & _
            "* Source: " & "SageAssistant.cmdCreateReport_Click()"
    
End Sub


'Setting ItemCount fires off a sequence of these events.

Private Sub gdxCCOrders_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
'Load the grid row specified by RowIndex
    Values(1) = m_arrayCCOrders(1, RowIndex)    'CCCharge flag
    Values(2) = m_arrayCCOrders(2, RowIndex)    'TranID
    Values(3) = m_arrayCCOrders(3, RowIndex)    'CustName
    Values(4) = m_arrayCCOrders(4, RowIndex)    'TranAmt
    Values(5) = m_arrayCCOrders(5, RowIndex)    'CrCardNo (masked)
    Values(6) = m_arrayCCOrders(6, RowIndex)    'CardHolderName
    Values(7) = m_arrayCCOrders(7, RowIndex)    'CrCardType
    Values(8) = m_arrayCCOrders(11, RowIndex)   'RespMsg
End Sub


'The checkbox in the first column can be toggled in the grid.
'If it is, write the change back into the associated array.

Private Sub gdxCCOrders_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    m_arrayCCOrders(1, RowIndex) = Values(1)    'ChargeCC flag
End Sub


Private Sub MarkCCOrderAsBad(ByRef oOrder As Order)  '(ByVal OPKey As Long)
    Dim aTrans As CCTransaction
    Dim output As String

    output = "Type Amount  CreateDate             PNRef         Comment1  Comment2    Result  RespMsg" + vbCrLf

    For Each aTrans In oOrder.CreditCard.Transactions
        output = output & aTrans.TranType '   Trim(rst.Fields("trantype").Value) & "    "
        output = output & aTrans.Amount '  rst.Fields("amount").Value & "  "
        output = output & aTrans.TimeStamp  '  rst.Fields("createdate").Value & "  "
        output = output & aTrans.PNREF ' rst.Fields("pnref").Value & "  "
        output = output & aTrans.Comment1 ' Trim(rst.Fields("comment1").Value) & "  "
        output = output & aTrans.Comment2  ' Trim(rst.Fields("comment2").Value) & "  "
        output = output & aTrans.Result  ' rst.Fields("result").Value & "  "
        output = output & aTrans.RESPMSG ' Trim(rst.Fields("respmsg").Value) & vbCrLf
    Next
    
    If Len(output) > 0 Then
        EMail.Send GetUserName & "@caseparts.com", "operations@caseparts.com", "Credit Card Billing Error", output, False 'TextFormat
    End If
    
End Sub



Private Sub cmdRefresh_Click()
    RefreshWillCallList
End Sub


Private Sub cmdFindWCOrder_Click()
    Dim Found As Boolean
    
    If Len(txtFindWCOrder.text) = 0 Then Exit Sub

    If Not IsNumeric(txtFindWCOrder.text) Then
        msg "Invalid OP number", vbCritical
        Exit Sub
    End If

    Found = gdxWillCallOrders.Find(1, jgexContains, txtFindWCOrder.text)
    If Not Found Then
        msg "OP number not found", vbExclamation
        Exit Sub
    End If
    gdxWillCallOrders.EnsureVisible gdxWillCallOrders.Row
End Sub


'This grid supports open MPK Will Call orders

Private Sub RefreshWillCallList()
    Dim i As Integer
    
    SetWaitCursor True

'    If cboWillCallWhse.ItemData(cboWillCallWhse.ListIndex) = 0 Then
'        Set m_orstWillCallOrders = CallSP("spcpcGetOpenWillCallOrders")
'    Else
'        Set m_orstWillCallOrders = CallSP("spcpcGetOpenWillCallOrders", "@_iWhseKey", cboWillCallWhse.ItemData(cboWillCallWhse.ListIndex))
'    End If
    
    Set m_orstWillCallOrders = CallSP("spcpcGetOpenWillCallOrders")
    
    lblOpenWCOrders.caption = m_orstWillCallOrders.RecordCount & " open Will Call orders"
    
    AttachGrid gdxWillCallOrders, m_orstWillCallOrders

    SetWaitCursor False
End Sub


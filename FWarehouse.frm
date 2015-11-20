VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "mmremark.ocx"
Begin VB.Form FWarehouse 
   Caption         =   "Warehouse Tool"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   9510
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   5895
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10398
      _Version        =   262144
      TabCount        =   5
      TagVariant      =   ""
      Tabs            =   "FWarehouse.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   5505
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   9710
         _Version        =   262144
         TabGuid         =   "FWarehouse.frx":0117
         Begin VB.CommandButton cmdMoveDown 
            Caption         =   "Move &Down"
            Height          =   375
            Left            =   7080
            TabIndex        =   8
            Top             =   5040
            Width           =   1215
         End
         Begin VB.CommandButton cmdMoveUp 
            Caption         =   "Move &Up"
            Height          =   375
            Left            =   5520
            TabIndex        =   7
            Top             =   5040
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "De&lete"
            Height          =   375
            Left            =   3960
            TabIndex        =   6
            Top             =   5040
            Width           =   1215
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "N&ew"
            Height          =   375
            Left            =   2400
            TabIndex        =   5
            Top             =   5040
            Width           =   1215
         End
         Begin VB.CommandButton cmdFindBin 
            Caption         =   "Fi&nd"
            Height          =   375
            Left            =   2880
            TabIndex        =   4
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtItemID 
            Height          =   375
            Left            =   1320
            TabIndex        =   3
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdNewBin 
            Caption         =   "Ne&w Bin"
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   5040
            Width           =   1215
         End
         Begin GridEX20.GridEX gdxBin 
            Height          =   4215
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   7435
            Version         =   "2.0"
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            ColumnsCount    =   3
            Column(1)       =   "FWarehouse.frx":013F
            Column(2)       =   "FWarehouse.frx":02E3
            Column(3)       =   "FWarehouse.frx":045B
            SortKeysCount   =   1
            SortKey(1)      =   "FWarehouse.frx":05B3
            FormatStylesCount=   6
            FormatStyle(1)  =   "FWarehouse.frx":061B
            FormatStyle(2)  =   "FWarehouse.frx":0753
            FormatStyle(3)  =   "FWarehouse.frx":0803
            FormatStyle(4)  =   "FWarehouse.frx":08B7
            FormatStyle(5)  =   "FWarehouse.frx":098F
            FormatStyle(6)  =   "FWarehouse.frx":0A47
            ImageCount      =   0
            PrinterProperties=   "FWarehouse.frx":0B27
         End
         Begin VB.Label lblItemID 
            Alignment       =   1  'Right Justify
            Caption         =   "Item ID"
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   195
            Width           =   735
         End
         Begin VB.Label lblBinManager 
            AutoSize        =   -1  'True
            Caption         =   "Item ID"
            Height          =   195
            Left            =   4425
            TabIndex        =   10
            Top             =   195
            Width           =   510
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   5505
         Left            =   30
         TabIndex        =   12
         Top             =   360
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   9710
         _Version        =   262144
         TabGuid         =   "FWarehouse.frx":0CFF
         Begin VB.Frame frmView 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   4212
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Visible         =   0   'False
            Width           =   9012
            Begin VB.CommandButton cmdRMAEdit 
               Caption         =   "Edit"
               Height          =   375
               Left            =   7920
               TabIndex        =   37
               Top             =   3720
               Width           =   1095
            End
            Begin GridEX20.GridEX gdxRMA 
               Height          =   3075
               Left            =   0
               TabIndex        =   38
               Top             =   480
               Width           =   9060
               _ExtentX        =   15981
               _ExtentY        =   5424
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
               ColumnsCount    =   16
               Column(1)       =   "FWarehouse.frx":0D27
               Column(2)       =   "FWarehouse.frx":0E67
               Column(3)       =   "FWarehouse.frx":0FB3
               Column(4)       =   "FWarehouse.frx":10E7
               Column(5)       =   "FWarehouse.frx":126B
               Column(6)       =   "FWarehouse.frx":13AB
               Column(7)       =   "FWarehouse.frx":14EF
               Column(8)       =   "FWarehouse.frx":1627
               Column(9)       =   "FWarehouse.frx":176B
               Column(10)      =   "FWarehouse.frx":18AF
               Column(11)      =   "FWarehouse.frx":1A2F
               Column(12)      =   "FWarehouse.frx":1B87
               Column(13)      =   "FWarehouse.frx":1D43
               Column(14)      =   "FWarehouse.frx":1E83
               Column(15)      =   "FWarehouse.frx":1FCF
               Column(16)      =   "FWarehouse.frx":20EF
               GroupCount      =   1
               Group(1)        =   "FWarehouse.frx":2203
               FormatStylesCount=   6
               FormatStyle(1)  =   "FWarehouse.frx":226B
               FormatStyle(2)  =   "FWarehouse.frx":23A3
               FormatStyle(3)  =   "FWarehouse.frx":2453
               FormatStyle(4)  =   "FWarehouse.frx":2507
               FormatStyle(5)  =   "FWarehouse.frx":25DF
               FormatStyle(6)  =   "FWarehouse.frx":2697
               ImageCount      =   0
               PrinterProperties=   "FWarehouse.frx":2777
            End
            Begin VB.Label Label24 
               Caption         =   "Please select RMA order that you want to edit receiving items"
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   0
               TabIndex        =   39
               Top             =   120
               Width           =   4575
            End
         End
         Begin VB.Frame frmEdit 
            Caption         =   "Edit Receiving Items"
            Height          =   4272
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Visible         =   0   'False
            Width           =   9015
            Begin VB.CommandButton cmdReceive 
               Caption         =   "&Receive Items"
               Height          =   372
               Left            =   7560
               TabIndex        =   25
               Top             =   3780
               Visible         =   0   'False
               Width           =   1272
            End
            Begin VB.CommandButton cmdPrintLabels 
               Caption         =   "Print Labels"
               Height          =   372
               Index           =   1
               Left            =   6060
               TabIndex        =   24
               Top             =   3780
               Width           =   1272
            End
            Begin GridEX20.GridEX gdxRMAEdit 
               Height          =   2412
               Left            =   120
               TabIndex        =   26
               Top             =   1140
               Width           =   8772
               _ExtentX        =   15478
               _ExtentY        =   4260
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
               ColumnsCount    =   16
               Column(1)       =   "FWarehouse.frx":294F
               Column(2)       =   "FWarehouse.frx":2AB3
               Column(3)       =   "FWarehouse.frx":2BFF
               Column(4)       =   "FWarehouse.frx":2D33
               Column(5)       =   "FWarehouse.frx":2E73
               Column(6)       =   "FWarehouse.frx":2F97
               Column(7)       =   "FWarehouse.frx":30DB
               Column(8)       =   "FWarehouse.frx":3213
               Column(9)       =   "FWarehouse.frx":3357
               Column(10)      =   "FWarehouse.frx":349B
               Column(11)      =   "FWarehouse.frx":35EB
               Column(12)      =   "FWarehouse.frx":373F
               Column(13)      =   "FWarehouse.frx":38B7
               Column(14)      =   "FWarehouse.frx":3A23
               Column(15)      =   "FWarehouse.frx":3B5F
               Column(16)      =   "FWarehouse.frx":3CA3
               FormatStylesCount=   6
               FormatStyle(1)  =   "FWarehouse.frx":3E07
               FormatStyle(2)  =   "FWarehouse.frx":3EE7
               FormatStyle(3)  =   "FWarehouse.frx":401F
               FormatStyle(4)  =   "FWarehouse.frx":40CF
               FormatStyle(5)  =   "FWarehouse.frx":4183
               FormatStyle(6)  =   "FWarehouse.frx":425B
               ImageCount      =   0
               PrinterProperties=   "FWarehouse.frx":4313
            End
            Begin MMRemark.RemarkViewer rvRMA 
               Height          =   810
               Left            =   7200
               TabIndex        =   27
               Top             =   300
               Visible         =   0   'False
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "ViewRMA"
               Caption         =   "RMA Remarks"
            End
            Begin VB.Label lblCustomer 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   960
               TabIndex        =   35
               Top             =   780
               Width           =   5655
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer"
               Height          =   375
               Left            =   120
               TabIndex        =   34
               Top             =   780
               Width           =   735
            End
            Begin VB.Label lblSO 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   5160
               TabIndex        =   33
               Top             =   300
               Width           =   1455
            End
            Begin VB.Label lblOP 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3000
               TabIndex        =   32
               Top             =   300
               Width           =   1335
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "SO"
               Height          =   255
               Left            =   4560
               TabIndex        =   31
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "OP"
               Height          =   255
               Left            =   2400
               TabIndex        =   30
               Top             =   300
               Width           =   495
            End
            Begin VB.Label lblRMA 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   960
               TabIndex        =   29
               Top             =   300
               Width           =   1215
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "RMA"
               Height          =   255
               Left            =   360
               TabIndex        =   28
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Find RMA via Customer, PartNbr, or both"
            Height          =   915
            Left            =   3240
            TabIndex        =   17
            Top             =   60
            Width           =   5895
            Begin VB.TextBox txtNewCust 
               Height          =   315
               Left            =   960
               TabIndex        =   20
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox txtNewItem 
               Height          =   315
               Left            =   3360
               TabIndex        =   19
               Top             =   360
               Width           =   1215
            End
            Begin VB.CommandButton cmdFind 
               Caption         =   "Find RMA"
               Height          =   315
               Index           =   1
               Left            =   4680
               TabIndex        =   18
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "Part Number"
               Height          =   255
               Left            =   2280
               TabIndex        =   21
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Find RMA via RMA, OP, or SO"
            Height          =   915
            Left            =   120
            TabIndex        =   13
            Top             =   60
            Width           =   3015
            Begin VB.CommandButton cmdFind 
               Caption         =   "Find RMA"
               Height          =   315
               Index           =   0
               Left            =   1800
               TabIndex        =   15
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtNewRMA 
               Height          =   315
               Left            =   720
               TabIndex        =   14
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "Find"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   495
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   5505
         Left            =   30
         TabIndex        =   40
         Top             =   360
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   9710
         _Version        =   262144
         TabGuid         =   "FWarehouse.frx":44EB
         Begin VB.Frame frmPOList 
            Height          =   3072
            Left            =   180
            TabIndex        =   58
            Top             =   660
            Visible         =   0   'False
            Width           =   8892
            Begin VB.CommandButton cmdFilterPartNbr 
               Caption         =   "Find"
               Height          =   312
               Left            =   2580
               TabIndex        =   60
               Top             =   2640
               Width           =   1152
            End
            Begin VB.TextBox txtPartNbr 
               Height          =   288
               Left            =   960
               TabIndex        =   59
               Top             =   2640
               Width           =   1452
            End
            Begin GridEX20.GridEX gdxPOs 
               Height          =   1992
               Left            =   120
               TabIndex        =   61
               Top             =   480
               Width           =   8652
               _ExtentX        =   15266
               _ExtentY        =   3519
               Version         =   "2.0"
               AllowRowSizing  =   -1  'True
               AutomaticSort   =   -1  'True
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               MethodHoldFields=   -1  'True
               GroupByBoxVisible=   0   'False
               ColumnHeaderHeight=   270
               IntProp1        =   0
               IntProp2        =   0
               IntProp7        =   0
               ColumnsCount    =   4
               Column(1)       =   "FWarehouse.frx":4513
               Column(2)       =   "FWarehouse.frx":4653
               Column(3)       =   "FWarehouse.frx":477B
               Column(4)       =   "FWarehouse.frx":4897
               FormatStylesCount=   6
               FormatStyle(1)  =   "FWarehouse.frx":49C3
               FormatStyle(2)  =   "FWarehouse.frx":4AA3
               FormatStyle(3)  =   "FWarehouse.frx":4BDB
               FormatStyle(4)  =   "FWarehouse.frx":4C8B
               FormatStyle(5)  =   "FWarehouse.frx":4D3F
               FormatStyle(6)  =   "FWarehouse.frx":4E17
               ImageCount      =   0
               PrinterProperties=   "FWarehouse.frx":4ECF
            End
            Begin VB.Label lblVendName 
               Height          =   252
               Left            =   120
               TabIndex        =   63
               Top             =   180
               Width           =   2592
            End
            Begin VB.Label Label2 
               Caption         =   "Part Nbr"
               Height          =   192
               Left            =   180
               TabIndex        =   62
               Top             =   2700
               Width           =   672
            End
         End
         Begin VB.TextBox txtFindVendor 
            Height          =   288
            Left            =   4440
            TabIndex        =   56
            Top             =   180
            Width           =   1752
         End
         Begin VB.CommandButton cmdLookup 
            Caption         =   "LookUp"
            Height          =   312
            Left            =   7080
            TabIndex        =   55
            Top             =   180
            Width           =   1152
         End
         Begin VB.TextBox txtPONum 
            Height          =   288
            Left            =   1560
            TabIndex        =   54
            Top             =   180
            Width           =   1152
         End
         Begin VB.Frame frmPOHdr 
            Height          =   3012
            Left            =   180
            TabIndex        =   41
            Top             =   660
            Visible         =   0   'False
            Width           =   8892
            Begin VB.CommandButton cmdPrintLabels 
               Caption         =   "Print Labels"
               Height          =   372
               Index           =   0
               Left            =   6660
               TabIndex        =   46
               Top             =   780
               Width           =   1032
            End
            Begin VB.TextBox txtPONbr 
               Appearance      =   0  'Flat
               Height          =   288
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   240
               Width           =   1092
            End
            Begin VB.TextBox txtPODate 
               Appearance      =   0  'Flat
               Height          =   288
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   840
               Width           =   1392
            End
            Begin VB.TextBox txtBuyer 
               Appearance      =   0  'Flat
               Height          =   288
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   840
               Width           =   972
            End
            Begin VB.TextBox txtVendName 
               Appearance      =   0  'Flat
               Height          =   288
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   540
               Width           =   3072
            End
            Begin MMRemark.RemarkViewer rvPO 
               Height          =   804
               Left            =   7860
               TabIndex        =   47
               Top             =   360
               Width           =   804
               _ExtentX        =   1429
               _ExtentY        =   1429
               ContextID       =   "viewpo"
            End
            Begin GridEX20.GridEX gdxPOLines 
               Height          =   1692
               Left            =   120
               TabIndex        =   48
               Top             =   1260
               Width           =   8592
               _ExtentX        =   15161
               _ExtentY        =   2990
               Version         =   "2.0"
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               TabKeyBehavior  =   1
               MethodHoldFields=   -1  'True
               Options         =   8
               RecordsetType   =   1
               AllowEdit       =   0   'False
               GroupByBoxVisible=   0   'False
               DataMode        =   1
               ColumnHeaderHeight=   270
               ColumnsCount    =   5
               Column(1)       =   "FWarehouse.frx":50A7
               Column(2)       =   "FWarehouse.frx":51CB
               Column(3)       =   "FWarehouse.frx":52EF
               Column(4)       =   "FWarehouse.frx":542B
               Column(5)       =   "FWarehouse.frx":5567
               FormatStylesCount=   6
               FormatStyle(1)  =   "FWarehouse.frx":568F
               FormatStyle(2)  =   "FWarehouse.frx":576F
               FormatStyle(3)  =   "FWarehouse.frx":58A7
               FormatStyle(4)  =   "FWarehouse.frx":5957
               FormatStyle(5)  =   "FWarehouse.frx":5A0B
               FormatStyle(6)  =   "FWarehouse.frx":5AE3
               ImageCount      =   0
               PrinterProperties=   "FWarehouse.frx":5B9B
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
               Height          =   312
               Left            =   2160
               TabIndex        =   53
               Top             =   240
               Width           =   5652
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "PO Nbr"
               Height          =   252
               Left            =   180
               TabIndex        =   52
               Top             =   300
               Width           =   672
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Date"
               Height          =   192
               Left            =   2100
               TabIndex        =   51
               Top             =   900
               Width           =   432
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Buyer"
               Height          =   252
               Left            =   240
               TabIndex        =   50
               Top             =   900
               Width           =   612
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Vendor"
               Height          =   252
               Left            =   120
               TabIndex        =   49
               Top             =   600
               Width           =   732
            End
         End
         Begin GridEX20.GridEX gdxSOLines 
            Height          =   1575
            Left            =   240
            TabIndex        =   57
            Top             =   3720
            Visible         =   0   'False
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   2778
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   270
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   6
            Column(1)       =   "FWarehouse.frx":5D73
            Column(2)       =   "FWarehouse.frx":5EBB
            Column(3)       =   "FWarehouse.frx":5FFB
            Column(4)       =   "FWarehouse.frx":610F
            Column(5)       =   "FWarehouse.frx":623F
            Column(6)       =   "FWarehouse.frx":638B
            FormatStylesCount=   6
            FormatStyle(1)  =   "FWarehouse.frx":660F
            FormatStyle(2)  =   "FWarehouse.frx":66EF
            FormatStyle(3)  =   "FWarehouse.frx":6827
            FormatStyle(4)  =   "FWarehouse.frx":68D7
            FormatStyle(5)  =   "FWarehouse.frx":698B
            FormatStyle(6)  =   "FWarehouse.frx":6A63
            ImageCount      =   0
            PrinterProperties=   "FWarehouse.frx":6B1B
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Vendor Name"
            Height          =   192
            Left            =   3300
            TabIndex        =   65
            Top             =   240
            Width           =   1032
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "CaseParts PO #"
            Height          =   252
            Left            =   180
            TabIndex        =   64
            Top             =   240
            Width           =   1272
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5505
         Left            =   30
         TabIndex        =   66
         Top             =   360
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   9710
         _Version        =   262144
         TabGuid         =   "FWarehouse.frx":6CF3
         Begin VB.Frame Frame3 
            Caption         =   "Check for Unexploded Kits"
            Height          =   3192
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   8652
            Begin VB.ComboBox cboWhse 
               Height          =   315
               Index           =   0
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   360
               Width           =   1032
            End
            Begin VB.CommandButton cmdFind 
               Caption         =   "Find"
               Height          =   312
               Index           =   2
               Left            =   2700
               TabIndex        =   68
               Top             =   360
               Width           =   1032
            End
            Begin GridEX20.GridEX gdxUnexplodedKit 
               Height          =   2112
               Left            =   180
               TabIndex        =   70
               Top             =   960
               Width           =   8352
               _ExtentX        =   14737
               _ExtentY        =   3731
               Version         =   "2.0"
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               MethodHoldFields=   -1  'True
               AllowEdit       =   0   'False
               GroupByBoxVisible=   0   'False
               ColumnHeaderHeight=   270
               IntProp1        =   0
               IntProp2        =   0
               IntProp7        =   0
               ColumnsCount    =   3
               Column(1)       =   "FWarehouse.frx":6D1B
               Column(2)       =   "FWarehouse.frx":6E63
               Column(3)       =   "FWarehouse.frx":6F97
               FormatStylesCount=   6
               FormatStyle(1)  =   "FWarehouse.frx":70C7
               FormatStyle(2)  =   "FWarehouse.frx":71A7
               FormatStyle(3)  =   "FWarehouse.frx":72DF
               FormatStyle(4)  =   "FWarehouse.frx":738F
               FormatStyle(5)  =   "FWarehouse.frx":7443
               FormatStyle(6)  =   "FWarehouse.frx":751B
               ImageCount      =   0
               PrinterProperties=   "FWarehouse.frx":75D3
            End
            Begin VB.Label Label8 
               Caption         =   "Warehouse"
               Height          =   252
               Left            =   180
               TabIndex        =   71
               Top             =   420
               Width           =   912
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   5505
         Left            =   30
         TabIndex        =   72
         Top             =   360
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   9710
         _Version        =   262144
         TabGuid         =   "FWarehouse.frx":77AB
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Re&fresh"
            Height          =   315
            Left            =   7080
            TabIndex        =   77
            Top             =   4920
            Width           =   915
         End
         Begin VB.CommandButton cmdPrintGrid 
            Caption         =   "&Print List"
            Height          =   315
            Left            =   8100
            TabIndex        =   76
            Top             =   4920
            Width           =   915
         End
         Begin VB.CheckBox chkOmitWillCall 
            Caption         =   "Omit Will Call Orders"
            Height          =   255
            Left            =   3360
            TabIndex        =   75
            Top             =   5040
            Width           =   1935
         End
         Begin VB.ComboBox cboWhse 
            Height          =   315
            Index           =   1
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   4920
            Width           =   1035
         End
         Begin VB.CheckBox chkOmitGsk 
            Caption         =   "Omit Orders Containing GSKs"
            Height          =   255
            Left            =   3360
            TabIndex        =   73
            Top             =   4800
            Width           =   2415
         End
         Begin GridEX20.GridEX gdxToPack 
            Height          =   4635
            Left            =   0
            TabIndex        =   78
            Top             =   120
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   8176
            Version         =   "2.0"
            AllowRowSizing  =   -1  'True
            AutomaticSort   =   -1  'True
            ShowToolTips    =   -1  'True
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            RowHeight       =   19
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   16
            Column(1)       =   "FWarehouse.frx":77D3
            Column(2)       =   "FWarehouse.frx":78E7
            Column(3)       =   "FWarehouse.frx":7A03
            Column(4)       =   "FWarehouse.frx":7B97
            Column(5)       =   "FWarehouse.frx":7D17
            Column(6)       =   "FWarehouse.frx":7E57
            Column(7)       =   "FWarehouse.frx":7F7B
            Column(8)       =   "FWarehouse.frx":80AB
            Column(9)       =   "FWarehouse.frx":81EB
            Column(10)      =   "FWarehouse.frx":833F
            Column(11)      =   "FWarehouse.frx":848B
            Column(12)      =   "FWarehouse.frx":85B3
            Column(13)      =   "FWarehouse.frx":870B
            Column(14)      =   "FWarehouse.frx":885B
            Column(15)      =   "FWarehouse.frx":89CF
            Column(16)      =   "FWarehouse.frx":8B2B
            SortKeysCount   =   1
            SortKey(1)      =   "FWarehouse.frx":8C67
            FormatStylesCount=   6
            FormatStyle(1)  =   "FWarehouse.frx":8CCF
            FormatStyle(2)  =   "FWarehouse.frx":8E07
            FormatStyle(3)  =   "FWarehouse.frx":8EB7
            FormatStyle(4)  =   "FWarehouse.frx":8F6B
            FormatStyle(5)  =   "FWarehouse.frx":9043
            FormatStyle(6)  =   "FWarehouse.frx":90FB
            ImageCount      =   0
            PrinterProperties=   "FWarehouse.frx":91DB
         End
      End
   End
End
Attribute VB_Name = "FWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FormWidth = 9600
Private Const FormHeight = 6504

Private Enum TabIndex
    tmiReceiving = 1
    tmiRMA = 2
    tmiKits = 3
    tmiReadyToPack = 4
    tmiBinManager = 5
End Enum

Private m_orstToPack As ADODB.Recordset

Private m_oRMAList As RMAList
Private m_colReason As Collection

Private WithEvents m_gwToPack As GridEXWrapper
Attribute m_gwToPack.VB_VarHelpID = -1
Private WithEvents m_gwRMALine As GridEXWrapper
Attribute m_gwRMALine.VB_VarHelpID = -1
Private WithEvents m_gwPOs As GridEXWrapper
Attribute m_gwPOs.VB_VarHelpID = -1
Private WithEvents m_gwPOLines As GridEXWrapper
Attribute m_gwPOLines.VB_VarHelpID = -1
Private WithEvents m_gwSOLines As GridEXWrapper
Attribute m_gwSOLines.VB_VarHelpID = -1

Private m_lCustKey As Long
Private m_sItemID As String
Private m_sItemDescr As String

Private m_lUserWhseKey As Long
Private m_lWhseKey As Long

'Private m_bRefreshPack As Boolean

Private m_lRMAKey As Long
Private m_lOPKey As Long
Private m_lVendKey As Long

Private m_bDataEntry As Boolean

Private m_bAssign As Boolean

Private m_lp As LabelPrinter

'PRN#148
Private WithEvents m_gw As GridEXWrapper
Attribute m_gw.VB_VarHelpID = -1
Private m_rstBin As ADODB.Recordset
Private m_lItemKey As Long


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



'*******************************************************************
'Std form events
'*******************************************************************

Private Sub Form_Load()
    Initialize
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
    TryToSetFocus txtNewRMA
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '3/31/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwToPack = Nothing
    Set m_gwRMALine = Nothing
    Set m_gwPOs = Nothing
    Set m_gwPOLines = Nothing
    Set m_gwSOLines = Nothing
    Set m_gw = Nothing
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.width < FormWidth Then Me.width = FormWidth
    If Me.Height < FormHeight Then Me.Height = FormHeight
End Sub


Public Sub DoShowHelp()
    ShowHelp "Warehouse"
End Sub


Private Sub Initialize()
    On Error GoTo Initialize_EH
    
    SetWaitCursor True

    SetCaption "Warehouse Tool"
    
    m_lUserWhseKey = GetUserWhseKey(GetUserKey(GetUserName))

    '***added filter 8/26/03 LR
    g_rstWhses.Filter = "transit = 0"
    LoadCombo cboWhse(0), g_rstWhses, "WhseID", "WhseKey", m_lUserWhseKey
    LoadCombo cboWhse(1), g_rstWhses, "WhseID", "WhseKey", m_lUserWhseKey
    g_rstWhses.Filter = adFilterNone
    
    m_lWhseKey = cboWhse(1).ItemData(cboWhse(1).ListIndex)  'Ready To Pack tab

    SetupRMAGrid

    Set m_gwToPack = New GridEXWrapper
    m_gwToPack.Grid = gdxToPack

    Set m_gwPOs = New GridEXWrapper
    m_gwPOs.Grid = gdxPOs

    Set m_gwPOLines = New GridEXWrapper
    m_gwPOLines.Grid = gdxPOLines
    
    Set m_gwSOLines = New GridEXWrapper
    m_gwSOLines.Grid = gdxSOLines

    cmdRefresh.Enabled = True
    cboWhse(0).Enabled = True
    cboWhse(1).Enabled = True
    cmdReceive.Visible = True

    TryToSetFocus txtNewRMA
    
    Set m_lp = New LabelPrinter
    
    'PRN#148
    Set m_gw = New GridEXWrapper
    m_gw.Grid = gdxBin
        
    lblBinManager.caption = "Manage Bins for " & GetWhseDescriptionFromWhseKey(GetUserWhseKey)
    SSActiveTabs1.Tabs(tmiBinManager).Visible = HasRight(k_sRightShowToolBins)

    Me.width = FormWidth
    Me.Height = FormHeight
    
    SetWaitCursor False
    Exit Sub
    
Initialize_EH:
    ErrorUI.FatalError "WarehouseTool.Init", _
                                  "Warehouse Tool initialization failed."
    ClearWaitCursor

End Sub


Private Sub SetupRMAGrid()
    Dim orst As ADODB.Recordset

    'setup the Disposition dropdown list in the RMA grid
    Set orst = LoadDiscRst("SELECT * FROM tcpRMADisposition")
    Dim colTemp As JSColumn
    Dim vl As JSValueList
    Set colTemp = gdxRMAEdit.Columns("Disposition")
    colTemp.HasValueList = True
    Set vl = colTemp.ValueList
    vl.Add 0, "- Select One -"
    Do While Not orst.EOF
        vl.Add orst.Fields("RMADispKey").value, orst.Fields("RMADispID").value
        orst.MoveNext
    Loop
    colTemp.EditType = jgexEditDropDown

    'setup the ReasonID collection for the RMA grid
'    Set orst = LoadDiscRst("SELECT * FROM tcpRMAReason WHERE Deprecate = 0 Order By RMAReasonID")
    Set orst = LoadDiscRst("SELECT * FROM tcpRMAReason Order By RMAReasonID")
    Set m_colReason = New Collection
    Do While Not orst.EOF
        m_colReason.Add orst.Fields("RMAReasonID").value, CStr(orst.Fields("RMAReasonKey").value)
        orst.MoveNext
    Loop

    Set m_gwRMALine = New GridEXWrapper
    m_gwRMALine.Grid = gdxRMA
    Set m_oRMAList = New RMAList
    RefreshRMAViewGrid

End Sub



Private Sub SSActiveTabs1_TabClick(ByVal NewTab As ActiveTabs.SSTab)
    If NewTab.Index = tmiRMA Then
          TryToSetFocus txtNewRMA
    ElseIf NewTab.Index = tmiReceiving Then
        cmdLookup.Default = True
        txtPONum.SetFocus
        m_bDataEntry = True
    End If
End Sub


Private Sub cmdRefresh_Click()
    RefreshPackList
End Sub


'*************************************************************************************
'Ready to Pack tab
'*************************************************************************************

Private Sub cboWhse_Click(Index As Integer)
    'Ready To Pack tab
    If Index = 1 Then
        With cboWhse(1)
            m_lWhseKey = .ItemData(.ListIndex)
            If .ItemData(.ListIndex) <> m_lUserWhseKey Then
                .BackColor = RGB(255, 255, 0)
            Else
                .BackColor = RGB(255, 255, 255)
            End If
        End With
    End If
End Sub

Private Sub cmdPrintGrid_Click()
    gdxToPack.PrinterProperties.Orientation = jgexPPLandscape
    gdxToPack.PrintGrid True
End Sub


Private Sub RefreshPackList()
    Dim sSQL As String
    Dim sWillCall As String
    
    SetWaitCursor True
'    m_bRefreshPack = True
    
    If chkOmitWillCall.value = vbChecked Then
        sWillCall = "%Will Call%"
    Else
        sWillCall = ""
    End If
    
    If chkOmitGsk.value = vbChecked Then
        Set m_orstToPack = CallSP("spCPCGetPackSlips3", "@_iWhseKey", cboWhse(1).ItemData(cboWhse(1).ListIndex), "@_iWillCall", sWillCall)
    Else
        Set m_orstToPack = CallSP("spCPCGetPackSlips2", "@_iWhseKey", cboWhse(1).ItemData(cboWhse(1).ListIndex), "@_iWillCall", sWillCall)
    End If
        
    With gdxToPack
        .HoldFields
        .SortKeys.Clear
        .SortKeys.Add 3, jgexSortDescending
        .HoldSortSettings = True
        Set .ADORecordset = m_orstToPack
    End With
    
    If gdxToPack.RowCount > 0 Then
        GroupFormatgsk
    End If
    
    cmdPrintGrid.Enabled = (m_orstToPack.RecordCount > 0)

'    m_bRefreshPack = False
    SetWaitCursor False
End Sub


Private Sub GroupFormatgsk()
    Dim fmtcon As JSFmtCondition
    Dim col As JSColumn
    
    Set col = gdxToPack.Columns("Gsk")
    Set fmtcon = gdxToPack.FmtConditions.Add(col.Index, jgexEqual, -1)
    fmtcon.FormatStyle.BackColor = vbYellow
End Sub


'*************************************************************************************
' RMA tab
'*************************************************************************************

Private Sub cmdRMAEdit_Click()
    SetWaitCursor True
    Set m_oRMAList = Billing.LoadOpenRMAByRMANumber(gdxRMA.value(1), True)
    frmEdit.Visible = True
    frmView.Visible = False
    refreshRMAEditGrid
    UpdateRMAMemo
    SetWaitCursor False
End Sub


Private Sub m_gwRMALine_RowChosen()
    If gdxRMA.value(1) = Empty Or IsNull(gdxRMA.value(1)) Then
        Exit Sub
    Else
        SetWaitCursor True
        Set m_oRMAList = Billing.LoadOpenRMAByRMANumber(gdxRMA.value(1), True)
        frmEdit.Visible = True
        frmView.Visible = False
        refreshRMAEditGrid
        UpdateRMAMemo
        SetWaitCursor False
    End If
End Sub


Private Sub RMASearchViaCustorPart()
    Dim sCustText As String
    Dim sPartText As String
    
    sCustText = Trim(txtNewCust.text)
    sPartText = Trim(txtNewItem.text)
    
    If Len(sCustText) = 0 And Len(sPartText) = 0 Then Exit Sub
    
    SetWaitCursor True
    
    If Len(sPartText) = 0 Then
        RMASearchViaCust
    ElseIf Len(sCustText) = 0 Then
        RMASearchViaPart
    Else
        RMASearchViaCustandPart
    End If
    
    SetWaitCursor False
    
End Sub


Private Sub RMASearchViaCustandPart()
    Dim sCustText As String
    Dim sPartText As String
    
    
    sCustText = Trim(txtNewCust.text)
    sPartText = Trim(txtNewItem.text)
    
    Set m_oRMAList = Billing.LoadOpenRMA(, sCustText, sPartText)
    
    
    If m_oRMAList.Count = 0 Then
        msg "No records satisfy this request"
    End If
    
    RefreshAfterSearch m_oRMAList.Count, txtNewCust
End Sub


'Recoded this to use the std CustSearch function in OA. See below.  5/20/03 LR

'Private Sub RMASearchViaCust()
'    Dim sCustText As String
'    Dim oFrm As frmSearchCust
'    Dim lCount As Long
'
'    sCustText = Trim(txtNewCust.Text)
'
'    Set oFrm = New frmSearchCust
'    oFrm.Caption = "Searching Customer For '" & sCustText & "'"
'    m_lCustKey = oFrm.Find(sCustText)
'
'    Set m_oRMAList = New RMAList
'    If m_lCustKey > 0 Then
'        lCount = m_oRMAList.LoadOpenRMA(m_lCustKey)
'    End If
'
'    RefreshAfterSearch lCount, txtNewCust
'End Sub


Private Sub RMASearchViaCust()
    Dim sCustText As String
    Dim oCustomer As Customer
    
    sCustText = Trim(txtNewCust.text)
    Set oCustomer = New Customer
    m_lCustKey = Search.FindCustomer(sCustText, 1, oCustomer)
    Set oCustomer = Nothing
    
    'vl do we need this setting ??
    Set m_oRMAList = New RMAList
    
    If m_lCustKey > 0 Then
        Set m_oRMAList = Billing.LoadOpenRMA(m_lCustKey)
    End If
    
    RefreshAfterSearch m_oRMAList.Count, txtNewCust
End Sub


Private Sub RefreshAfterSearch(ByVal lCount As Long, oCtr As Control)
    If lCount < 1 Then
        ResetRMACtrls oCtr
        frmView.Visible = False
        frmEdit.Visible = False
        Set m_oRMAList = Nothing
    ElseIf lCount = 1 Then
        frmEdit.Visible = True
        frmView.Visible = False
        refreshRMAEditGrid
        UpdateRMAMemo
    Else
        frmView.Visible = True
        frmEdit.Visible = False
        RefreshRMAViewGrid
    End If
End Sub


Private Sub RMASearchViaPart()
    Dim sPartText As String
    Dim oFrm As FRMAItemSearch
    
    sPartText = Trim(txtNewItem.text)
    
    Set oFrm = New FRMAItemSearch
    oFrm.Find "Find", sPartText, m_sItemID, m_sItemDescr
    
    Set m_oRMAList = New RMAList
    
    If m_sItemID <> "" Or m_sItemDescr <> "" Then
        Set m_oRMAList = Billing.LoadOpenRMA(, , m_sItemID, m_sItemDescr)
    End If
    
    RefreshAfterSearch m_oRMAList.Count, txtNewItem
End Sub


Private Sub RMASearchViaNumber()
    On Error GoTo ErrorHandler
    
    Dim sText As String
        
    sText = Trim(txtNewRMA.text)
    
    If sText = "" Then Exit Sub
    
    If Not IsNumeric(sText) Then
        msg "Please enter valid Order#, RMA#, or SOID for searching", vbOKOnly + vbExclamation, "Find RMA"
        ResetRMACtrls txtNewRMA
        Exit Sub
    End If
    
    SetWaitCursor True
    
    Set m_oRMAList = Billing.LoadOpenRMAByRMANumber(CLng(sText))
    
    
    If m_oRMAList.Count = 0 Then
        msg "No Open RMA Lines meeting your specification was found. Check your input and search again.", vbExclamation + vbOKOnly, "Find RMA"
    End If
    
    RefreshAfterSearch m_oRMAList.Count, txtNewRMA
    
    SetWaitCursor False
        
    Exit Sub
    
ErrorHandler:
    msg Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, Err.Source
End Sub


Private Sub ResetRMACtrls(oCtrl As Control)
    On Error Resume Next
    
    oCtrl.SelStart = 0
    oCtrl.SelLength = Len(oCtrl.text)
    oCtrl.SetFocus
End Sub


Private Sub txtNewCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtNewCust.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            cmdFind_Click (1)
        End If
    End If
End Sub


Private Sub txtNewCust_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub


Private Sub txtNewItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtNewItem.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            cmdFind_Click (1)
        End If
    End If
End Sub


Private Sub txtNewItem_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub


Private Sub txtNewRMA_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtNewRMA.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            RMASearchViaNumber
        End If
    End If
End Sub


Private Sub RefreshReturns()
    Dim i_sText As String
    Dim lCustKey As Long
    Dim lCount As Long
    
    Set m_oRMAList = Billing.LoadOpenRMAByRMANumber(m_lRMAKey, True)
   
    
    If m_oRMAList.Count < 1 Then
        frmEdit.Visible = False
        Set m_oRMAList = Nothing
    Else
        refreshRMAEditGrid
        UpdateRMAMemo
    End If
End Sub


'TODO: enable when the grid is 'dirty'

Private Sub cmdReceive_Click()
    Dim lIndex As Long
    Dim lStatus As Long
    Dim lUpdateCnt As Long
    Dim cmd As ADODB.Command

    If m_oRMAList.Count < 1 Then Exit Sub
    
    'Use the current login NT user as default receive user  01/17/02 TeddyX
    
    'Precheck:
    'Was a UserID selected?
    'scan all items in the list
    'if none have received quantities, report this
    'if any are received without a disposition, report this
'    If cboUserID.ListIndex = 0 Then
'        Msg "Select a UserID.", vbInformation, "Receive Items"
'        Exit Sub
'    End If

    With m_oRMAList
        For lIndex = 1 To .Count
            If .Item(lIndex).QtyRcvd > 0 Then
                lUpdateCnt = lUpdateCnt + 1
                If .Item(lIndex).Disposition = 0 Then
                    SetWaitCursor False
                    msg "Specify a disposition for each return.", vbExclamation + vbOKOnly, "Warehouse Tool"
                    Exit Sub
                'smr 01/21/2005 - elseif disposition is return to stock and part# is S-P
                ElseIf .Item(lIndex).Disposition = 2 And .Item(lIndex).ItemID = "S-P" Then
                    SetWaitCursor False
                    msg "Please change the disposition for PartNbr 'S-P'.", vbExclamation + vbOKOnly, "Warehouse Tool"
                    Exit Sub
                End If
            End If
        Next
        If lUpdateCnt = 0 Then
            SetWaitCursor False
            msg "Specify the quantity received for returned items.", vbExclamation + vbOKOnly, "Warehouse Tool"
            Exit Sub
        End If
  
        'NOTE: this is assuming only one RMA is getting loaded
        'This will need to change to support "All RMAs"
        m_lRMAKey = .Item(1).RMAKey
        m_lOPKey = .Item(1).OPKey
    End With
    
'    If vbNo = Msg("Are you sure you want to save your changes?", vbExclamation + vbYesNo, "Receive RMA items") Then
'        Exit Sub
'    End If
    
    SetWaitCursor True
    
    With m_oRMAList
        For lIndex = 1 To .Count
            If .Item(lIndex).QtyRcvd > 0 Then
                'PRN#96
                Set cmd = CreateCommandSP("spcpcRMAReceiveItem")
                cmd.Parameters("@_iRMAKey").value = .Item(lIndex).RMAKey
                cmd.Parameters("@_iSOLineKey").value = .Item(lIndex).OPLineKey
                cmd.Parameters("@_iRMALineKey").value = .Item(lIndex).RmaLineKey
                cmd.Parameters("@_iQtyAuth").value = .Item(lIndex).QtyRcvd
                cmd.Parameters("@_iReason").value = .Item(lIndex).Reason
                cmd.Parameters("@_iRestock").value = .Item(lIndex).Restock
                cmd.Parameters("@_iDisposition").value = .Item(lIndex).Disposition
                cmd.Parameters("@_iQtyRcvd").value = .Item(lIndex).QtyRcvd
                cmd.Parameters("@_iUserID").value = GetUserID
                cmd.Parameters("@_iCreditFreight").value = .Item(lIndex).CreditFreight
                cmd.Execute
                lStatus = cmd.Parameters("@_oStatusCode").value
                If lStatus < 0 Then
                    msg "Enter received item for RMA " & .Item(lIndex).RMAKey & " failed.", _
                        vbExclamation + vbOKOnly, "Entering Failed"
                End If
                Set cmd = Nothing
            End If
        Next
    End With
    
    RefreshReturns
    SetWaitCursor False
End Sub


'The set MM remark control state
Private Sub UpdateRMAMemo()
    If gdxRMAEdit.value(3) = Empty Or IsNull(gdxRMAEdit.value(3)) Then
        rvRMA.Visible = False
        Exit Sub
    Else
        rvRMA.Visible = True
        rvRMA.OwnerID = ""
        rvRMA.OwnerID = gdxRMAEdit.value(3)
    End If
End Sub


Private Sub refreshRMAEditGrid()
    Dim i As Integer
    
    With m_oRMAList.Item(1)
        lblRMA.caption = .RMAKey
        lblOP.caption = .OPKey
        lblCustomer.caption = .CustID & " " & .CustName
        lblSO.caption = .SOID
    End With
      
    With gdxRMAEdit
        .HoldFields
        .HoldSortSettings = True
        .ItemCount = m_oRMAList.Count
        .Refetch
        .Row = 1
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
End Sub


Private Sub RefreshRMAViewGrid()
    Dim i As Integer
    With gdxRMA
        .HoldFields
        .HoldSortSettings = True
        .ItemCount = m_oRMAList.Count
        .Refetch
        .Row = 1
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
End Sub


'************************************************************************
' gdxRMAEdit grid events
'************************************************************************

Private Sub gdxRMAEdit_Click()
    UpdateRMAMemo
End Sub


Private Sub gdxRMAEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        UpdateRMAMemo
    End If
End Sub


Private Sub gdxRMAEdit_LostFocus()
    gdxRMAEdit.Update
End Sub


Private Sub gdxRMAEdit_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    gdxRMAEdit.Update
End Sub


Private Sub gdxRMAEdit_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_oRMAList Is Nothing Then Exit Sub
    If RowIndex > m_oRMAList.Count Then Exit Sub
    
    With m_oRMAList.Item(RowIndex)
        Values(1) = .RMAKey
        Values(2) = .SOLineKey
        Values(3) = .OPKey
        Values(5) = .QtyRcvd
        Values(6) = .ItemID
        Values(7) = .Descr
        Values(8) = .QtyAuthorized
        Values(9) = .AuthBy
        Values(10) = .AuthDate
        Values(11) = .QtyPreRcvd
        Values(12) = .QtyPreCred
        Values(13) = .OPLineKey
        Values(14) = .Disposition
        Values(15) = m_colReason(CStr(.Reason))
        Values(16) = .CreditFreight
    End With
End Sub


'************************************************************************
' gdxRMA grid events
'************************************************************************

Private Sub gdxRMA_Click()
    If gdxRMA.value(1) = Empty Or IsNull(gdxRMA.value(1)) Then
        cmdRMAEdit.Enabled = False
    Else
        cmdRMAEdit.Enabled = True
    End If
End Sub


Private Sub gdxRMA_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If gdxRMA.value(1) = Empty Or IsNull(gdxRMA.value(1)) Then
            cmdRMAEdit.Enabled = False
        Else
            cmdRMAEdit.Enabled = True
        End If
    End If
End Sub


Private Sub gdxRMA_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If m_oRMAList Is Nothing Then Exit Sub
    If RowIndex > m_oRMAList.Count Then Exit Sub
    
    With m_oRMAList.Item(RowIndex)
        Values(1) = .RMAKey
        Values(2) = .SOLineKey
        Values(3) = .OPKey
        Values(5) = .QtyRcvd
        Values(6) = .ItemID
        Values(7) = .Descr
        Values(8) = .QtyAuthorized
        Values(9) = .AuthBy
        Values(10) = .AuthDate
        Values(11) = .QtyPreRcvd
        Values(12) = .QtyPreCred
        Values(13) = .RMAInfo
        Values(14) = .OPLineKey
        Values(15) = .Disposition
        Values(16) = m_colReason(CStr(.Reason))
    End With
End Sub


Private Sub gdxRMAEdit_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    m_oRMAList.Item(RowIndex).Disposition = Values(14)

    If Not IsNumeric(Values(5)) Then Exit Sub
    
    If CLng(Values(5)) > (CLng(Values(8)) - CLng(Values(11))) Then
        msg "Sorry. QtyRcvd must be less than or equal to " & (CLng(Values(8)) - CLng(Values(11))) _
            & vbCrLf & "Please re-enter.", vbExclamation + vbOKOnly, "Warehouse Tool"
    Else
        m_oRMAList.Item(RowIndex).QtyRcvd = Values(5)
    End If
    
End Sub


Private Function IsRMA(ByVal i_sText As String) As Boolean
    If Left(i_sText, 1) = "R" Then 'we can assume our caller is forcing uppercase
        If Len(i_sText) = 1 Then
            IsRMA = True
        ElseIf IsNumeric(Mid(i_sText, 2)) Then
            IsRMA = True
        End If
    End If
End Function


Private Function IsOPID(ByVal i_sText As String) As Boolean
    If Left(i_sText, 1) = "O" Then 'we can assume our caller is forcing uppercase
        If Len(i_sText) = 1 Then
            IsOPID = True
        ElseIf IsNumeric(Mid(i_sText, 2)) Then
            IsOPID = True
        End If
    End If
End Function


Private Function IsNumeric(ByVal i_sText As String) As Boolean
    Dim i As Long
    Dim chr As Long

    If Len(i_sText) = 0 Then Exit Function
    
    For i = 1 To Len(i_sText)
        chr = Asc(Mid(i_sText, i, 1))
        If chr < Asc("0") Or chr > Asc("9") Then
            IsNumeric = False
            Exit Function
        End If
    Next
    IsNumeric = True
End Function


'*************************************************************************************
' Receiving tab (3)
'*************************************************************************************

Private Sub txtPONum_Change()
    If m_bAssign Then Exit Sub
    m_bAssign = True
    txtFindVendor = vbNullString
    m_bAssign = False
    TerminatePOList
End Sub


'if the user types something in the Vendor Name box, clear the PO# box
Private Sub txtFindVendor_Change()
    If m_bAssign Then Exit Sub
    m_bAssign = True
    txtPONum.text = vbNullString
    m_bAssign = False
    TerminatePOList
End Sub


Private Sub txtFindVendor_GotFocus()
    SelectText txtFindVendor
End Sub

Private Sub txtPartNbr_Change()
    cmdLookup.Default = False
    cmdFilterPartNbr.Default = True
End Sub


Private Sub TerminatePOList()
    cmdLookup.Default = True
    cmdFilterPartNbr.Default = False
    frmPOList.Visible = False
    txtPartNbr.text = ""
End Sub


Private Sub SelectText(ByRef i_oTxtBox As TextBox)
    i_oTxtBox.SelStart = 0
    i_oTxtBox.SelLength = Len(i_oTxtBox.text)
End Sub

'The user has performed a vendor search and gotten back a list of POs (gdxPOs/spcpcPOGetOpenVendPO)
'Now the user wants to filter the list to those containing a specific PartNo.

Private Sub cmdFilterPartNbr_Click()
    Dim orst As ADODB.Recordset

    If Len(txtPartNbr.text) = 0 Then Exit Sub
    
    Set orst = CallSP("spcpcPOGetOpenVendPObyPartNo", "@_iVendKey", m_lVendKey, "@_iItemID", Trim(txtPartNbr.text))
    If orst.EOF Then
        msg "None of these POs contain that PartNo."
    Else
        AttachGrid gdxPOs, orst
        gdxPOs.SetFocus
    End If

End Sub


Private Sub cmdLookUp_Click()
    Dim sPONum As String
    Dim orst As ADODB.Recordset
    Dim oFrm As FSearch
    Dim sVendName As String

'doing an exact PO# search
'no partial PO#s
'no wildcards

    'clear the form
    lblWarning.caption = vbNullString
    txtVendName.text = vbNullString
    txtBuyer.text = vbNullString
    txtPODate.text = vbNullString
   
    sPONum = Trim(txtPONum.text)
    'if PO# was specified
    If Len(sPONum) > 0 Then
    
        LoadPODetail sPONum
        
    Else
        'if a vendor was specified
        If Len(txtFindVendor.text) > 0 Then
            Set oFrm = New FSearch
            m_lVendKey = oFrm.Find(txtFindVendor.text, sVendName)
            Set oFrm = Nothing
            If m_lVendKey > 0 Then
                lblVendName = sVendName
                'find all open POs for selected vendor
                Set orst = CallSP("spcpcPOGetOpenVendPO", "@_iVendKey", m_lVendKey)
                Select Case orst.RecordCount
                    Case 0
                        msg "We have no open POs for this vendor.", vbInformation
                        
                    Case 1
                        txtFindVendor.text = ""
                        LoadPODetail orst.Fields("tranno").value
                        
                    Case Else
                        AttachGrid gdxPOs, orst
                        
                        'update the display
                        txtFindVendor.text = ""
                        frmPOList.Visible = True
                        gdxPOs.SetFocus
                        frmPOHdr.Visible = False
                        gdxSOLines.Visible = False
                        
                End Select
            End If
        End If

    End If

End Sub


Private Sub LoadPODetail(ByVal sPONum As String)
    Dim orst As ADODB.Recordset

    Set orst = CallSP("spcpcGetPOInfo", "@_iPONum", sPONum)
    
    If orst.EOF Then
        'update the display
        frmPOList.Visible = False
        frmPOHdr.Visible = False
        gdxSOLines.Visible = False
    
        msg "This PO does not exist."
        SelectText txtPONum
        
        Exit Sub
    End If
    
    If orst.Fields("DfltDropShip").value = 1 Then
        lblWarning.caption = "Warning: This PO is a DropShip."
    End If
    
    If orst.Fields("WhseKey").value <> m_lUserWhseKey Then
        lblWarning.caption = "Warning: This package belongs in " & orst.Fields("Description")
    End If

    LoadPO sPONum

    'update the display
    frmPOList.Visible = False
    txtPONum.text = ""
    txtFindVendor.text = ""
    lblVendName.caption = ""
    frmPOHdr.Visible = True
    gdxPOLines.SetFocus
    gdxSOLines.Visible = True

End Sub


Private Sub LoadPO(ByVal i_sPONum As String)
    Dim orstHeader As ADODB.Recordset
    Dim oRstLines As ADODB.Recordset
    Dim lPOKey As Long

    Set orstHeader = CallSP("spcpcGetPOHeader", "@_iPONum", i_sPONum)
    With orstHeader
        txtPONbr.text = StripLeadingZeros(.Fields("TranNo").value)
        txtVendName.text = .Fields("VendName").value
        txtBuyer.text = .Fields("BuyerID").value
        txtPODate.text = FormatDateTime(.Fields("TranDate").value, vbShortDate)
        lPOKey = .Fields("POKey").value
    End With

    rvPO.OwnerID = i_sPONum
    
'   po line info
    Set oRstLines = CallSP("spcpcGetPOLines2", "@_iPOKey", lPOKey, "@_iWhseKey", m_lUserWhseKey)

    AttachGrid gdxPOLines, oRstLines
    
    If oRstLines.RecordCount = 0 Then
        AttachGrid gdxSOLines, Nothing
    End If
End Sub


Private Sub gdxPOLines_SelectionChange()
    Dim sSQL As String
    Dim orst As ADODB.Recordset
    Dim s As String
    
    Dim OPKey As Long
    

    If Len(Trim(m_gwPOLines.value("ItemID"))) > 0 Then
    
        'if the item is a SPO or a SHF, it's been frozen to its SO
        If InStr(1, m_gwPOLines.value("ItemID"), "SPO-") Or InStr(1, m_gwPOLines.value("ItemID"), "SHF-") Then
            Set orst = CallSP("spCPCWASalesOrderforSPOPO", "@_iPOLineKey", m_gwPOLines.value("POLineKey"))
            'Why would it be empty?
            If Not orst.EOF Then
                OPKey = CLng(orst.Fields("OPKey").value)
                s = "OP# " & OPKey & "  " & Trim(orst.Fields("ShipMethDesc").value)
                If orst.Fields("ShipComplete").value <> 0 Then s = s & "  S/C"
                m_lp.Clear
                m_lp.AddLine s
                m_lp.AddLine "PO# " & txtPONbr.text
                m_lp.AddLine m_gwPOLines.value("ItemID")
                m_lp.AddLine m_gwPOLines.value("Description")
                m_lp.AddLine "Received: " & Format(Now, "ddd mmm dd, yyyy")
            End If
            m_lp.NumLabels = m_gwPOLines.value("QtyOrd")
            
            LogDB.LogActivity "SA", "PO " & txtPONbr.text & " received. [" & m_gwPOLines.value("ItemID") & "] " & m_gwPOLines.value("Description"), _
                OPKey, , orst.Fields("TranNo").value, , , txtPONbr.text, m_gwPOLines.value("POLineKey")
                
        Else
            Set orst = CallSP("spCPCWASalesOrderforStockPO", "@_iItemID", Trim(m_gwPOLines.value("ItemID")))
                m_lp.Clear
                m_lp.AddLine m_gwPOLines.value("ItemID")
                m_lp.AddLine m_gwPOLines.value("Description")
                m_lp.AddLine Format(m_gwPOLines.value("WhseBinID"))
                m_lp.AddLine "Received: " & Format(Now, "ddd mmm dd, yyyy")
                m_lp.NumLabels = m_gwPOLines.value("QtyOrd")
        End If
        AttachGrid gdxSOLines, orst
        Set orst = Nothing
    End If
End Sub


Private Sub cmdPrintLabels_Click(Index As Integer)
    On Error GoTo EH

    Select Case Index
        Case 0:     'Receiving
            m_lp.PrintLabel
            
        Case 1:     'RMA
        'PRN 278 1/29/04 LR: if it's a stock part, find the bin location
            Dim Location As String
            Dim ocmd As ADODB.Command
            Set ocmd = CreateCommandSP("spcpcIMGetPartLocation")
            With ocmd
                .Parameters("@_iItemID").value = gdxRMAEdit.value(6)        'Part#
                .Parameters("@_iWhseKey").value = m_lWhseKey
                .Execute
                Location = IIf(IsNull(.Parameters("@_oLocation").value), vbNullString, .Parameters("@_oLocation").value)
            End With
            Set ocmd = Nothing
            
            m_lp.Clear
            m_lp.AddLine gdxRMAEdit.value(6)        'Part#
            m_lp.AddLine gdxRMAEdit.value(7)        'Descr
            m_lp.AddLine Location
            m_lp.AddLine "Received: " & Format(Now, "ddd mmm dd, yyyy")
            m_lp.NumLabels = gdxRMAEdit.value(5)    'QtyRcvd
            m_lp.PrintLabel
    End Select
    Exit Sub
EH:
    msg Err.Description, vbCritical, Err.Source
End Sub


Private Sub m_gwPOs_RowChosen()
    LoadPODetail m_gwPOs.value(1)
End Sub


'************************************************************************************
'   Tab(4) Unexploded Kits
'   This should be enabled for Purchasing Group only
'************************************************************************************

'Control Array
Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
    Case 0:     'Find RMA
            RMASearchViaNumber
    
    Case 1:     'Find RMA
            RMASearchViaCustorPart
    
    Case 2:     'Kit explode
        Dim orst As ADODB.Recordset
        Set orst = CallSP("spcpcGetUnexplodedKits", "@_iWhseKey", cboWhse(0).ItemData(cboWhse(0).ListIndex))
        AttachGrid gdxUnexplodedKit, orst
    End Select
End Sub



'************************************************************************************
'   Private Functions
'************************************************************************************

'This function returns the string result of each value of LineStatus
Private Function GetLineStatus(oValue As LineStatus) As String
    Select Case oValue
        Case IsInvoiced:
            GetLineStatus = "Invoiced"
        Case IsShipComplete:
            GetLineStatus = "Shipped"
        Case IsShipBackorders:
            GetLineStatus = "Shipped with Backorders"
        Case IsReadyToShip:
            GetLineStatus = "Ready to be Shipped"
        Case IsOnOrder:
            GetLineStatus = "On Order"
        Case IsNeedsToBeOrder:
            GetLineStatus = "Needs to Be Ordered"
        Case IsDropShipInActive:
            GetLineStatus = "DropShip In Active"
        Case IsDropShipCancelled:
            GetLineStatus = "DropShip Cancelled"
        Case IsDropShipClosed:
            GetLineStatus = "DropShip Closed"
        Case LineStatus.IsGskNew:
            GetLineStatus = "Not Yet Started"
        Case LineStatus.IsGskBegin:
            GetLineStatus = "Being Cut"
        Case LineStatus.IsGskCut:
            GetLineStatus = "Being Molded"
        Case LineStatus.IsGskMold:
            GetLineStatus = "Being Trimmed"
        Case LineStatus.IsGskTrim:
            GetLineStatus = "Complete"
        Case LineStatus.IsGskNotAvail:
            GetLineStatus = "Gsk Status Not Avail"
    End Select
End Function


' Bin Manager

'PRN#148
Private Sub cmdNew_Click()
    Dim oFrm As FChooseBin
    Dim lNewBin As Long
    Dim sSQL As String
    
    If m_rstBin Is Nothing Then Exit Sub
    
    Set oFrm = New FChooseBin
    lNewBin = oFrm.LoadNewBin(m_rstBin)
    Set oFrm = Nothing
    
    If lNewBin > 0 Then
        InsertNewBin lNewBin
        Set m_rstBin = Nothing
        SearchBins
    End If
    With gdxBin
        .Row = .ItemCount
        TryToSetFocus gdxBin
    End With
End Sub


Private Sub InsertNewBin(lBin As Long)
    Dim lPrefNo As Long
    SetWaitCursor True
    If m_rstBin.RecordCount = 0 Then
        lPrefNo = 1
    Else
        m_rstBin.MoveLast
        lPrefNo = m_rstBin.Fields("PrefNo").value + 1
    End If
    
    CallSP "spCPCAddNewBin", "@WhseKey", GetUserWhseKey, "@ItemKey", m_lItemKey, "@PrefNo", lPrefNo, "@WhseBinKey", lBin
    SetWaitCursor False
End Sub


Private Sub cmdDelete_Click()
    Dim lPrefNo As Long
    SetWaitCursor True
    If gdxBin.ItemCount > 0 Then
        lPrefNo = m_gw.value("PrefNo")
        CallSP "spCPCDeleteBin", "@WhseKey", GetUserWhseKey, "@ItemKey", m_lItemKey, "@PrefNo", lPrefNo, "@WhseBinKey", m_gw.value("WhseBinKey")
      
        If lPrefNo < m_rstBin.RecordCount Then
            m_rstBin.Filter = "PrefNo > " & lPrefNo
            m_rstBin.MoveFirst
            While Not m_rstBin.EOF
                CallSP "spCPCUpdateBin", "@WhseKey", GetUserWhseKey, "@ItemKey", m_lItemKey, "@PrefNo", m_rstBin("PrefNo").value, "@WhseBinKey", m_rstBin.Fields("WhseBinKey"), "@NewPrefNo", (m_rstBin.Fields("PrefNo") - 1)
                m_rstBin.MoveNext
            Wend
        End If
        
        Set m_rstBin = Nothing
        SearchBins
    End If
    SetWaitCursor False
End Sub


Private Sub cmdMoveUp_Click()
    Dim lPrefNo As String
    Dim lWhseBinKey As String
    
    SetWaitCursor True
    If gdxBin.ItemCount > 0 Then
        If m_gw.value("PrefNo") > 1 Then
            lPrefNo = m_gw.value("PrefNo")
            lWhseBinKey = m_gw.value("WhseBinKey")
            ExecUpDown lPrefNo, lPrefNo - 1, lWhseBinKey
            Set m_rstBin = Nothing
            SearchBins
            
            With gdxBin
                .Row = lPrefNo - 1
                TryToSetFocus gdxBin
            End With
        End If
    End If
    SetWaitCursor False
End Sub


Private Sub cmdMoveDown_Click()
    Dim lPrefNo As String
    Dim lWhseBinKey As String
    SetWaitCursor True
    If gdxBin.ItemCount > 0 Then
        If m_gw.value("PrefNo") < m_rstBin.RecordCount Then
            lPrefNo = m_gw.value("PrefNo")
            lWhseBinKey = m_gw.value("WhseBinKey")
            ExecUpDown lPrefNo, lPrefNo + 1, lWhseBinKey
            Set m_rstBin = Nothing
            SearchBins
        
            With gdxBin
                .Row = lPrefNo + 1
                TryToSetFocus gdxBin
            End With
        End If
    End If
    SetWaitCursor False
End Sub


Private Sub ExecUpDown(ByVal lOldPrefNo As Long, ByVal lNewPrefNo As Long, ByVal lWhseBinKey As Long)
    CallSP "spCPCDeleteBin", "@WhseKey", GetUserWhseKey, "@ItemKey", m_lItemKey, "@PrefNo", m_gw.value("PrefNo"), "@WhseBinKey", m_gw.value("WhseBinKey")
    
    m_rstBin.Filter = "PrefNo = " & lNewPrefNo
    CallSP "spCPCUpdateBin", "@WhseKey", GetUserWhseKey, "@ItemKey", m_lItemKey, "@PrefNo", lNewPrefNo, "@WhseBinKey", m_rstBin.Fields("WhseBinKey"), "@NewPrefNo", lOldPrefNo
    CallSP "spCPCAddNewBin", "@WhseKey", GetUserWhseKey, "@ItemKey", m_lItemKey, "@PrefNo", lNewPrefNo, "@WhseBinKey", lWhseBinKey

End Sub


Private Sub cmdFindBin_Click()
    On Error GoTo ErrorHandler
    Dim eItemType As ItemTypeCode
    Dim sOriginalItemID As String
    Dim sRefSource As String
    Dim frmItemSearch As FItemSearch
    Dim rstItemID As ADODB.Recordset
    Dim ItemID As String
    
    Dim sSQL As String
    Dim itemIDSQL As String
    'if the field is blank, no look-up is needed
    If Len(txtItemID.text) = 0 Then
        Exit Sub 'no error on empty field
    End If

    Set frmItemSearch = New FItemSearch
    Load frmItemSearch
    Dim bCancelSearch As Boolean
    frmItemSearch.Find k_sPartNbr, txtItemID.text, m_lItemKey, eItemType, sOriginalItemID, sRefSource, bCancelSearch, GetUserWhseKey, True, 1 'CustType :: 1=End User, default to End User to get list price
    If bCancelSearch Then Exit Sub

    If m_lItemKey > 0 Then
        itemIDSQL = "select PartNbr from vwOPItemSearch where ItemKey = " & m_lItemKey
        Set rstItemID = LoadDiscRst(itemIDSQL)
        If rstItemID.RecordCount > 0 Then
            txtItemID.text = Trim(rstItemID.Fields("PartNbr").value)
        End If
        
        sSQL = "select timInvtBinList.PrefNo, timWhseBin.WhseBinID, timWhseBin.WhseBinKey, timInvtBinList.ItemKey " _
                & "from timInventory inner join timInvtBinList on " _
                & "timInventory.ItemKey = timInvtBinList.ItemKey and " _
                & "timInventory.WhseKey = timInvtBinList.WhseKey " _
                & "inner join timWhseBin on " _
                & "timInvtBinList.WhseBinKey = timWhseBin.WhseBinKey " _
                & "where timInvtBinList.ItemKey = " & m_lItemKey _
                & " and timInvtBinList.WhseKey = " & GetUserWhseKey
        Set m_rstBin = LoadDiscRst(sSQL)
        
        If m_rstBin.RecordCount = 0 Then
            CheckInventory
            Exit Sub
        End If
        TryToSetFocus gdxBin
    Else
        TryToSetFocus txtItemID
        txtItemID.SelStart = 0
        txtItemID.SelLength = Len(txtItemID.text)
        Set m_rstBin = Nothing
    End If
    UpdateGrid
    Exit Sub
    
ErrorHandler:
    ClearWaitCursor
    msg Err.Description, vbOKOnly + vbExclamation, Err.Source
End Sub


Private Sub SearchBins()
    Dim sSQL As String
    
    sSQL = "select timInvtBinList.PrefNo, timWhseBin.WhseBinID, timWhseBin.WhseBinKey, timInvtBinList.ItemKey " _
            & "from timInventory inner join timInvtBinList on " _
            & "timInventory.ItemKey = timInvtBinList.ItemKey and " _
            & "timInventory.WhseKey = timInvtBinList.WhseKey " _
            & "inner join timWhseBin on " _
            & "timInvtBinList.WhseBinKey = timWhseBin.WhseBinKey " _
            & "where timInvtBinList.ItemKey = " & m_lItemKey _
            & " and timInvtBinList.WhseKey = " & GetUserWhseKey
    Set m_rstBin = LoadDiscRst(sSQL)
    
    TryToSetFocus gdxBin
    UpdateGrid
End Sub


Private Sub UpdateGrid()
    With gdxBin
        .HoldFields
        Set .ADORecordset = m_rstBin
        .Row = 1
    End With
End Sub


Private Sub CheckInventory()
    Dim sSQL As String
    Dim rstInventory As ADODB.Recordset
    
    sSQL = "select * from timInventory where " _
            & "ItemKey = " & m_lItemKey _
            & "and WhseKey = " & GetUserWhseKey
    Set rstInventory = LoadDiscRst(sSQL)
            
    If rstInventory.RecordCount > 0 Then
        If vbYes = msg("There are no bin related with this item at present. " & vbCrLf & _
                    "Would you like to create a new bin for this item?", vbYesNo + vbExclamation, "Create New Bin") Then
            cmdNew_Click
        End If
    Else
        msg "No records satisfy this request."
        TryToSetFocus txtItemID
        txtItemID.SelStart = 0
        txtItemID.SelLength = Len(txtItemID.text)
        Set m_rstBin = Nothing
    End If
    UpdateGrid
End Sub


Private Sub txtItemID_KeyDown(KeyCode As Integer, Shift As Integer)
     If Len(txtItemID.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            cmdFindBin_Click
        End If
    End If
End Sub


Private Sub cmdNewBin_Click()
    Dim ocmd As ADODB.Command
    Dim sBinID As String
    Dim lBinKey As Long
    On Error GoTo EH
    SetWaitCursor True
    sBinID = Trim(InputBox("Please enter the new bin.", "New Bin"))
    If sBinID <> "" Then
        lBinKey = database.GetSurrogateKey("timWhseBin")
        
        ' insert item in Seattle inventory
        Set ocmd = CreateCommandSP("spCPCimInsertBin")
        With ocmd
            .Parameters("@WhseBinKey") = lBinKey
            .Parameters("@WhseBinID") = sBinID
            .Parameters("@WhseKey") = m_lUserWhseKey
            .Execute
        End With
        Set ocmd = Nothing
    End If
    SetWaitCursor False
    Exit Sub
EH:
    Set ocmd = Nothing
    SetWaitCursor False
    If Err.Number = -2147217873 Then 'Cannot insert duplicate key row in object 'timWhseBin' with unique index 'XAK1timWhseBin'.
        MsgBox "Bin '" & sBinID & "' already exists. Please make appropriate change." & vbCrLf & vbCrLf & "Technical Error Description: " & Err.Description, vbInformation
    Else
        MsgBox "Can't insert new bin due to error " & Err.Number & " '" & Err.Description & "'. Please make appropriate change.", vbInformation
    End If
End Sub




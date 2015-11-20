VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FOrderManager 
   Caption         =   "Order Manager"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleMode       =   0  'User
   ScaleWidth      =   9600.624
   Visible         =   0   'False
   Begin ActiveTabs.SSActiveTabs tabMain 
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   10927
      _Version        =   262144
      TabCount        =   2
      TagVariant      =   ""
      Tabs            =   "FOrderManager.frx":0000
      Begin ActiveTabs.SSActiveTabPanel tpCreateOrder 
         Height          =   5805
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   10239
         _Version        =   262144
         TabGuid         =   "FOrderManager.frx":0088
         Begin VB.Frame frmCustOrders 
            Height          =   5775
            Left            =   60
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   9255
            Begin VB.CommandButton cmdFilterByPart 
               Caption         =   "Filter"
               Height          =   300
               Left            =   2520
               TabIndex        =   48
               Top             =   1020
               Width           =   570
            End
            Begin VB.TextBox txtFilterByPart 
               Height          =   315
               Left            =   1380
               TabIndex        =   47
               Top             =   1013
               Width           =   1035
            End
            Begin VB.CommandButton cmdContactMgr 
               Caption         =   "Contacts"
               Height          =   315
               Left            =   2880
               TabIndex        =   44
               Top             =   5340
               Width           =   1275
            End
            Begin VB.ComboBox cboOrderStatus 
               Height          =   315
               ItemData        =   "FOrderManager.frx":00B0
               Left            =   5940
               List            =   "FOrderManager.frx":00ED
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   5340
               Width           =   1635
            End
            Begin VB.CommandButton cmdNewSearch 
               Caption         =   "New Se&arch"
               Height          =   315
               Left            =   7740
               TabIndex        =   42
               Top             =   5340
               Width           =   1215
            End
            Begin VB.CommandButton cmdNewOrder 
               Caption         =   "&New Order"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   41
               Top             =   5340
               Width           =   1215
            End
            Begin VB.CommandButton cmdLoadOrder 
               Caption         =   "Load Ord&er"
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   1500
               TabIndex        =   40
               Top             =   5340
               Width           =   1215
            End
            Begin VB.CheckBox chkShowOrdersForShipAddr 
               Caption         =   "Show orders for all shipping addresses"
               Height          =   192
               Left            =   3360
               TabIndex        =   21
               Top             =   1320
               Width           =   3168
            End
            Begin GridEX20.GridEX gdxCustOrders 
               Height          =   3615
               Left            =   120
               TabIndex        =   22
               Top             =   1620
               Width           =   9060
               _ExtentX        =   15981
               _ExtentY        =   6376
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
               Column(1)       =   "FOrderManager.frx":01D9
               Column(2)       =   "FOrderManager.frx":03B1
               Column(3)       =   "FOrderManager.frx":0D8D
               Column(4)       =   "FOrderManager.frx":0F51
               Column(5)       =   "FOrderManager.frx":10B1
               Column(6)       =   "FOrderManager.frx":1371
               Column(7)       =   "FOrderManager.frx":151D
               Column(8)       =   "FOrderManager.frx":1675
               Column(9)       =   "FOrderManager.frx":180D
               Column(10)      =   "FOrderManager.frx":1995
               Column(11)      =   "FOrderManager.frx":1AFD
               Column(12)      =   "FOrderManager.frx":1C95
               Column(13)      =   "FOrderManager.frx":1DED
               Column(14)      =   "FOrderManager.frx":1F5D
               Column(15)      =   "FOrderManager.frx":2099
               Column(16)      =   "FOrderManager.frx":220D
               Column(17)      =   "FOrderManager.frx":2331
               Column(18)      =   "FOrderManager.frx":2479
               SortKeysCount   =   1
               SortKey(1)      =   "FOrderManager.frx":2591
               FmtConditionsCount=   1
               FmtCondition(1) =   "FOrderManager.frx":25F9
               FormatStylesCount=   6
               FormatStyle(1)  =   "FOrderManager.frx":2751
               FormatStyle(2)  =   "FOrderManager.frx":2831
               FormatStyle(3)  =   "FOrderManager.frx":2969
               FormatStyle(4)  =   "FOrderManager.frx":2A19
               FormatStyle(5)  =   "FOrderManager.frx":2ACD
               FormatStyle(6)  =   "FOrderManager.frx":2BA5
               ImageCount      =   0
               PrinterProperties=   "FOrderManager.frx":2C5D
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Filter by Part #"
               Height          =   195
               Left            =   120
               TabIndex        =   46
               Top             =   1073
               Width           =   1095
            End
            Begin VB.Label lblOrderStatus 
               Caption         =   "Order Status"
               Height          =   315
               Left            =   4920
               TabIndex        =   45
               Top             =   5340
               Width           =   975
            End
            Begin VB.Label lblCustInfo 
               Caption         =   "Order         Ship Address"
               Height          =   375
               Index           =   1
               Left            =   5820
               TabIndex        =   29
               Top             =   240
               UseMnemonic     =   0   'False
               Visible         =   0   'False
               Width           =   1035
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
               Left            =   180
               TabIndex        =   28
               Top             =   660
               Width           =   2115
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
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   2115
            End
            Begin VB.Label lblOrderAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblOrderAddress"
               ForeColor       =   &H80000008&
               Height          =   960
               Left            =   6960
               TabIndex        =   26
               Top             =   240
               Visible         =   0   'False
               Width           =   2190
            End
            Begin VB.Label lblCustInfo 
               Caption         =   "Customer Ship Address"
               Height          =   495
               Index           =   0
               Left            =   2280
               TabIndex        =   25
               Top             =   240
               UseMnemonic     =   0   'False
               Width           =   975
            End
            Begin VB.Label lblCustAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblCustAddress"
               ForeColor       =   &H80000008&
               Height          =   960
               Left            =   3360
               TabIndex        =   24
               Top             =   240
               Width           =   2190
            End
            Begin VB.Label lblOrderCount 
               Caption         =   "lblOrderCount"
               Height          =   315
               Left            =   120
               TabIndex        =   23
               Top             =   1380
               Width           =   1935
            End
         End
         Begin VB.Frame frmCreateOrder 
            Height          =   5775
            Left            =   60
            TabIndex        =   30
            Top             =   0
            Width           =   9255
            Begin VB.CommandButton cmdMiscOrder 
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
               Left            =   720
               TabIndex        =   35
               Top             =   3420
               Width           =   1815
            End
            Begin VB.CommandButton cmdWalkupOrder 
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
               Left            =   720
               TabIndex        =   34
               Top             =   2520
               Width           =   1815
            End
            Begin VB.CommandButton cmdNewCustomer 
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
               Left            =   720
               TabIndex        =   33
               Top             =   1620
               Width           =   1815
            End
            Begin VB.ComboBox cboSearchType 
               Height          =   315
               ItemData        =   "FOrderManager.frx":2E35
               Left            =   5220
               List            =   "FOrderManager.frx":2E4B
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   780
               Width           =   1935
            End
            Begin VB.CommandButton cmdFindAccount 
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
               Left            =   720
               TabIndex        =   31
               Top             =   720
               Width           =   1815
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtCustSearch 
               Height          =   315
               Left            =   2880
               TabIndex        =   36
               ToolTipText     =   "Enter the first few characters of the customer name"
               Top             =   780
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
            Begin VB.Label lblExplainMisc 
               Caption         =   $"FOrderManager.frx":2E90
               Height          =   615
               Left            =   2880
               TabIndex        =   39
               Top             =   3420
               Width           =   4155
            End
            Begin VB.Label lblExplainWalkup 
               Caption         =   $"FOrderManager.frx":2F3B
               Height          =   615
               Left            =   2880
               TabIndex        =   38
               Top             =   2520
               Width           =   4155
            End
            Begin VB.Label lblExplainNew 
               Caption         =   "Use this to create a quote for a customer you intend to setup with an account before committing your order."
               Height          =   615
               Left            =   2880
               TabIndex        =   37
               Top             =   1620
               Width           =   4155
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel tpSelectOrder 
         Height          =   5805
         Left            =   30
         TabIndex        =   2
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   10239
         _Version        =   262144
         TabGuid         =   "FOrderManager.frx":2FE7
         Begin VB.Frame frmFindWithInfo 
            Caption         =   "Find Order via Related Information"
            ClipControls    =   0   'False
            Height          =   1212
            Index           =   7
            Left            =   1920
            TabIndex        =   7
            Top             =   120
            Width           =   7380
            Begin VB.CommandButton cmdClearOrder 
               Caption         =   "R&eset"
               Height          =   315
               Left            =   4740
               TabIndex        =   14
               Top             =   720
               Width           =   900
            End
            Begin VB.ComboBox cboFindStatus 
               Height          =   315
               ItemData        =   "FOrderManager.frx":300F
               Left            =   3000
               List            =   "FOrderManager.frx":304C
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   750
               Width           =   1632
            End
            Begin VB.ComboBox cboFindCSR 
               Height          =   315
               ItemData        =   "FOrderManager.frx":312F
               Left            =   3000
               List            =   "FOrderManager.frx":3131
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   300
               Width           =   1632
            End
            Begin VB.CommandButton cmdFindOrders 
               Caption         =   "Fi&nd"
               Height          =   315
               Index           =   1
               Left            =   4740
               TabIndex        =   11
               Top             =   300
               Width           =   900
            End
            Begin VB.TextBox txtFindText 
               Height          =   288
               Left            =   960
               TabIndex        =   10
               Top             =   750
               Width           =   1332
            End
            Begin VB.TextBox txtFindCust 
               Height          =   288
               Left            =   960
               TabIndex        =   9
               Top             =   300
               Width           =   1332
            End
            Begin VB.ComboBox cboTimeInterval 
               Height          =   315
               ItemData        =   "FOrderManager.frx":3133
               Left            =   5880
               List            =   "FOrderManager.frx":3151
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   300
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status"
               Height          =   252
               Index           =   79
               Left            =   2460
               TabIndex        =   18
               Top             =   750
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "CSR"
               Height          =   252
               Index           =   78
               Left            =   2400
               TabIndex        =   17
               Top             =   300
               Width           =   492
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Keywords"
               Height          =   252
               Index           =   77
               Left            =   180
               TabIndex        =   16
               Top             =   750
               Width           =   732
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer"
               Height          =   252
               Index           =   4
               Left            =   180
               TabIndex        =   15
               Top             =   300
               Width           =   732
            End
         End
         Begin VB.Frame frmFindByNum 
            Caption         =   "Find by OP/SO/RMA"
            ClipControls    =   0   'False
            Height          =   1212
            Index           =   9
            Left            =   60
            TabIndex        =   4
            Top             =   120
            Width           =   1800
            Begin VB.TextBox txtFindOrder 
               Height          =   288
               Left            =   120
               TabIndex        =   6
               Top             =   300
               Width           =   1140
            End
            Begin VB.CommandButton cmdFindOrders 
               Caption         =   "Fin&d"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   720
               Width           =   900
            End
         End
         Begin VB.CommandButton cmdLoadOrder 
            Caption         =   "Load Ord&er"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   3
            Top             =   5400
            Width           =   1215
         End
         Begin GridEX20.GridEX gdxOrders 
            Height          =   3915
            Left            =   60
            TabIndex        =   19
            Top             =   1380
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   6906
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
            Column(1)       =   "FOrderManager.frx":318D
            Column(2)       =   "FOrderManager.frx":339D
            Column(3)       =   "FOrderManager.frx":3585
            Column(4)       =   "FOrderManager.frx":36CD
            Column(5)       =   "FOrderManager.frx":40DD
            Column(6)       =   "FOrderManager.frx":42A5
            Column(7)       =   "FOrderManager.frx":442D
            Column(8)       =   "FOrderManager.frx":455D
            Column(9)       =   "FOrderManager.frx":481D
            Column(10)      =   "FOrderManager.frx":49C9
            Column(11)      =   "FOrderManager.frx":4B21
            Column(12)      =   "FOrderManager.frx":4C69
            Column(13)      =   "FOrderManager.frx":4DBD
            Column(14)      =   "FOrderManager.frx":4F11
            Column(15)      =   "FOrderManager.frx":50E5
            Column(16)      =   "FOrderManager.frx":526D
            Column(17)      =   "FOrderManager.frx":53D5
            Column(18)      =   "FOrderManager.frx":56B9
            Column(19)      =   "FOrderManager.frx":5821
            Column(20)      =   "FOrderManager.frx":5995
            Column(21)      =   "FOrderManager.frx":5AB9
            Column(22)      =   "FOrderManager.frx":5C39
            Column(23)      =   "FOrderManager.frx":5DB5
            Column(24)      =   "FOrderManager.frx":5F59
            SortKeysCount   =   1
            SortKey(1)      =   "FOrderManager.frx":6071
            FormatStylesCount=   6
            FormatStyle(1)  =   "FOrderManager.frx":60D9
            FormatStyle(2)  =   "FOrderManager.frx":61B9
            FormatStyle(3)  =   "FOrderManager.frx":62F1
            FormatStyle(4)  =   "FOrderManager.frx":63A1
            FormatStyle(5)  =   "FOrderManager.frx":6455
            FormatStyle(6)  =   "FOrderManager.frx":652D
            ImageCount      =   0
            PrinterProperties=   "FOrderManager.frx":65E5
         End
      End
   End
   Begin MSComctlLib.ImageList imglStatus16 
      Left            =   0
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
            Picture         =   "FOrderManager.frx":67BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":6BA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":6F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":735E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":7717
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":7AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":7E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":8186
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":8DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FOrderManager.frx":912A
            Key             =   ""
         EndProperty
      EndProperty
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
End
Attribute VB_Name = "FOrderManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_skSource = "FOrderManager"

Private Const k_lMinWidth = 9690
Private Const k_lMinHeight = 6840


Private Enum TabMainIndexes
    tmicreateorder = 1
    tmiSelectOrder = 2
End Enum

'!!! This is temp for testing
Private m_oOrder As Order

Private m_bLoading As Boolean

Private WithEvents m_oBrokenRules As BrokenRules
Attribute m_oBrokenRules.VB_VarHelpID = -1

'These objects wrap the GridEX controls to give a more convenient
'interface for getting the events we're interested in.
Private WithEvents m_gwCustOrders As GridEXWrapper
Attribute m_gwCustOrders.VB_VarHelpID = -1
Private WithEvents m_gwOrders As GridEXWrapper
Attribute m_gwOrders.VB_VarHelpID = -1

Private m_bFindOrderFlag As Boolean

Private m_colWCGridPrefs As Collection

Private m_lWindowID As Long

Private m_oCustomer As Customer

Private m_lCustKey As Long
Private m_lDefaultBillAddrKey As Long
Private m_lDefaultShipAddrKey As Long

Public Property Get Customer() As Customer
    Set Customer = m_oCustomer
End Property

Public Property Let Customer(oNewValue As Customer)
    Set m_oCustomer = oNewValue
End Property


Public Property Get BrokenRules() As BrokenRules
    Set BrokenRules = m_oBrokenRules
End Property


'*******************************************************************
'Extended form property & method
'*******************************************************************

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


Private Sub cmdFilterByPart_Click()
    If cmdFilterByPart.caption = "Filter" Then
        If Len(Trim$(txtFilterByPart)) > 0 Then
            cmdFilterByPart.caption = "Reset"
            txtFilterByPart.Enabled = False
        End If
    Else
        txtFilterByPart.text = ""
        txtFilterByPart.Enabled = True
        cmdFilterByPart.caption = "Filter"
    End If
    
End Sub


'*******************************************************************
'Std form events
'*******************************************************************

Private Sub Form_Load()
    
    'MDI child forms set to 2 (Sizable) are displayed within the MDI form in a default size
    'defined by the Windows operating environment at run time. For any other setting,
    'the form is displayed in the size specified at design time.
    
    Me.Height = k_lMinHeight
    Me.width = k_lMinWidth

    SetCaption "Order Manager"

    Set m_oBrokenRules = New BrokenRules
    m_oBrokenRules.Form = Me
    LoadValidationRules
    
    With m_oBrokenRules
        .EnableClass ControlClass.ccCustomer, True
        .Validate
    End With
    
    tabMain.Tabs(tmicreateorder).Selected = True
    
    LoadImageList imglStatus16, gdxCustOrders
    LoadImageList imglStatus16, gdxOrders
    
    User.LoadActiveCSRs cboFindCSR
    cboFindCSR.AddItem "<Any>"
    cboFindCSR.AddItem "<MPK>"
    cboFindCSR.AddItem "<SEA>"
    cboFindCSR.AddItem "<STL>"
    
    'support a Click event handler on cboSearchType
    m_bLoading = True
    cboSearchType.ListIndex = 0
    m_bLoading = False

    cmdNewCustomer.Visible = Not g_bWillCallUser
    cmdMiscOrder.Visible = Not g_bWillCallUser
    lblExplainNew.Visible = Not g_bWillCallUser
    lblExplainMisc.Visible = Not g_bWillCallUser
    
    'disable Load Order button on Select Order grid at start up because grid will be empty
    cmdLoadOrder(1).Enabled = False
    
    'Initialize grid wrappers
    Set m_gwOrders = New GridEXWrapper
    m_gwOrders.Grid = gdxOrders
    Set m_gwCustOrders = New GridEXWrapper
    m_gwCustOrders.Grid = gdxCustOrders
    
    m_gwOrders.InitGridLayout GetUserKey, g_OrderGridRev
    m_gwCustOrders.InitGridLayout GetUserKey, g_CustOrderGridRev

    SetSearchDefaults

    SetComboByText cboTimeInterval, "30 days"
    
    lblOrderStatus.Visible = False
    cboOrderStatus.Visible = False
    
    If InStr(1, GetUserName, "WillCall", vbTextCompare) > 0 Then
        LoadWCGridPrefs
    End If
    
    chkShowOrdersForShipAddr.Visible = False
            
End Sub


Private Sub Form_Activate()
    If tabMain.SelectedTab.Index = tmicreateorder Then
        txtCustSearch.SetFocus
    End If
    
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Set m_gwCustOrders = Nothing
    Set m_gwOrders = Nothing
    
    MDIMain.UnloadTool m_lWindowID
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyW And Shift = vbCtrlMask Then
        Call cmdWalkupOrder_Click
    Else
        MDIMain.GlobalKeyDownProcessing KeyCode, Shift
    End If
End Sub


Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.width < k_lMinWidth Then
        Me.width = k_lMinWidth
    End If
    
    If Me.Height < k_lMinHeight Then
        Me.Height = k_lMinHeight
    End If

    tabMain.width = Me.width - 225
    tabMain.Height = Me.Height - 645
    
    'CreateOrder
    tpCreateOrder.width = tabMain.width - 110
    tpCreateOrder.Height = tabMain.Height - 390
    
    'View 1
    frmCreateOrder.width = tpCreateOrder.width - 120
    frmCreateOrder.Height = tpCreateOrder.Height - 30
    
    'View 2
    frmCustOrders.width = tpCreateOrder.width - 120
    frmCustOrders.Height = tpCreateOrder.Height - 30
    
    gdxCustOrders.width = frmCustOrders.width - 195
    gdxCustOrders.Height = frmCustOrders.Height - 2160
    
    cmdNewOrder.Top = gdxCustOrders.Top + gdxCustOrders.Height + 105
    cmdLoadOrder(0).Top = cmdNewOrder.Top
    cmdNewSearch.Top = cmdNewOrder.Top
    lblOrderStatus.Top = cmdNewOrder.Top
    cboOrderStatus.Top = cmdNewOrder.Top
    cmdContactMgr.Top = cmdNewOrder.Top
    
    'SelectOrder
    tpSelectOrder.width = tabMain.width - 110
    tpSelectOrder.Height = tabMain.Height - 390
    
    gdxOrders.width = tpSelectOrder.width - 120
    gdxOrders.Height = tpSelectOrder.Height - 1890
    
    cmdLoadOrder(1).Top = gdxOrders.Top + gdxOrders.Height + 105
    
    DoEvents
    MDIMain.DoRefresh
End Sub



Private Sub TransitionTabs(ByVal bFindMode As Boolean, Optional strSBText As String = vbNullString)
'a place holder to get this to compile for now
End Sub


Private Sub tabMain_TabClick(ByVal NewTab As ActiveTabs.SSTab)
    Select Case NewTab.Index
        
        Case tmicreateorder
            If frmCreateOrder.Visible Then
                TryToSetFocus txtCustSearch
            End If
        
        Case tmiSelectOrder
            TryToSetFocus txtFindOrder
    End Select
    
    With m_oBrokenRules
        .EnableClass ControlClass.ccCustomer, (NewTab.Index = tmicreateorder)
        .Validate
    End With
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

'!!! Some or all of this should be in FOrder2
    m_oOrder.PricePackList = SetPPL(m_oCustomer)
'    chkPricePackList.value = IIf(m_oOrder.PricePackList, vbChecked, vbUnchecked)
   
    TransitionTabs False

'!!! This too should be in FOrder2
'    LoadUPSCtrl m_oCustomer.Key
    
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
' Assumes the bound recordset for each of the above 3 grids contains
' OPKey, StatusCode, SOKey, SOID

'This is the calling hierarchy.
' cmdLoadOrder()
'     ContinueLoadOrder()
'         ContinueEditSageOrder()
'             CheckSageOrder()
'                 EligibleForDelete()
'                     spcpcEligibleForDelete()
'
' This needs to be reviewed and factored.
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

    m_bLoading = True

'!!! Change this

    'SOStatusCode = 0 for now
    If LoadOrder(oGW.value("OPKey"), _
                    oGW.value("StatusCode"), _
                    oGW.value("SOKey"), _
                    oGW.value("SOID"), _
                    0) Then
        
        m_bLoading = True
    
        'How about keeping current customer search result. (?)
        
        'Clear customer search if order selected from Select Order tab
        If Index = 1 Then
            ClearCustomerSearch
        End If

        'cache module level references to the order's customer and item collection objects
        Set m_oCustomer = m_oOrder.Customer
        
'!!!commented out for now
'        Set m_oItems = m_oOrder.Items
        
        TransitionTabs False
    
'!!! This should be in FOrder2
'        LoadUPSCtrl m_oCustomer.Key
    
        m_bLoading = False
    End If
    
End Sub


'!!!This will probably be in FOrder2
'It's here for the moment because it's called by FindOrderByNumber and cmdLoadOrder_Click

Private Function LoadOrder(ByVal i_OPKey As Long, _
                            ByVal i_OPStatusCode As Long, _
                            ByVal i_SOKey As Long, _
                            i_SOID As String, _
                            ByVal i_SOStatusCode As SOStatusCode) As Boolean

    LoadOrder = False
                
'!!! Think about how this fits in with the new refactoring
'    'control the caption of toolbar's commit button
'    m_bRecommit = False

'!!! Should all of this be done in FOrder2?
'
'    'i_OPStatusCode was generated by the SQL select
'    'it is not the same as tcpSO.StatusCode (aka ItemStatusCode)
'    Dim lOPStatusCode As ItemStatusCode
'    lOPStatusCode = RestoreItemStatusCode(i_OPStatusCode)
'
'    Select Case lOPStatusCode
'
'        Case ItemStatusCode.iscPendingCommit
'            msg "This order is in the process of being saved to Sage and can't be opened at this time."
'            Exit Function
'
'        Case ItemStatusCode.iscCommitted
'            If Not ContinueEditSageOrder(i_OPKey, i_SOKey, i_SOID, i_SOStatusCode) Then
'                Exit Function
'            End If
'
'        Case ItemStatusCode.iscHasRMA
'            Set m_oOrder = New Order        'Why ???!!!
'            m_oOrder.Load i_OPKey
'
'        Case Else
'            m_oOrder.Load i_OPKey
'
'    End Select
    
    LoadOrder = True

End Function



Private Sub cmdContactMgr_Click()
    If Not m_oCustomer.Contacts Is Nothing Then
        m_oCustomer.Contacts.Edit GetUserName
    End If
End Sub


'Click new search button to return to Select Customer search mode

Private Sub cmdNewSearch_Click()
    Set m_oCustomer = New Customer
    SetCaption "Order Manager"          '???
    ClearCustomerSearch
End Sub


Private Sub cboOrderStatus_Click()
    If m_bLoading Then Exit Sub
    m_bLoading = True
    LoadCustomerOrders (chkShowOrdersForShipAddr.value = vbChecked), cboOrderStatus.text
    m_bLoading = False
End Sub


'************************************************************************************
' Create Order Tab
'************************************************************************************

'Called by:
'   ClearCustomerSearch()   IsNewSearch = True
'   FillSelectCustTab()     IsNewSearch = False

Private Sub SetSelCustCtrls(ByVal b_IsNewSearch As Boolean)
    Dim bLoading As Boolean
    bLoading = m_bLoading
    m_bLoading = True

    frmCreateOrder.Visible = b_IsNewSearch
    frmCustOrders.Visible = Not b_IsNewSearch

    chkShowOrdersForShipAddr.Visible = Not b_IsNewSearch
    chkShowOrdersForShipAddr.value = vbChecked
    
    lblOrderStatus.Visible = Not b_IsNewSearch
    cboOrderStatus.Visible = Not b_IsNewSearch
    SetComboByText cboOrderStatus, "<Any>"
    
    lblExplainNew.Visible = Not g_bWillCallUser
    lblExplainMisc.Visible = Not g_bWillCallUser
    cmdNewCustomer.Visible = Not g_bWillCallUser
    cmdMiscOrder.Visible = Not g_bWillCallUser
    
    m_bLoading = bLoading
End Sub


'Select a customer with an existing account.
'Find a customer account and load a list of the customer's past orders
'This subroutine loads customer orders if finding customer succeeds.
'Otherwise, set focus to cust search textbox for new search
    
Private Sub cmdFindAccount_Click()

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
    
End Sub


'Create an order for a customer that will open an account with this order.

Private Sub cmdNewCustomer_Click()
    
 '!!! Change this

    With m_oCustomer
        .Clear
        .IsTemp = True
    End With
    
    m_oOrder.Create
    m_oOrder.Customer = m_oCustomer
    
    m_bLoading = True

'!!! I think this is on FOrder2
'    txtCustID.Text = ""
'    txtCustName.Text = ""
    
    TransitionTabs False
    
    m_bLoading = False
            
End Sub


Private Sub cmdWalkupOrder_Click()

'!!! The initial template for creating an order

'   instantiate a Customer object and initialize it as a Walkup
'   instantiate a new FOrder
'   passing it a customer object

    Set m_oCustomer = New Customer
    m_oCustomer.InitWalkup CreateMISC_CustID

    Dim oFrm As FOrder2
    Set oFrm = New FOrder2
    MDIMain.AddNewWindow oFrm
    With oFrm
        .Show
        .Order.Create
        .Order.Customer = m_oCustomer
        .Order.IsWalkup = True
        .Order.ShipMethKey = GetWillCallShipMethodKey
        .Order.SalesTax.Init m_oCustomer
        If .Order.IsWillCall Then
            .Order.SalesTax.WillCallTaxOverride .Order.whseid
        End If
        .EnteringOrderMode
        .ResizeOrderMode
        .UpdateOrderStatusBar
    
    End With
    MDIMain.UpdateToolbarStatus
    
End Sub


Private Sub cmdMiscOrder_Click()

'!!! Change this

    'Set m_oCustomer = New Customer
    m_oCustomer.InitMiscCustomer CreateMISC_CustID
            
    m_oOrder.Create
    m_oOrder.Customer = m_oCustomer

    'NOTE: Contact ownerkey = 0, Contact ownrtype = 0
    m_oOrder.IsMisc = True
    
    m_oOrder.SalesTax.Init m_oCustomer
    If m_oOrder.IsWillCall Then
        m_oOrder.SalesTax.WillCallTaxOverride m_oOrder.whseid
    End If
    
    m_bLoading = True
    
    TransitionTabs False
    
    m_bLoading = False
            
End Sub



Private Sub txtCustSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtCustSearch.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            Call cmdFindAccount_Click
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


'this event handler returns focus to txtCustSearch after selecting a new Search type
'Added Loading flag logic to Form_Load because the SetFocus method throws an error during Form_Load
Private Sub cboSearchType_Click()
    If m_bLoading Then Exit Sub
    txtCustSearch.SetFocus
End Sub


Private Sub LoadValidationRules()
    Dim oCtlWrapper As ControlWrapper

    With m_oBrokenRules
        'Customer search fields
        Set oCtlWrapper = .AddControl(txtCustSearch, k_sCustNameOrID, True, False)
        oCtlWrapper.AddRuleRequired "", ccCustomer, True, "Enter a value to search for a customer."
        .EnableClass ccCustomer, True
    End With
    
End Sub


'FOrder.LoadCustomerOrders
'populates the Customer Orders grid (gdxCustOrders)
'it does this in 3 different ways

'Called by:
'   chkShowOrdersForShipAddr_Click()    bShowAllShipAddr = value of chkShowOrdersForShipAddr, cboOrderStatus.Text
'   cboOrderStatus_Click()              bShowAllShipAddr = value of chkShowOrdersForShipAddr, cboOrderStatus.Text
'   FillSelectCustTab()                 bShowAllShipAddr = True, cboOrderStatus.Text

Private Sub LoadCustomerOrders(ByVal bShowAllShipAddr As Boolean, ByVal sOrderStatus As String)
    Dim orst As ADODB.Recordset
    Dim sSQL As String
    Dim sWhere As String
    
    SetWaitCursor True
    
'    If sOrderStatus = "Promises" Then
'
'        If Not bShowAllShipAddr Then
'            Set oRst = CallSP("spcpcCustPromiseOpenOrder", "@_iCustKey", m_oCustomer.Key, _
'                "@_iShipAddrKey", m_oCustomer.ShipAddr.AddrKey)
'        Else
'            Set oRst = CallSP("spCPCCustPromiseOpenOrder", "@_iCustKey", m_oCustomer.Key)
'        End If
'
'    Else
    If sOrderStatus = "Open RMA" Then
    
        sSQL = "SELECT OPKey, StatusCode, CreateDate, RTRIM(UserID) as UserID, WhseKey, CustKey, OrderedBy, SOKey, Summary, PurchOrd AS CustPO, ShipAddrKey, Info as Note " & _
           "FROM tcpSO WHERE 1 = 2"
        Set orst = LoadDiscRst(sSQL)
        
    Else
    
        sWhere = "WHERE o.CustKey=" & m_oCustomer.Key & " AND StatusCode < " & iscDeleted
        
        If Not bShowAllShipAddr Then
            sWhere = sWhere & " AND ShipAddrKey=" & m_oCustomer.ShipAddr.AddrKey
        End If
        
        AddStatusFilterClause sWhere, cboOrderStatus

        'Load special research status to the grid instead of general research status if applicable
        'Add SOID to Customer order recordset
        'PRN#3 Add tsoSalesOrder to retrieve Open/Closed status

        sSQL = "SELECT o.OPKey as OPID, o.TranKey as SOID, o.CreateDate, o.ShipAddrXML, rma.rmakey, o.Info as Note, " _
             & "Case StatusCode when 0 then 0 when 1 then " _
             & "Case when ResearchStatus is null then 1001 " _
             & "when ResearchStatus = 0 then 1001 " _
             & "Else 1000+ResearchStatus End " _
             & "Else 2000+StatusCode End As StatusCode, " _
             & "(CASE WHEN o.flags&0x01 = 0x01 THEN -1 ELSE 0 END) AS Dropship, " _
             & "o.UpdateDate, RTRIM(o.UserID) as UserID, o.WhseKey, o.CustKey, " _
             & "RTRIM(ISNULL(c.Name, '')) as OrderedBy, o.OPKey, " _
             & "o.SOKey, o.Summary, RTRIM(o.PurchOrd) as CustPO, ShipAddrKey " _
             & ", (CASE dbo.tsoSalesOrder.Status WHEN 1 THEN 'Open' when 4 then 'Closed' ELSE 'Other' END)  Status "

        sSQL = sSQL & "FROM tcpSO o LEFT OUTER JOIN " & _
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



'After user changes selection in the grid, update the header address accordingly.

Private Sub gdxCustOrders_SelectionChange()
    Dim lShipAddrKey As Long

    'Exit if the grid is empty
    If m_gwCustOrders.value("OPKey") = Empty Then Exit Sub
    
    lShipAddrKey = m_gwCustOrders.value("ShipAddrKey")
    'If lShipAddrKey = 0 Then Exit Sub 'get out if no record defined
    
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
'    Msg "There is not a valid shipping address associated with the selected order." & vbCrLf _
'      & "Please report this problem to OP including the selected order number.", _
'        vbOKOnly + vbInformation, "Invalid Shipping Address"
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
    Set orst = CallSP("cpopGetInvoiceDetail", "@_iOPKey", m_gwCustOrders.value("OPKey"))
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
    Dim bLoading As Boolean
    bLoading = m_bLoading
    m_bLoading = True
    
    m_lCustKey = 0
    txtCustSearch.text = ""
    ClearCustomerOrders
    SetSelCustCtrls True
    
    If tabMain.SelectedTab.Index = tmicreateorder Then
        TryToSetFocus txtCustSearch
        m_oBrokenRules.Validate txtCustSearch
    End If
    
    m_bLoading = bLoading
End Sub


Private Sub ClearCustomerOrders()
    Dim orst As ADODB.Recordset
    Dim sSQL As String

    'generate a zero-record recordset

    sSQL = "SELECT OPKey, StatusCode, CreateDate, RTRIM(UserID) as UserID, WhseKey, CustKey, OrderedBy, SOKey, Summary, PurchOrd AS CustPO, ShipAddrKey, Info as Note " & _
           "FROM tcpSO WHERE 1 = 2"

    Set orst = LoadDiscRst(sSQL)

    With gdxCustOrders
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = orst
        cmdNewOrder.Enabled = False
        cmdLoadOrder(0).Enabled = False
        .Refetch
    End With
    Set orst = Nothing
End Sub


'both modes load a Customer object by passing in a key
'if called from FindMode,





'Load the customer with the found key.
'Load the controls on Select Customer tab.
'
'Called by:
'   cmdCustSearch_Click()           FindMode = True
'   mnuCustOrderRefresh_Click()     FindMode = True
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
        
        lblCustID(1).caption = .ID
        lblCustType(1).caption = .CustType
        
        m_lDefaultBillAddrKey = .BillAddr.AddrKey
        m_lDefaultShipAddrKey = .ShipAddr.AddrKey
        
        lblCustAddress.caption = .ShipAddr.CompleteAddr
    End With

    'change the default value of the checkbox
    'Load orders to grid
    LoadCustomerOrders (chkShowOrdersForShipAddr.value = vbChecked), cboOrderStatus.text

    cmdNewOrder.Enabled = HasRight(k_sRightShowToolOP)

    'if the customer has orders
    If gdxCustOrders.RowCount > 0 Then
        cmdLoadOrder(0).Enabled = True
        TryToSetFocus gdxCustOrders
    Else
        cmdLoadOrder(0).Enabled = False
        TryToSetFocus cmdNewOrder
        If vbYes = msg("Would you like to create a new order for " & m_oCustomer.Name & "?", vbYesNo, "New Order?") Then
            cmdNewOrder_Click
        End If
    End If
    
    SetWaitCursor False
End Sub


'Called by Form Load

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

' Called by LoadResultGrid()
' This is used to expand only the columns needed for Will Call users.
' If the collection is not populated on form load, all rows will be expanded.

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


'************************************************************************************
' Events and subroutines on Select Order tab
'************************************************************************************

Private Sub cmdFindOrders_Click(Index As Integer)
    Select Case Index
'This event finds orders via OPID, SOID(Trankey), RMA Key, or Customer PO
'Called by:
'   mnuOrderRefresh_Click
'   OrderDelete
'   txtFindOrder_KeyDown
        Case 0:

            FindOrderByNumber

'This click event is used to search orders
'via combination of many criteria: Customer, CSR, Status, etc.
'Called by
'   txtFindCust_KeyDown
'   txtFindText_KeyDown
'   cmdMyOrders_Click
'   mnuOrderRefresh_Click
'   CancelButton
'   SaveButton
'   Commit
        Case 1:
            FindOrdersByCriteria
    End Select
    
End Sub


Private Sub FindOrderByNumber()

    Dim sWhere As String
    Dim sInput As String
    Dim bLoading As Boolean
    Dim orst As ADODB.Recordset
    Dim lResult As Long
    Dim lSOID As Long

'1/31/05 LR Replaced
'Require a numeric input (we're checking for only OP, SO and RMA #)
            
    sInput = Trim$(PrepSQLText(txtFindOrder.text))
    
    If Len(sInput) = 0 Then
        Exit Sub
    End If
    
    ' added PO lookup 8/19/2010 LR
    
    On Error GoTo ErrorHandler
            
    If InStr(LCase(sInput), "p") = 1 Then
    
        m_bFindOrderFlag = True
        
        SetWaitCursor True
        
        Set orst = CallSP("spcpcFindOrderByCustPO", "@_iCustPO", Mid$(sInput, 2))
        
        SetWaitCursor False
        
    Else
    
        If Not IsNumeric(sInput) Then
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
        
            msg "No records satisfy this request"
            txtFindOrder.SelStart = 0
            txtFindOrder.SelLength = Len(txtFindOrder.text)
            TryToSetFocus txtFindOrder
            
        Case 1:
        
'!!! Change this
            'SOStatusCode = 0 for now
            If LoadOrder(orst("OPKey").value, _
                            orst("StatusCode").value, _
                            orst("SOKey").value, _
                            orst("SOID").value, _
                            0) Then
                
                bLoading = m_bLoading
                m_bLoading = True
            
                ClearCustomerSearch  'invoked from Select Order tab
            
                'loading a MISC order with OrderedBy value
                'Contact state went from New to New + Dirty
                Set m_oCustomer = m_oOrder.Customer
'!!! This needs to be in FOrder2
'                Set m_oItems = m_oOrder.Items
                
                TransitionTabs False
            
'!!! This should be in FOrder2
'                 LoadUPSCtrl m_oCustomer.Key
            
                m_bLoading = bLoading
            
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


Private Sub FindOrdersByCriteria()

    Dim sWhere As String
    Dim sInput As String
    Dim lResult As Long
    Dim sTemp As String
    Dim rst As ADODB.Recordset
    
    SetWaitCursor True
    
    m_bFindOrderFlag = False
      
'    If cboFindStatus.text = "Promises" Then
'
'        Select Case cboFindCSR.text
'            Case "<Any>"
'                Set rst = CallSP("spcpcPromiseOpenOrder")
'            Case "<STL>"
'                Set rst = CallSP("spcpcPromiseOpenOrder", "@_iWhseKey", 25)
'            Case "<MPK>"
'                Set rst = CallSP("spcpcPromiseOpenOrder", "@_iWhseKey", 23)
'            Case "<SEA>"
'                Set rst = CallSP("spcpcPromiseOpenOrder", "@_iWhseKey", 24)
'            Case Else
'                Set rst = CallSP("spcpcPromiseOpenOrder", "@_iUserID", cboFindCSR.text)
'        End Select
'
'    Else
    'Show the Open RMA Orders
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
        
'4/21/08 LR use the new time interval combo in our query
' NOTE: we've coded the data directly into the control's List and ItemData properties using the design
' time property pages
'***DH 5/13/08 Changed the where clause to use UpdateDate instead of CreateDate

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


'Called by
' FindOrdersByCriteria()
' SetSearchDefaults()

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

'*** 7/18/08 DH added MPK WillCall exception.
    If InStr(1, GetUserName, "LAWillCall", vbTextCompare) > 0 Then
        SetComboByText cboFindCSR, "<MPK>"
        SetComboByText cboFindStatus, "<Any>"
    Else
        '4/21/08 LR changed cbo selected values
        SetComboByText cboFindCSR, GetUserName
        SetComboByText cboFindStatus, "All Quotes"
    End If
    
    TryToSetFocus txtFindCust
    
    'generate a zero-record recordset to clear the grid
    sSQL = "SELECT OPKey FROM tcpSO WHERE 1 = 2"
    Set rst = LoadDiscRst(sSQL)
    LoadResultGrid rst
End Sub




'This event is used to reset searching criteria
Private Sub cmdClearOrder_Click()
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
    'gdxOrders.LoadLayoutString m_gwOrders.BkLayout
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
            'PRN#92
            'Case "Open (any of above)": AppendClause "StatusCode <= 7", i_sWhere
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


'Displays the intenational country caution based on the customer's shipping address

Private Sub DisplayIntrlCaution()
    Dim sSQL As String
    Dim sMsg As String
    Dim orst As ADODB.Recordset
    Dim sMsgTitle As String
    
    sMsgTitle = m_oCustomer.Name
    
    'Get Caution for International Country
    sSQL = "select Cautions, CSWCountrySymbol from tcpcswcountry where countryid = '" & m_oCustomer.ShipAddr.CountryID & "'"
    Set orst = New ADODB.Recordset
    Set orst = LoadDiscRst(sSQL, , adLockBatchOptimistic)

    If orst.EOF Then
        'there are no cautions for this country
    Else
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
        
    orst.Close
    Set orst = Nothing
End Sub


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


Private Sub txtFindText_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtFindText.text) > 0 Then
        If KeyCode = vbKeyReturn Then
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

Private Sub txtFindCust_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
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



Private Sub cmdCollapseAll_Click()
    gdxOrders.CollapseAll
End Sub

Private Sub cmdExpandAll_Click()
    gdxOrders.ExpandAll
End Sub

Private Sub mnuCustOrderCollapseAll_Click()
    gdxCustOrders.CollapseAll
End Sub

Private Sub txtFindCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtFindCust.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            FindOrdersByCriteria
        End If
    End If
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
         & "o.CustName, o.SOKey, o.Summary, RTRIM(o.PurchOrd) as CustPO, s.Hold, rma.RMAKey " _
         & ", (CASE dbo.tsoSalesOrder.Status WHEN 1 THEN 'Open' when 4 then 'Closed' ELSE 'Other' END) as Status " _
         & ", tciShipMethod.ShipMethID, tciPaymentTerms.PmtTermsID, o.Info as Note " _
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
         & sSearchingText & vbCrLf & "ORDER BY o.UpdateDate DESC"
    
    Set SearchOrder = LoadDiscRst(sSQL)
End Function


Private Sub AppendClause(ByVal i_sClause As String, ByRef io_sWhere As String)
    If Len(io_sWhere) = 0 Then
        io_sWhere = vbCrLf & "WHERE (" & i_sClause & ")"
    Else
        io_sWhere = io_sWhere & vbCrLf & "  AND (" & i_sClause & ")"
    End If
End Sub







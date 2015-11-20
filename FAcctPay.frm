VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "MMRemark.ocx"
Object = "{0FA91D91-3062-44DB-B896-91406D28F92A}#54.0#0"; "SOTACalendar.ocx"
Begin VB.Form FAcctPay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Payable"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   9855
   Begin ActiveTabs.SSActiveTabs TabMain 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11880
      _Version        =   262144
      TabCount        =   4
      Tabs            =   "FAcctPay.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   68
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FAcctPay.frx":00FA
         Begin MSComctlLib.TreeView tvwPO 
            Height          =   3855
            Left            =   240
            TabIndex        =   77
            Top             =   1680
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   6800
            _Version        =   393217
            Style           =   2
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Frame frmPendInvoices 
            Caption         =   "Pending Invoice"
            Height          =   1455
            Left            =   6000
            TabIndex        =   80
            Top             =   120
            Width           =   2775
            Begin VB.CommandButton cmdPIDisplay 
               Caption         =   "&Display"
               Height          =   375
               Left            =   240
               TabIndex        =   75
               Top             =   960
               Width           =   1095
            End
            Begin VB.CommandButton cmdAddPendInv 
               Caption         =   "&Add"
               Height          =   375
               Left            =   1440
               TabIndex        =   76
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox txtPendInvoice 
               Height          =   285
               Left            =   840
               TabIndex        =   73
               Top             =   240
               Width           =   1455
            End
            Begin SOTACalendarControl.SOTACalendar sotaInvDate 
               Height          =   255
               Left            =   840
               TabIndex        =   74
               Top             =   600
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   450
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
            Begin VB.Label lblInvDate 
               Caption         =   "Inv Date"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label5 
               Caption         =   "Invoice"
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame frmPurchOrd 
            Caption         =   "Purchase Order"
            Height          =   1455
            Left            =   120
            TabIndex        =   69
            Top             =   120
            Width           =   5775
            Begin VB.CommandButton cmdFindPO 
               Caption         =   "F&ind"
               Height          =   255
               Left            =   3000
               TabIndex        =   72
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtPO 
               Height          =   285
               Left            =   1320
               TabIndex        =   71
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lblVIPODate 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   4320
               TabIndex        =   79
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblVIVendorName 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1320
               TabIndex        =   78
               Top             =   600
               Width           =   4215
            End
            Begin VB.Label lblPO 
               Caption         =   "PO Number"
               Height          =   255
               Left            =   240
               TabIndex        =   70
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   39
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FAcctPay.frx":0122
         Begin VB.Frame frmAssignVend 
            Caption         =   "Assign Buyers to a Parts Vendor"
            Height          =   1572
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   8655
            Begin VB.ComboBox cboVendors 
               Height          =   315
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   360
               Width           =   2652
            End
            Begin VB.CommandButton cmdSaveBuyers 
               Caption         =   "Save"
               Height          =   312
               Left            =   7260
               TabIndex        =   49
               Top             =   360
               Width           =   912
            End
            Begin VB.ComboBox cboSEABuyer 
               Height          =   315
               Left            =   5580
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   720
               Width           =   1392
            End
            Begin VB.ComboBox cboSTLBuyer 
               Height          =   315
               Left            =   5580
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   1080
               Width           =   1392
            End
            Begin VB.ComboBox cboMPKBuyer 
               Height          =   315
               Left            =   5580
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   360
               Width           =   1392
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "Vendor"
               Height          =   192
               Left            =   240
               TabIndex        =   41
               Top             =   420
               Width           =   552
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Seattle"
               Height          =   252
               Left            =   4800
               TabIndex        =   45
               Top             =   780
               Width           =   612
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "St. Louis"
               Height          =   252
               Left            =   4740
               TabIndex        =   47
               Top             =   1140
               Width           =   672
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Monterey Park"
               Height          =   252
               Left            =   4320
               TabIndex        =   42
               Top             =   420
               Width           =   1092
            End
         End
         Begin VB.Frame FrmPrefVendor 
            Caption         =   "Preferred Vendor Maintenance"
            Height          =   3852
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   8655
            Begin VB.CommandButton cmdGet 
               Caption         =   "Get"
               Enabled         =   0   'False
               Height          =   312
               Left            =   6840
               TabIndex        =   88
               Top             =   360
               Width           =   912
            End
            Begin VB.ComboBox cboWhse 
               Height          =   315
               Index           =   0
               Left            =   4800
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   360
               Width           =   915
            End
            Begin VB.Frame frmVendDetail 
               Caption         =   "Vendor Detail"
               Height          =   2955
               Left            =   120
               TabIndex        =   53
               Top             =   780
               Width           =   8295
               Begin VB.CommandButton cmdPrefUpdate 
                  Caption         =   "&Update"
                  Enabled         =   0   'False
                  Height          =   312
                  Left            =   6720
                  TabIndex        =   87
                  Top             =   360
                  Width           =   912
               End
               Begin VB.ComboBox cboVendor 
                  Height          =   315
                  Left            =   1560
                  Style           =   2  'Dropdown List
                  TabIndex        =   85
                  Top             =   360
                  Width           =   3552
               End
               Begin MMRemark.RemarkViewer rvVendor 
                  Height          =   915
                  Left            =   6720
                  TabIndex        =   54
                  Top             =   1080
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   1614
                  ContextID       =   "ViewVendor"
                  Caption         =   "Vendor Remarks"
               End
               Begin VB.Label Label3 
                  Caption         =   "Vendor"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   86
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Caption         =   "Address"
                  Height          =   255
                  Index           =   2
                  Left            =   600
                  TabIndex        =   55
                  Top             =   1080
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Name"
                  Height          =   255
                  Index           =   1
                  Left            =   600
                  TabIndex        =   56
                  Top             =   720
                  Width           =   495
               End
               Begin VB.Label lblVendName 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   58
                  Top             =   780
                  Width           =   4335
               End
               Begin VB.Label lblVendEMail 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   4320
                  TabIndex        =   65
                  Top             =   2580
                  Width           =   1575
               End
               Begin VB.Label lblVendFax 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   4320
                  TabIndex        =   66
                  Top             =   2280
                  Width           =   1575
               End
               Begin VB.Label Label61 
                  Caption         =   "EMail"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   63
                  Top             =   2580
                  Width           =   495
               End
               Begin VB.Label Label60 
                  Caption         =   "Fax"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   62
                  Top             =   2280
                  Width           =   375
               End
               Begin VB.Label lblVendPhone 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   64
                  Top             =   2580
                  Width           =   1575
               End
               Begin VB.Label Label59 
                  Caption         =   "Phone"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   61
                  Top             =   2580
                  Width           =   495
               End
               Begin VB.Label Label58 
                  Caption         =   "Contact"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   60
                  Top             =   2280
                  Width           =   735
               End
               Begin VB.Label lblVendContact 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   59
                  Top             =   2280
                  Width           =   1575
               End
               Begin VB.Label lblAddress 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   1095
                  Left            =   1560
                  TabIndex        =   57
                  Top             =   1080
                  Width           =   4335
               End
            End
            Begin VB.ComboBox cboMake 
               Height          =   315
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   360
               Width           =   1395
            End
            Begin VB.Label Label4 
               Caption         =   "Warehouse"
               Height          =   255
               Left            =   3720
               TabIndex        =   83
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Make"
               Height          =   255
               Left            =   300
               TabIndex        =   51
               Top             =   360
               Width           =   1095
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   2
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FAcctPay.frx":014A
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   372
            Left            =   4860
            TabIndex        =   67
            Top             =   5220
            Width           =   972
         End
         Begin VB.Frame Frame2 
            Caption         =   "Confirm DropShip Info"
            Height          =   912
            Left            =   6120
            TabIndex        =   35
            Top             =   4620
            Width           =   2652
            Begin VB.CommandButton cmdDisplay 
               Caption         =   "Display"
               Height          =   375
               Left            =   1800
               TabIndex        =   38
               Top             =   360
               Width           =   732
            End
            Begin VB.TextBox txtTenKeyPONbr 
               Height          =   315
               Left            =   720
               TabIndex        =   36
               Top             =   360
               Width           =   912
            End
            Begin VB.Label Label1 
               Caption         =   "PONbr"
               Height          =   252
               Index           =   0
               Left            =   120
               TabIndex        =   37
               Top             =   420
               Width           =   552
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Options"
            Height          =   1212
            Left            =   6120
            TabIndex        =   31
            Top             =   3300
            Width           =   2655
            Begin VB.CheckBox ckbDecimal 
               Caption         =   "Automatic Decimal"
               Height          =   255
               Left            =   300
               TabIndex        =   34
               Top             =   840
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox ckbEnter 
               Caption         =   "'Enter' = '+'"
               Height          =   255
               Left            =   300
               TabIndex        =   33
               Top             =   540
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox ckbSubtraction 
               Caption         =   "Allow Subtraction"
               Height          =   255
               Left            =   300
               TabIndex        =   32
               Top             =   240
               Value           =   1  'Checked
               Width           =   1575
            End
         End
         Begin VB.CommandButton cmdKey 
            Caption         =   "--"
            Height          =   1095
            Index           =   14
            Left            =   8280
            TabIndex        =   29
            Top             =   900
            Width           =   495
         End
         Begin VB.CommandButton cmdKey 
            Caption         =   "+"
            Height          =   1095
            Index           =   13
            Left            =   8280
            TabIndex        =   30
            Top             =   2100
            Width           =   495
         End
         Begin VB.CommandButton cmdKey 
            Caption         =   "Clear"
            Height          =   375
            Index           =   12
            Left            =   7560
            TabIndex        =   17
            Top             =   900
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "."
            Height          =   375
            Index           =   11
            Left            =   7560
            TabIndex        =   28
            Top             =   2820
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "Backspace"
            Height          =   375
            Index           =   10
            Left            =   6120
            TabIndex        =   16
            Top             =   900
            Width           =   1335
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "9"
            Height          =   375
            Index           =   9
            Left            =   7560
            TabIndex        =   20
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "8"
            Height          =   375
            Index           =   8
            Left            =   6840
            TabIndex        =   19
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "7"
            Height          =   375
            Index           =   7
            Left            =   6120
            TabIndex        =   18
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "6"
            Height          =   375
            Index           =   6
            Left            =   7560
            TabIndex        =   23
            Top             =   1860
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "5"
            Height          =   375
            Index           =   5
            Left            =   6840
            TabIndex        =   22
            Top             =   1860
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "4"
            Height          =   375
            Index           =   4
            Left            =   6120
            TabIndex        =   21
            Top             =   1860
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "3"
            Height          =   375
            Index           =   3
            Left            =   7560
            TabIndex        =   26
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "2"
            Height          =   375
            Index           =   2
            Left            =   6840
            TabIndex        =   25
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "1"
            Height          =   375
            Index           =   1
            Left            =   6120
            TabIndex        =   24
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton cmdKey 
            Appearance      =   0  'Flat
            Caption         =   "0"
            Height          =   375
            Index           =   0
            Left            =   6120
            TabIndex        =   27
            Top             =   2820
            Width           =   1335
         End
         Begin VB.CommandButton cmdPrintTape 
            Caption         =   "Print"
            Height          =   375
            Left            =   3720
            TabIndex        =   14
            Top             =   5220
            Width           =   975
         End
         Begin VB.CommandButton cmdReconcile 
            Caption         =   "Reconcile"
            Height          =   375
            Left            =   4860
            TabIndex        =   13
            Top             =   4800
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
            Left            =   6120
            TabIndex        =   15
            Top             =   360
            Width           =   2652
         End
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Load"
            Height          =   375
            Left            =   3720
            TabIndex        =   12
            Top             =   4800
            Width           =   975
         End
         Begin VB.TextBox txtBatch 
            Height          =   285
            Left            =   960
            TabIndex        =   11
            Top             =   4860
            Width           =   1335
         End
         Begin VB.OptionButton optAP 
            Caption         =   "A/P"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   4860
            Width           =   615
         End
         Begin VB.OptionButton optPO2 
            Caption         =   "PO"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   5100
            Width           =   615
         End
         Begin MSComctlLib.ListView lvwBatch 
            Height          =   3912
            Left            =   180
            TabIndex        =   4
            Top             =   360
            Width           =   3348
            _ExtentX        =   5900
            _ExtentY        =   6906
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
            Height          =   3912
            Left            =   3600
            TabIndex        =   6
            Top             =   360
            Width           =   2292
            _ExtentX        =   4048
            _ExtentY        =   6906
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
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
            Height          =   372
            Left            =   180
            TabIndex        =   7
            Top             =   4260
            Width           =   3348
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
            Height          =   372
            Left            =   3600
            TabIndex        =   8
            Top             =   4260
            Width           =   2292
         End
         Begin VB.Label Label11 
            Caption         =   "Batch Detail"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Tape Detail"
            Height          =   255
            Left            =   3600
            TabIndex        =   5
            Top             =   120
            Width           =   975
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6345
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11192
         _Version        =   262144
         TabGuid         =   "FAcctPay.frx":0172
         Begin VB.CommandButton cmdfindDropShipPO 
            Caption         =   "Find"
            Height          =   315
            Left            =   5760
            TabIndex        =   104
            Top             =   5760
            Width           =   855
         End
         Begin VB.TextBox txtPoNumber 
            Height          =   285
            Left            =   4680
            TabIndex        =   103
            Top             =   5760
            Width           =   975
         End
         Begin VB.TextBox txtGridRows 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8460
            Locked          =   -1  'True
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   2760
            Width           =   672
         End
         Begin VB.CommandButton cmdProcess 
            Caption         =   "Post"
            Height          =   312
            Left            =   8260
            TabIndex        =   94
            Top             =   100
            Width           =   972
         End
         Begin VB.ComboBox cboWhse 
            Height          =   315
            Index           =   1
            ItemData        =   "FAcctPay.frx":019A
            Left            =   4680
            List            =   "FAcctPay.frx":019C
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Pre-Post"
            Height          =   312
            Index           =   1
            Left            =   7200
            TabIndex        =   90
            Top             =   120
            Width           =   972
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   312
            Left            =   6120
            TabIndex        =   89
            Top             =   100
            Width           =   972
         End
         Begin GridEX20.GridEX gdxDSOrders 
            Height          =   2295
            Left            =   0
            TabIndex        =   96
            Top             =   480
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4048
            Version         =   "2.0"
            ShowToolTips    =   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            ColumnsCount    =   25
            Column(1)       =   "FAcctPay.frx":019E
            Column(2)       =   "FAcctPay.frx":03DA
            Column(3)       =   "FAcctPay.frx":05E6
            Column(4)       =   "FAcctPay.frx":0842
            Column(5)       =   "FAcctPay.frx":0A26
            Column(6)       =   "FAcctPay.frx":0C26
            Column(7)       =   "FAcctPay.frx":0E32
            Column(8)       =   "FAcctPay.frx":1002
            Column(9)       =   "FAcctPay.frx":11FA
            Column(10)      =   "FAcctPay.frx":13E6
            Column(11)      =   "FAcctPay.frx":15D2
            Column(12)      =   "FAcctPay.frx":17BA
            Column(13)      =   "FAcctPay.frx":19A2
            Column(14)      =   "FAcctPay.frx":1B2A
            Column(15)      =   "FAcctPay.frx":1CB2
            Column(16)      =   "FAcctPay.frx":1E3A
            Column(17)      =   "FAcctPay.frx":1FC2
            Column(18)      =   "FAcctPay.frx":213A
            Column(19)      =   "FAcctPay.frx":22B6
            Column(20)      =   "FAcctPay.frx":2446
            Column(21)      =   "FAcctPay.frx":25D6
            Column(22)      =   "FAcctPay.frx":28A6
            Column(23)      =   "FAcctPay.frx":2A42
            Column(24)      =   "FAcctPay.frx":2B5A
            Column(25)      =   "FAcctPay.frx":2C8E
            FormatStylesCount=   6
            FormatStyle(1)  =   "FAcctPay.frx":2DDA
            FormatStyle(2)  =   "FAcctPay.frx":2EBA
            FormatStyle(3)  =   "FAcctPay.frx":2FF2
            FormatStyle(4)  =   "FAcctPay.frx":30A2
            FormatStyle(5)  =   "FAcctPay.frx":3156
            FormatStyle(6)  =   "FAcctPay.frx":322E
            ImageCount      =   0
            PrinterProperties=   "FAcctPay.frx":32E6
         End
         Begin GridEX20.GridEX gdxPO 
            Height          =   1035
            Left            =   0
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   4440
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   1826
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
            ColumnsCount    =   4
            Column(1)       =   "FAcctPay.frx":34BE
            Column(2)       =   "FAcctPay.frx":3606
            Column(3)       =   "FAcctPay.frx":372A
            Column(4)       =   "FAcctPay.frx":3866
            FormatStylesCount=   6
            FormatStyle(1)  =   "FAcctPay.frx":39EA
            FormatStyle(2)  =   "FAcctPay.frx":3ACA
            FormatStyle(3)  =   "FAcctPay.frx":3C02
            FormatStyle(4)  =   "FAcctPay.frx":3CB2
            FormatStyle(5)  =   "FAcctPay.frx":3D66
            FormatStyle(6)  =   "FAcctPay.frx":3E3E
            ImageCount      =   0
            PrinterProperties=   "FAcctPay.frx":3EF6
         End
         Begin GridEX20.GridEX gdxSO 
            Height          =   1035
            Left            =   0
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   3120
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   1826
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
            ColumnsCount    =   4
            Column(1)       =   "FAcctPay.frx":40CE
            Column(2)       =   "FAcctPay.frx":4216
            Column(3)       =   "FAcctPay.frx":433A
            Column(4)       =   "FAcctPay.frx":4476
            FormatStylesCount=   6
            FormatStyle(1)  =   "FAcctPay.frx":45FE
            FormatStyle(2)  =   "FAcctPay.frx":46DE
            FormatStyle(3)  =   "FAcctPay.frx":4816
            FormatStyle(4)  =   "FAcctPay.frx":48C6
            FormatStyle(5)  =   "FAcctPay.frx":497A
            FormatStyle(6)  =   "FAcctPay.frx":4A52
            ImageCount      =   0
            PrinterProperties=   "FAcctPay.frx":4B0A
         End
         Begin MSComctlLib.ImageList imglRemarks 
            Left            =   960
            Top             =   5520
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
                  Picture         =   "FAcctPay.frx":4CE2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FAcctPay.frx":5134
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label lblPoNumber 
            Caption         =   "PO#:"
            Height          =   255
            Left            =   4200
            TabIndex        =   102
            Top             =   5760
            Width           =   495
         End
         Begin VB.Label Label28 
            Caption         =   "Purchase Order Line Items"
            Height          =   195
            Left            =   0
            TabIndex        =   101
            Top             =   4200
            Width           =   2115
         End
         Begin VB.Label Label29 
            Caption         =   "Sales Order Line Items"
            Height          =   195
            Left            =   0
            TabIndex        =   98
            Top             =   2880
            Width           =   2115
         End
         Begin VB.Label Label26 
            Caption         =   "Order"
            Height          =   195
            Left            =   7980
            TabIndex        =   97
            Top             =   2820
            Width           =   435
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "Warehouse"
            Height          =   255
            Left            =   3600
            TabIndex        =   93
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label27 
            Caption         =   "Drop Ship Orders"
            Height          =   192
            Left            =   0
            TabIndex        =   92
            Top             =   240
            Width           =   2112
         End
      End
   End
End
Attribute VB_Name = "FAcctPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWindowID As Long

Private m_bNew As Boolean
Private m_bLoading As Boolean

Private m_dBalance As Double
Private m_dCurEntry As Double
Private m_sBatchCmnt As String

Private m_bNewVendor As Boolean
Private m_lVendKey As Long
Private m_vMatchToleranceKey As Variant

'To Pass mouse down location to HitTest
Private m_sngX As Single
Private m_sngY As Single

'This is the list item if the user has double
'clicked and wants to edit a tape entry
Private m_EditItem As ListItem

Private m_sSearchText As String
Private m_bVendorLoad As Boolean
Private m_bPrefVendor As Boolean

'Display Mask
'Private Const ksDisplayMask = "###,###,###.00"
Private Const klVScrollBarWidth = 350
Private Const klReconcileColWidth = 300

'For Vendor Invoice Tab
Private m_iPOKey As Long
Private m_iPoIdToFind As Long

'For DropShip tab
Private m_oRstDSOrders As ADODB.Recordset
Private m_oRstSO As ADODB.Recordset
Private m_oRstPO As ADODB.Recordset
Private WithEvents m_gwDSOrders As GridEXWrapper
Attribute m_gwDSOrders.VB_VarHelpID = -1
Private m_arrayDSOrders As Variant
Private m_iDSOrderCount As Integer
Private m_lX As Long
Private m_lY As Long
Public WithEvents oFrm As FProvisionalShipment
Attribute oFrm.VB_VarHelpID = -1
Private dict As Dictionary
Private m_bDictLoaded As Boolean

'Tab Enumerations
Private Enum TabMainIndexes
    tmiDropShip = 1
    tmiTurboTenKey = 2
    tmiVendMaint = 3
    tmiVendInv = 4
End Enum



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


Private Sub cmdAddPendInv_Click()
    If Len(txtPendInvoice) = 0 Then
        MsgBox "You must supply a vendor invoice."
        TryToSetFocus txtPendInvoice
        Exit Sub
    End If
    
    If Not IsDate(sotaInvDate.text) Then
        MsgBox "You must supply an Invoice date."
        TryToSetFocus sotaInvDate
        Exit Sub
    End If
    
    If m_iPOKey = 0 Then
        TryToSetFocus txtPO
        MsgBox "You must select a PO."
        Exit Sub
    End If
    
    'Check to make sure this entry doesn't already exist
    Dim rst As ADODB.Recordset
    Set rst = LoadDiscRst("SELECT POKey From dbo.tcpAPRcvdInv WHERE POKey = " & m_iPOKey & " And VendInvNbr LIKE '" & Trim(Left$(txtPendInvoice, 10)) & "%'")
    If rst.EOF Then
        ExecuteSP "spCPCAPInsertRcvdInv", "@POKey", m_iPOKey, "@VendInvNbr", Left$(txtPendInvoice, 10), "@InvDate", sotaInvDate.text, "@UserID", GetUserID
        txtPendInvoice.text = ""
        sotaInvDate.text = ""
        BuildPOTV m_iPOKey
    Else
        MsgBox "This invoice is already recorded for this PO."
        TryToSetFocus txtPendInvoice
    End If
    
    CloseRst rst
    
End Sub


Private Sub cmdfindDropShipPO_Click()
    Dim Found As Boolean
    
        If Len(txtPoNumber.text) = 0 Then Exit Sub

        If Not IsNumeric(txtPoNumber.text) Then
            msg "Invalid PO number", vbCritical
            Exit Sub
        End If
    
        Found = gdxDSOrders.Find(9, jgexContains, txtPoNumber.text)
        If Not Found Then
            msg "PO number not found", vbExclamation
            Exit Sub
        End If
        gdxDSOrders.EnsureVisible gdxDSOrders.Row
End Sub

Private Sub cmdFindPO_Click()
    Dim frmFindPO As FPOSearch
    
    Set frmFindPO = New FPOSearch
    
    m_iPOKey = 0
    m_iPOKey = frmFindPO.GetPOKey
    Unload frmFindPO
    If m_iPOKey > 0 Then
        LoadPO (m_iPOKey)
    End If
End Sub

Private Sub cmdTest_Click()
    BuildPOTV m_iPOKey
End Sub


Private Sub cmdPIDisplay_Click()
    BuildRecTV
End Sub



Private Sub Form_Load()
    SetCaption "Accounts Payable"
    SetUpTape
    SetUpBatch
    
    'if the DropShip tab is enabled
    If TabMain.Tabs(tmiDropShip).Visible Then
        'Set m_gwDSOrders = New GridEXWrapper
        'm_gwDSOrders.Grid = gdxDSOrders
        'LoadImageList imglRemarks, gdxDSOrders
    End If
    
    LoadControls
    
    'VL
    Dim fmsYellow As JSFormatStyle
    Set fmsYellow = gdxDSOrders.FormatStyles.Add("Non-AP")
'    fmsYellow.BackColor = RGB(253, 253, 150)
    fmsYellow.BackColor = RGB(255, 255, 191)
     
    Dim fmsGreen As JSFormatStyle
    Set fmsGreen = gdxDSOrders.FormatStyles.Add("AP")
'    fmsGreen.BackColor = RGB(119, 190, 119)
    fmsGreen.BackColor = RGB(128, 215, 193)
End Sub


Private Sub LoadControls()
    Dim bSaveFlag As Boolean
    
    bSaveFlag = m_bLoading
    m_bLoading = True
    
    'LoadCombo cboBuyers, orst, "BuyerID", "BuyerKey", , 1
    'SetComboByKey cboBuyers, GetBuyerKey(GetUserName)
    LoadCombo cboMPKBuyer, g_rstBuyers, "BuyerID", "BuyerKey", , 1
    LoadCombo cboSTLBuyer, g_rstBuyers, "BuyerID", "BuyerKey", , 1
    LoadCombo cboSEABuyer, g_rstBuyers, "BuyerID", "BuyerKey", , 1
    
    cmdSaveBuyers.Enabled = False

    LoadCombo cboVendors, g_rstVendors, "VendName", "VendKey", , True
    
    LoadCombo cboMake, g_rstMakes, "MakeText", "MakeID"

    '***changed rst and added filter 8/26/03 LR
    g_rstWhses.Filter = "transit = 0"
    LoadCombo cboWhse(0), g_rstWhses, "WhseID", "WhseKey"
    cboWhse(0).ListIndex = 0
    
    LoadCombo cboWhse(1), g_rstWhses, "WhseID", "WhseKey"
    cboWhse(1).AddItem "All", 0
    cboWhse(1).ListIndex = 0
    
    g_rstWhses.Filter = adFilterNone
    
    LoadCombo cboVendor, g_rstVendors, "VendName", "VendKey", , True
    
    LoadImageList imglRemarks, gdxDSOrders
        
    m_bLoading = bSaveFlag
End Sub


Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
  
    TabMain.width = Me.width - 255
    TabMain.Height = Me.Height - 500
    
    gdxDSOrders.width = TabMain.width - 240
    gdxSO.width = gdxDSOrders.width
    gdxPO.width = gdxDSOrders.width
    gdxDSOrders.Height = (TabMain.Height - 2490) * (2235 / 4305)
    gdxSO.Height = (TabMain.Height - 2490) * (1035 / 4305)
    gdxPO.Height = gdxSO.Height

    Label29.Top = gdxDSOrders.Top + gdxDSOrders.Height + 120
    Label26.Top = Label29.Top
    txtGridRows.Top = Label26.Top - 90
    
    gdxSO.Top = Label29.Top + Label29.Height + 60
    Label28.Top = gdxSO.Top + gdxSO.Height + 60
    gdxPO.Top = Label28.Top + Label28.Height + 30
    
    'Frame5.Top = gdxPO.Top + gdxPO.Height + 30
    'Frame8.Top = Frame5.Top

    gdxDSOrders.Refresh
    gdxSO.Refresh
    gdxPO.Refresh

    DoEvents
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_gwDSOrders = Nothing
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Public Sub DoShowHelp()
    ShowHelp "FAcctPay", True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    'Key preview is turned on.
    'This sub passes all keystrokes to the correct
    'element in the command button array for processing.
    
    If TabMain.SelectedTab.Index <> 2 Then Exit Sub  'Ignore if not Turbo 10 Key
    If Me.ActiveControl.Name = "txtBatch" Or Me.ActiveControl.Name = "txtTenKeyPONbr" Then Exit Sub
    
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


'************************************************************************************
' Private functions
'************************************************************************************

'Convert UserName to BuyerKey.
'If the user is not a buyer in Sage, then return 0.
Private Function GetBuyerKey(username As String) As Long
    Dim orst As Recordset
    Dim sSQL As String
    
    sSQL = "SELECT BuyerKey FROM timBuyer WHERE BuyerID='" & username & " '"
    Set orst = LoadDiscRst(sSQL)
    If orst.EOF Then
        GetBuyerKey = 0
    Else
        GetBuyerKey = orst.Fields("BuyerKey")
    End If
    Set orst = Nothing
End Function


Private Sub AttachGrid(ByRef i_oGrid As GridEX, ByRef i_orst As ADODB.Recordset)
    With i_oGrid
        Dim i As Long

        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = i_orst

        For i = 1 To .Columns.Count
            If .Columns(i).Key <> "Remarks" Then
                .Columns(i).AutoSize
            End If
        Next
    End With
End Sub


'*****************************************************************************
' Costs/Part Numbers tab
'*****************************************************************************

Private Function IsWildcard(sInput As String) As Boolean
    Dim sTemp As String
    Dim sWildChar1
    Dim sWildChar2
    
    IsWildcard = True
    
    sWildChar1 = Split(sInput, "?", -1, vbTextCompare)
    If UBound(sWildChar1) = 0 Then
        sWildChar2 = Split(sInput, "*", -1, vbTextCompare)
        If UBound(sWildChar2) = 0 Then
            IsWildcard = False
            Exit Function
        End If
    End If
End Function


Private Function FormatBatch(sInput As String, sMask As String) As String
    Dim sTemp As String
    
    sTemp = sMask

    sTemp = Left$(sTemp, Len(sTemp) - Len(sInput)) & sInput

    FormatBatch = sTemp
End Function


Private Sub cmdLoad_Click()
    On Error GoTo ErrorHandler
    
    Dim sSQL As String
    Dim sBatchID As String
    Dim lstItm As ListItem
    Dim subItm As ListSubItem
    Dim dAmt As Double
    
    If (Not optAP) And (Not optPO2) Then
         msg "Please select one of the batch types."
         Exit Sub
    End If
    
    If Len(Trim(txtBatch.text)) = 0 Then
        TryToSetFocus txtBatch
        Exit Sub
    End If
        
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    
    SetWaitCursor True
    If optAP Then
        sBatchID = FormatBatch(txtBatch.text, "APVO-0000000")
        txtBatch.text = sBatchID
    
        sSQL = "SELECT tapBatch.BatchCmnt, tapPendVoucher.TranNo, tapPendVoucher.TranAmt " & _
            "FROM tapBatch INNER JOIN tapPendVoucher ON " & _
            "tapBatch.BatchKey = tapPendVoucher.BatchKey INNER JOIN " & _
            "tciBatchLog ON tapBatch.BatchKey = tciBatchlog.BatchKey " & _
            "WHERE tciBatchlog.BatchID = '" & sBatchID & "' Order by tapPendVoucher.SeqNo"
    End If
    
    
    If optPO2 Then
        sBatchID = FormatBatch(txtBatch.text, "POVO-0000000")
        txtBatch.text = sBatchID
    
        sSQL = "SELECT tpoBatch.BatchCmnt, tapPendVoucher.TranNo, tapPendVoucher.TranAmt " & _
            "FROM tpoBatch INNER JOIN tapPendVoucher ON " & _
            "tpoBatch.BatchKey = tapPendVoucher.BatchKey INNER JOIN " & _
            "tciBatchLog ON tpoBatch.BatchKey = tciBatchlog.BatchKey " & _
            "WHERE tciBatchlog.BatchID = '" & sBatchID & "' Order by tapPendVoucher.SeqNo"
    End If
    
    m_sBatchCmnt = ""
    lvwBatch.ListItems.Clear
    
    With rst
        .Open sSQL, g_DB.Connection, adOpenDynamic, adLockReadOnly
        
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
        MsgBox "No items on this batch."
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
    
'   This code prints the batch
    For i = 1 To lvwBatch.ListItems.Count
        Set lstItm = lvwBatch.ListItems(i)
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
    Printer.CurrentX = 1200 - TextWidth(Format$(Val(Replace(lblBatchTotal.caption, ",", "")), g_DisplayMask))

    Printer.Print Format$(Val(Replace(lblBatchTotal.caption, ",", "")), g_DisplayMask)
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



Private Sub gdxDSOrders_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowType = jgexRowTypeRecord Then
        If RowBuffer(25) = 1 Then 'Check the value of the CreateShipment flag coming from the DB
            RowBuffer.RowStyle = "Non-AP"
        ElseIf RowBuffer(25) = 2 Then
            RowBuffer.RowStyle = "AP"
        End If
        

        'TODO: SET THE CELL AS READ ONLY
'        If RowBuffer(1) = 1 Then 'Check if the invoice flag comes in as true
'            RowBuffer.CellStyle (1)
'        End If
        
    End If
End Sub

'*******************************************************************************
' Turbo TenKey tab
'*******************************************************************************

Private Sub lvwTape_DblClick()
    Dim lstItm As ListItem
    Dim subItm As ListSubItem
    
    Set lstItm = lvwTape.HitTest(m_sngX, m_sngY)
    If Not lstItm Is Nothing Then
        'if multiselected, clear all selections to reduce user ambiguity
        ClearSelections
        lvwTape.Refresh
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

Private Sub cmdClear_Click()
    lvwTape.ListItems.Clear
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


Private Sub UpdateVendorInfo(i_lVendKey As Long, i_lItemKey As Long, i_dNewCost As Double, i_sNewVPN As String)
    Dim sSQL As String
    Dim cmd As ADODB.Command
    
    sSQL = "UPDATE timVendItem SET RplcmntUnitCost = " & CStr(i_dNewCost) & " , VendItemID= '" & PrepSQLText(i_sNewVPN) & "' " & _
        "WHERE VendKey = " & CStr(i_lVendKey) & " AND ItemKey = " & CStr(i_lItemKey)
        
    
    Set cmd = CreateCommandSP(sSQL, adCmdText)
    cmd.Execute
    Set cmd = Nothing
End Sub

Private Sub txtBatch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtBatch.text)) > 0 And KeyCode = vbKeyReturn Then
        cmdLoad_Click
    End If
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - (Asc("a") - Asc("A"))
    End If
End Sub

Private Sub txtPO_LostFocus()
    If Len(txtPO.text) > 0 Then
        '10/14/2002     TeddyX
        'Add PrepSQLText to guard if user inputs single quote
        LoadPO 0, PrepSQLText(Trim(txtPO.text))
    End If
End Sub

'****************************************************************************
' Confirm Drop Ship Info
'****************************************************************************

Private Sub txtTenKeyPONbr_Change()
    If Len(txtTenKeyPONbr.text) > 0 Then
        If Not IsNumeric(txtTenKeyPONbr.text) Then
            msg "This must be a number.", vbCritical
            txtTenKeyPONbr.text = ""
        End If
    End If
End Sub

Private Sub cmdDisplay_Click()
    Dim orst As ADODB.Recordset
    
    If Len(txtTenKeyPONbr.text) = 0 Then Exit Sub
    
    'Guard if the txtTenKeyPONbr.text is numeric
    If Not IsNumeric(txtTenKeyPONbr.text) Then
        msg "This must be a number.", vbCritical
        Exit Sub
    End If
    
    Set orst = New ADODB.Recordset
    
    orst.Open "SELECT DISTINCT tcpSO.OPKey FROM tpoPurchOrder" _
        & " INNER JOIN tpoPOLine ON tpoPOLine.POKey = tpoPurchOrder.POKey" _
        & " INNER JOIN tsoSOLine ON tpoPOLine.POLineKey = tsoSOLine.POLineKey" _
        & " INNER JOIN tsoSalesOrder ON tsoSOLine.SOKey = tsoSalesOrder.SOKey" _
        & " INNER JOIN tcpSO ON tsoSalesOrder.SOKey = tcpSO.SOKey" _
        & " WHERE tpoPurchOrder.tranno LIKE '%" & CLng(txtTenKeyPONbr.text) & "'", g_DB.Connection
    If Not orst.EOF Then
        EditRemarks orst.Fields("OPKey")
    Else
        msg "DropShip PO not found.", vbInformation
    End If
End Sub





'************************************************************************************
'   Vendor Maint tab
'************************************************************************************

Private Sub cboMPKBuyer_Click()
    If m_bLoading Then Exit Sub
    
    m_bLoading = True
    If cboVendors.ListIndex <> 0 Then
        cmdSaveBuyers.Enabled = True
    End If
    m_bLoading = False
End Sub

Private Sub cboSEABuyer_Click()
    If m_bLoading Then Exit Sub
    
    m_bLoading = True
    If cboVendors.ListIndex <> 0 Then
        cmdSaveBuyers.Enabled = True
    End If
    m_bLoading = False
End Sub

Private Sub cboSTLBuyer_Click()
    If m_bLoading Then Exit Sub
    
    m_bLoading = True
    If cboVendors.ListIndex <> 0 Then
        cmdSaveBuyers.Enabled = True
    End If
    m_bLoading = False
End Sub

Private Sub cboVendors_Click()
    With cboVendors
    If m_bLoading Then Exit Sub
    
    LoadBuyers
    End With
End Sub


Private Sub LoadBuyers()
    Dim sSQL As String
    Dim orst As ADODB.Recordset
    Dim bSaveFlag As Boolean

    'if <none> selected reset interface and exit
    
    bSaveFlag = m_bLoading
    m_bLoading = True
    
    If cboVendors.ListIndex = 0 Then
        cboMPKBuyer.ListIndex = 0
        cboSEABuyer.ListIndex = 0
        cboSTLBuyer.ListIndex = 0
        cmdSaveBuyers.Enabled = False
        m_bLoading = bSaveFlag
        Exit Sub
    End If

    'get vendor info
    sSQL = "SELECT MatchToleranceKey FROM tapVendor WHERE VendKey=" & cboVendors.ItemData(cboVendors.ListIndex)
    Set orst = LoadRst(sSQL)
    
    m_lVendKey = cboVendors.ItemData(cboVendors.ListIndex)
    m_vMatchToleranceKey = orst.Fields("MatchToleranceKey").value
    
    m_bNewVendor = False
    
    sSQL = "SELECT WhseKey, BuyerKey FROM tcpWhseVendBuyer WHERE VendKey=" & m_lVendKey
    Set orst = LoadRst(sSQL)

    With orst
        If .EOF Then
            m_bNewVendor = True
            cboMPKBuyer.ListIndex = 0
            cboSEABuyer.ListIndex = 0
            cboSTLBuyer.ListIndex = 0
        Else
            Do While Not .EOF
                'I know I really shouldn't use key values here but...
                Select Case .Fields("WhseKey").value
                    Case g_MPKWhseKey
                        SetComboByKey cboMPKBuyer, .Fields("BuyerKey").value, vWithNone:=True
                    Case g_SEAWhseKey
                        SetComboByKey cboSEABuyer, .Fields("BuyerKey").value, vWithNone:=True
                    Case g_STLWhseKey
                        SetComboByKey cboSTLBuyer, .Fields("BuyerKey").value, vWithNone:=True
                End Select
                .MoveNext
            Loop
        End If
    End With
    m_bLoading = bSaveFlag
    
    cmdSaveBuyers.Enabled = False
End Sub


Private Sub cmdSaveBuyers_Click()
    On Error GoTo ErrorHandler

    Dim ocmd As ADODB.Command
    Dim lMPKBuyer As Long
    Dim lSTLBuyer As Long
    Dim lSEABuyer As Long
    
    m_bLoading = True
    
    Set ocmd = New ADODB.Command
    
    ocmd.ActiveConnection = g_DB.Connection
    ocmd.CommandType = adCmdStoredProc
    
    lMPKBuyer = cboMPKBuyer.ItemData(cboMPKBuyer.ListIndex)
    lSEABuyer = cboSEABuyer.ItemData(cboSEABuyer.ListIndex)
    lSTLBuyer = cboSTLBuyer.ItemData(cboSTLBuyer.ListIndex)
    
    If m_bNewVendor Then
        With ocmd
            .CommandText = "spCPCInsertWhseVendBuyer"
            If lMPKBuyer > 0 Then .Parameters("@_iMPKBuyerKey") = lMPKBuyer
            If lSEABuyer > 0 Then .Parameters("@_iSEABuyerKey") = lSEABuyer
            If lSTLBuyer > 0 Then .Parameters("@_iSTLBuyerKey") = lSTLBuyer
            .Parameters("@_iVendKey") = m_lVendKey
            'PRN#96
            .Execute
        End With
        
        'After inserting, the flag comes to be false
        m_bNewVendor = False
        msg "Done."
    Else
        'prompt about updating
        If vbYes = msg("This vendor already has buyers assigned. Update them?", vbQuestion + vbYesNo) Then
            With ocmd
                .CommandText = "spCPCUpdateWhseVendBuyer"
                If lMPKBuyer > 0 Then .Parameters("@_iMPKBuyerKey") = lMPKBuyer
                If lSEABuyer > 0 Then .Parameters("@_iSEABuyerKey") = lSEABuyer
                If lSTLBuyer > 0 Then .Parameters("@_iSTLBuyerKey") = lSTLBuyer
                .Parameters("@_iVendKey") = m_lVendKey
                'PRN#96
                .Execute
            End With
            
            msg "Done."
        End If
    End If
    
'*** NOTE: This is a HACK.
' While we're at it, check the Vendor's MatchToleranceKey.
' A Null value chokes the Sage POAPI when POWizard commits POs.
' If Null, assign it the magic number 13.
    If IsNull(m_vMatchToleranceKey) Then
        ocmd.CommandType = adCmdText
        ocmd.CommandText = "UPDATE tapVendor SET MatchToleranceKey=13 WHERE VendKey=" & m_lVendKey
        ocmd.Execute
    End If
    m_bLoading = False
    cmdSaveBuyers.Enabled = False
    cboVendors.SetFocus
    Exit Sub
    
ErrorHandler:
    m_bLoading = False
    msg Err.Number & " - " & Err.Description, vbOKOnly + vbExclamation, Err.Source
End Sub


Private Sub cmdGet_Click()
    Dim sSQL As String
    Dim rst As ADODB.Recordset
    
    If m_bLoading Then Exit Sub
    
    SetWaitCursor True
    m_bLoading = True
    
    sSQL = "Select * from tcpPrefVendor where MakeID = " & cboMake.ItemData(cboMake.ListIndex) & _
        " and WhseKey = " & cboWhse(0).ItemData(cboWhse(0).ListIndex)
    
    Set rst = LoadDiscRst(sSQL)
    
    m_bPrefVendor = (rst.RecordCount > 0)
    
    If m_bPrefVendor Then
        rst.MoveFirst
        SetComboByKey cboVendor, rst.Fields("PrefVendKey").value, True
    Else
        cboVendor.ListIndex = 0
    End If
    
    DisplayVendorInfo cboVendor.ItemData(cboVendor.ListIndex)
    
    m_bLoading = False
    SetWaitCursor False
End Sub


Private Sub cboMake_Click()
    
    If m_bLoading Then Exit Sub
    
    m_bLoading = True
    
    If cboMake.ItemData(cboMake.ListIndex) = 1 Then
        cmdGet.Enabled = False
        cmdPrefUpdate.Enabled = False
        cboVendor.ListIndex = 0
        m_bPrefVendor = False
    Else
        cmdGet.Enabled = True
        cmdPrefUpdate.Enabled = True
    End If
    m_bLoading = False
End Sub


Private Sub cboVendor_Click()
    If m_bVendorLoad Then Exit Sub
    
    SetWaitCursor True
    m_bVendorLoad = True
    DisplayVendorInfo cboVendor.ItemData(cboVendor.ListIndex)
    m_bVendorLoad = False
    SetWaitCursor False
End Sub


Private Sub cmdPrefUpdate_Click()
    m_bVendorLoad = True
    
    If m_bPrefVendor Then
        If vbYes = msg("Are you sure that you want to update preferred vendor for " _
                & cboMake.list(cboMake.ListIndex) & " in " & cboWhse(0).text & "?", vbExclamation + vbYesNo, "Update Preferred Vendor") Then
            DeletePrefVendor
            If cboVendor.ItemData(cboVendor.ListIndex) > 0 Then
                InsertPrefVendor
            End If
            msg "Preferred vendor has been updated successfully!"
        End If
    Else
        If cboVendor.ItemData(cboVendor.ListIndex) = 0 Then
            msg "Please choose preferred vendor first.", vbExclamation + vbOKOnly, "Choose Vendor"
            TryToSetFocus cboVendor
        Else
            If vbYes = msg("Are you sure that you want to insert preferred vendor for " _
                    & cboMake.list(cboMake.ListIndex) & " in " & cboWhse(0).text & "?", vbExclamation + vbYesNo, "Insert Preferred Vendor") Then
                InsertPrefVendor
                msg "Preferred vendor has been inserted successfully!"
            End If
        End If
    End If
        
    m_bVendorLoad = False
End Sub


Private Sub DisplayVendorInfo(ByVal iVendorKey As Long)
    If iVendorKey = 0 Then
        lblVendName.caption = ""
        lblAddress.caption = ""
        lblVendContact.caption = ""
        lblVendPhone.caption = ""
        lblVendFax.caption = ""
        lblVendEMail.caption = ""
        rvVendor.OwnerID = ""
        rvVendor.Visible = False
        Exit Sub
    End If
    
    Dim orst As ADODB.Recordset
    Set orst = CallSP("spcpcGetVendInfo", "@_iVendorKey", iVendorKey)
    
    If orst.RecordCount > 0 Then
        With orst
            lblVendName.caption = .Fields("VendName").value
            lblAddress.caption = CompAddr(.Fields("AddrName").value, _
                .Fields("AddrLine1").value, _
                .Fields("AddrLine2").value, _
                .Fields("City").value, _
                .Fields("StateID").value, _
                .Fields("PostalCode").value, _
                .Fields("CountryID").value)
            lblVendContact.caption = .Fields("Name").value
            lblVendPhone.caption = FormatPhoneNumber(.Fields("Phone").value, .Fields("PhoneExt").value)
            lblVendFax.caption = FormatPhoneNumber(.Fields("Fax").value, .Fields("FaxExt").value)
            lblVendEMail.caption = .Fields("EmailAddr").value
            rvVendor.Visible = True
            rvVendor.OwnerID = ""
            rvVendor.OwnerID = .Fields("VendID").value
        End With
    End If
End Sub


Private Sub DeletePrefVendor()
    Dim cmd As ADODB.Command
    Dim sSQL As String
    Dim lMakeID As Long
    
    SetWaitCursor True
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = g_DB.Connection
    cmd.CommandType = adCmdText
    
    lMakeID = cboMake.ItemData(cboMake.ListIndex)
    sSQL = "Delete from tcpPrefVendor where MakeID = " & lMakeID & " and WhseKey = " & cboWhse(0).ItemData(cboWhse(0).ListIndex)
    cmd.CommandText = sSQL
    cmd.Execute
    Set cmd = Nothing
    
    SetWaitCursor False
End Sub


Private Sub InsertPrefVendor()
    Dim cmd As ADODB.Command
    Dim sSQL As String
    Dim lMakeID As Long
    Dim lPrefVendKey As Long
    
    SetWaitCursor True
    lMakeID = cboMake.ItemData(cboMake.ListIndex)
    lPrefVendKey = cboVendor.ItemData(cboVendor.ListIndex)
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = g_DB.Connection
    cmd.CommandType = adCmdText
    
    sSQL = "INSERT tcpPrefVendor (MakeID, PrefVendKey, WhseKey) VALUES (" & lMakeID & "," & lPrefVendKey & "," & cboWhse(0).ItemData(cboWhse(0).ListIndex) & ")"
    With cmd
        .CommandText = sSQL
        .Execute
    End With
    
    Set cmd = Nothing
    SetWaitCursor False
End Sub


'*************************
'   Vendor Invoice Tab
'*************************
Private Sub BuildPOTV(POKey As Long)
'smr - 11/19/2004
'Private Sub BuildPOTV(POKey As Integer)
    Dim orst As ADODB.Recordset
    Dim sHoldTrx As String
    Dim iRcvdInvCnt As Integer
    Dim sSQL As String
    Dim aNode As Node
    
    Set orst = LoadDiscRst("Exec spCPCAPGetPOTrxInfo @i_POKey = " & CStr(POKey))
    
    With tvwPO.Nodes
        .Clear
        .Add , , "R", "Purchase Order " & Trim(txtPO.text) & " - " & lblVIVendorName.caption
        .Add "R", tvwChild, "PO", "Purchase Order Lines"
        .Add "R", tvwChild, "RI", "Received Vendor Invoices"
        .Add "R", tvwChild, "EV", "Existing Vouchers"
        .Add "R", tvwChild, "CR", "Completed Receivers"
        .Add "R", tvwChild, "OR", "Open Receivers"
        .Add "R", tvwChild, "PR", "Pending Receivers"
    End With
    
    
    'Process Existing Vouchers
    sHoldTrx = ""
    With orst
    .Filter = "VouchNo <> Null"
    .Sort = "VouchNo Desc"
    
    Do While Not .EOF
        If sHoldTrx <> .Fields("VouchNo").value Then
            sHoldTrx = .Fields("VouchNo").value
            'Add a Voucher Header Node
            tvwPO.Nodes.Add "EV", tvwChild, "VN" & .Fields("VouchNo").value, _
                "Voucher: " & StripLeadingZeros(.Fields("VouchNo").value) & " " & _
                Format$(.Fields("VendInvDate").value, "MM/DD/YY") & " " & _
                "InvNo: " & .Fields("VendInvNo").value & " " & _
                FormatMoney(.Fields("VendInvAmt").value)
        End If

        'Add a Voucher Detail Node
        tvwPO.Nodes.Add "VN" & .Fields("VouchNo").value, tvwChild, "VLDK" & CStr(.Fields("VoucherLineDistKey").value), _
            "(" & Format$(.Fields("VoucherQty").value, "@@@@") & ") " & _
            FormatMoney(.Fields("UnitCost").value) & " " & _
            Trim(.Fields("ItemID").value) & "  " & _
            .Fields("Description").value

        .MoveNext
    Loop
    
    'Process Closed Receivers
    sHoldTrx = ""
    .Sort = "RecID Desc"
    If Not .EOF Then
        .MoveFirst
    End If
    
    Do While Not .EOF
        If sHoldTrx <> .Fields("RecID").value Then
            sHoldTrx = .Fields("RecID").value
            'Add a Receiver Header Node
            tvwPO.Nodes.Add "CR", tvwChild, "CRN" & .Fields("RecID").value, _
                "Receiver : " & StripLeadingZeros(.Fields("RecID").value) & " " & _
                Format$(.Fields("RecDate").value, "MM/DD/YY") & " " & _
                .Fields("WhseID").value
        End If
        
        'Add a Receiver Line
        tvwPO.Nodes.Add "CRN" & .Fields("RecID").value, tvwChild, "CRL" & CStr(.Fields("RcvrLineDistKey").value) & .Fields("VouchNo").value, _
            "(" & Format$(.Fields("QtyRcvd").value, "@@@@") & ") " & _
            "Vchr: " & .Fields("VouchNo").value & " " & _
            Left$(.Fields("ItemID").value, 16) & "  " & _
            .Fields("Description").value
    
        .MoveNext
    Loop
    
    
    'Process Open Receivers
    sHoldTrx = ""
    .Filter = "VouchNo = Null and PendWhseID  =  'None'"
    .Sort = "RecID Desc"
    Do While Not .EOF
        If sHoldTrx <> .Fields("RecID").value Then
            sHoldTrx = .Fields("RecID").value
            'Add a Receiver Header Node
            tvwPO.Nodes.Add "OR", tvwChild, "ORN" & .Fields("RecID").value, _
                "Receiver: " & StripLeadingZeros(.Fields("RecID").value) & " " & _
                Format$(.Fields("RecDate").value, "MM/DD/YY") & " " & _
                .Fields("WhseID").value
        End If
        
        'Add a Receiver Line
        tvwPO.Nodes.Add "ORN" & .Fields("RecID").value, tvwChild, "CRL" & CStr(.Fields("RcvrLineDistKey").value), _
            "(" & Format$(.Fields("QtyRcvd").value, "@@@@") & ") " & _
            Trim(.Fields("ItemID").value) & "  " & _
            .Fields("Description").value
    
        .MoveNext
    Loop
    

    'Process Pending Receivers
    sHoldTrx = ""
    .Filter = "VouchNo = Null and PendWhseID <> 'None'"
    .Sort = "RecID Desc"
    Do While Not .EOF
        If sHoldTrx <> .Fields("RecID").value Then
            sHoldTrx = .Fields("RecID").value
            'Add a Receiver Header Node
            tvwPO.Nodes.Add "PR", tvwChild, "PRN" & .Fields("RecID").value, _
                "Receiver: " & .Fields("RecID").value & " " & _
                Format$(.Fields("RecDate").value, "MM/DD/YY") & " " & _
                .Fields("PendWhseID").value
        End If
        
        'Add a Receiver Line
        tvwPO.Nodes.Add "PRN" & .Fields("RecID").value, tvwChild, "CRL" & CStr(.Fields("RcvrLineDistKey").value), _
            "(" & Format$(.Fields("QtyRcvd").value, "####") & ") " & _
            Left$(.Fields("ItemID").value, 16) & "  " & _
            .Fields("Description").value
    
        .MoveNext
    Loop
    
    End With
    
    tvwPO.Nodes("R").Expanded = True
    tvwPO.Nodes("OR").Expanded = True
    tvwPO.Nodes("PR").Expanded = True
    

    CloseRst orst
    Set orst = LoadDiscRst("Select * from tcpAPRcvdInv where POKey = " & CStr(POKey))
    
    
    'Process Received Invoices
    With orst

    Do While Not .EOF
        iRcvdInvCnt = iRcvdInvCnt + 1   'Prevent duplicate rcvd invoice from blowing up TV
        'Add a Received Invoice Node
        tvwPO.Nodes.Add "RI", tvwChild, "VN" & CStr(iRcvdInvCnt) & .Fields("VendInvNbr").value, _
            .Fields("VendInvNbr").value & "    " & _
            Format$(.Fields("InvoiceDate").value, "MM/DD/YY") & "  (" & _
            Trim(.Fields("UserID").value) & " on " & _
            Format$(.Fields("DateEntered").value, "MM/DD/YY") & ") "
        .MoveNext
    Loop
    End With
    CloseRst orst
    
    'PO Lines
    sSQL = "SELECT dbo.tpoPOLineDist.QtyOrd, dbo.timItem.ItemID, dbo.tpoPOLine.Description, dbo.tpoPOLine.UnitCost " & _
                "FROM dbo.tpoPOLine INNER JOIN dbo.tpoPOLineDist ON dbo.tpoPOLine.POLineKey = dbo.tpoPOLineDist.POLineKey INNER JOIN dbo.timItem ON dbo.tpoPOLine.ItemKey = dbo.timItem.ItemKey " & _
                "WHERE dbo.tpoPOLine.POKey = " & CStr(POKey)
    Set orst = LoadDiscRst(sSQL)
    With orst
        Do While Not .EOF
            iRcvdInvCnt = iRcvdInvCnt + 1
            tvwPO.Nodes.Add "PO", tvwChild, "POL" & CStr(iRcvdInvCnt) & .Fields("ItemID").value, _
            "(" & Format$(.Fields("QtyOrd").value, "@@@@") & ") " & _
            FormatMoney(CStr(.Fields("UnitCost").value)) & "  " & _
            Left$(.Fields("ItemID").value, 16) & "    " & .Fields("Description").value
        .MoveNext
        Loop
    End With
    CloseRst orst
    
    tvwPO.Nodes("RI").Expanded = True
End Sub

Private Sub LoadPO(POKey As Long, Optional PONbr As String)
'smr - 11/19/2004
'Private Sub LoadPO(POKey As Integer, Optional PONbr As String)
    Dim sSQL As String
    Dim orst As ADODB.Recordset
    
    sSQL = "SELECT dbo.tpoPurchOrder.POKey, dbo.tpoPurchOrder.TranID, dbo.tpoPurchOrder.TranDate, dbo.tapVendor.VendID, dbo.tapVendor.VendName " & _
            "FROM dbo.tapVendor INNER JOIN dbo.tpoPurchOrder ON dbo.tapVendor.VendKey = dbo.tpoPurchOrder.VendKey " & _
            "WHERE (dbo.tpoPurchOrder.Status = 1) AND (dbo.tpoPurchOrder.CompanyID = 'CPC')"
            
    If POKey > 0 Then
        sSQL = sSQL & " AND dbo.tpoPurchOrder.POKey = " & CStr(POKey)
    Else
        sSQL = sSQL & " AND dbo.tpoPurchOrder.TranID Like '%" & Trim(PONbr) & "%'"
    End If
    
    Set orst = LoadDiscRst(sSQL)
    
    With orst
        If Not .EOF Then
            txtPO.text = .Fields("TranID").value
            lblVIVendorName.caption = Trim(.Fields("VendID").value) & " - " & Trim(.Fields("VendName").value)
            lblVIPODate.caption = Format$(.Fields("TranDate").value, "MM/DD/YYYY")
            BuildPOTV .Fields("POKey").value
            m_iPOKey = .Fields("POKey").value
        Else
            MsgBox "Sorry, cannot find this open PO."
            lblVIVendorName.caption = ""
            lblVIPODate.caption = ""
            tvwPO.Nodes.Clear
            '01/03/2003     TeddyX
            'Use TryToSetFocus txtPO to replace txtPO.setFocus
            TryToSetFocus txtPO
        End If
        
    End With
    
    CloseRst orst
    
End Sub

Private Sub BuildRecTV()
    Dim orst As ADODB.Recordset
    Dim sHoldPO As String
    Dim sHoldInv As String
    Dim sHoldRec As String
    Dim sHoldRecLine As String
    Dim iRcvdInvCnt As Integer
    
    Set orst = LoadDiscRst("Exec spCPCAPGetOpenRcvrs")
    
    With tvwPO.Nodes
        .Clear
        .Add , , "R", "Open Receivers"
    End With
    

    sHoldPO = ""
    sHoldInv = ""
    
    
    With orst
    
    Do While Not .EOF
        'If this is a new PO...
        If sHoldPO <> .Fields("TranID").value Then
            sHoldPO = .Fields("TranID").value
            'Add a PO Header Node
            tvwPO.Nodes.Add "R", tvwChild, .Fields("TranID").value, _
                "PO: " & .Fields("TranID").value & " (" & _
                Format$(.Fields("TranDate").value, "MM/DD/YY") & ") " & _
                Trim(.Fields("VendID").value) & " - " & Trim(.Fields("VendName").value)
            'add sub nodes for each PO
            tvwPO.Nodes.Add .Fields("TranID").value, tvwChild, "PR" & .Fields("TranID").value, "Pending Receivers"
            tvwPO.Nodes.Add .Fields("TranID").value, tvwChild, "OR" & .Fields("TranID").value, "Open Receivers"
            tvwPO.Nodes.Add .Fields("TranID").value, tvwChild, "RI" & .Fields("TranID").value, "Received Invoices"
        End If
        
        'if this is a new Vendor Invoice, add it
        If sHoldInv <> .Fields("VendInvNbr").value Then
            sHoldInv = .Fields("VendInvNbr").value
            'Add a a received vendor invoice
            tvwPO.Nodes.Add "RI" & .Fields("TranID").value, tvwChild, "VI" & .Fields("TranID").value & .Fields("VendInvNbr").value, _
            .Fields("VendInvNbr").value & "    " & _
            Format$(.Fields("InvoiceDate").value, "MM/DD/YY") & "  (" & _
            Trim(.Fields("UserID").value) & " on " & _
            Format$(.Fields("DateEntered").value, "MM/DD/YY") & ") "
        End If
        .MoveNext
    Loop
    
    
    'Process Open Receivers
    sHoldRec = ""
    sHoldRecLine = ""
    .Filter = "PendWhseID  =  'None'"
    .Sort = "RecID Desc"
    Do While Not .EOF
        If sHoldRec <> .Fields("RecID").value Then
            sHoldRec = .Fields("RecID").value
            'Add a Receiver Header Node
            tvwPO.Nodes.Add "OR" & .Fields("TranID").value, tvwChild, "ORN" & .Fields("RecID").value, _
                "Receiver: " & StripLeadingZeros(.Fields("RecID").value) & " " & _
                Format$(.Fields("RecDate").value, "MM/DD/YY") & " " & _
                .Fields("WhseID").value
        End If
        
        'Add a Receiver Line
        If sHoldRecLine <> .Fields("RcvrLineDistKey").value Then
            sHoldRecLine = .Fields("RcvrLineDistKey")
            iRcvdInvCnt = iRcvdInvCnt + 1
            tvwPO.Nodes.Add "ORN" & .Fields("RecID").value, tvwChild, "CRL" & CStr(.Fields("RcvrLineDistKey").value) & CStr(iRcvdInvCnt), _
                "(" & Format$(.Fields("QtyRcvd").value, "####") & ") " & _
                Trim(.Fields("ItemID").value) & "  " & _
                .Fields("Description").value
        End If
        .MoveNext
    Loop
    

    'Process Pending Receivers
    sHoldRec = ""
    sHoldRecLine = ""
    .Filter = "PendWhseID <> 'None'"
    .Sort = "RecID Desc"
    Do While Not .EOF
        If sHoldRec <> .Fields("RecID").value Then
            sHoldRec = .Fields("RecID").value
            'Add a Receiver Header Node
            tvwPO.Nodes.Add "PR" & .Fields("TranID").value, tvwChild, "PRN" & .Fields("RecID").value, _
                "Receiver: " & .Fields("RecID").value & " " & _
                Format$(.Fields("RecDate").value, "MM/DD/YY") & " " & _
                .Fields("PendWhseID").value
        End If
        
        'Add a Receiver Line
        If sHoldRecLine <> .Fields("RcvrLineDistKey").value Then
            sHoldRecLine = .Fields("RcvrLineDistKey")
            iRcvdInvCnt = iRcvdInvCnt + 1
            tvwPO.Nodes.Add "PRN" & .Fields("RecID").value, tvwChild, "CRL" & CStr(.Fields("RcvrLineDistKey").value) & CStr(iRcvdInvCnt), _
                "(" & Format$(.Fields("QtyRcvd").value, "####") & ") " & _
                Trim(.Fields("ItemID").value) & "  " & _
                .Fields("Description").value
        End If
        .MoveNext
    Loop
    
    End With
    
    CloseRst orst
    tvwPO.Nodes("R").Expanded = True
    
    
End Sub

Private Function FormatMoney(sMoneyString As String) As String
    Dim sDollars As String
    Dim sCents As String
    Dim iDecimalPos As Integer
    
    iDecimalPos = InStr(sMoneyString, ".")
    If iDecimalPos = 0 Then
        FormatMoney = Format(sMoneyString, "@@@@@") + ".00"
        Exit Function
    End If
    sDollars = Left$(sMoneyString, iDecimalPos - 1)
    sCents = Mid$(sMoneyString, iDecimalPos + 1, 99)
    FormatMoney = Format(sDollars, "@@@@@") + Format$("." + sCents, ".00")
End Function


'****************************************************************************
' Drop Ship Tab
'****************************************************************************

Private Sub cmdPrint_Click(Index As Integer)
    Dim currentPSKey As String
    Dim lastPSKey As String
    Dim shipments As String
    
    'VL to get around GridEX limitation of updateing last row
    'This way any changes are immediately applied to the current row
    'http://stackoverflow.com/questions/726965/janus-gridex-problem
    If gdxDSOrders.EditMode = jgexEditModeOn Then gdxDSOrders.Update
    
    Dim i As Integer
    For i = 0 To m_iDSOrderCount - 1
        If CBool(m_arrayDSOrders(0, i)) Then 'Only process the rows with invoice checked
            currentPSKey = m_arrayDSOrders(15, i)

            If Not currentPSKey = lastPSKey And i > 0 Then
                shipments = shipments & "," & currentPSKey
            Else
                shipments = currentPSKey
            End If

            lastPSKey = currentPSKey
        End If
    Next

    Debug.Print "Shipments: " & shipments
    
    If Len(shipments) > 0 Then
        Screen.MousePointer = vbHourglass
        
        Dim oFrm As FDSPreflight
        
        
        Set oFrm = New FDSPreflight
        oFrm.DataSource = m_arrayDSOrders
                        
        oFrm.Show vbModal
       
    End If
    
End Sub


Private Sub cmdProcess_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo EH
    
    'VL to get around GridEX limitation of updateing last row
    'This way any changes are immediately applied to the current row
    'http://stackoverflow.com/questions/726965/janus-gridex-problem
    If gdxDSOrders.EditMode = jgexEditModeOn Then gdxDSOrders.Update
    
    Dim pskey As Integer
    Dim invoice As Boolean
    Dim voucher As Boolean
    Dim process As Boolean
    Dim freight As Double
    Dim handling As Double
    Dim tax As Double
    Dim packing As Double
    Dim trackingNo As String
    Dim TranNo As String
    Dim i As Integer
    Dim Result As String
    
    For i = 0 To m_iDSOrderCount - 1
       
        TranNo = m_arrayDSOrders(9, i)
        pskey = m_arrayDSOrders(15, i)
        process = m_arrayDSOrders(21, i)
        
        If m_arrayDSOrders(0, i) = Null Then
            invoice = False
            Debug.Print "Null : " & invoice
        ElseIf m_arrayDSOrders(0, i) = Empty Then
            invoice = False
            Debug.Print "Empty : " & invoice
        Else
            invoice = CBool(m_arrayDSOrders(0, i))
            Debug.Print "Not Empty or Null: " & m_arrayDSOrders(0, i)
        End If
        

        If invoice = True And pskey > 0 Then
            Debug.Print "TranNo: " & TranNo; " PSKey: " & pskey
            
'            If invoice = True And pskey > 0 Then
                Debug.Print "Create Invoice for TranNo: " & TranNo & " Tracking No: " & trackingNo & " Freight: " & freight & " Handling: " & handling & " Tax: " & tax & " Packing: " & packing
                Debug.Print "Updating Shipment to show being Invoiced"
                    
                Dim sql As String
                Dim cmd As ADODB.Command
                
                sql = "update tcpProvisionalShipment " & vbCrLf & _
                    "set Invoice=1, Status=1 " & vbCrLf & _
                    "where pskey=" & pskey
                
                'Debug.Print sql
                Set cmd = CreateCommandSP(sql, adCmdText)
                cmd.Execute
                Set cmd = Nothing
                
           ' End If
            
            
'            If voucher = True And pskey > 0 Then
'                Debug.Print "Create Voucher for TranNo: " & tranNo & " Tracking No: " & trackingNo & " Freight: " & freight & " Handling: " & handling & " Tax: " & tax & " Packing: " & packing
'
'                Debug.Print "Updating Shipment to show being Vouchered"
'
'                Dim sql2 As String
'                Dim cmd2 As ADODB.Command
'
'                sql2 = "update tcpProvisionalShipment " & vbCrLf & _
'                    "set Voucher=1 " & vbCrLf & _
'                    "where pskey=" & pskey
'
'                Debug.Print sql2
'                Set cmd2 = CreateCommandSP(sql2, adCmdText)
'                cmd2.Execute
'                Set cmd2 = Nothing
'            End If
                        
            'proxy.ReleaseOrder tranNo, trackingNo, freight, voucherKey
            
            
        End If
    
        
    Next
    
    RefreshDropShipList
        
    Screen.MousePointer = vbDefault
    
    Exit Sub
EH:
    
    msg "* Runtime error" & vbCrLf & _
            vbTab & "Err Num: " & vbTab & Err.Number & vbCrLf & _
            vbTab & "Err Source: " & vbTab & Err.Source & vbCrLf & _
            vbTab & "Err Desc: " & vbTab & Err.Description & vbCrLf & _
            vbCrLf & _
            "* Source: " & "SageAssistant.Billing.cmdProcess_Click()" & vbCrLf & _
            vbCrLf & _
            "*Timestamp: " & Now & " " & Timer
   
    SetWaitCursor False
End Sub





'Private Sub PrintDropShipReport()
'    Dim remarks As String
'    Dim s As String
'
'    SetWaitCursor True
'
'    Printer.Font = "Arial"
'    Printer.FontSize = 14
'    Printer.Print "Open Drop Ship Order Report" & "     " & Date
'    Printer.Print
'    Printer.FontSize = 10
'    With m_oRstDSOrders
'        .MoveFirst
'        Do While Not .EOF
'            remarks = FetchRemarks(.Fields("OPKey"))
'            If Len(remarks) > 0 Then
'                s = .Fields("SOID") & ", " & .Fields("SODate") & ", "
'                s = s & Trim(.Fields("CustID")) & " " & Trim(.Fields("CustName")) & vbCrLf
'                s = s & .Fields("POID") & ", " & .Fields("PODate") & ", "
'                s = s & Trim(.Fields("VendID")) & " " & Trim(.Fields("VendName")) & " " & FormatPhoneNumber(.Fields("Phone"), vbNullString) & vbCrLf
'                s = s & remarks
'                Printer.Print s
'                Printer.Print
'            End If
'            .MoveNext
'        Loop
'    End With
'    Printer.EndDoc
'
'    SetWaitCursor False
'    msg "Report printed on " & Printer.DeviceName
'End Sub





Private Sub cmdSave_Click()
'    Dim oRC As RemarkContext
'    Dim s As String
'
'    'build the remark string
'    If chkShipComp.value = vbChecked Then
'        s = s & "Shipped Complete"
'    End If
'    If Len(txtFreight) > 0 Then
'        If Len(s) > 0 Then s = s & "; "
'        s = s & "Freight = " & txtFreight
'    End If
'    If Len(txtTrackNo) > 0 Then
'        If Len(s) > 0 Then s = s & "; "
'        s = s & txtTrackNo
'    End If
'
'    Set oRC = New RemarkContext
'    oRC.Load "DropShipUpdate", m_gwDSOrders.value("OPKey")
'    oRC.AddRemark "Order.DropShip", s
'    oRC.Save True
'    Set oRC = Nothing
'
'    'Write another remark and attach it to the PO - JEJ 10/11/04
'    Set oRC = New RemarkContext
'    oRC.Load "ViewPO", m_gwDSOrders.value("POID")
'    oRC.AddRemark "PO.Status", s
'    oRC.Save True
'
'    Set oRC = Nothing


    'Dim proxy As MSSOAPLib30.SoapClient30
    'Dim result As String
    
    
    'If Len(txtFreight) > 0 Then
    '    Screen.MousePointer = vbHourglass
        
    '    Set proxy = New MSSOAPLib30.SoapClient30
    '    proxy.MSSoapInit "http://10.6.4.1/AutoPickWeb/Service.asmx?WSDL"
    
     '   result = proxy.PickAndShipOrder(m_gwDSOrders.value("OPKey"), cboWhse(1).ItemData(cboWhse(1).ListIndex), CDbl(Trim(txtFreight.text)), Trim(txtTrackNo.text))
        
     '   Screen.MousePointer = vbDefault
   ' End If
    
       

    'ClearRemark
    'RefreshDropShipList
End Sub


Private Sub cmdRefresh_Click()
    ' RESET THE VALUE TO FIND POS IN THE GRID
    m_iPoIdToFind = 0
    
    RefreshDropShipList
End Sub

Private Sub oFrm_OnClose()
    RefreshDropShipList
End Sub
    
Private Sub RefreshDropShipList()
        SetWaitCursor True
        
        Set dict = New Dictionary
        m_bDictLoaded = False
                       
        If cboWhse(1).ItemData(cboWhse(1).ListIndex) = 0 Then
            Set m_oRstDSOrders = CallSP("spcpcGetDropShipsForAP")
        Else
            Set m_oRstDSOrders = CallSP("spcpcGetDropShipsForAP", "@_iWhseKey", cboWhse(1).ItemData(cboWhse(1).ListIndex))
        End If
            
        'AttachGrid gdxDSOrders, m_oRstDSOrders
            
        If m_oRstDSOrders.RecordCount > 0 Then
            m_arrayDSOrders = m_oRstDSOrders.GetRows
            m_iDSOrderCount = UBound(m_arrayDSOrders, 2) + 1

        Else
            m_arrayDSOrders = Empty
            m_iDSOrderCount = 0
            
            Set gdxSO.ADORecordset = Nothing
            Set gdxPO.ADORecordset = Nothing
        End If
       
    
        
        Dim i As Integer
        With gdxDSOrders
            .Row = -1
            .HoldFields
            .HoldSortSettings = True
            .ItemCount = m_iDSOrderCount
            .Refetch
            .Row = -1
            For i = 1 To .Columns.Count
                If .Columns(i).Key = "TrackingNo" Then
                    .Columns(i).AutoSize
                End If
            Next
        End With
        
        TryToSetFocus gdxDSOrders

        ' THIS IS TO PUT FOCUS ON THE SELECTED ROW BEFORE USER
        ' ACCESSED THE PROVISIONAL SHIPMENT SCREEN
        If (m_iPoIdToFind > 0) Then
            txtPoNumber.text = m_iPoIdToFind
            cmdfindDropShipPO_Click
            
            'RESET VALUES
            txtPoNumber.text = ""
            m_iPoIdToFind = 0
        End If

        If Not m_oRstDSOrders.EOF Then cmdPrint(1).Enabled = True
        
        Debug.Print dict.Count
        
        SetWaitCursor False
End Sub



Private Sub EditRemarks(OPKey As Long)
    Dim oRC As RemarkContext
    
    Set oRC = New RemarkContext
    oRC.Edit "DropShipUpdate", OPKey
End Sub


Private Sub RowChosen()
    'SetWaitCursor True
    
    'If IsEmpty(m_gwDSOrders.Value("OPKey")) Then Exit Sub
    Debug.Print "OpKey: " & gdxDSOrders.value(13)
    m_iPoIdToFind = gdxDSOrders.value(9)
    
    'Dim oFrm As FProvisionalShipment
    Set oFrm = New FProvisionalShipment
    oFrm.OPKey = gdxDSOrders.value(13)
    oFrm.CreateShipment Me
        
    'SetWaitCursor False
End Sub

Private Sub LoadSpecialHandling()
    Dim m_oOrder As Order
    Set m_oOrder = New Order
    m_oOrder.Load gdxDSOrders.value(13)
    
    Dim oFrm As FSpecialHandling
    
    Set oFrm = New FSpecialHandling
    oFrm.ShipMethod = m_oOrder.ShipMethod
    oFrm.Load m_oOrder
End Sub

Private Sub gdxDSOrders_SelectionChange()
    Dim RowIndex As Long
    Dim soKey As String
    Dim POKey As String
    
    With gdxDSOrders
        RowIndex = .Row
    End With

    If RowIndex < 1 Then Exit Sub
    
    soKey = m_arrayDSOrders(13, RowIndex - 1)
    POKey = m_arrayDSOrders(14, RowIndex - 1)
    
    
    If IsEmpty(soKey) Or IsEmpty(POKey) Then Exit Sub


    Set m_oRstSO = LoadDiscRst("SELECT ItemID, Description, QtyOrd, UnitPrice FROM tsoSOLine " _
                & "inner join tsoSOLineDist on tsoSOLine.SOLineKey = tsoSOLineDist.SOLineKey " _
                & "inner join timitem on tsoSOLine.ItemKey = timitem.itemkey " _
               & "WHERE SOKey=" & soKey)
   AttachGrid gdxSO, m_oRstSO

    Set m_oRstPO = LoadDiscRst("SELECT ItemID, Description, QtyOrd, UnitCost FROM tpoPOLine " _
                & "inner join timitem on tpoPOLine.ItemKey = timitem.itemkey " _
                & "inner join tpoPOLineDist on tpoPOLine.POLineKey = tpoPOLineDist.POLineKey " _
                & "WHERE POKey=" & POKey)
   AttachGrid gdxPO, m_oRstPO

    txtGridRows.text = gdxDSOrders.Row & "/" & CStr(m_iDSOrderCount)

    Set m_oRstSO = Nothing
    Set m_oRstPO = Nothing
    
   
End Sub

Private Sub gdxDSOrders_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
       
    If m_iDSOrderCount = 0 Then Exit Sub
    
    If RowIndex > m_iDSOrderCount Then
        m_bDictLoaded = True
        Exit Sub
    End If
    
    Values(1) = m_arrayDSOrders(0, RowIndex - 1)    'Invoice
    Values(2) = m_arrayDSOrders(1, RowIndex - 1)    'Voucher
    Values(3) = m_arrayDSOrders(2, RowIndex - 1)    'CreateShipment
    Values(4) = m_arrayDSOrders(3, RowIndex - 1)    'TrackingNo
    Values(5) = m_arrayDSOrders(4, RowIndex - 1)    'FreightAmt
    Values(6) = m_arrayDSOrders(5, RowIndex - 1)    'Handling
    Values(7) = m_arrayDSOrders(6, RowIndex - 1)    'Tax
    Values(8) = m_arrayDSOrders(7, RowIndex - 1)    'Packing
    Values(9) = m_arrayDSOrders(8, RowIndex - 1)    'POID
    Values(10) = m_arrayDSOrders(9, RowIndex - 1)   'SOID
    Values(11) = m_arrayDSOrders(10, RowIndex - 1)  'CustID
    Values(12) = m_arrayDSOrders(11, RowIndex - 1)  'VendID
    Values(13) = m_arrayDSOrders(12, RowIndex - 1)  'OPKey
    Values(14) = m_arrayDSOrders(13, RowIndex - 1)  'Sokey
    Values(15) = m_arrayDSOrders(14, RowIndex - 1)  'POKey
    Values(16) = m_arrayDSOrders(15, RowIndex - 1)  'PSKey
    Values(17) = m_arrayDSOrders(16, RowIndex - 1)  'Customer Name
    Values(18) = m_arrayDSOrders(17, RowIndex - 1)  'Vendor Name
    Values(19) = m_arrayDSOrders(18, RowIndex - 1)  'PODate
    Values(20) = m_arrayDSOrders(19, RowIndex - 1)  'SODate
    Values(21) = m_arrayDSOrders(20, RowIndex - 1)  'Remarks
    'Values(22) = m_arrayDSOrders(21, RowIndex - 1)  'Process
    'Values(23) = m_arrayDSOrders(22, RowIndex - 1)  'Status
    'Values(24) = m_arrayDSOrders(23, RowIndex - 1)  'StatusId
    Values(22) = m_arrayDSOrders(24, RowIndex - 1)  'VendorPays
    Values(23) = m_arrayDSOrders(25, RowIndex - 1) 'Note
    Values(24) = m_arrayDSOrders(26, RowIndex - 1) 'CreatedBy
    Values(25) = m_arrayDSOrders(27, RowIndex - 1) 'UserType
    
'    If m_bDictLoaded = False And CBool(m_arrayDSOrders(0, RowIndex - 1)) And Not dict.Exists(CInt(m_arrayDSOrders(15, RowIndex - 1))) Then
'        dict.Add CInt(m_arrayDSOrders(15, RowIndex - 1)), CStr(m_arrayDSOrders(15, RowIndex - 1))
'    End If
    
End Sub

Private Sub gdxDSOrders_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim pskey As Integer
    pskey = Values(16)
'
'    Dim prevInvoiceValue As Boolean
'    Dim prevVoucherValue As Boolean
'
'    prevInvoiceValue = m_arrayDSOrders(0, RowIndex - 1)
'    prevVoucherValue = m_arrayDSOrders(1, RowIndex - 1)
'
'    If Not dict.Exists(pskey) Then
'
'        If prevInvoiceValue = False And Values(1) = True Then
'            m_arrayDSOrders(0, RowIndex - 1) = Values(1)
'        End If
'
'        If prevVoucherValue = False And Values(2) = True Then
'            m_arrayDSOrders(1, RowIndex - 1) = Values(2)
'        End If
'
'    End If
'
    
    m_arrayDSOrders(0, RowIndex - 1) = Values(1)    'Invoice
    'm_arrayDSOrders(1, RowIndex - 1) = Values(2)    'Voucher
    'm_arrayDSOrders(21, RowIndex - 1) = Values(22)  'Process checkbox
    
'    If IsNumeric(Values(11)) Then
'        m_arrayDSOrders(15, RowIndex - 1) = Values(11)
'    End If

    'm_arrayDSOrders(16, RowIndex - 1) = Values(8)
    'm_arrayDSOrders(17, RowIndex - 1) = Values(13)
    
    If pskey > 0 Then SavePSInvoiceFlag pskey, Values(1)
    
End Sub

'BORROWED FROM GRIDEXWRAPPER CLASS
Private Sub gdxDSOrders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'cache this for use by DblClick
    m_lX = X
    m_lY = Y
End Sub

'BORROWED FROM GRIDEXWRAPPER CLASS
Private Sub gdxDSOrders_DblClick()
    Dim RowIndex As Integer

    With gdxDSOrders
        RowIndex = .RowIndex(.Row)
    End With
    
    If RowIndex = 0 Then Exit Sub
    
    With gdxDSOrders
        Select Case .HitTest(m_lX, m_lY)
        
                Case jgexHTBackGround, jgexHTColumnHeader, jgexHTGroupByBox, jgexHTNewRow
                    'Debug.Print "No Double-Click. HitTest = " & .HitTest(m_lX, m_lY) & " for (" & m_lX & ", " & m_lY & ")"
                    Exit Sub
                Case jgexHTCell
                    Dim colClicked As JSColumn
                    Set colClicked = .ColFromPoint(m_lX, m_lY)
                    
                    If Not colClicked Is Nothing Then
                        If colClicked.caption = "Tracking" Then
                            RowChosen
                        'ElseIf colClicked.caption = "Remarks" And gdxDSOrders.value(21) Then
                        ElseIf colClicked.caption = "Remarks" Then
                            LoadSpecialHandling
                        End If
                    End If
                Case jgexHTNoWhere 'JJC: for some reason, gdxItems returns this code on valid hits
                    'Debug.Print "GOT Double-Click. HitTest = " & .HitTest(m_lX, m_lY) & " for (" & m_lX & ", " & m_lY & ")"
                    'RaiseEvent RowChosen
                    'RowChosen
                    Exit Sub
                Case Else
                    'Debug.Print "GOT Double-Click. HitTest = " & .HitTest(m_lX, m_lY) & " for (" & m_lX & ", " & m_lY & ")"
                    'RaiseEvent RowChosen
                    'RowChosen
                    Exit Sub
        End Select
    End With
End Sub


'Display dropship PO remarks

Public Sub LoadCombo(cboCombo As ComboBox, rst As ADODB.Recordset, sDisplayField As String, Optional vKeyField As Variant, Optional vDfltKeyValue As Variant, Optional vWithNone As Variant)
    Dim lIdx As Long
    
    If IsMissing(vWithNone) Then
        vWithNone = False
    End If

    cboCombo.Clear
    
    If vWithNone Then
        cboCombo.AddItem "<none>"
        cboCombo.ItemData(cboCombo.NewIndex) = 0
    End If
    
    On Error Resume Next
    rst.MoveFirst
    On Error GoTo 0
    lIdx = 0
    If Not rst.EOF Then
        With rst
            Do While Not .EOF
                cboCombo.AddItem Trim(.Fields(sDisplayField).value)
    
                If Not IsMissing(vKeyField) Then
                    cboCombo.ItemData(cboCombo.NewIndex) = .Fields(vKeyField).value
                End If
    
                If Not IsMissing(vDfltKeyValue) Then
                    If Not IsMissing(vKeyField) Then
                        If vDfltKeyValue = .Fields(vKeyField).value Then
                            lIdx = cboCombo.NewIndex
                        End If
                    Else
                        If vDfltKeyValue = Trim(.Fields(sDisplayField).value) Then
                            lIdx = cboCombo.NewIndex
                        End If
                    End If
                End If
    
                .MoveNext
            Loop
        End With
    End If
    
    If cboCombo.ListCount > 0 Then
        cboCombo.ListIndex = lIdx
    End If
End Sub

Public Sub LoadImageList(ByRef i_oImageList As ImageList, ByRef o_oGridEX As GridEX)
    Dim i As Long
    
    o_oGridEX.GridImages.Clear
    For i = 1 To i_oImageList.ListImages.Count
        o_oGridEX.GridImages.Add i_oImageList.ListImages(i).Picture
    Next
End Sub

Private Sub SavePSInvoiceFlag(pskey As Integer, invoice As Boolean)
    Dim sSQL As String
    
    sSQL = "UPDATE tcpProvisionalShipment " & _
        "SET invoice=" & IIf(invoice, 1, 0) & " " & _
        "WHERE pskey=" & pskey
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP(sSQL, adCmdText)
    cmd.Execute
    Set cmd = Nothing
    
End Sub


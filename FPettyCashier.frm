VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FPettyCashier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Petty Cash Program"
   ClientHeight    =   8610
   ClientLeft      =   2820
   ClientTop       =   2775
   ClientWidth     =   8385
   Icon            =   "FPettyCashier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   8385
   Begin VB.CommandButton cmdBrokenRules 
      Caption         =   "Help"
      Height          =   855
      Left            =   8520
      TabIndex        =   65
      Top             =   960
      Width           =   735
   End
   Begin TabDlg.SSTab tabPettyCash 
      Height          =   7215
      Left            =   240
      TabIndex        =   36
      Top             =   1200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Transactions"
      TabPicture(0)   =   "FPettyCashier.frx":27A2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdCashInOut"
      Tab(0).Control(1)=   "frmCustCashTrx"
      Tab(0).Control(2)=   "cmdAccept"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Transaction Log"
      TabPicture(1)   =   "FPettyCashier.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDisplay"
      Tab(1).Control(1)=   "lblLogRowCount"
      Tab(1).Control(2)=   "grdTrxLog"
      Tab(1).Control(3)=   "cmdRefresh"
      Tab(1).Control(4)=   "UpDownTrxLog"
      Tab(1).Control(5)=   "cmdTransLogRptDisplay"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Transfer Amount"
      TabPicture(2)   =   "FPettyCashier.frx":27DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmTran"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Report"
      TabPicture(3)   =   "FPettyCashier.frx":27F6
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Research"
      TabPicture(4)   =   "FPettyCashier.frx":2812
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).Control(1)=   "grxResearch"
      Tab(4).ControlCount=   2
      Begin VB.CommandButton cmdTransLogRptDisplay 
         Caption         =   "Display"
         Height          =   375
         Left            =   -68520
         TabIndex        =   77
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "End of Quarter - Transaction Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   4080
         TabIndex        =   67
         Top             =   600
         Width           =   3375
         Begin VB.CommandButton cmdTransRptDisplay 
            BackColor       =   &H00000000&
            Caption         =   "Display"
            Height          =   375
            Left            =   1200
            TabIndex        =   72
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtTransRptStartDt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1200
            TabIndex        =   71
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtTransRptEndDt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1200
            TabIndex        =   70
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdTransRptStartDt 
            Caption         =   "..."
            Height          =   288
            Left            =   2640
            TabIndex        =   69
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdTransRptEndtDt 
            Caption         =   "..."
            Height          =   288
            Left            =   2640
            TabIndex        =   68
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label16 
            Caption         =   "Start Date:"
            Height          =   255
            Left            =   360
            TabIndex        =   74
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "End Date:"
            Height          =   255
            Left            =   360
            TabIndex        =   73
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Application Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   360
         TabIndex        =   59
         Top             =   600
         Width           =   3375
         Begin VB.CommandButton cmdRptEndDt 
            Caption         =   "..."
            Height          =   288
            Left            =   2640
            TabIndex        =   20
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton cmdRptStartDt 
            Caption         =   "..."
            Height          =   288
            Left            =   2640
            TabIndex        =   18
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtRptEndDt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1200
            TabIndex        =   19
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtRptStartDt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1200
            TabIndex        =   17
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton cmdRptDisplay 
            Caption         =   "Display"
            Height          =   375
            Left            =   1200
            TabIndex        =   21
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "End Date:"
            Height          =   255
            Left            =   360
            TabIndex        =   61
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Start Date:"
            Height          =   255
            Left            =   360
            TabIndex        =   60
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   -74760
         TabIndex        =   52
         Top             =   480
         Width           =   7455
         Begin VB.CommandButton cmdReEndDt 
            Caption         =   "..."
            Height          =   288
            Left            =   2280
            TabIndex        =   25
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton cmdReStartDt 
            Caption         =   "..."
            Height          =   288
            Left            =   2280
            TabIndex        =   23
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmdRFind 
            BackColor       =   &H80000012&
            Caption         =   "Find"
            Height          =   375
            Left            =   6120
            TabIndex        =   31
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtReStartDt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1080
            TabIndex        =   22
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtReAmt 
            Height          =   288
            Left            =   1080
            TabIndex        =   26
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtReDocNbr 
            Height          =   288
            Left            =   4080
            TabIndex        =   28
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtReEndDt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1080
            TabIndex        =   24
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtReCustName 
            Height          =   288
            Left            =   4080
            TabIndex        =   27
            Top             =   360
            Width           =   3135
         End
         Begin MSComCtl2.UpDown UpDownReNbrRecs 
            Height          =   285
            Left            =   4560
            TabIndex        =   30
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   50
            Max             =   200
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.Label lblReNoRecs 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4080
            TabIndex        =   29
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Nbr of Records:"
            Height          =   255
            Left            =   2880
            TabIndex        =   58
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Start Date:"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "OP Number:"
            Height          =   255
            Left            =   2880
            TabIndex        =   55
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "End Date:"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblRCustName 
            Caption         =   "Customer Name:"
            Height          =   255
            Left            =   2880
            TabIndex        =   53
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSComCtl2.UpDown UpDownTrxLog 
         Height          =   255
         Left            =   -67560
         TabIndex        =   13
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   50
         Max             =   200
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Frame frmTran 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   48
         Top             =   480
         Width           =   7455
         Begin VB.CommandButton cmdTransAmtKeep 
            Caption         =   "Keep $120"
            Height          =   375
            Left            =   4800
            TabIndex        =   66
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtTrxAmtNotes 
            Height          =   645
            Left            =   960
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   840
            Width           =   6255
         End
         Begin VB.CommandButton cmdTranAccept 
            Caption         =   "Accept"
            Height          =   372
            Left            =   6120
            TabIndex        =   16
            Top             =   1680
            Width           =   1092
         End
         Begin VB.TextBox txtAmountToTran 
            Height          =   288
            Left            =   960
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Notes:"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblDrawerToTran 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5040
            TabIndex        =   14
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Transfer to:"
            Height          =   255
            Left            =   4175
            TabIndex        =   50
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   375
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdCashInOut 
         Caption         =   "Cash In"
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   577
         Width           =   975
      End
      Begin VB.Frame frmCustCashTrx 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   43
         Top             =   3120
         Visible         =   0   'False
         Width           =   7455
         Begin VB.CommandButton cmdLookUp 
            Caption         =   "Look Up"
            Height          =   252
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   972
         End
         Begin VB.TextBox txtCustID 
            Height          =   288
            Left            =   840
            TabIndex        =   9
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtOpNo 
            Height          =   288
            Left            =   840
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
         Begin GridEX20.GridEX grxApp 
            Height          =   1935
            Left            =   240
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1440
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   3413
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MultiSelect     =   -1  'True
            HideSelection   =   2
            UseEvenOddColor =   -1  'True
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            ColumnHeaderHeight=   285
            ColumnsCount    =   8
            Column(1)       =   "FPettyCashier.frx":282E
            Column(2)       =   "FPettyCashier.frx":29A6
            Column(3)       =   "FPettyCashier.frx":2AF6
            Column(4)       =   "FPettyCashier.frx":2C3A
            Column(5)       =   "FPettyCashier.frx":2DAA
            Column(6)       =   "FPettyCashier.frx":2EF6
            Column(7)       =   "FPettyCashier.frx":304E
            Column(8)       =   "FPettyCashier.frx":31DE
            FormatStylesCount=   6
            FormatStyle(1)  =   "FPettyCashier.frx":3372
            FormatStyle(2)  =   "FPettyCashier.frx":3452
            FormatStyle(3)  =   "FPettyCashier.frx":358A
            FormatStyle(4)  =   "FPettyCashier.frx":363A
            FormatStyle(5)  =   "FPettyCashier.frx":36EE
            FormatStyle(6)  =   "FPettyCashier.frx":37C6
            ImageCount      =   0
            PrinterProperties=   "FPettyCashier.frx":387E
         End
         Begin VB.Label lblAmount2 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6360
            TabIndex        =   47
            Top             =   3480
            Width           =   855
         End
         Begin VB.Label lblAmt2 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   5760
            TabIndex        =   46
            Top             =   3480
            Width           =   735
         End
         Begin VB.Label lblCustID 
            Caption         =   "Cust ID:"
            Height          =   252
            Left            =   240
            TabIndex        =   45
            Top             =   720
            Width           =   1092
         End
         Begin VB.Label lblOpNo 
            Caption         =   "OP No:"
            Height          =   252
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H80000012&
         Caption         =   "Refresh"
         Height          =   252
         Left            =   -74760
         TabIndex        =   11
         Top             =   600
         Width           =   972
      End
      Begin MSFlexGridLib.MSFlexGrid grdTrxLog 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   840
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   9975
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         RowHeightMin    =   1
         MergeCells      =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Height          =   372
         Left            =   -68640
         TabIndex        =   7
         Top             =   2640
         Width           =   1092
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   37
         Top             =   720
         Width           =   7455
         Begin VB.TextBox txtDescr 
            Height          =   288
            Left            =   960
            TabIndex        =   5
            Top             =   720
            Width           =   6255
         End
         Begin VB.TextBox txtNotes 
            Height          =   615
            Left            =   960
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   1080
            Width           =   6255
         End
         Begin VB.ComboBox cboTrxType 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   2412
         End
         Begin VB.TextBox txtAmount 
            Height          =   288
            Left            =   6000
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Descr:"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   765
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Notes:"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1095
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Trx Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   435
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   5400
            TabIndex        =   38
            Top             =   435
            Width           =   735
         End
      End
      Begin GridEX20.GridEX grxResearch 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2280
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8281
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         HideSelection   =   2
         UseEvenOddColor =   -1  'True
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   5
         Column(1)       =   "FPettyCashier.frx":3A56
         Column(2)       =   "FPettyCashier.frx":3BC2
         Column(3)       =   "FPettyCashier.frx":3D0E
         Column(4)       =   "FPettyCashier.frx":3E6E
         Column(5)       =   "FPettyCashier.frx":3FCE
         FormatStylesCount=   6
         FormatStyle(1)  =   "FPettyCashier.frx":4112
         FormatStyle(2)  =   "FPettyCashier.frx":41F2
         FormatStyle(3)  =   "FPettyCashier.frx":432A
         FormatStyle(4)  =   "FPettyCashier.frx":43DA
         FormatStyle(5)  =   "FPettyCashier.frx":448E
         FormatStyle(6)  =   "FPettyCashier.frx":4566
         ImageCount      =   0
         PrinterProperties=   "FPettyCashier.frx":461E
      End
      Begin VB.Label lblLogRowCount 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -67920
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblDisplay 
         Caption         =   "Display Max:"
         Height          =   255
         Left            =   -68880
         TabIndex        =   51
         Top             =   630
         Width           =   975
      End
   End
   Begin VB.ComboBox cboDrawer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblMainDrawer 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   64
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8880
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   35
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblUserID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   34
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label lblBal 
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4560
      TabIndex        =   33
      Top             =   630
      Width           =   735
   End
   Begin VB.Label lblDrawer 
      Caption         =   "Drawer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   32
      Top             =   630
      Width           =   735
   End
End
Attribute VB_Name = "FPettyCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents m_oDrawers As Drawers
Attribute m_oDrawers.VB_VarHelpID = -1
Dim m_oAppl As Application
Dim mbIsCashIn As Boolean

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


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub cmdTransLogRptDisplay_Click()
    SetWaitCursor True

    Dim oFrm As FViewer
    Set oFrm = New FViewer
    
    Call oFrm.ViewReportByType("Transaction Log")
    Set oFrm = Nothing
        
    SetWaitCursor False
End Sub

Private Sub cmdTransRptDisplay_Click()
    Call TransRptDisplay
End Sub



Private Sub TransRptDisplay()

    Dim ldBegBalDrawer1 As Variant
    Dim ldBegBalDrawer2 As Variant
    Dim ldEndBalDrawer1 As Double
    Dim ldEndBalDrawer2 As Double

    ldBegBalDrawer1 = Null: ldBegBalDrawer2 = Null: ldEndBalDrawer1 = 0: ldEndBalDrawer2 = 0

   
        Dim lsSql As String
        lsSql = "select tcpPCTransaction.PCTrxTypeKey, tcpPCTransaction.PCTrxKey, tcpPCTransaction.TrxDate, tcpPCTransaction.Amount, "
        lsSql = lsSql & "tcpPCTransaction.DrawerKey, tcpPCTransaction.Balance, tcpPCTrxTypeDef.PCTrxTypeDescr from tcpPCTransaction "
        lsSql = lsSql & "inner join tcpPCTrxTypeDef on tcpPCTransaction.PCTrxTypeKey = tcpPCTrxTypeDef.PCTrxTypeKey "
        lsSql = lsSql & "where trxdate >= '" & txtTransRptStartDt.text & "' and trxdate <= '"
        lsSql = lsSql & txtTransRptEndDt.text & " 23:59:59.000' "
        'lsSql = lsSql & "order by tcpPCTransaction.PCTrxKey ASC, trxdate asc "
        lsSql = lsSql & "order by trxdate asc "

        Dim lors As ADODB.Recordset
        Set lors = New ADODB.Recordset
        lors.Source = lsSql
        Set lors.ActiveConnection = g_DB.Connection
        lors.Open

        If lors.EOF = False Then
            'get & save beg/end balance to send to report
            While lors.EOF = False
            Select Case lors!Drawerkey
                Case 1
                If IsNull(ldBegBalDrawer1) Then ldBegBalDrawer1 = lors!Balance - lors!Amount
                ldEndBalDrawer1 = lors!Balance
                Case 2
                If IsNull(ldBegBalDrawer2) Then ldBegBalDrawer2 = lors!Balance - lors!Amount
                ldEndBalDrawer2 = lors!Balance
            End Select
            lors.MoveNext
            Wend
        End If

        Dim begBal As Double
        Dim endBal As Double
        Dim startDate As Date
        Dim endDate As Date
                
        begBal = IIf(IsNull(ldBegBalDrawer1), 0, ldBegBalDrawer1) + IIf(IsNull(ldBegBalDrawer2), 0, ldBegBalDrawer2)
        endBal = ldEndBalDrawer1 + ldEndBalDrawer2
        startDate = CDate(txtTransRptStartDt.text + " 00:00:00")
        endDate = CDate(txtTransRptEndDt.text + " 23:59:59")

    
    
    SetWaitCursor True
    
    Dim oFrm As FViewer
    Set oFrm = New FViewer
       
    Call oFrm.ParamAdd(1, "BegBal", begBal)
    Call oFrm.ParamAdd(2, "EndBal", endBal)
    Call oFrm.ParamAdd(3, "TrxStartDate", startDate)
    Call oFrm.ParamAdd(4, "TrxEndDate", endDate)
    
    Call oFrm.ViewReportByType("PCEndofQuarter")
    Set oFrm = Nothing
        
    SetWaitCursor False
End Sub


Private Sub cmdTransRptStartDt_Click()
    frmDate.Show 1
    Me.txtTransRptStartDt.text = frmDate.mthView.value
    Unload frmDate
End Sub
Private Sub cmdTransRptEndtDt_Click()
    frmDate.Show 1
    Me.txtTransRptEndDt.text = frmDate.mthView.value
    Unload frmDate
End Sub


''''''''''''''''''''''''''''''''''''''''''''''
' Form Events
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
    'Note all users in the Cashier group must also be setup in tcpPCDrawer table
    
    SetCaption "Petty Cashier Tool"
    Me.width = 8475
    Me.Height = 8985

    'Call OpenDbConnection
    mbIsCashIn = True
    Set m_oDrawers = New Drawers
    
    Call DisplayForm
    
    'Transaction tab
    If m_oDrawers.Drawers.Count = 1 Then
        Call m_oDrawers.CurDrawer.StandardTrx.Init(m_oDrawers.CurDrawer.StandardTrx.PCTrxKey, m_oDrawers.CurDrawer.Balance)
    Else
        Call m_oDrawers.CurDrawer.StandardTrx.Init(cboTrxType.ItemData(cboTrxType.ListIndex), m_oDrawers.CurDrawer.Balance)
    End If
    Call DisplayTrx
    
    'Transfer tab
    Call m_oDrawers.CurDrawer.TransTrx.Init(TRANSFERIN, m_oDrawers.CurDrawer.TransTrx.TrxBalance)
    m_oDrawers.CurDrawer.TransTrx.TrxDescr = "Transfer"
    m_oDrawers.CurDrawer.TransTrx.TrxAmt = 0
     
    'Need to call this once since event isn't raised in init
    Me.cmdAccept.Enabled = m_oDrawers.CurDrawer.StandardTrx.IsValid
    Me.cmdTranAccept.Enabled = m_oDrawers.CurDrawer.TransTrx.IsValid
    
    For liCounter = 1 To grxApp.Columns.Count
        grxApp.Columns(liCounter).AutoSize
    Next
    
    lblReNoRecs.caption = UpDownReNbrRecs.value
    lblLogRowCount.caption = UpDownTrxLog.value
    
    'Set report start and end date to today's date
    txtRptStartDt.text = Format(Now, "MM/DD/YYYY")
    txtRptEndDt.text = Format(Now, "MM/DD/YYYY")
    txtTransRptStartDt.text = Format(Now, "MM/DD/YYYY")
    txtTransRptEndDt.text = Format(Now, "MM/DD/YYYY")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub m_oDrawers_CurDrawerStdTrxValidityChanged(IsValid As Boolean)
    cmdAccept.Enabled = IsValid
End Sub


Private Sub m_oDrawers_CurDrawerTransTrxValidityChanged(IsValid As Boolean)
    cmdTranAccept.Enabled = IsValid
End Sub


Private Sub cboDrawer_Click()
    m_oDrawers.CurDrawer = m_oDrawers.Drawers(cboDrawer.ListIndex + 1)
    
    'Reset the current transaction balance when the drawer changes (LA Vault <> LA Working.LA Vault)
    m_oDrawers.CurDrawer.TranDrawer.SetBalance
    
    'Transaction tab: After curdrawer is selected, show current transaction data
    txtAmount.text = m_oDrawers.CurDrawer.StandardTrx.TrxAmt
    txtDescr.text = m_oDrawers.CurDrawer.StandardTrx.TrxDescr
    cmdAccept.Enabled = m_oDrawers.CurDrawer.StandardTrx.IsValid
    
    'Transfer tab: After curdrawer is selected, show trandrawer on transfer tab
    lblDrawerToTran.caption = m_oDrawers.CurDrawer.TranDrawer.Descr
    txtAmountToTran.text = m_oDrawers.CurDrawer.TransTrx.TrxAmt
    cmdTranAccept.Enabled = m_oDrawers.CurDrawer.TransTrx.IsValid
    
    Call FillTrxTypeCbo
    Call FillTrxGrid
End Sub


Private Sub cboTrxType_Click()
    Call SetTrxTypeProps
    With m_oDrawers.CurDrawer.StandardTrx.TrxType
        If .TrxTypeKey = CASHRECEIPT Or .TrxTypeKey = CASHREFUND Then
            frmCustCashTrx.Visible = True
        Else
            frmCustCashTrx.Visible = False
        End If
    End With
End Sub


Private Sub cmdAccept_Click()
    Call SetTrxTypeProps
    
    'Get current balance from the database for database concurrency reasons
    Call m_oDrawers.CurDrawer.SetBalance
    m_oDrawers.CurDrawer.StandardTrx.TrxBalance = m_oDrawers.CurDrawer.Balance
    
    'Save transaction
    Call m_oDrawers.CurDrawer.StandardTrx.Save(m_oDrawers.CurDrawer.Drawerkey, m_oDrawers.UserID)
    
    'Save application(s)
    If grxApp.SelectedItems.Count > 0 Then
        'GetSelCol: retrieves PCTrxKey, TranType, DocNbr, CustName, TranAmt to store in a collection
        'SetApplication: Calls Application.Insert for each item in the col (to insert into tcpPCApplication)
        Call m_oDrawers.CurDrawer.StandardTrx.SetApplications(GetSelCol)
    End If
        
    'Grab new Balance and display it
    m_oDrawers.CurDrawer.Balance = m_oDrawers.CurDrawer.StandardTrx.TrxBalance
    lblBalance.caption = " $" & Format(m_oDrawers.CurDrawer.Balance, "#0.00")
    
    'Reset the trx
    m_oDrawers.CurDrawer.StandardTrx.Reset cboTrxType.ItemData(cboTrxType.ListIndex), m_oDrawers.CurDrawer.Balance
    
    'Display the reset transaction
    Call DisplayTrx
    Call ClearTransactionCtrls
End Sub


Private Sub grxApp_SelectionChange()
    txtAmount.text = Format(GetSelAmt, "0#.00")
End Sub


Private Sub cmdBrokenRules_Click()
    Dim aRule As Rule
    
    If m_oDrawers.CurDrawer.StandardTrx.GetBrokenRules.Count = 0 Then
        MsgBox "no broken rules"
    Else
        For Each aRule In m_oDrawers.CurDrawer.StandardTrx.GetBrokenRules
            MsgBox aRule.Descr
        Next
    End If
End Sub


Private Sub cmdCashInOut_Click()
    If cmdCashInOut.caption = "Cash In" Then
        cmdCashInOut.caption = "Cash Out"
        mbIsCashIn = False
    Else
        cmdCashInOut.caption = "Cash In"
        mbIsCashIn = True
    End If
    
    Call FillTrxTypeCbo
    cboTrxType.SetFocus
    
    Call ClearTransactionCtrls
End Sub


Private Sub cmdLookUp_Click()
    Dim liDateCounter As Integer
    liDateCounter = 0

    If Len(Trim(txtCustID.text)) > 0 Then
        Dim lors As ADODB.Recordset
        Dim liCounter As Integer
        
        SetWaitCursor True
        
        'returns recordset and pop grid
        If Len(Trim(txtOpNo.text)) = 0 Then
            Set lors = m_oDrawers.CurDrawer.StandardTrx.GetInvAppls(txtCustID.text)
        Else
            Set lors = m_oDrawers.CurDrawer.StandardTrx.GetInvApplsUnion(txtOpNo.text, txtCustID.text)
        End If
    
        SetWaitCursor False
        
        'fill in grid with recordset
        grxApp.HoldFields
        grxApp.HoldSortSettings = True
        Set grxApp.ADORecordset = lors
        
        If lors.RecordCount > 0 Then
            txtDescr.text = grxApp.value(6) & "; Customer ID: " & Trim(grxApp.value(5))
            
            'check transaction date - display a message is more than 3 months old
            If DateAdd("m", 3, grxApp.value(4)) <= Now Then
                liDateCounter = liDateCounter + 1
            End If
        End If
        
        If IsNull(grxApp.value(7)) Then
           txtAmount.text = 0
        Else
            txtAmount.text = grxApp.value(7)
        End If

        For liCounter = 1 To grxApp.Columns.Count
            grxApp.Columns(liCounter).AutoSize
        Next
        
        Set lors = Nothing
    Else
        ClearTransactionCtrls
    End If
    
    If liDateCounter > 0 Then MsgBox "Transaction date is more than three months old.", vbInformation, "Petty Cashier"
End Sub


Private Sub cmdReEndDt_Click()
    frmDate.Show 1
    Me.txtReEndDt.text = frmDate.mthView.value
    Unload frmDate
End Sub


Private Sub cmdRefresh_Click()
    m_oDrawers.CurDrawer.TrxLogData (lblLogRowCount.caption)
    Call FillTrxGrid
End Sub


Private Sub cmdReStartDt_Click()
    frmDate.Show 1
    Me.txtReStartDt.text = frmDate.mthView.value
    Unload frmDate
End Sub


Private Sub cmdRFind_Click()
    Dim lors As ADODB.Recordset
    Dim liCounter As Integer
    
    'returns recordset and pop grid
    Set lors = m_oDrawers.CurDrawer.StandardTrx.GetTransResearch(lblReNoRecs.caption, txtReCustName.text, txtReStartDt.text, txtReEndDt.text, txtReDocNbr.text, txtReAmt.text)

    'fill in grid with recordset
    grxResearch.HoldFields
    grxResearch.HoldSortSettings = True
    Set grxResearch.ADORecordset = lors
    
    For liCounter = 1 To grxResearch.Columns.Count
        grxResearch.Columns(liCounter).AutoSize
    Next
        
    Set lors = Nothing
End Sub


Private Sub cmdRptEndDt_Click()
    frmDate.Show 1
    Me.txtRptEndDt.text = frmDate.mthView.value
    Unload frmDate
End Sub


Private Sub cmdRptDisplay_Click()
    SetWaitCursor True
    Call RptDisplay
    SetWaitCursor False
End Sub

Private Sub cmdRptStartDt_Click()
    frmDate.Show 1
    Me.txtRptStartDt.text = frmDate.mthView.value
    Unload frmDate
End Sub


Private Sub cmdTranAccept_Click()
    Dim aDrawer As Drawer
    
    Call m_oDrawers.CurDrawer.TransferBal(CDbl(txtAmountToTran.text), CStr(txtTrxAmtNotes.text), m_oDrawers.UserID)
    txtAmountToTran.text = 0
    txtTrxAmtNotes.text = ""
    
   'Display new balance
    lblBalance.caption = " $" & Format(m_oDrawers.CurDrawer.Balance, "#0.00")
    
    'walk through all the drawers for this user
    For Each aDrawer In m_oDrawers.Drawers
        'if the drawer key matches key of transfer drawer
        If aDrawer.Drawerkey = m_oDrawers.CurDrawer.TranDrawer.Drawerkey Then
            aDrawer.Balance = m_oDrawers.CurDrawer.TranDrawer.Balance
        End If
    Next
End Sub


Private Sub optCashIn_Click()
    Call FillTrxTypeCbo
    cboTrxType.SetFocus
End Sub


Private Sub optCashOut_Click()
    Call FillTrxTypeCbo
    cboTrxType.SetFocus
End Sub


Private Sub tabPettyCash_Click(PreviousTab As Integer)
    If tabPettyCash.Tab = 1 Then
        m_oDrawers.CurDrawer.TrxLogData (lblLogRowCount.caption)
        Call FillTrxGrid
    End If
End Sub


Private Sub txtAmount_Change()
    If Not IsNumeric(txtAmount.text) Or Len(txtAmount.text) = 0 Then
        txtAmount.text = 0
        txtAmount.SetFocus
        txtAmount.SelStart = 0
        txtAmount.SelLength = 1
        Exit Sub
    End If

    m_oDrawers.CurDrawer.StandardTrx.TrxAmt = CDbl(txtAmount.text)
    lblAmount2.caption = txtAmount.text
End Sub


Private Sub txtAmount_LostFocus()
    m_oDrawers.CurDrawer.StandardTrx.TrxAmt = CDbl(txtAmount.text)
End Sub


Private Sub txtAmountToTran_Change()
    If Not IsNumeric(txtAmountToTran.text) Or Len(txtAmountToTran.text) = 0 Then
        txtAmountToTran.text = 0
        txtAmountToTran.SetFocus
        txtAmountToTran.SelStart = 0
        txtAmountToTran.SelLength = 1
        Exit Sub
    End If
    
    m_oDrawers.CurDrawer.TransTrx.TrxType.TrxTypeKey = TRANSFEROUT
    m_oDrawers.CurDrawer.TransTrx.TrxAmt = txtAmountToTran.text
    m_oDrawers.CurDrawer.TransTrx.TrxDescr = m_oDrawers.CurDrawer.TranDrawer.Descr 'lblDrawerToTran.Caption
    
    'added
    cmdTranAccept.Enabled = m_oDrawers.CurDrawer.TransTrx.IsValid
End Sub


Private Sub txtDescr_Change()
    m_oDrawers.CurDrawer.StandardTrx.TrxDescr = txtDescr.text
End Sub


Private Sub txtNotes_Change()
    m_oDrawers.CurDrawer.StandardTrx.TrxNotes = txtNotes.text
End Sub

Private Sub txtOpNo_LostFocus()
    If Len(txtOpNo.text) > 0 And Not IsNumeric(txtOpNo) Then
        MsgBox "Op number needs to be numeric.", vbInformation, "Petty Cash"
        txtOpNo.SetFocus
        Exit Sub
    End If

    If Len(Trim(txtOpNo)) > 0 Then
        Dim lors As ADODB.Recordset
        Dim liCounter As Integer
        
        'returns recordset and pop grid
        Set lors = m_oDrawers.CurDrawer.StandardTrx.GetInvApplsPO(txtOpNo.text)
    
        With grxApp
            'fill in grid with recordset
            .HoldFields
            .HoldSortSettings = True
                       
            Set .ADORecordset = lors
                        
            For liCounter = 1 To .Columns.Count
                .Columns(liCounter).AutoSize
            Next
            
            If lors.RecordCount = 1 Then
                txtCustID.text = .value(5)
                txtDescr.text = .value(6) & "; " & .value(2) & "/Doc Number: " & .value(3) & "; Customer ID: " & Trim(.value(5))
                If IsNull(.value(7)) Then
                   txtAmount.text = 0
                Else
                    txtAmount.text = .value(7)
                End If
                
                'PRN473 Inbound Freight msgbox - passing in the OPKey as string
                If Trim(.value(8)) = 1 Then MsgBox "Inbound Freight needs to be applied for " & Trim(.value(5)) & "."
                
            End If
        End With
        
        Set lors = Nothing
    End If
End Sub

Private Sub txtReAmt_Change()
    Call ResearchValid
End Sub


Private Sub txtReCustName_Change()
    Call ResearchValid
End Sub


Private Sub txtReDocNbr_Change()
    Call ResearchValid
End Sub


Private Sub txtReEndDt_Change()
    Call ResearchValid
End Sub


Private Sub txtReStartDt_Change()
    Call ResearchValid
End Sub


Private Sub txtRptEndDt_Change()
    Call RptValid
End Sub


Private Sub txtRptStartDt_Change()
    Call RptValid
End Sub


Private Sub cmdTransAmtKeep_Click()
    txtAmountToTran.text = Format(m_oDrawers.CurDrawer.Balance - 120, "#0.00")
End Sub





Private Sub UpDownReNbrRecs_Change()
    lblReNoRecs.caption = UpDownReNbrRecs.value
End Sub


Private Sub UpDownTrxLog_Change()
    lblLogRowCount.caption = UpDownTrxLog.value
End Sub



''''''''''''''''''''''''''''''''''
'PRIVATE SUBROUTINES
''''''''''''''''''''''''''''''''''

Private Sub RptValid()
    Dim lbRptValid As Boolean

    If Len(txtRptStartDt.text) = 0 Or Len(txtRptEndDt.text) = 0 Then
        lbRptValid = False
    ElseIf CDate(txtRptEndDt.text) > DateAdd("d", 7, txtRptStartDt.text) Then
        lbRptValid = False
        MsgBox "End Date can only be a maximum of 7 day after Start Date.", vbInformation, "Petty Cash"
    ElseIf CDate(txtRptEndDt.text) < CDate(txtRptStartDt.text) Then
        lbRptValid = False
        MsgBox "End Date must be after Start Date.", vbInformation, "Date Range Error"
    Else
        lbRptValid = True
    End If
    
    cmdRptDisplay.Enabled = lbRptValid
End Sub


Private Sub ResearchValid()
    Dim lbResearchValid As Boolean
    lbResearchValid = True

    If Len(txtReStartDt.text) > 0 And Len(txtReEndDt.text) > 0 Then
        If CDate(txtReEndDt.text) <= CDate(txtReStartDt.text) Then
            MsgBox "End Date must be after Start Date.", vbInformation, "Date Range Error"
            lbResearchValid = False
        End If
    ElseIf Len(txtReStartDt.text) = 0 Or Len(txtReEndDt.text) = 0 Then
        lbResearchValid = False
    End If
    
    If IsNumeric(txtReAmt.text) = False And Len(txtReAmt.text) > 0 Then
        lbResearchValid = False
    End If
    
    cmdRFind.Enabled = lbResearchValid
End Sub


Private Sub SetTrxTypeProps()
    'To set transaction type properties
    If mbIsCashIn = True Then   'debit
        m_oDrawers.CurDrawer.StandardTrx.TrxType.TrxTypeKey = m_oDrawers.CurDrawer.PossDebitTrx.Item(cboTrxType.ListIndex + 1).TrxTypeKey
        m_oDrawers.CurDrawer.StandardTrx.TrxType.IsDebit = True
        m_oDrawers.CurDrawer.StandardTrx.TrxType.Descr = m_oDrawers.CurDrawer.PossDebitTrx.Item(cboTrxType.ListIndex + 1).Descr
    Else    'credit
        m_oDrawers.CurDrawer.StandardTrx.TrxType.TrxTypeKey = m_oDrawers.CurDrawer.PossCreditTrx.Item(cboTrxType.ListIndex + 1).TrxTypeKey
        m_oDrawers.CurDrawer.StandardTrx.TrxType.IsDebit = False
        m_oDrawers.CurDrawer.StandardTrx.TrxType.Descr = m_oDrawers.CurDrawer.PossCreditTrx.Item(cboTrxType.ListIndex + 1).Descr
    End If
End Sub


Private Sub ClearTransactionCtrls()
    txtOpNo.text = ""
    txtCustID.text = ""
    txtDescr.text = ""
    txtAmount.text = 0
    txtNotes.text = ""
    
    'clear transaction grid
    Dim lors As ADODB.Recordset
    Set lors = Nothing
    grxApp.HoldFields
    Set grxApp.ADORecordset = lors
End Sub


Private Sub RptDisplay()
    SetWaitCursor True
    
    Dim oFrm As FViewer
    Set oFrm = New FViewer
    
    FillDbTable txtRptStartDt.text, DateAdd("d", 1, txtRptEndDt)
    
    Call oFrm.ViewReportByType("ApplWorking")
    Set oFrm = Nothing

        
    SetWaitCursor False
End Sub


Private Sub DisplayForm()
    Me.lblUserID.caption = m_oDrawers.UserID
    Me.lblDate.caption = Format(Now, "MM/DD/YYYY")
    Me.tabPettyCash.Tab = 0

    'startup for report and research command buttons
    Call RptValid
    Call ResearchValid

    Call DisplayDrawers
End Sub


Private Sub DisplayTrx()
    txtAmount.text = m_oDrawers.CurDrawer.StandardTrx.TrxAmt
    txtNotes.text = m_oDrawers.CurDrawer.StandardTrx.TrxNotes
    txtDescr.text = m_oDrawers.CurDrawer.StandardTrx.TrxDescr
End Sub


Private Sub DisplayDrawers()
    Dim PrefIndex As Integer
    Dim aDrawer As Drawer
        
    lblMainDrawer.caption = ""
    cboDrawer.Clear
    
    For Each aDrawer In m_oDrawers.Drawers
        'This logic is for only 1 drawer shown in a text box
            'therefore, code will not call the cbo clicked event
        If m_oDrawers.Drawers.Count = 1 Then
            lblMainDrawer.Visible = True
            lblMainDrawer.caption = aDrawer.Descr
            m_oDrawers.CurDrawer = m_oDrawers.Drawers(1)
            
            'After curdrawer is selected, show trandrawer on transfer tab
            lblDrawerToTran.caption = m_oDrawers.CurDrawer.TranDrawer.Descr

            Call FillTrxTypeCbo
            Call FillTrxGrid
            Exit Sub
        Else
            cboDrawer.Visible = True
            cboDrawer.AddItem aDrawer.Descr
                
            If (aDrawer Is m_oDrawers.CurDrawer) Then
                'make this the selected drawer in the cbo
                PrefIndex = cboDrawer.ListCount - 1
            End If
        End If
    Next
    
    cboDrawer.ListIndex = PrefIndex
End Sub


Private Sub FillTrxTypeCbo()
    Dim aType As TransactionType
  
    With cboTrxType
        .Clear
        If UCase(cmdCashInOut.caption) = "CASH OUT" Then
            For Each aType In m_oDrawers.CurDrawer.PossCreditTrx
                .AddItem aType.Descr
                .ItemData(.NewIndex) = aType.TrxTypeKey
            Next
        Else
            For Each aType In m_oDrawers.CurDrawer.PossDebitTrx
                .AddItem aType.Descr
                .ItemData(.NewIndex) = aType.TrxTypeKey
            Next
        End If

        .ListIndex = 0
    End With
End Sub


Private Sub FillTrxGrid()
    Dim aTrx As Transaction
    Dim liCounter As Integer
    Dim ldBalance As Double
    
    liCounter = 0

    With grdTrxLog
        .Clear
        .Rows = 1
        .Cols = 4
        
        .col = 0
        .ColWidth(0) = 900    'date
        .text = "Date"
        
        .col = 1
        .ColWidth(1) = 850     'amt
        .text = "Amount"
        
        .col = 2
        .ColWidth(2) = 4512     'descr
        .text = "Description"
        
        .col = 3
        .ColWidth(3) = 850     'balance
        .text = "Balance"
    End With
    'load grid
    For Each aTrx In m_oDrawers.CurDrawer.TrxLog
        liCounter = liCounter + 1
        'fill grid
        With grdTrxLog
            .Rows = liCounter + 1
            .Row = liCounter
            
            .col = 0
            .text = Format(aTrx.TrxDate, "MM/DD/YYYY")
            .col = 1
            .text = "$" & Format(aTrx.TrxAmt, "#0.00")
            .col = 2
            .text = aTrx.TrxDescr
            .col = 3
            .text = "$" & Format(aTrx.TrxBalance, "#0.00")
            .Refresh
        End With
    Next
    lblBalance.caption = " $" & Format(m_oDrawers.CurDrawer.Balance, "#0.00")
End Sub


''''''''''''''''''''''''''''''''''
'PRIVATE FUNCTIONS
''''''''''''''''''''''''''''''''''

Private Function GetSelCol() As Collection
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    Dim i As Long
    Dim loAppl As Application
    Dim l_objSelAppCol As Collection  '(selected items on grid)
    
    Set l_objSelAppCol = New Collection

    For Each simTemp In grxApp.SelectedItems
        Set RowData = grxApp.GetRowData(simTemp.RowPosition)
        
        'need to add application objects (loAppl) to a collection (l_objSelAppCol)
        Set loAppl = New Application

        loAppl.AppKey = m_oDrawers.CurDrawer.StandardTrx.PCTrxKey  'Transaction Key
        'loAppl.AppKey = RowData(1)        'Unique Key from OP# and CustID queries
        loAppl.AppTranType = RowData(2)    'TranType
        loAppl.AppDocNbr = RowData(3)      'Doc Number
        loAppl.AppTranDate = RowData(4)    'Date
        loAppl.AppCustID = RowData(5)      'Cust ID
        loAppl.AppCustName = RowData(6)    'Cust Name
        
        If IsNull(RowData(7)) Then
            loAppl.AppAmt = 0
        Else
            loAppl.AppAmt = RowData(7) 'Amount
        End If

        'smr IBF
        loAppl.AppIBF = RowData(8)    'IBF (0 or 1)

        'add data to Sel App Collection''''''''''''''
        l_objSelAppCol.Add loAppl
        Set GetSelCol = l_objSelAppCol
    Next
End Function


Private Function GetSelAmt() As Double
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    GetSelAmt = 0

    On Error Resume Next    'this fixes a runtime bug in the grid (only on TS)
    For Each simTemp In grxApp.SelectedItems
        Set RowData = grxApp.GetRowData(simTemp.RowPosition)
        If IsNull(RowData(7)) Then
            GetSelAmt = 0
        Else
            GetSelAmt = GetSelAmt + (RowData(7))   'Amount
        End If
    Next
End Function

Private Sub FillDbTable(startDate As String, endDate As String)
    ExecuteSP "spCPC_PCApplWorking", "@InParamStartDate", startDate, "@InParamEndDate", endDate
End Sub

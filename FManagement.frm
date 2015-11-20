VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FManagement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   9210
   Begin ActiveTabs.SSActiveTabs tabManagement 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8070
      _Version        =   262144
      TabCount        =   2
      TagVariant      =   ""
      Tabs            =   "FManagement.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4185
         Left            =   30
         TabIndex        =   26
         Top             =   360
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   7382
         _Version        =   262144
         TabGuid         =   "FManagement.frx":0086
         Begin VB.Frame Frame1 
            Caption         =   "Result"
            Height          =   3015
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Width           =   8295
            Begin GridEX20.GridEX gdxCSRMetrics 
               Height          =   2175
               Left            =   120
               TabIndex        =   25
               Top             =   720
               Width           =   7935
               _ExtentX        =   13996
               _ExtentY        =   3836
               Version         =   "2.0"
               AutomaticSort   =   -1  'True
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               MethodHoldFields=   -1  'True
               AutomaticArrange=   0   'False
               AllowEdit       =   0   'False
               GroupByBoxVisible=   0   'False
               ColumnHeaderHeight=   285
               IntProp1        =   0
               IntProp2        =   0
               IntProp7        =   0
               ColumnsCount    =   3
               Column(1)       =   "FManagement.frx":00AE
               Column(2)       =   "FManagement.frx":020E
               Column(3)       =   "FManagement.frx":040E
               FormatStylesCount=   5
               FormatStyle(1)  =   "FManagement.frx":05EA
               FormatStyle(2)  =   "FManagement.frx":0722
               FormatStyle(3)  =   "FManagement.frx":07D2
               FormatStyle(4)  =   "FManagement.frx":0886
               FormatStyle(5)  =   "FManagement.frx":095E
               ImageCount      =   0
               PrinterProperties=   "FManagement.frx":0A16
            End
            Begin VB.TextBox txtTotalSales 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4200
               TabIndex        =   24
               Top             =   330
               Width           =   1455
            End
            Begin VB.TextBox txtTotalOrder 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               TabIndex        =   23
               Top             =   330
               Width           =   1455
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Total Order"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   32
               Top             =   360
               Width           =   795
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Total Sales"
               Height          =   195
               Index           =   1
               Left            =   3240
               TabIndex        =   31
               Top             =   360
               Width           =   795
            End
         End
         Begin VB.CommandButton cmdRefresh_CSRMetrics 
            Caption         =   "&Refresh"
            Height          =   375
            Left            =   6840
            TabIndex        =   22
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cboWarehouse_CSRMetrics 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker dtpBeginTime_CSRMetrics 
            Height          =   315
            Left            =   1680
            TabIndex        =   21
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62062593
            CurrentDate     =   37312
         End
         Begin MSComCtl2.DTPicker dtpEndTime_CSRMetrics 
            Height          =   315
            Left            =   4440
            TabIndex        =   33
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62062593
            CurrentDate     =   37312
         End
         Begin VB.Label Label1 
            Caption         =   "Warehouse"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "From"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "To"
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   27
            Top             =   720
            Width           =   495
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4185
         Left            =   30
         TabIndex        =   15
         Top             =   360
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   7382
         _Version        =   262144
         TabGuid         =   "FManagement.frx":0BEE
         Begin VB.CheckBox chkInvoicedOrders 
            Caption         =   "Invoiced Orders"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Show Me the Money"
            Height          =   375
            Left            =   6840
            TabIndex        =   9
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optShipments 
            Caption         =   "Shipped From"
            Height          =   255
            Left            =   1800
            TabIndex        =   2
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optOrders 
            Caption         =   "Ordered From"
            Height          =   255
            Left            =   360
            TabIndex        =   1
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboWarehouse 
            Height          =   315
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame Frame2 
            Caption         =   "Result"
            Height          =   1335
            Left            =   360
            TabIndex        =   10
            Top             =   2040
            Width           =   8175
            Begin VB.TextBox txtTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   375
               Left            =   5880
               TabIndex        =   19
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtFreight 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   375
               Left            =   5880
               TabIndex        =   18
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txtSalesTax 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   375
               Left            =   1800
               TabIndex        =   17
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtSales 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   375
               Left            =   1800
               TabIndex        =   16
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label11 
               Caption         =   "Sales"
               Height          =   375
               Index           =   0
               Left            =   600
               TabIndex        =   11
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label10 
               Caption         =   "Sales Tax"
               Height          =   375
               Index           =   0
               Left            =   600
               TabIndex        =   12
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label9 
               Caption         =   "Freight"
               Height          =   375
               Left            =   4560
               TabIndex        =   13
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label8 
               Caption         =   "Total"
               Height          =   255
               Left            =   4560
               TabIndex        =   14
               Top             =   840
               Width           =   855
            End
         End
         Begin MSComCtl2.DTPicker dtpBeginTime 
            Height          =   315
            Left            =   2280
            TabIndex        =   6
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62062593
            CurrentDate     =   37312
         End
         Begin MSComCtl2.DTPicker dtpEndTime 
            Height          =   315
            Left            =   5040
            TabIndex        =   8
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62062593
            CurrentDate     =   37312
         End
         Begin VB.Label Label3 
            Caption         =   "To"
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   7
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "From"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   5
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Warehouse"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "FManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWindowID As Long


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


Private Sub chkInvoicedOrders_Click()
    'smr - added to support "invoiced orders" when "Show Me the Money" is checked
    If chkInvoicedOrders = vbChecked Then
        dtpBeginTime.Enabled = False
        dtpEndTime.Enabled = False
    Else
        dtpBeginTime.Value = Format(Now, "mm/01/yy")
        dtpEndTime.Value = Format(Now, "mm/dd/yy 23:59:59")
        dtpBeginTime.Enabled = True
        dtpEndTime.Enabled = True
    End If
End Sub

Private Sub cmdRefresh_Click()
    Dim rst As ADODB.Recordset
    
    SetWaitCursor True
    
    If chkInvoicedOrders = vbChecked Then
        'smr - added to support "invoiced orders" that are either "ordered from" or "shipped from"
        Set rst = CallSP("spCPCGetMTDSales", "@IsCSR", Int(optOrders.Value), "@WhseID", cboWarehouse.List(cboWarehouse.ListIndex))
    Else
        Set rst = CallSP("spCPCGetSalesSummary", "@StartDate", dtpBeginTime.Value, _
                        "@EndDate", dtpEndTime.Value, "@WhseID", cboWarehouse.List(cboWarehouse.ListIndex), _
                        "@ByShipment", CInt(optShipments.Value))
    End If

    If rst.EOF Then
        txtSales.Text = FormatCurrency(0)
        txtSalesTax.Text = FormatCurrency(0)
        txtFreight.Text = FormatCurrency(0)
        txtTotal.Text = FormatCurrency(0)
    Else
        txtSales.Text = FormatCurrency(rst.Fields("Sales").Value)
        txtFreight.Text = FormatCurrency(rst.Fields("Freight").Value)
        txtSalesTax.Text = FormatCurrency(rst.Fields("SalesTax").Value)
        txtTotal.Text = FormatCurrency(rst.Fields("Sales").Value + rst.Fields("Freight").Value + rst.Fields("SalesTax").Value)
        
        'smr - added to support "invoiced orders" - to and from date are set based on a recordset field value
        If chkInvoicedOrders = vbChecked Then
            dtpBeginTime.Value = Format(rst.Fields("DateTimeRecorded").Value, "mm/01/yy")
            dtpEndTime.Value = Format(rst.Fields("DateTimeRecorded").Value, "mm/dd/yy 23:59:59")
        End If
    End If
    
    SetWaitCursor False
End Sub



Private Sub cmdRefresh_CSRMetrics_Click()
    'PRN#204 Add a CSR metrics tool
    Dim rst As ADODB.Recordset
    On Error GoTo EH
    
    SetWaitCursor True
    
    Set rst = CallSP("spCPCGetCSRMetrics", "@StartDate", dtpBeginTime_CSRMetrics.Value, _
                    "@EndDate", dtpEndTime_CSRMetrics.Value, "@WhseID", cboWarehouse_CSRMetrics.List(cboWarehouse_CSRMetrics.ListIndex))

    With gdxCSRMetrics
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = rst
    End With
    
    Set rst = Nothing
    
    Dim i As Integer
    Dim lTotalOrder As Long
    Dim cTotalSales As Currency
    For i = 1 To gdxCSRMetrics.RowCount
        lTotalOrder = lTotalOrder + gdxCSRMetrics.GetRowData(i).GetSubTotal(2, jgexSum)
        cTotalSales = cTotalSales + gdxCSRMetrics.GetRowData(i).GetSubTotal(3, jgexSum)
    Next
    
    txtTotalOrder.Text = Format(lTotalOrder, "#,###")
    txtTotalSales.Text = Format(cTotalSales, "$#,###.00")
    
    SetWaitCursor False
    Exit Sub
EH:
    SetWaitCursor False
    MsgBox "Failed to load CSR Metrics due to error " & Err.number & " (" & Err.Description & ")"
End Sub

Private Sub dtpBeginTime_CSRMetrics_LostFocus()
    Dim TempMonth, TempYear, TempDate
    
    If dtpBeginTime_CSRMetrics.Value > dtpEndTime_CSRMetrics.Value Then
        Msg "Sorry. The beginning date should be earlier than the ending date"
        TempMonth = Month(dtpEndTime.Value)
        TempYear = Year(dtpEndTime.Value)
        TempDate = TempMonth & "/1/" & TempYear
        dtpBeginTime_CSRMetrics.Value = CDate(TempDate)
        TryToSetFocus dtpBeginTime_CSRMetrics
    End If

End Sub

Private Sub dtpBeginTime_LostFocus()
    Dim TempMonth, TempYear, TempDate
    
    If dtpBeginTime.Value > dtpEndTime.Value Then
        Msg "Sorry. The beginning date should be earlier than the ending date"
        TempMonth = Month(dtpEndTime.Value)
        TempYear = Year(dtpEndTime.Value)
        TempDate = TempMonth & "/1/" & TempYear
        dtpBeginTime.Value = CDate(TempDate)
        TryToSetFocus dtpBeginTime
    End If
End Sub

Private Sub dtpEndTime_CSRMetrics_LostFocus()
    If dtpBeginTime_CSRMetrics.Value > dtpEndTime_CSRMetrics.Value Then
        Msg "Sorry. The ending date should be later than the beginning date"
        dtpEndTime_CSRMetrics.Value = Format(Now, "mm/dd/yy 23:59:59")
        TryToSetFocus dtpEndTime_CSRMetrics
    End If

End Sub

Private Sub dtpEndTime_LostFocus()
    If dtpBeginTime.Value > dtpEndTime.Value Then
        Msg "Sorry. The ending date should be later than the beginning date"
        'PRN#96 OK
        dtpEndTime.Value = Format(Now, "mm/dd/yy 23:59:59")
        TryToSetFocus dtpEndTime
    End If
End Sub


Private Sub Form_Load()
    SetCaption "Management Tool"
    InitSalesSummary
    
    'smr load values when command button "Show Me the Money" is clicked and not on form load
    'cmdRefresh_Click
End Sub


Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    tabManagement.Height = Me.Height - 630
    tabManagement.Width = Me.Width - 285
    
    Frame1.Height = tabManagement.Height - 1600
    gdxCSRMetrics.Height = Frame1.Height - 900
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.UnloadTool m_lWindowID
End Sub


Public Sub DoShowHelp()
    ShowHelp "FManagement", True
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub InitSalesSummary()
    Dim TempMonth, TempYear, TempDate
    optOrders.Value = True

    'added filter 8/26/03 LR
    g_rstWhses.Filter = "transit = 0"
    LoadCombo cboWarehouse, g_rstWhses, "WhseID", "WhseKey", GetUserWhseKey
    g_rstWhses.Filter = adFilterNone

    'PRN#96 OK
    dtpEndTime.Value = Format(Now, "mm/dd/yy 23:59:59")
    TempMonth = Month(Now)
    TempYear = Year(Now)
    TempDate = TempMonth & "/1/" & TempYear
    dtpBeginTime.Value = CDate(TempDate)
    
    'PRN#204
    'added filter 8/26/03 LR
    g_rstWhses.Filter = "transit = 0"
    LoadCombo cboWarehouse_CSRMetrics, g_rstWhses, "WhseID", "WhseKey", GetUserWhseKey
    g_rstWhses.Filter = adFilterNone
    
    dtpEndTime_CSRMetrics.Value = Format(Now, "mm/dd/yy 23:59:59")
    dtpBeginTime_CSRMetrics.Value = CDate(TempDate)
End Sub



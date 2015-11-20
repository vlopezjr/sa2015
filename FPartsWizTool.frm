VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "MMRemark.ocx"
Begin VB.Form FPartsWizTool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parts Wiz"
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   10200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   10200
   Begin VB.Frame Frame3 
      Caption         =   "Detail"
      Height          =   4815
      Left            =   4920
      TabIndex        =   19
      Top             =   2220
      Width           =   5175
      Begin MMRemark.RemarkViewer rvItem 
         Height          =   1032
         Left            =   960
         TabIndex        =   14
         Top             =   2700
         Width           =   1092
         _ExtentX        =   1931
         _ExtentY        =   1826
         ContextID       =   "ViewOrderLine"
         Caption         =   "Item Remarks"
      End
      Begin VB.TextBox txtPrivateRem 
         Appearance      =   0  'Flat
         Height          =   552
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3540
         Width           =   4752
      End
      Begin VB.TextBox txtPublicRem 
         Appearance      =   0  'Flat
         Height          =   552
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2700
         Width           =   4812
      End
      Begin VB.Label lblVendor 
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   960
         TabIndex        =   55
         Top             =   2040
         Width           =   3972
      End
      Begin VB.Label Label7 
         Caption         =   "Vendor"
         Height          =   252
         Left            =   120
         TabIndex        =   54
         Top             =   2040
         Width           =   672
      End
      Begin VB.Label Label8 
         Caption         =   "CustID"
         Height          =   252
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   612
      End
      Begin VB.Label lblCustID 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   50
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label lblWarning 
         Caption         =   "WARNING: THIS PART WAS SUBSEQUENTLY RETURNED BY THIS CUSTOMER."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   4200
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label lblCSR 
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   3120
         TabIndex        =   39
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "CSR"
         Height          =   252
         Left            =   2580
         TabIndex        =   38
         Top             =   1680
         Width           =   372
      End
      Begin VB.Label lblVendPN 
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   3120
         TabIndex        =   37
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblSerial 
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   3120
         TabIndex        =   36
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblModel 
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   3120
         TabIndex        =   35
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblMake 
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   3120
         TabIndex        =   34
         Top             =   240
         Width           =   1812
      End
      Begin VB.Label lblCost 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   33
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   32
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label lblOrder 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   31
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label lblAcct 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   30
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Vend P/N"
         Height          =   252
         Left            =   2220
         TabIndex        =   29
         Top             =   1320
         Width           =   732
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial"
         Height          =   252
         Left            =   2457
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Model"
         Height          =   252
         Left            =   2457
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Make"
         Height          =   252
         Left            =   2457
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Cost"
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   372
      End
      Begin VB.Label lblPrivateRem 
         Caption         =   "Private Remarks"
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   3300
         Width           =   1452
      End
      Begin VB.Label lblPublicRem 
         Caption         =   "Public Remarks"
         Height          =   252
         Left            =   120
         TabIndex        =   23
         Top             =   2460
         Width           =   1572
      End
      Begin VB.Label Label6 
         Caption         =   "Date"
         Height          =   252
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   492
      End
      Begin VB.Label Label5 
         Caption         =   "Order"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   492
      End
      Begin VB.Label Label4 
         Caption         =   "VAX Acct"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   795
      End
   End
   Begin GridEX20.GridEX gdxParts 
      Height          =   4812
      Left            =   120
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2220
      Width           =   4692
      _ExtentX        =   8281
      _ExtentY        =   8493
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      ShowToolTips    =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "FPartsWizTool.frx":0000
      Column(2)       =   "FPartsWizTool.frx":0190
      Column(3)       =   "FPartsWizTool.frx":0300
      Column(4)       =   "FPartsWizTool.frx":0468
      Column(5)       =   "FPartsWizTool.frx":05D4
      SortKeysCount   =   2
      SortKey(1)      =   "FPartsWizTool.frx":0748
      SortKey(2)      =   "FPartsWizTool.frx":07B0
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPartsWizTool.frx":0818
      FormatStyle(2)  =   "FPartsWizTool.frx":08F8
      FormatStyle(3)  =   "FPartsWizTool.frx":0A30
      FormatStyle(4)  =   "FPartsWizTool.frx":0AE0
      FormatStyle(5)  =   "FPartsWizTool.frx":0B94
      FormatStyle(6)  =   "FPartsWizTool.frx":0C6C
      ImageCount      =   0
      PrinterProperties=   "FPartsWizTool.frx":0D24
   End
   Begin VB.Frame Frame2 
      Caption         =   "Model/Serial/Description"
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   60
      Width           =   9975
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   7320
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   8640
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   8640
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Frame Frame4 
         Height          =   400
         Left            =   2880
         TabIndex        =   45
         Top             =   990
         Width           =   4095
         Begin VB.OptionButton optSerialOmit 
            Caption         =   "Omit"
            Height          =   255
            Left            =   3120
            TabIndex        =   7
            Top             =   120
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optSContains 
            Caption         =   "Contains"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optSBegins 
            Caption         =   "Begins With"
            Height          =   255
            Left            =   1560
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   400
         Left            =   2880
         TabIndex        =   44
         Top             =   600
         Width           =   2895
         Begin VB.OptionButton optMBegins 
            Caption         =   "Begins With"
            Height          =   255
            Left            =   1560
            TabIndex        =   3
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton optMContains 
            Caption         =   "Contains"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox txtDescr3 
         Height          =   285
         Left            =   7200
         TabIndex        =   10
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtDescr2 
         Height          =   285
         Left            =   4200
         TabIndex        =   9
         Top             =   1560
         Width           =   2535
      End
      Begin VB.ComboBox cboMake 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtDescr 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtSerial 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtModel 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblTooMany 
         Alignment       =   1  'Right Justify
         Caption         =   "WARNING:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   0
         Left            =   2880
         TabIndex        =   49
         Top             =   180
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label lblTooMany 
         Caption         =   "Only the first 100 matches are shown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   216
         Index           =   2
         Left            =   3960
         TabIndex        =   48
         Top             =   396
         Visible         =   0   'False
         Width           =   3372
      End
      Begin VB.Label lblTooMany 
         Caption         =   "Your search is too general.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   1
         Left            =   3960
         TabIndex        =   47
         Top             =   180
         Visible         =   0   'False
         Width           =   3492
      End
      Begin VB.Label Label22 
         Caption         =   "OR"
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
         Left            =   3840
         TabIndex        =   43
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "OR"
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
         Left            =   6840
         TabIndex        =   42
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "Make"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Serial"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Model"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "FPartsWizTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bModal As Boolean
Private m_oRst As ADODB.Recordset
Private m_lWindowID As Long
Private m_bSelectedPart As Boolean
Private m_sSelectedPartNbr As String


'Public method of the Modal version of PartsWiz
'Invoked by FOrder button press

Public Function FindPart( _
        ByVal i_sDescr As String, _
        ByVal i_sModel As String, _
        ByVal i_sSerial As String, _
        ByVal i_lMakeKey As Long, _
        ByRef o_sPartNbr As String _
) As Boolean
    txtDescr.Text = i_sDescr
    txtModel = i_sModel
    txtSerial = i_sSerial
    SetComboByKey cboMake, i_lMakeKey
    m_bModal = True
    cmdOK.Visible = True
    cmdOK.Enabled = False
    cmdCancel.Visible = True
    
    Show vbModal
    FindPart = m_bSelectedPart
    o_sPartNbr = m_sSelectedPartNbr
    Unload Me
End Function


Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property


Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property


Public Sub SetCaption(ByRef i_sTitle As String)
    Me.Caption = i_sTitle
    If m_lWindowID > 0 Then
        MDIMain.UpdateCaption Me
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    m_bSelectedPart = False
    
    SetCaption "Parts Wiz"
    Helpers.LoadCombo cboMake, g_rstMakes, "MakeText", "MakeID"
    
    'turn off all Remark controls
    lblPublicRem.Visible = False
    txtPublicRem.Visible = False
    lblPrivateRem.Visible = False
    txtPrivateRem.Visible = False
    rvItem.Visible = False

    Exit Sub

ErrorHandler:
    msg Err.Description, vbCritical, "Oops"
End Sub


Private Sub Form_Activate()
    If m_lWindowID > 0 Then
        MDIMain.UpdateWindowListSelection Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If m_lWindowID > 0 Then
        MDIMain.UnloadTool m_lWindowID
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Form.KeyPreview needs to be set to TRUE
    If KeyCode = vbKeyF1 Then
        DoShowHelp
    Else
        MDIMain.GlobalKeyDownProcessing KeyCode, Shift
    End If
End Sub

Public Sub DoShowHelp()
    ShowHelp "PartsWiz", m_bModal
End Sub


Private Sub cmdCancel_Click()
    CloseSelf
End Sub

Private Sub cmdOK_Click()
    m_bSelectedPart = True
    CloseSelf
End Sub

Private Sub CloseSelf()
    If m_bModal Then
        Hide
    Else
        Unload Me
    End If
End Sub

Private Sub cmdGo_Click()
    Dim sSQL As String
    Dim sWhere As String

    SetWaitCursor True


    sWhere = WherePhrase(cboMake.ItemData(cboMake.ListIndex), ScrubText(txtModel.Text), ScrubText(txtSerial.Text), PrepSQLText(txtDescr.Text), PrepSQLText(txtDescr2.Text), PrepSQLText(txtDescr3.Text))
    
    If sWhere <> "" Then
        sSQL = "SET rowcount 100 SELECT ID, PartNbr, Descr, CabModel, CabSerial, DataSource FROM vwOPLineData WHERE " & sWhere & " ORDER BY CabModel, CabSerial SET ROWCOUNT 0 "
    Else
        sSQL = "SET rowcount 100 SELECT ID, PartNbr, Descr, CabModel, CabSerial, DataSource FROM vwOPLineData ORDER BY CabModel, CabSerial SET ROWCOUNT 0 "
    End If

'   9/18/09 LR a timestamped userID and query string every time Parts Wiz is interrogated
    LogEvent "FPartsWizTool", "cmdGo_Click", Left(sSQL, 512)

    Set m_oRst = LoadDiscRst(sSQL)
    With gdxParts
        .HoldFields
        Set .ADORecordset = m_oRst
    End With

    ClearDetails

    PopLowerGrid
        
    If m_oRst.RecordCount = 100 Then
        lblTooMany(0).Visible = True
        lblTooMany(1).Visible = True
        lblTooMany(2).Visible = True
    Else
        lblTooMany(0).Visible = False
        lblTooMany(1).Visible = False
        lblTooMany(2).Visible = False
    End If
    
    SetWaitCursor False

End Sub


Private Function WherePhrase(lMakeID As Long, sModel As String, sSerial As String, sDescr1 As String, sDescr2 As String, sDescr3 As String) As String
    Dim sTemp As String

    'The first selection in the Make combobox is "<none>"
    If lMakeID <> 1 Then
        sTemp = "MakeID = " & CStr(lMakeID)
    End If
    
    If sModel <> "" Then
        If sTemp <> "" Then
            If optMContains Then
                sTemp = sTemp & " AND ModModel Like '%" & sModel & "%'"
            Else
                sTemp = sTemp & " AND ModModel Like '" & sModel & "%'"
            End If
        Else
            If optMContains Then
                sTemp = "ModModel Like '%" & sModel & "%'"
            Else
                sTemp = "ModModel Like '" & sModel & "%'"
            End If
        End If
    End If
    
    If Not optSerialOmit.value Then
        If sSerial <> "" Then
            If sTemp <> "" Then
                If optSContains Then
                    sTemp = sTemp & " AND ModSerial Like '%" & sSerial & "%'"
                Else
                    sTemp = sTemp & " AND ModSerial Like '" & sSerial & "%'"
                End If
            Else
                If optSContains Then
                    sTemp = "ModSerial Like '%" & sSerial & "%'"
                Else
                    sTemp = "ModSerial Like '" & sSerial & "%'"
                End If
            End If
        End If
    End If
    
    If sDescr1 <> "" Then
        If sTemp <> "" Then
            sTemp = sTemp & " AND ((" & AndDescr(sDescr1) & ")"
            If sDescr2 <> "" Then
                sTemp = sTemp & " OR (" & AndDescr(sDescr2) & ")"
                If sDescr3 <> "" Then
                    sTemp = sTemp & " OR (" & AndDescr(sDescr3) & ")"
                End If
            End If
            sTemp = sTemp & ")"
        Else
            sTemp = " (" & AndDescr(sDescr1) & ")"
            If sDescr2 <> "" Then
                sTemp = sTemp & " OR (" & AndDescr(sDescr2) & ")"
                If sDescr3 <> "" Then
                    sTemp = sTemp & " OR (" & AndDescr(sDescr3) & ")"
                End If
            End If
        
        End If
        
    End If
    WherePhrase = sTemp
End Function


Private Function AndDescr(sInput As String) As String
    Dim sTemp As String
    sTemp = sInput
    sTemp = Trim(sTemp)
    sTemp = Replace(sTemp, " ", "%' AND Descr Like '%")
    sTemp = "Descr Like '%" & sTemp & "%'"
    AndDescr = sTemp
End Function


Private Sub PopLowerGrid()
    Dim rst As ADODB.Recordset
    
    'If no data in grid, bail...
    With gdxParts
        If .Row > 0 Then
            Dim lRow As Long
            Dim lRowIndex As Long
            
            For lRow = .Row To .Row + 3
                lRowIndex = .RowIndex(lRow)
                If lRowIndex > 0 Then
                    GoTo GetData
                End If
            Next
        End If

GetData:
        If lRowIndex <= 0 Then
            m_sSelectedPartNbr = ""
            If cmdOK.Visible Then cmdOK.Enabled = False
            ClearDetails
            Exit Sub
        Else
            m_oRst.Bookmark = .RowBookmark(lRowIndex)
            m_sSelectedPartNbr = m_oRst.Fields("PartNbr").value
            If cmdOK.Visible Then cmdOK.Enabled = True
        End If
    End With

    'Set rst = LoadRst("SELECT * FROM topLineData WHERE ID = " & m_oRst(0).Value, , , , adOpenDynamic)
    Set rst = LoadRst("SELECT * FROM vwOPLineData " _
        & "WHERE ID=" & m_oRst("ID").value & " AND " _
        & "DataSource=" & m_oRst("DataSource"), , , , adOpenDynamic)
    
    With rst
        If Not .EOF Then
            lblCustID.Caption = Format(.Fields("CustID").value)
            lblAcct.Caption = Format(.Fields("AcctNbr").value)
            lblOrder.Caption = Format(.Fields("OrderNbr").value)
            lblDate.Caption = IIf(IsNull(.Fields("OrdDate")), "", Format$(.Fields("OrdDate"), "mm/dd/yy"))
            lblCost.Caption = IIf(IsNull(.Fields("Cost")), "", Format$(.Fields("Cost"), "####.00"))
            lblVendor.Caption = Format(.Fields("VendorName").value)

            If .Fields("DataSource").value = 1 Then
                rvItem.Visible = False
                lblPublicRem.Visible = True
                txtPublicRem.Visible = True
                txtPublicRem.Text = Format(.Fields("DisRem1").value) _
                                        & Format(.Fields("DisRem2").value)
                lblPrivateRem.Visible = True
                txtPrivateRem.Visible = True
                txtPrivateRem.Text = Format(.Fields("InsRem1")) _
                                        & Format(.Fields("InsRem2"))
            Else
                lblPublicRem.Visible = False
                txtPublicRem.Visible = False
                lblPrivateRem.Visible = False
                txtPrivateRem.Visible = False
                rvItem.Visible = True
                rvItem.OwnerID = .Fields("ID").value
            End If
            
            lblMake.Caption = Format(.Fields("CabMake"))
            lblModel.Caption = Format(.Fields("CabModel"))
            lblSerial.Caption = Format(.Fields("CabSerial"))
            lblVendPN.Caption = Format(.Fields("VendPartNbr"))
            lblCSR.Caption = Format(.Fields("CSR"))
            
            lblWarning.Visible = (IIf(IsNull(.Fields("WasReturned").value), 0, .Fields("WasReturned").value) > 0)
            'If .Fields("WasReturned").Value > 0 Then
                'lblWarning.Visible = True
            'End If
        Else
            ClearDetails
        End If
    End With
    
    CloseRst rst
End Sub


Private Sub ClearDetails()
    lblCustID.Caption = ""
    lblAcct.Caption = ""
    lblOrder.Caption = ""
    lblDate.Caption = ""
    lblCost.Caption = ""
    lblCSR.Caption = ""
    lblVendor.Caption = ""
    
    txtPublicRem.Text = ""
    txtPrivateRem.Text = ""
    
    lblMake.Caption = ""
    lblModel.Caption = ""
    lblSerial.Caption = ""
    lblVendPN.Caption = ""
    lblCSR.Caption = ""
    lblWarning.Visible = False
End Sub


Private Sub gdxParts_SelectionChange()
    PopLowerGrid
End Sub


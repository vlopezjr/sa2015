VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{24AB01F6-31FE-4657-A1CE-602D689527F9}#1.0#0"; "MMRemark.ocx"
Begin VB.Form FPartsWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts Wiz"
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   10200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Detail"
      Height          =   4692
      Left            =   4920
      TabIndex        =   19
      Top             =   2280
      Width           =   5175
      Begin MMRemark.RemarkViewer rvItem 
         Height          =   1152
         Left            =   1560
         TabIndex        =   14
         Top             =   2460
         Width           =   1212
         _ExtentX        =   2143
         _ExtentY        =   2037
         ContextID       =   "ViewOrderLine"
         Caption         =   "Item Remarks"
      End
      Begin VB.TextBox txtPrivateRem 
         Appearance      =   0  'Flat
         Height          =   492
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3420
         Width           =   4692
      End
      Begin VB.TextBox txtPublicRem 
         Appearance      =   0  'Flat
         Height          =   492
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2700
         Width           =   4692
      End
      Begin VB.Label Label19 
         Caption         =   "Vendor"
         Height          =   192
         Left            =   120
         TabIndex        =   53
         Top             =   2100
         Width           =   612
      End
      Begin VB.Label lblVendor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   52
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label lblCustID 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   51
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "CustID"
         Height          =   252
         Left            =   120
         TabIndex        =   50
         Top             =   300
         Width           =   612
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
         Caption         =   "CSR"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   1680
         Width           =   375
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
         Width           =   1815
      End
      Begin VB.Label lblCost 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   33
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   32
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label lblOrder 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   31
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lblAcct 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   30
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label15 
         Caption         =   "Vend P/N"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Serial"
         Height          =   255
         Left            =   2520
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Model"
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Make"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Cost"
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   372
      End
      Begin VB.Label lblPrivateRem 
         Caption         =   "Private Remarks"
         Height          =   252
         Left            =   240
         TabIndex        =   24
         Top             =   3240
         Width           =   1452
      End
      Begin VB.Label lblPublicRem 
         Caption         =   "Public Remarks"
         Height          =   252
         Left            =   240
         TabIndex        =   23
         Top             =   2520
         Width           =   1572
      End
      Begin VB.Label Label6 
         Caption         =   "Date"
         Height          =   252
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   492
      End
      Begin VB.Label Label5 
         Caption         =   "Order"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   492
      End
      Begin VB.Label Label4 
         Caption         =   "VAX Acct"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   660
         Width           =   795
      End
   End
   Begin GridEX20.GridEX gdxParts 
      Height          =   4692
      Left            =   120
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2280
      Width           =   4692
      _ExtentX        =   8281
      _ExtentY        =   8281
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
      Column(1)       =   "FPartsWiz.frx":0000
      Column(2)       =   "FPartsWiz.frx":0190
      Column(3)       =   "FPartsWiz.frx":0300
      Column(4)       =   "FPartsWiz.frx":0468
      Column(5)       =   "FPartsWiz.frx":05D4
      SortKeysCount   =   2
      SortKey(1)      =   "FPartsWiz.frx":0748
      SortKey(2)      =   "FPartsWiz.frx":07B0
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPartsWiz.frx":0818
      FormatStyle(2)  =   "FPartsWiz.frx":08F8
      FormatStyle(3)  =   "FPartsWiz.frx":0A30
      FormatStyle(4)  =   "FPartsWiz.frx":0AE0
      FormatStyle(5)  =   "FPartsWiz.frx":0B94
      FormatStyle(6)  =   "FPartsWiz.frx":0C6C
      ImageCount      =   0
      PrinterProperties=   "FPartsWiz.frx":0D24
   End
   Begin VB.Frame Frame2 
      Caption         =   "Model/Serial/Description"
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   120
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
         Left            =   7260
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
Attribute VB_Name = "FPartsWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oRst As ADODB.Recordset
Private m_lWindowID As Long
Private m_bSelectedPart As Boolean
Private m_sSelectedPartNbr As String
Private m_sSelectedModal As String
Private m_sSelectedDescr As String
Private m_sSelectedSerialNbr As String
Private m_lSelectedMake As Long
Private m_lSelectedVendor As Long
Private m_lOPID As Long

'Public method of the Modal version of PartsWiz
'Invoked by FOrder button press


'10/08/02           TeddyX
'Add new parameters in this function to load the modal, serial nbr, partnbr,
'descr, and make to FOrder.

'10/21/02
'Only descr, part number, and make are needed in PartzWiz
Public Function FindPart( _
        ByVal i_sDescr As String, _
        ByVal i_sModel As String, _
        ByVal i_sSerial As String, _
        ByVal i_lMakeKey As Long, _
        ByVal i_lOPID As Long, _
        ByRef o_sPartNbr As String, _
        ByRef o_lMake As Long, _
        ByRef o_sDescr As String _
) As Boolean
    txtDescr.text = i_sDescr
    txtModel = i_sModel
    txtSerial = i_sSerial
    m_lOPID = i_lOPID
    SetComboByKey cboMake, i_lMakeKey
    cmdOK.Visible = True
    cmdOK.Enabled = False
    cmdCancel.Visible = True
    
    Show vbModal
    FindPart = m_bSelectedPart
    o_sPartNbr = m_sSelectedPartNbr
    o_sDescr = m_sSelectedDescr
    o_lMake = m_lSelectedMake
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
    ShowHelp "PartsWiz", True
End Sub


Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    m_bSelectedPart = True
    Hide
End Sub


Private Function EventString() As String
    Dim sEvent As String
    
    sEvent = "PartsWiz Search: " & Trim(cboMake.text) & ", " & Trim(txtModel.text) & ", " & Trim(txtDescr.text)
    
    If Trim(txtDescr2.text) <> "" Then
        sEvent = sEvent & " or " & Trim(txtDescr2.text)
    End If
    
    If Trim(txtDescr3.text) <> "" Then
        sEvent = sEvent & " or " & Trim(txtDescr3.text)
    End If
    
    EventString = sEvent
End Function


Private Sub cmdGo_Click()
    Dim sSQL As String
    Dim sWhere As String

    SetWaitCursor True
    
    '10/15/2002         TeddyX
    'Use PrepSQLText() to guard the error of inputting single input
    
    sWhere = WherePhrase(cboMake.ItemData(cboMake.ListIndex), ScrubText(txtModel.text), ScrubText(txtSerial.text), PrepSQLText(txtDescr.text), PrepSQLText(txtDescr2.text), PrepSQLText(txtDescr3.text))
    
    '10/08/2002         TeddyX
    'Add the MakeID and Vendor to the select clause
    If sWhere <> "" Then
        'sSQL = "SET rowcount 100 SELECT ID, PartNbr, Descr, CabModel, CabSerial FROM topLineData WHERE " & sWhere & " ORDER BY CabModel, CabSerial"
        sSQL = "SET rowcount 100 SELECT ID, isnull(Vendor, 0) as Vendor, MakeID, PartNbr, isnull(Descr, '') as Descr, isnull(CabModel, '') as CabModel, isnull(CabSerial, '') as CabSerial, DataSource FROM vwOPLineData WHERE " & sWhere & " ORDER BY CabModel, CabSerial SET ROWCOUNT 0 "
    Else
        'sSQL = "SET rowcount 100 SELECT ID, PartNbr, Descr, CabModel, CabSerial FROM topLineData ORDER BY CabModel, CabSerial"
        sSQL = "SET rowcount 100 SELECT ID, isnull(Vendor, 0) as Vendor, MakeID, PartNbr, isnull(Descr, '') as Descr, isnull(CabModel, '') as CabModel, isnull(CabSerial, '') as CabSerial, DataSource FROM vwOPLineData ORDER BY CabModel, CabSerial SET ROWCOUNT 0 "
    End If
    
'   9/18/09 LR a timestamped userID and query string every time Parts Wiz is interrogated
    LogEvent "FPartsWiz", "cmdGo_Click", Left(sSQL, 512)
    
    Set m_oRst = LoadDiscRst(sSQL)

    'PRN#96 Log Event for PartsWiz searching
    LogOAEvent "Order", GetUserID, m_lOPID, , , EventString
   
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
            m_sSelectedModal = ""
            m_sSelectedDescr = ""
            m_sSelectedSerialNbr = ""
            m_lSelectedMake = 0
            m_lSelectedVendor = 0
            If cmdOK.Visible Then cmdOK.Enabled = False
            ClearDetails
            Exit Sub
        Else
            '10/08/2002         TeddyX
            'Retrieve more item information from the recordset
            m_oRst.Bookmark = .RowBookmark(lRowIndex)
            m_sSelectedPartNbr = Trim(m_oRst.Fields("PartNbr").value)
            m_sSelectedModal = Trim(m_oRst.Fields("CabModel").value)
            m_sSelectedDescr = Trim(m_oRst.Fields("Descr").value)
            m_sSelectedSerialNbr = Trim(m_oRst.Fields("CabSerial").value)
            m_lSelectedMake = m_oRst.Fields("MakeID").value
            m_lSelectedVendor = m_oRst.Fields("Vendor").value
            If cmdOK.Visible Then cmdOK.Enabled = True
        End If
    End With

    'Set rst = LoadRst("SELECT * FROM topLineData WHERE ID = " & m_oRst(0).Value, , , , adOpenDynamic)
    Set rst = LoadRst("SELECT * FROM vwOPLineData " _
        & "WHERE ID=" & m_oRst("ID").value & " AND " _
        & "DataSource=" & m_oRst("DataSource"), , , , adOpenDynamic)
 
    With rst
         If Not .EOF Then
            lblCustID.Caption = Format(.Fields("CustID"))
            lblAcct.Caption = Format(.Fields("AcctNbr"))
            lblOrder.Caption = Format(.Fields("OrderNbr"))
            lblDate.Caption = IIf(IsNull(.Fields("OrdDate")), "", Format$(.Fields("OrdDate"), "mm/dd/yy"))
            lblCost.Caption = IIf(IsNull(.Fields("Cost")), "", Format$(.Fields("Cost"), "####.00"))
            lblVendor.Caption = Format(.Fields("VendorName"))

            If .Fields("DataSource").value = 1 Then
                rvItem.Visible = False
                lblPublicRem.Visible = True
                txtPublicRem.Visible = True
                txtPublicRem.text = Format(.Fields("DisRem1")) _
                                        & Format(.Fields("DisRem2"))
                lblPrivateRem.Visible = True
                txtPrivateRem.Visible = True
                txtPrivateRem.text = Format(.Fields("InsRem1")) _
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

    txtPublicRem.text = ""
    txtPrivateRem.text = ""
    
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



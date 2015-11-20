VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FPOSearch 
   Caption         =   "PO Search"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optSTL 
      Caption         =   "STL"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.OptionButton optSEA 
      Caption         =   "SEA"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton optMPK 
      Caption         =   "MPK"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton optALL 
      Caption         =   "All"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtPartNbr 
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin GridEX20.GridEX gdxPOs 
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7223
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FPOSearch.frx":0000
      Column(2)       =   "FPOSearch.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FPOSearch.frx":016C
      FormatStyle(2)  =   "FPOSearch.frx":024C
      FormatStyle(3)  =   "FPOSearch.frx":0384
      FormatStyle(4)  =   "FPOSearch.frx":0434
      FormatStyle(5)  =   "FPOSearch.frx":04E8
      FormatStyle(6)  =   "FPOSearch.frx":05C0
      ImageCount      =   0
      PrinterProperties=   "FPOSearch.frx":0678
   End
   Begin VB.ComboBox cboVendors 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblPartNbr 
      Caption         =   "Part Number"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Vendor"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FPOSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'smr - 11/19/2004
'Private m_POKey As Integer
Private m_POKey As Long

Private WithEvents m_gw As GridEXWrapper
Attribute m_gw.VB_VarHelpID = -1


Public Function GetPOKey() As Long
'smr - 11/19/2004
'Public Function GetPOKey() As Integer
    m_POKey = 0
    Me.Show vbModal
    GetPOKey = m_POKey
    Unload Me
End Function


Private Sub cmdFind_Click()
    Dim orst As ADODB.Recordset
    Dim vVendKey As Variant
    Dim vWhseKey As Variant
    Dim vItemID As Variant
    Dim i As Integer
    
    Dim sSQL As String
    
    sSQL = "Exec spCPCAPFindPO"
    
    If cboVendors.Text <> "<none>" Then
        sSQL = sSQL & " @i_VendKey = " & CStr(cboVendors.ItemData(cboVendors.ListIndex)) & ","
    End If

    If optMPK.value Then
        sSQL = sSQL & " @i_WhseKey = " & g_MPKWhseKey & ","
    End If
    
    If optSEA.value Then
        sSQL = sSQL & " @i_WhseKey = " & g_SEAWhseKey & ","
    End If
    
    If optSTL.value Then
        sSQL = sSQL & " @i_WhseKey = " & g_STLWhseKey & ","
    End If
    
    If Len(txtPartNbr.Text) > 0 Then
        sSQL = sSQL & " @i_ItemID = '" & PrepSQLText(Trim(txtPartNbr.Text)) & "',"
    End If
    
    'remove the last comma
    sSQL = Left$(sSQL, Len(sSQL) - 1)
    
    Set orst = LoadDiscRst(sSQL)
    
    Set gdxPOs.ADORecordset = orst
    
    gdxPOs.Columns(1).Visible = False
    For i = 2 To gdxPOs.Columns.Count
        gdxPOs.Columns(i).AutoSize
    Next
    
    Set orst = Nothing
End Sub

Private Sub Form_Load()
    Helpers.LoadCombo cboVendors, g_rstVendors, "VendName", "VendKey", , True
    Set m_gw = New GridEXWrapper
    m_gw.Grid = gdxPOs
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '4/7/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gw = Nothing
End Sub


Private Sub gdxPOs_DblClick()
    m_POKey = m_gw.value("POKey")
    Me.Hide
End Sub

Private Sub m_gw_RowChosen()
    m_POKey = m_gw.value("POKey")
    Me.Hide
End Sub

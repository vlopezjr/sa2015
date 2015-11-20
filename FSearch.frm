VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2565
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   5220
      TabIndex        =   1
      Top             =   2100
      Width           =   972
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Default         =   -1  'True
      Height          =   372
      Left            =   4200
      TabIndex        =   0
      Top             =   2100
      Width           =   972
   End
   Begin GridEX20.GridEX gdxVendors 
      Height          =   1872
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6132
      _ExtentX        =   10821
      _ExtentY        =   3307
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   4
      Column(1)       =   "FSearch.frx":0000
      Column(2)       =   "FSearch.frx":014C
      Column(3)       =   "FSearch.frx":0264
      Column(4)       =   "FSearch.frx":0380
      FormatStylesCount=   6
      FormatStyle(1)  =   "FSearch.frx":049C
      FormatStyle(2)  =   "FSearch.frx":05D4
      FormatStyle(3)  =   "FSearch.frx":0684
      FormatStyle(4)  =   "FSearch.frx":0738
      FormatStyle(5)  =   "FSearch.frx":0810
      FormatStyle(6)  =   "FSearch.frx":08C8
      ImageCount      =   0
      PrinterProperties=   "FSearch.frx":08E8
   End
End
Attribute VB_Name = "FSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_gwVendors As GridEXWrapper
Attribute m_gwVendors.VB_VarHelpID = -1

Private m_bSelected As Boolean


Private Sub Form_Activate()
    'Before form displays, set focus to the first row of the grid
    With gdxVendors
        .SetFocus
        If .RowCount >= 1 Then
            .Row = 2
        End If
    End With
    Set m_gwVendors = New GridEXWrapper
    m_gwVendors.Grid = gdxVendors
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '4/7/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gwVendors = Nothing
End Sub


'NOTE: The output parameter o_sVendName is an expedient kludge for now.
'Better to return a Vendor object.

Public Function Find(ByVal i_sInput As String, ByRef o_sVendName As String) As Long
    Dim orst As ADODB.Recordset
    Dim sSQL As String
        
    SetWaitCursor True
    
    Me.Caption = "Vendors matching '" & Trim(i_sInput) & "*'"

    sSQL = "select vendkey, vendname, city, stateid, postalcode " & _
        "from tapvendor inner join tciaddress on tapvendor.dfltpurchaddrkey = tciaddress.addrkey " & _
        "where vendname like '" & PrepSQLText(Trim(i_sInput)) & "%'"
    Set orst = LoadDiscRst(sSQL)
    
    SetWaitCursor False
    
    Select Case orst.RecordCount
    Case 0
        Msg "We have no matching vendor."
        Find = 0
    Case 1
        o_sVendName = orst("VendName").value
        Find = orst("VendKey").value
    Case Else
        Find = SelectFromGrid(orst, "VendKey", o_sVendName)
    End Select
    
    Set orst = Nothing
    
    Unload Me
    
End Function


'NOTE: the o_sVendName parameter is an expedient kludge passed on from above (LR 9/9/02)

Private Function SelectFromGrid(ByRef i_orst As ADODB.Recordset, ByVal i_sKeyField As String, o_sVendName As String) As Long
    With gdxVendors
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = i_orst
        Dim i As Integer
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    
    Me.Show vbModal     'program blocks here waiting for user event

    If m_bSelected Then
        SelectFromGrid = m_gwVendors.value(i_sKeyField)
        o_sVendName = m_gwVendors.value(1)              'NOTE: expedient kludge
    End If
End Function


Private Sub cmdSelect_Click()
    With gdxVendors
        If .RowIndex(.Row) <= 0 Then
            Msg "Please select the desired vendor from the grid.", , "Select Vendor"
            Exit Sub
        End If
    End With
    m_bSelected = True
    Me.Hide
End Sub


Private Sub m_gwVendors_RowChosen()
    cmdSelect_Click
End Sub


Private Sub cmdCancel_Click()
    m_bSelected = False
    Me.Hide
End Sub



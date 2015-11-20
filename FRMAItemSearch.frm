VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FRMAItemSearch 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   210
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   210
      Width           =   1335
   End
   Begin GridEX20.GridEX gdxItemSearch 
      Height          =   3852
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7332
      _ExtentX        =   12938
      _ExtentY        =   6800
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      ScrollToolTips  =   -1  'True
      ShowToolTips    =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      CursorLocation  =   3
      ColumnAutoResize=   -1  'True
      ReadOnly        =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      CardCaptionPrefix=   "Customer Information"
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   270
      ColumnsCount    =   4
      Column(1)       =   "FRMAItemSearch.frx":0000
      Column(2)       =   "FRMAItemSearch.frx":0124
      Column(3)       =   "FRMAItemSearch.frx":023C
      Column(4)       =   "FRMAItemSearch.frx":0360
      FormatStylesCount=   6
      FormatStyle(1)  =   "FRMAItemSearch.frx":0494
      FormatStyle(2)  =   "FRMAItemSearch.frx":0574
      FormatStyle(3)  =   "FRMAItemSearch.frx":06AC
      FormatStyle(4)  =   "FRMAItemSearch.frx":075C
      FormatStyle(5)  =   "FRMAItemSearch.frx":0810
      FormatStyle(6)  =   "FRMAItemSearch.frx":08E8
      ImageCount      =   0
      PrinterProperties=   "FRMAItemSearch.frx":09A0
   End
   Begin VB.Label lblSearchRem 
      Caption         =   "Item Search"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   570
      Width           =   4335
   End
   Begin VB.Label lblTooMany 
      Caption         =   "Only the first 50 matches are shown"
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
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   330
      Visible         =   0   'False
      Width           =   3375
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
      Height          =   210
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FRMAItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lMinWidth = 7800
Private Const k_lMinHeight = 2400
Private m_bLoad As Boolean      'Did the user select 'Load'?
Private WithEvents m_gw As GridEXWrapper
Attribute m_gw.VB_VarHelpID = -1
Private m_bResize As Boolean


Public Sub Find(i_sCaption As String, i_sInput As String, i_sItemID As String, i_sItemDescr As String)
    Dim sSQL As String
    Dim sWhere As String
    Dim orst As ADODB.Recordset
    Dim sSearch As String
    Dim sItemID As String
    
    lblSearchRem.Caption = FormatCaption(i_sCaption & ": " & UCase(i_sInput))
            
    SetWaitCursor True
    Set orst = CallSP("spCPCRMASearchPart", "@_iPartInput", i_sInput)
    SetWaitCursor False
        
    Select Case orst.RecordCount
        Case Is = 0
            Msg "No records satisfy this request"
        Case Is = 1
            i_sItemID = orst.Fields("PartNbr").Value
            i_sItemDescr = orst.Fields("Descr").Value
        Case Else
            ChooseFromGrid orst, i_sItemID, i_sItemDescr
       End Select
    Set orst = Nothing
    
    Unload Me
End Sub


Private Sub cmdCancel_Click()
    m_bLoad = False
    Me.Hide
End Sub


Private Sub cmdLoad_Click()
    m_bLoad = True
    Me.Hide
End Sub


Private Sub Form_Activate()
    'Before form displays, set focus to first row of grid
    TryToSetFocus gdxItemSearch
    gdxItemSearch.Row = 1
End Sub


Private Sub Form_Load()
    Set m_gw = New GridEXWrapper
    m_gw.Grid = gdxItemSearch
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '4/7/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gw = Nothing
End Sub


Private Sub ChooseFromGrid(i_rst As ADODB.Recordset, i_sItemID As String, i_sItemDescr As String)
    If i_rst.RecordCount >= g_MaxItemRows Then
        lblTooMany(0).Visible = True
        lblTooMany(1).Visible = True
        lblTooMany(2).Visible = True
    End If
    
    With gdxItemSearch
        Dim i As Long
        
        .HoldFields
        .SortKeys.Clear
        .SortKeys.Add 1, jgexSortAscending 'sort by PartNbr
        .HoldSortSettings = True
        .ColumnAutoResize = False
        Set .ADORecordset = i_rst
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    
    Me.Show vbModal
    If m_bLoad Then
        i_sItemID = m_gw.Value("PartNbr")
        i_sItemDescr = m_gw.Value("Descr")
    End If
    Unload Me
End Sub


Private Sub Form_Resize()
    '09/09/2002 TeddyX
    'I doubt that the following error is caused by Form_Resize logic in FItemSearch.
    'I add protect flag to prevent Form_Resize from cascading calling itself.
    'I also disable the MinButton of this form because it will cause error in \
    'Form_Resize logic
    'Error Report Date: 09-06-2002 11:50:21
    'Module: Database.bas          Sub: LoadDiscRst          User: TerryR
    
    If m_bResize Then Exit Sub
    
    
    Dim i As Long
    Dim lBorder As Long

    m_bResize = True
    If Me.Width < k_lMinWidth Then Me.Width = k_lMinWidth
    If Me.Height < k_lMinHeight Then Me.Height = k_lMinHeight
    lBorder = 120
    
    With gdxItemSearch
        .Width = Me.Width - (2 * lBorder + 120)
        .Height = Me.Height - (cmdLoad.Top + cmdLoad.Height + 580 + lBorder)
        For i = 1 To gdxItemSearch.Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    
    cmdCancel.Left = Me.Width - (cmdCancel.Width + 2 * lBorder)
    cmdLoad.Left = cmdCancel.Left - (cmdLoad.Width + lBorder)
    m_bResize = False
End Sub


Private Sub m_gw_RowChosen()
    cmdLoad_Click
End Sub


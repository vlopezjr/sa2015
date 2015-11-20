VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FItemSearch 
   Caption         =   "Item Search"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   7590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin GridEX20.GridEX gdxItemSearch 
      Height          =   3852
      Left            =   120
      TabIndex        =   2
      Top             =   720
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
      ColumnsCount    =   3
      Column(1)       =   "FItemSearch.frx":0000
      Column(2)       =   "FItemSearch.frx":0124
      Column(3)       =   "FItemSearch.frx":023C
      FormatStylesCount=   6
      FormatStyle(1)  =   "FItemSearch.frx":0360
      FormatStyle(2)  =   "FItemSearch.frx":0440
      FormatStyle(3)  =   "FItemSearch.frx":0578
      FormatStyle(4)  =   "FItemSearch.frx":0628
      FormatStyle(5)  =   "FItemSearch.frx":06DC
      FormatStyle(6)  =   "FItemSearch.frx":07B4
      ImageCount      =   0
      PrinterProperties=   "FItemSearch.frx":086C
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
      Left            =   120
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   972
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
      Height          =   216
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   36
      Visible         =   0   'False
      Width           =   3492
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
      Height          =   252
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   3372
   End
   Begin VB.Label lblSearchRem 
      Caption         =   "Item Search"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4332
   End
End
Attribute VB_Name = "FItemSearch"
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



Public Sub Find( _
    i_sCaption As String, _
    i_sInput As String, _
    o_lItemKey As Long, _
    o_eItemType As ItemTypeCode, _
    o_sOriginalItemID As String, _
    o_sRefSource As String, _
    ByRef o_bCancelSearch As Boolean, _
    ByVal lWhseKey As Long, _
    ByVal bDoNotShowIfNoXRef As Boolean, _
    ByVal yCustType As Byte _
    )
    Dim orst As ADODB.Recordset
    Dim sSearch As String
    Dim frmXRef As FXRef
    Dim sItemId As String
    
    ' Format caption display
    lblSearchRem.Caption = FormatCaption(i_sCaption & ": " & UCase(i_sInput))
    
    ' Search in our active inventory, limited to finished good and BTO Kit
    sSearch = PrepSQLText(i_sInput)
    
    Select Case i_sCaption
    
        Case Is = k_sPartNbr
        
            SetWaitCursor True
            Set orst = CallSP("spcpcLookupPartNumber", "@ItemId", i_sInput, "@CleanItemId", Replace$(i_sInput, "-", ""))
            SetWaitCursor False
            
        Case Is = k_sPartDescr
        
            Dim sSQL As String
            sSQL = sSQL & "SET ROWCOUNT " & g_MaxItemRows _
                & " SELECT * FROM vwOPItemSearch WHERE Descr LIKE '%" & sSearch & "%'" _
                & " SET ROWCOUNT 0 "
            SetWaitCursor True
            Set orst = LoadDiscRst(sSQL)
            SetWaitCursor False
            
    End Select
        
    Set frmXRef = New FXRef
        
    ' Process the search result
    Select Case orst.RecordCount
        '
        Case Is = 0
            ' Test case: Item ID 60-109
            'Set ofrm = New FCrossRef
            'ofrm.XRefSearch sSearch, o_lItemKey, o_eItemType, o_sOriginalItemID, o_sRefSource
            frmXRef.XRefSearch sSearch, o_lItemKey, o_eItemType, o_sOriginalItemID, o_sRefSource, o_bCancelSearch, lWhseKey, bDoNotShowIfNoXRef, yCustType
            
        ' Found an exact match for this part
        Case Is = 1
        
            ' 11/13/03  We query with a wild card above, so this might not be an exact match.
            '           Check it before loading o_eItemType and o_lItemKey
            
            '2/4/15 LR
            'If UCase$(sSearch) = UCase$(Trim(orst.Fields("PartNbr").Value)) Then
            If UCase$(sSearch) = UCase$(Trim(orst.Fields("PartNbr").value)) Then
                ' FOUND EXACT MATCH
            
                ' Test case: Item ID 17170
                ' Set the item type & key
                o_eItemType = ConvertSageItemType(orst.Fields("ItemType").value)
                o_lItemKey = orst.Fields("ItemKey").value
                
                ' Send to XRef for processing
                '2/4/15 LR
                'frmXRef.XRefSearch orst.Fields("PartNbr").Value, o_lItemKey, o_eItemType, o_sOriginalItemID, o_sRefSource, o_bCancelSearch, lWhseKey, bDoNotShowIfNoXRef, yCustType
                frmXRef.XRefSearch orst.Fields("PartNbr").value, o_lItemKey, o_eItemType, o_sOriginalItemID, o_sRefSource, o_bCancelSearch, lWhseKey, bDoNotShowIfNoXRef, yCustType
            Else
                ' 01/06/04
                
                ' FOUND A WILDCARD MATCH. Load result into grid, let user decide whether she wants it or not
                ChooseFromGrid orst, o_lItemKey, o_eItemType, sItemId
                frmXRef.XRefSearch sItemId, o_lItemKey, o_eItemType, o_sOriginalItemID, o_sRefSource, o_bCancelSearch, lWhseKey, bDoNotShowIfNoXRef, yCustType
                ' 01/06/04
            End If
        Case Else
            ' Found multiple matches in inventory, let user select an item before proceeding to xref
            ChooseFromGrid orst, o_lItemKey, o_eItemType, sItemId
            frmXRef.XRefSearch sItemId, o_lItemKey, o_eItemType, o_sOriginalItemID, o_sRefSource, o_bCancelSearch, lWhseKey, bDoNotShowIfNoXRef, yCustType
            
    End Select
    Set orst = Nothing
    
    Set frmXRef = Nothing
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Form.KeyPreveiw needs to be set to TRUE
    If KeyCode = vbKeyF1 Then
        DoShowHelp
    Else
        MDIMain.GlobalKeyDownProcessing KeyCode, Shift
    End If
End Sub

Public Sub DoShowHelp()
    ShowHelp "ItemSearch", True
End Sub

Private Sub Form_Load()
    Set m_gw = New GridEXWrapper
    m_gw.Grid = gdxItemSearch
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '4/7/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gw = Nothing
End Sub


Private Sub ChooseFromGrid(i_rst As ADODB.Recordset, o_lItemKey As Long, o_eItemType As ItemTypeCode, o_sItemID As String)
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
        o_lItemKey = m_gw.value("ItemKey")
        o_eItemType = ConvertSageItemType(m_gw.value("ItemType"))
        o_sItemID = Trim(m_gw.value("PartNbr"))
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
    If Me.width < k_lMinWidth Then Me.width = k_lMinWidth
    If Me.Height < k_lMinHeight Then Me.Height = k_lMinHeight
    lBorder = 120
    
    With gdxItemSearch
        .width = Me.width - (2 * lBorder + 120)
        .Height = Me.Height - (cmdLoad.Top + cmdLoad.Height + 580 + lBorder)
        For i = 1 To gdxItemSearch.Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    
    cmdCancel.Left = Me.width - (cmdCancel.width + 2 * lBorder)
    cmdLoad.Left = cmdCancel.Left - (cmdLoad.width + lBorder)
    m_bResize = False
End Sub

Private Sub m_gw_RowChosen()
    cmdLoad_Click
End Sub

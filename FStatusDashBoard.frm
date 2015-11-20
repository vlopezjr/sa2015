VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FStatusDashBoard 
   Caption         =   "Quote Status"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8820
   Begin VB.CommandButton cmdToggleView 
      Caption         =   "<>"
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.Timer timerRefresh 
      Interval        =   60000
      Left            =   1200
      Top             =   420
   End
   Begin VB.ComboBox cboEntity 
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   60
      Width           =   2175
   End
   Begin GridEX20.GridEX gdxDetail 
      Height          =   1995
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3519
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FStatusDashBoard.frx":0000
      Column(2)       =   "FStatusDashBoard.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FStatusDashBoard.frx":016C
      FormatStyle(2)  =   "FStatusDashBoard.frx":024C
      FormatStyle(3)  =   "FStatusDashBoard.frx":0384
      FormatStyle(4)  =   "FStatusDashBoard.frx":0434
      FormatStyle(5)  =   "FStatusDashBoard.frx":04E8
      FormatStyle(6)  =   "FStatusDashBoard.frx":05C0
      ImageCount      =   0
      PrinterProperties=   "FStatusDashBoard.frx":0678
   End
   Begin GridEX20.GridEX gdxSummary 
      Height          =   915
      Left            =   2340
      TabIndex        =   3
      Top             =   60
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   1614
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FStatusDashBoard.frx":0850
      Column(2)       =   "FStatusDashBoard.frx":0918
      FormatStylesCount=   6
      FormatStyle(1)  =   "FStatusDashBoard.frx":09BC
      FormatStyle(2)  =   "FStatusDashBoard.frx":0A9C
      FormatStyle(3)  =   "FStatusDashBoard.frx":0BD4
      FormatStyle(4)  =   "FStatusDashBoard.frx":0C84
      FormatStyle(5)  =   "FStatusDashBoard.frx":0D38
      FormatStyle(6)  =   "FStatusDashBoard.frx":0E10
      ImageCount      =   0
      PrinterProperties=   "FStatusDashBoard.frx":0EC8
   End
End
Attribute VB_Name = "FStatusDashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lMinWidth = 8940
Private Const k_lMinHeight = 1300

Private m_lWindowID As Long

Private isLoading As Boolean

Private showDetail As Boolean
Private detailGridTop As Integer

Private minutesElapsed As Integer

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

Private Sub cmdToggleView_Click()
    If showDetail Then
        showDetail = False
        'shrink form
        Me.Height = 525 + gdxSummary.Height + 120
    Else
        showDetail = True
        'expand form
        Me.Height = Me.Height + gdxDetail.Height
    End If
    cboEntity_Click
End Sub

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Load()
    
    SetCaption "Quote Status Dashboard"
    
    timerRefresh.Interval = 60000 '60 seconds
    
    showDetail = False
    Me.width = k_lMinWidth
    
    isLoading = True
    
    'load only the active CSRs
    g_rstCSRs.Filter = "IsActive=-1"
    LoadCombo cboEntity, g_rstCSRs, "UserID", "UserKey"
    g_rstUsers.Filter = adFilterNone
    
    cboEntity.AddItem "<ALL>"
    cboEntity.AddItem "<MPK>"
    cboEntity.AddItem "<SEA>"
    cboEntity.AddItem "<STL>"
    
    isLoading = False
    
    If InStr(1, GetUserName, "LAWillCall", vbTextCompare) > 0 Then
        SetComboByText cboEntity, "<MPK>"
    Else
        SetComboByText cboEntity, GetUserName
    End If

    cboEntity_Click
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub

    'If Me.width < k_lMinWidth Then Me.width = k_lMinWidth
    'If Me.Height < k_lMinHeight Then Me.Height = k_lMinHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.FormUnregister Me
End Sub


'map recordset directly to grid
'no control over column headers or order

Private Sub cboEntity_Click()
    
    If isLoading Then Exit Sub

    LoadGrids
    
    'adjust the height of the summary grid
    gdxSummary.Height = 318 + gdxSummary.RowCount * (gdxSummary.RowHeight + 4)
    
    'alter the top of the detail grid
    detailGridTop = 60 + gdxSummary.Height + 120
    gdxDetail.Top = detailGridTop
    
    'adjust the height of the detail grid
    gdxDetail.Height = 318 + gdxDetail.RowCount * (gdxDetail.RowHeight + 4)
    
    If showDetail Then
        Me.Height = 675 + gdxSummary.Height + 240 + gdxDetail.Height + 120
    Else
        Me.Height = 675 + gdxSummary.Height
    End If
End Sub


Private Sub LoadGrids()
    Dim orst As ADODB.Recordset
    Dim i As Integer
    Dim entity As String
    Dim mode As String
    
    On Error GoTo EH
    
    entity = cboEntity.text
    entity = Replace(entity, "<", "")
    entity = Replace(entity, ">", "")
    
    Set orst = CallSP("spCPCGetSOStatusDash", "@Entity", entity, "@Mode", "S")
    
    With gdxSummary
        Set .ADORecordset = orst
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With

    Set orst = CallSP("spCPCGetSOStatusDash", "@Entity", entity, "@Mode", "D")
    With gdxDetail
        Set .ADORecordset = orst
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    
    Exit Sub
EH:
    LogDB.LogError "FStatusDashBoard", "LoadGrids", "Data Access Error", Err.Source, Err.Number, Err.Description

End Sub

Private Sub timerRefresh_Timer()

    minutesElapsed = minutesElapsed + 1
    
    If minutesElapsed = 5 Then
        timerRefresh.Enabled = False
        LoadGrids
        minutesElapsed = 0
        timerRefresh.Enabled = True
    End If
End Sub


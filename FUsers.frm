VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FUsers 
   Caption         =   "Users"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   2910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   2910
   Begin VB.Frame frmRefresh 
      Caption         =   "AutoRefresh"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   2655
      Begin VB.CheckBox chkAutoRefresh 
         Caption         =   "On"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         OrigLeft        =   360
         OrigTop         =   720
         OrigRight       =   600
         OrigBottom      =   975
         Max             =   30
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMinutes 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Refresh interval (minutes)"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   3240
   End
   Begin VB.ListBox lstUsers 
      Height          =   3180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2595
   End
   Begin VB.CommandButton cmdRefreshUsers 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
End
Attribute VB_Name = "FUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lMinWidth = 3030
Private Const k_lMinHeight = 5700
Private Const k_lHeightDiff = 2250
Private Const k_lLeftDiff = 720

Private m_lWindowID As Long
Private m_lMinute As Long


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


Private Sub chkAutoRefresh_Click()
    Timer1.Enabled = (chkAutoRefresh.value = vbChecked)
    ResetTimer
End Sub

Private Sub cmdRefreshUsers_Click()
    RefreshUserList
End Sub

Private Sub RefreshUserList()
    Dim orst As ADODB.Recordset
    Set orst = LoadDiscRst("exec spCPCGetUsers")
    LoadList lstUsers, orst, "UserID", "UserKey"
    ResetTimer
    Set orst = Nothing
End Sub

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
    cmdRefreshUsers_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        DoShowHelp
    Else
        MDIMain.GlobalKeyDownProcessing KeyCode, Shift
    End If
End Sub


Public Sub DoShowHelp()
    ShowHelp "FUsers", True
End Sub

Private Sub Form_Load()
    SetCaption "Who's Logged On?"
    With Me
        .Height = k_lMinHeight
        .Width = k_lMinWidth
    End With
    
    UpDown1.value = 2
    lblMinutes.Caption = UpDown1.value
    chkAutoRefresh.value = vbChecked
End Sub


Private Sub Form_Resize()
    Dim lBorder As Long
    lBorder = 120
    
    If Me.WindowState = 1 Then Exit Sub
    
    Me.Width = k_lMinWidth
    
    With Me
        'If .Width < k_lMinWidth Then .Width = k_lMinWidth
        If .Height < k_lMinHeight Then .Height = k_lMinHeight
    End With
    
    With lstUsers
        .Width = Me.Width - (3 * lBorder)
        .Height = Me.Height - .Top - k_lHeightDiff
    End With
    
    With cmdRefreshUsers
        .Top = lstUsers.Height + lstUsers.Top + 200
        .Left = (lstUsers.Width - .Width) / 2 + lBorder
    End With
    
    With frmRefresh
        .Top = cmdRefreshUsers.Height + cmdRefreshUsers.Top + 70
    End With
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.UnloadTool m_lWindowID
End Sub

Private Sub Timer1_Timer()
    m_lMinute = m_lMinute + 1
    If m_lMinute >= UpDown1.value Then
        RefreshUserList
    End If
End Sub


Private Sub UpDown1_Change()
    lblMinutes.Caption = UpDown1.value
End Sub

Private Sub UpDown1_LostFocus()
    ResetTimer
End Sub

Private Sub ResetTimer()
    Timer1.Enabled = (chkAutoRefresh.value = vbChecked)
    m_lMinute = 0
End Sub

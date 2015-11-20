VERSION 5.00
Begin VB.Form FLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Developer Logon"
   ClientHeight    =   2385
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSelectEnvironment 
      Caption         =   "Select Environment"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   3195
   End
   Begin VB.ComboBox cboUserId 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "FLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This is form is only loaded for developers to allow choice of target platform and impersonating endusers

Private m_userconfig As Configuration

Private m_servers As Dictionary
Private m_servername As String

Private m_bLoggedOn As Boolean
Private m_isLoading As Boolean


Public Property Get TargetServer() As String
    TargetServer = m_servername
End Property


Public Function Logon(userconfig As Configuration) As Boolean
    Dim i As Integer
    Dim opt As OptionButton
    
    Set m_userconfig = userconfig
    
    Set m_servers = userconfig.GetServerList
    For i = 0 To m_servers.Count - 1
        Set opt = Controls.Add("VB.OptionButton", "opt" & i)
        Set opt.Container = frmSelectEnvironment
        opt.Top = 300 + i * 350
        opt.Left = 200
        opt.width = 2000
        opt.Visible = True
        opt.Caption = m_servers.Keys()(i)
        If i = 0 Then
            opt.value = True
        End If
    Next i
    
    Show vbModal
    
    Logon = m_bLoggedOn
    Unload Me
    
End Function


Private Sub cmdOK_Click()
    Dim c As Control
    Dim i As Integer

    For Each c In Me.Controls
        If TypeOf c Is OptionButton Then
            If c.value Then
                m_servername = m_servers.Items()(i)
            End If
            i = i + 1
        End If
    Next
    
    m_bLoggedOn = True
    
'update the selected DB in the persistance structure
    
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    m_bLoggedOn = False
    Me.Hide
End Sub


Private Sub Form_Load()
    m_isLoading = True
    
    LoadCombo cboUserId, GetAllActiveUsers, "userid"
    cboUserId.Enabled = True
    SetComboByText cboUserId, GetUserName
    
    m_isLoading = False
End Sub


Private Function GetAllActiveUsers() As ADODB.Recordset
    Dim ocon As ADODB.Connection
    Dim sql As String
    Dim server As String
    Dim database As String

    server = m_userconfig.GetKeyValue("activeusers", "server")
    database = m_userconfig.GetKeyValue("activeusers", "database")
    
    Set ocon = New ADODB.Connection
    ocon.Open "Provider=SQLOLEDB.1;Server=" & server & ";Database=" & database & ";Trusted_Connection=yes;"
    sql = "SELECT lower(username) as userid FROM staff WHERE active = 1 ORDER BY username"
    Set GetAllActiveUsers = LoadDiscRst(sql, ocon, adUseClient)
    ocon.Close
End Function


Private Sub cboUserId_Click()
    If m_isLoading Then Exit Sub
    User.LoggedInUserId = cboUserId.text
End Sub


Private Sub optDatabase_DblClick(index As Integer)
    cmdOK_Click
End Sub

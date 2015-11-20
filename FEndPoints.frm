VERSION 5.00
Begin VB.Form FEndPoints 
   Caption         =   "Connections and EndPoints"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrinterName 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   1440
      TabIndex        =   44
      Top             =   8760
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterLocation 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   3180
      TabIndex        =   43
      Top             =   8760
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterWarehouse 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   4920
      TabIndex        =   42
      Top             =   8760
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterName 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   41
      Top             =   8400
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterLocation 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   3180
      TabIndex        =   40
      Top             =   8400
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterWarehouse 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   4920
      TabIndex        =   39
      Top             =   8400
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterName 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   38
      Top             =   8040
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterLocation 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   3180
      TabIndex        =   37
      Top             =   8040
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterWarehouse 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   4920
      TabIndex        =   36
      Top             =   8040
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterWarehouse 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   4920
      TabIndex        =   34
      Top             =   7680
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterWarehouse 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   33
      Top             =   7320
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterWarehouse 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   4920
      TabIndex        =   32
      Top             =   6960
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterLocation 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   3180
      TabIndex        =   31
      Top             =   7320
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterLocation 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   3180
      TabIndex        =   30
      Top             =   6960
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterLocation 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   3180
      TabIndex        =   29
      Top             =   7680
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterName 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   28
      Top             =   7680
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterName 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   27
      Top             =   7320
      Width           =   1635
   End
   Begin VB.TextBox txtPrinterName 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   26
      Top             =   6960
      Width           =   1635
   End
   Begin VB.Frame Frame2 
      Caption         =   "GlobalConfig"
      Height          =   2055
      Left            =   180
      TabIndex        =   13
      Top             =   2220
      Width           =   7515
      Begin VB.TextBox txtGlobalRemotePath 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1320
         TabIndex        =   23
         Top             =   1440
         Width           =   5595
      End
      Begin VB.TextBox txtGlobalLocalPath 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1320
         TabIndex        =   22
         Top             =   1020
         Width           =   5595
      End
      Begin VB.TextBox txtGlobalTargetServer 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1260
         TabIndex        =   16
         Top             =   300
         Width           =   1635
      End
      Begin VB.Image imgGlobalRemote 
         Height          =   390
         Left            =   6960
         Picture         =   "FEndPoints.frx":0000
         Top             =   1380
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgGlobalLocal 
         Height          =   390
         Left            =   6960
         Picture         =   "FEndPoints.frx":011E
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remote"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   1500
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   1140
         Width           =   390
      End
      Begin VB.Label lblFilePath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FilePath"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label4 
         Caption         =   "TargetServer"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "UserConfig"
      Height          =   2055
      Left            =   180
      TabIndex        =   10
      Top             =   120
      Width           =   7515
      Begin VB.TextBox txtUserRemotePath 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1260
         TabIndex        =   18
         Top             =   1440
         Width           =   5595
      End
      Begin VB.TextBox txtUserLocalPath 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1260
         TabIndex        =   17
         Top             =   1020
         Width           =   5595
      End
      Begin VB.TextBox txtUserTargetServer 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1260
         TabIndex        =   12
         Top             =   300
         Width           =   1635
      End
      Begin VB.Image imgUserRemote 
         Height          =   390
         Left            =   6960
         Picture         =   "FEndPoints.frx":023C
         Top             =   1380
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgUserLocal 
         Height          =   390
         Left            =   6960
         Picture         =   "FEndPoints.frx":035A
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Remote"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Local"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1125
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "FilePath"
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   740
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "TargetServer"
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.TextBox txtAutoPickUrl 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   5430
      Width           =   8895
   End
   Begin VB.TextBox txtMailUri 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   6285
      Width           =   8895
   End
   Begin VB.TextBox txtPayPalUrl 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   5850
      Width           =   8895
   End
   Begin VB.TextBox txtAutoCommitUrl 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   4995
      Width           =   8895
   End
   Begin VB.TextBox txtDBServer 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   4575
      Width           =   1635
   End
   Begin VB.Label Label9 
      Caption         =   "Printers"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "AutoPickUrl"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   5460
      Width           =   855
   End
   Begin VB.Label lblMail 
      Caption         =   "MailUrl"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   6300
      Width           =   855
   End
   Begin VB.Label lblQuoteEmailPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PayPalUrl"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Width           =   690
   End
   Begin VB.Label lblAutoCommitUrl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AutoCommitUrl"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   1035
   End
   Begin VB.Label lblDBServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DBServer"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   4620
      Width           =   690
   End
End
Attribute VB_Name = "FEndPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private Const k_lMinWidth = 8940
'Private Const k_lMinHeight = 1300


Private m_lWindowID As Long

Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property


Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property


Public Sub SetCaption(ByRef i_sTitle As String)
    Me.caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub

Private Sub Form_Load()
    Dim orst As ADODB.Recordset
    Dim i As Integer
    
    SetCaption "EndPoints"
    
    txtUserTargetServer.text = g_UserConfig.TargetServer
    txtUserLocalPath.text = g_UserConfig.LocalPathname
    txtUserRemotePath.text = g_UserConfig.RemotePathname
    
    imgUserLocal.Visible = IIf(g_UserConfig.LocalConfigFound, True, False)
    imgUserRemote.Visible = IIf(g_UserConfig.LocalConfigFound, False, True)
    
    txtGlobalTargetServer.text = g_GlobalConfig.TargetServer
    txtGlobalLocalPath.text = g_GlobalConfig.LocalPathname
    txtGlobalRemotePath.text = g_GlobalConfig.RemotePathname
    
    imgGlobalLocal.Visible = IIf(g_GlobalConfig.LocalConfigFound, True, False)
    imgGlobalRemote.Visible = IIf(g_GlobalConfig.LocalConfigFound, False, True)
    
    txtDBServer.text = g_DB.server
    txtAutoCommitUrl.text = g_AutoCommitUrl
    txtPayPalUrl.text = g_PayPalUri
    txtMailUri.text = g_MailServiceUrl
    txtAutoPickUrl.text = g_AutoPickUrl
    
    Set orst = GetPrinters
    i = 0
    Do While Not orst.EOF
        txtPrinterName(i).text = orst.Fields("Name")
        txtPrinterLocation(i).text = orst.Fields("PrinterLocation")
        txtPrinterName(i).text = orst.Fields("Name")
        txtPrinterWarehouse(i).text = orst.Fields("WareHouseId")
        i = i + 1
        orst.MoveNext
    Loop
    orst.Close
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub

    'If Me.width < k_lMinWidth Then Me.width = k_lMinWidth
    'If Me.Height < k_lMinHeight Then Me.Height = k_lMinHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.FormUnregister Me
End Sub

Private Function GetPrinters() As ADODB.Recordset
    Set GetPrinters = LoadDiscRst("select * from tcpPrinters")
End Function

    
    

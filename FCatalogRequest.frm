VERSION 5.00
Begin VB.Form FCatalogRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Our New Catalog"
   ClientHeight    =   2175
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmQty 
      Height          =   1452
      Left            =   2220
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   2352
      Begin VB.TextBox txtNumPriceLists 
         Height          =   312
         Left            =   180
         TabIndex        =   8
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox txtNumNotebook 
         Enabled         =   0   'False
         Height          =   312
         Left            =   180
         TabIndex        =   7
         Text            =   "0"
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox txtNumBound 
         Height          =   312
         Left            =   180
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   372
      End
      Begin VB.Label lblPriceLists 
         Caption         =   "Price Lists"
         Height          =   312
         Left            =   660
         TabIndex        =   11
         Top             =   1020
         Visible         =   0   'False
         Width           =   1032
      End
      Begin VB.Label lblNumNotebook 
         Caption         =   "Notebook Catalogs"
         Enabled         =   0   'False
         Height          =   312
         Left            =   660
         TabIndex        =   10
         Top             =   660
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Softbound Catalogs"
         Height          =   255
         Left            =   660
         TabIndex        =   9
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1740
      TabIndex        =   4
      Top             =   1680
      Width           =   972
   End
   Begin VB.OptionButton optAskAgain 
      Caption         =   "Ask again later"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Value           =   -1  'True
      Width           =   1992
   End
   Begin VB.OptionButton optDoNotSend 
      Caption         =   "Doesn't want a catalog"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1992
   End
   Begin VB.OptionButton optBulkMail 
      Caption         =   "Send later in bulk mail"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   1992
   End
   Begin VB.OptionButton optShip 
      Caption         =   "Ship with this order"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1992
   End
End
Attribute VB_Name = "FCatalogRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The OrderPad that spawned this dialog
Private m_FParent As FOrder

Private m_lCustKey As Long
Private m_lAddrKey As Long
Private m_sCustType As String

Private Sub cmdOK_Click()
    'do nothing if "AskAgain"
    If optAskAgain.Value = False Then
        If optShip.Value = True Then
            'if no qtys are indicated, ignore this input and treat like "AskAgain"
            If CInt(txtNumBound.Text) > 0 Or CInt(txtNumNotebook.Text) > 0 Or CInt(txtNumPriceLists.Text) > 0 Then
                m_FParent.AddMarketingItem CInt(txtNumBound.Text), CInt(txtNumNotebook.Text), CInt(txtNumPriceLists.Text)
                LogRequest RespCode:=1
            End If
        ElseIf optBulkMail.Value = True Then
            LogRequest RespCode:=2
        ElseIf optDoNotSend.Value = True Then
            LogRequest RespCode:=3
        End If
    End If
    Set m_FParent = Nothing
    Unload Me
End Sub


Private Sub LogRequest(ByVal RespCode As Integer)
    Dim cmd As ADODB.Command

    Set cmd = CreateCommandSP("spcpcInsertCatalogRequest")
    With cmd
        .Parameters("@_iCustKey").Value = m_lCustKey
        .Parameters("@_iAddrKey").Value = m_lAddrKey
        .Parameters("@_iRespCode").Value = RespCode
        .Parameters("@_iQtyBound").Value = CInt(txtNumBound.Text)
        .Parameters("@_iQtyNotebook").Value = CInt(txtNumNotebook.Text)
        .Parameters("@_iQtyPricelist").Value = CInt(txtNumPriceLists.Text)
        .Execute
    End With
    Set cmd = Nothing
End Sub


Private Sub optAskAgain_Click()
    frmQty.Visible = False
End Sub

Private Sub optBulkMail_Click()
    frmQty.Visible = False
End Sub

Private Sub optDoNotSend_Click()
    frmQty.Visible = False
End Sub

Private Sub optShip_Click()
'removed 8/24/06 LR. we're not supporting pricelists this time.
'    If m_sCustType = "EndUser" Then
'        txtNumPriceLists.Visible = False
'        lblPriceLists.Visible = False
'    Else
'        txtNumPriceLists.Visible = True
'        lblPriceLists.Visible = True
'    End If
    frmQty.Visible = True
    txtNumBound.TabIndex = 1
    txtNumNotebook.TabIndex = 2
    txtNumPriceLists.TabIndex = 3
    txtNumBound.SetFocus
End Sub

Public Sub Init(ByVal CustKey As Long, ByVal AddrKey As Long, _
                ByVal CustType As String, ByVal FParent As FOrder)
    m_lCustKey = CustKey
    m_sCustType = CustType
    m_lAddrKey = AddrKey
    Set m_FParent = FParent

    If g_SupportNotebooks Then
  txtNumNotebook.Enabled = True
  lblNumNotebook.Enabled = True
    Else
  txtNumNotebook.Enabled = False
  lblNumNotebook.Enabled = False
    End If
    
    Me.Show vbModal
End Sub

Private Sub txtNumBound_GotFocus()
    txtNumBound.SelStart = 0
    txtNumBound.SelLength = Len(txtNumBound.Text)
End Sub

Private Sub txtNumBound_LostFocus()
    If Not IsNumeric(txtNumBound.Text) Then
        Msg "Must be a number.", vbCritical
        txtNumBound.Text = "0"
    End If
End Sub

Private Sub txtNumNotebook_GotFocus()
    txtNumNotebook.SelStart = 0
    txtNumNotebook.SelLength = Len(txtNumNotebook.Text)
End Sub

Private Sub txtNumNotebook_LostFocus()
    If Not IsNumeric(txtNumNotebook.Text) Then
        Msg "Must be a number.", vbCritical
        txtNumNotebook.Text = "0"
    End If
End Sub

Private Sub txtNumPriceLists_GotFocus()
    txtNumPriceLists.SelStart = 0
    txtNumPriceLists.SelLength = Len(txtNumPriceLists.Text)
End Sub

Private Sub txtNumPriceLists_LostFocus()
    If Not IsNumeric(txtNumPriceLists.Text) Then
        Msg "Must be a number.", vbCritical
        txtNumPriceLists.Text = "0"
    End If
End Sub


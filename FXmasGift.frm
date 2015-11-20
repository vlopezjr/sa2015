VERSION 5.00
Begin VB.Form FXmasGift 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Christmas Gifts "
   ClientHeight    =   2445
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3732
      Begin VB.CommandButton cmdOKF1 
         Caption         =   "OK"
         Height          =   312
         Left            =   1380
         TabIndex        =   5
         Top             =   1920
         Width           =   972
      End
      Begin VB.OptionButton optNoF1 
         Caption         =   "No/Unsure"
         Height          =   255
         Left            =   420
         TabIndex        =   4
         Top             =   1500
         Width           =   1152
      End
      Begin VB.OptionButton optYesF1 
         Caption         =   "Yes"
         Height          =   255
         Left            =   420
         TabIndex        =   3
         Top             =   1140
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Step 1 of 2 :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCutoffMsg 
         Caption         =   "Is all or part of this order going to ship on or before "
         Height          =   495
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   3435
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   60
      TabIndex        =   1
      Top             =   -60
      Visible         =   0   'False
      Width           =   3732
      Begin VB.CommandButton cmdOKF2 
         Caption         =   "OK"
         Height          =   312
         Left            =   1380
         TabIndex        =   10
         Top             =   1920
         Width           =   972
      End
      Begin VB.OptionButton optUnsureF2 
         Caption         =   "No/Unsure - It's OK to ask again"
         Height          =   252
         Left            =   420
         TabIndex        =   9
         Top             =   1500
         Width           =   3012
      End
      Begin VB.OptionButton optNoF2 
         Caption         =   "No - Send Later"
         Enabled         =   0   'False
         Height          =   252
         Left            =   480
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.OptionButton optYesF2 
         Caption         =   "Yes"
         Height          =   252
         Left            =   420
         TabIndex        =   7
         Top             =   1140
         Value           =   -1  'True
         Width           =   1152
      End
      Begin VB.Label Label4 
         Caption         =   "Step 2 of 2 :"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Would this customer like to have a gift included with this order?"
         Height          =   435
         Left            =   180
         TabIndex        =   6
         Top             =   540
         Width           =   2895
      End
   End
End
Attribute VB_Name = "FXmasGift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The OrderPad that spawned this dialog
Private m_FParent As FOrder

Private m_lCustKey As Long
Private m_lAddrKey As Long

Private m_sMailTo As String


Public Sub Init(ByVal CustKey As Long, ByVal AddrKey As Long, ByVal FParent As FOrder)
    m_lCustKey = CustKey
    m_lAddrKey = AddrKey
    Set m_FParent = FParent
    
    Frame1.Visible = True
    cmdOKF1.Default = True
    Frame2.Visible = False
    Me.Caption = Me.Caption & Year(Now)
    lblCutoffMsg.Caption = lblCutoffMsg.Caption & g_CutOffDate & "?"
    m_sMailTo = "BobG@caseparts.com;JoannaR@caseparts.com;" & GetUserName & "@caseparts.com"
    Me.Show vbModal
End Sub


Private Sub cmdOKF1_Click()
    If optYesF1.Value = True Then
        Frame1.Visible = False
        cmdOKF1.Default = False
        Frame2.Visible = True
        cmdOKF2.Default = True
    Else
        Dim body As String
        body = "OP" & m_FParent.Order.OPKey & " for " & m_FParent.Order.Customer.ID & ", " & m_FParent.Order.UserID & vbCrLf & "In Step 1 selected 'No/Unsure that this order will ship before <cutoff date>'"
        EMail.Send "op@caseparts.com", m_sMailTo, "Gift Alert", body, False
        Set m_FParent = Nothing
        Unload Me
    End If
End Sub


Private Sub cmdOKF2_Click()
    Dim body As String

    If optYesF2.Value = True Then
        m_FParent.AddGiftItem
        UpdateTable RespCode:=1
    ElseIf optNoF2.Value = True Then
        UpdateTable RespCode:=2
        body = "OP" & m_FParent.Order.OPKey & " for " & m_FParent.Order.Customer.ID & ", " & m_FParent.Order.UserID & vbCrLf & "In Step 2 selected 'No - Customer does not want gift in their order. Send separately.'"

        EMail.Send "op@caseparts.com", m_sMailTo, "Gift Alert", body, False
    ElseIf optUnsureF2.Value = True Then
        body = "OP" & m_FParent.Order.OPKey & " for " & m_FParent.Order.Customer.ID & ", " & m_FParent.Order.UserID & vbCrLf & "In Step 2 selected 'No/Unsure - Customer does not want gift in this order. It's OK to ask again.'"

        EMail.Send "op@caseparts.com", m_sMailTo, "Gift Alert", body, False
    End If
    Set m_FParent = Nothing
    Unload Me

End Sub


Private Sub UpdateTable(ByVal RespCode As Integer)
    Dim ocmd As ADODB.Command
    Set ocmd = CreateCommandSP("spcpcUpdateXmasGift")
    With ocmd
        .Parameters("@_iCustKey").Value = m_lCustKey
        .Parameters("@_iAddrKey").Value = m_lAddrKey
        .Parameters("@_iRespCode").Value = RespCode
        .Execute
    End With
    Set ocmd = Nothing
End Sub

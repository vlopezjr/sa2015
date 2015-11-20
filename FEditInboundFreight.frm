VERSION 5.00
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Begin VB.Form FEditInboundFreight 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Inbound Freight Amount"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3300
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin NEWSOTALib.SOTACurrency txtInboundFreightAmt 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
      text            =   "           0.00"
      sDecimalPlaces  =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Inbound Freight Amt"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FEditInboundFreight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_bEdit As Boolean


Public Sub EnterInboundFreightAmt(ByVal i_OPID As Long)
    Dim cmd As ADODB.Command
    Dim dInboundFreight As Double
    
    Set cmd = CreateCommandSP("spCPCWAGetInboundFreightAmt")
    cmd.Parameters("@_iOPID").value = i_OPID
    cmd.Execute
    dInboundFreight = cmd.Parameters("@_iRetVal").value
    
    txtInboundFreightAmt.Amount = dInboundFreight
    
    Show vbModal
    
    If m_bEdit Then
        If txtInboundFreightAmt.Amount - dInboundFreight <> 0 Then
            UpdateInboundFreightAmt i_OPID, txtInboundFreightAmt.Amount
            'PRN#96
            LogOAEvent "Order", GetUserID, i_OPID, , , "Change Order Inbound Freight Amount from $" & dInboundFreight & " to $" & txtInboundFreightAmt.Amount
        End If
    End If
End Sub


Private Sub UpdateInboundFreightAmt(ByVal i_OPID As Long, ByVal d_InboundFreightAmt As Double)
    Dim cmd As ADODB.Command
    
    Set cmd = CreateCommandSP("spCPCWAUpdateInboundFreightAmt")
    cmd.Parameters("@_iOPID").value = i_OPID
    cmd.Parameters("@_iInboundFreightAmt") = d_InboundFreightAmt
    cmd.Execute
    Set cmd = Nothing
End Sub


Private Sub cmdCancel_Click()
    m_bEdit = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    If txtInboundFreightAmt.Amount < 0 Then
        msg "The amount of Inbound Freight should be bigger than 0", vbOKOnly + vbExclamation, "Negative Inbound Freight Amt"
        txtInboundFreightAmt.SetFocus
        Exit Sub
    End If
    
    m_bEdit = True
    Me.Hide
End Sub

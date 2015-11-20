VERSION 5.00
Begin VB.Form FNumLabels 
   Caption         =   "Print Labels"
   ClientHeight    =   1632
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3972
   LinkTopic       =   "Form1"
   ScaleHeight     =   1632
   ScaleWidth      =   3972
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumLabels 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   780
      TabIndex        =   2
      Top             =   240
      Width           =   432
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   432
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   432
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label lblType 
      Caption         =   "SPO item labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   2472
   End
   Begin VB.Label Label1 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   432
   End
End
Attribute VB_Name = "FNumLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is a custom dialog box for LabelPrinter.cls


Public Property Get NumLabels(iNewVal As Integer) As Integer
    txtNumLabels.Text = iNewVal
    lblType.Caption = "labels"
    Me.Show vbModal
    If IsNumeric(txtNumLabels.Text) Then
        NumLabels = CInt(txtNumLabels.Text)
    Else
        NumLabels = 0
    End If
End Property


Private Sub cmdCancel_Click()
    txtNumLabels.Text = 0
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    If CInt(txtNumLabels.Text) > 5 Then
        If vbYes = Msg("Are you sure you want to print " & txtNumLabels.Text & " labels?", vbYesNo, "LabelPrinter") Then
            Me.Hide
        Else
            SelectText txtNumLabels
        End If
    Else
        Me.Hide
    End If
End Sub


Private Sub Form_Activate()
    SelectText txtNumLabels
End Sub


Private Sub SelectText(tb As TextBox)
    tb.SetFocus
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
End Sub

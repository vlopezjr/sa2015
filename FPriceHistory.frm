VERSION 5.00
Begin VB.Form FPriceHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price History"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   972
   End
   Begin VB.TextBox txtHistory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Effective Date"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label lblPriceType 
      AutoSize        =   -1  'True
      Caption         =   "Price"
      Height          =   195
      Left            =   1860
      TabIndex        =   1
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "FPriceHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowPriceHistory(ByVal sItemID As String, ByVal lCustType As Long)
    Dim orst As ADODB.Recordset
    On Error GoTo EH
    
    Select Case lCustType
        Case 2
            lblPriceType.Caption = "Dealer Price"
        Case 3
            lblPriceType.Caption = "Wholesale Price"
        Case Else
            lblPriceType.Caption = "List Price"
    End Select
    
    Set orst = CallSP("cpoaGetPriceHistory", "@ItemID", sItemID, "@CustType", lCustType)
    Do Until orst.EOF
        txtHistory.Text = txtHistory.Text & orst.Fields("EffectiveDateEx") & vbTab & FormatCurrency(orst.Fields("EffectivePrice"), 2) & vbCrLf & vbCrLf
        orst.MoveNext
    Loop
    Set orst = Nothing
    
    Show vbModal
    
    Exit Sub
EH:
    Set orst = Nothing
    MsgBox "Load price history failed due to error '" & Err.Description & "'", vbInformation
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FPriceHistory = Nothing
End Sub

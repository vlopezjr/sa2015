VERSION 5.00
Begin VB.Form FSpecialHandlingWarnings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warnings"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMsg 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   60
      Width           =   4515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&No"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2220
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      TabIndex        =   0
      Top             =   2220
      Width           =   975
   End
End
Attribute VB_Name = "FSpecialHandlingWarnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Result As VbMsgBoxResult

Public Function ShowMessage(msg As String) As VbMsgBoxResult
    txtMsg.text = msg
    'block here
    Me.Show vbModal
    'on the way out
    ShowMessage = m_Result
    Unload Me
End Function

Private Sub cmdCancel_Click()
    m_Result = vbCancel
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    m_Result = vbOK
    Me.Hide
End Sub


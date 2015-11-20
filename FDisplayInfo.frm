VERSION 5.00
Begin VB.Form FDisplayInfo 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FDisplayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sBody As String


Public Property Let Body(ByVal sNewVal As String)
    m_sBody = sNewVal
End Property


Private Sub Form_Activate()
    Me.Cls
    Print m_sBody
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub


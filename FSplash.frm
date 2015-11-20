VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2172
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4524
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   252
         Left            =   240
         TabIndex        =   1
         Top             =   1656
         Width           =   4092
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Order Pad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00990000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   885
         Width           =   4050
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "Initializing..."
         ForeColor       =   &H00990000&
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   1416
         Width           =   4092
      End
      Begin VB.Label lblAppName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Order Pad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00990000&
         Height          =   492
         Left            =   300
         TabIndex        =   2
         Top             =   360
         Width           =   3996
      End
   End
End
Attribute VB_Name = "FSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lPercent As Long

Private Sub Form_Load()
    Me.Caption = App.ProductName & " loading..."
    lblAppName.Caption = App.ProductName
    lblVersion.Caption = "version " & GlobalFunctions.Version
End Sub


Public Function Progress(ByVal lPercent As Long, Optional sMessage As String = "Initializing")
    If lPercent < 0 Then
        lPercent = 0
    ElseIf lPercent > 100 Then
        lPercent = 100
    End If

    m_lPercent = lPercent
    ProgressBar1.value = m_lPercent
    lblMessage = sMessage & "..."
    Debug.Print "Progress " & m_lPercent & " " & sMessage
    DoEvents
End Function


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FProgress 
   Caption         =   "Progress"
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   10000
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait ... ..."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lMaxValue = 10000

Public Sub FirstStepProgress(ByVal dProgress As Double)
    pbProgress.Value = dProgress * (k_lMaxValue / 2)
    Refresh
End Sub

Public Sub SecondStepProgress(ByVal dProgress As Double)
    pbProgress.Value = (dProgress / 2 + 0.5) * k_lMaxValue
    Refresh
End Sub

Public Sub HideProgress()
    Unload Me
End Sub


Private Sub Form_Load()
    pbProgress.Value = 0
End Sub

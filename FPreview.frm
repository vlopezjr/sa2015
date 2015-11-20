VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form FPreview 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GEXPreview grPrev 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   2990
      BeginProperty ToolbarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PageSetupText   =   "Page Set&up..."
      PrintText       =   "&Print..."
      CloseButtonText =   "&Close"
   End
End
Attribute VB_Name = "FPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        grPrev.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub

Private Sub grPrev_OnCloseClick()
    Unload Me
 End Sub


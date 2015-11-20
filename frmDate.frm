VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please select a date."
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   2925
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mthView 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   19267585
      CurrentDate     =   38139
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    mthView.value = Now
End Sub

Private Sub mthView_DateClick(ByVal DateClicked As Date)
    Me.Hide
End Sub

VERSION 5.00
Begin VB.Form FSaveContact 
   Caption         =   "Save Contact"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   4995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Do you want to change the current contact or add a new contact?"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "FSaveContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum SaveContactOptions
    scoUpdateContact = 0
    scoAddContact = 1
    scoCancelContact = 2
End Enum

Private m_eSaveOption As SaveContactOptions


Public Function SaveContact() As SaveContactOptions
    Show vbModal
    
    SaveContact = m_eSaveOption
    
    Unload Me
End Function


Private Sub cmdAdd_Click()
    m_eSaveOption = scoAddContact
    Me.Hide
End Sub


Private Sub cmdCancel_Click()
    m_eSaveOption = scoCancelContact
    Me.Hide
End Sub


Private Sub cmdUpdate_Click()
    m_eSaveOption = scoUpdateContact
    Me.Hide
End Sub


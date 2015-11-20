VERSION 5.00
Begin VB.Form FSaveAddress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Address"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5580
      TabIndex        =   5
      Top             =   660
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   5580
      TabIndex        =   4
      Top             =   240
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Update Options"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optSaveOption 
         Caption         =   "Use this address with this order only. Do not include it on the customer's address list"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   4695
      End
      Begin VB.OptionButton optSaveOption 
         Caption         =   "Add this as a new address to the customer's list of addresses"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   780
         Width           =   4575
      End
      Begin VB.OptionButton optSaveOption 
         Caption         =   "Update the address to reflect these changes"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4515
      End
   End
End
Attribute VB_Name = "FSaveAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum SaveAddressOptions
    saoCancel = 0
    saoChangeAddress = 1
    saoAddAddress = 2
    saoThisOrderOnly = 3
End Enum


Private m_eSaveOption As SaveAddressOptions

Private Sub cmdCancel_Click()
    m_eSaveOption = saoCancel
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    If optSaveOption(0).value = True Then
        m_eSaveOption = saoChangeAddress
    ElseIf optSaveOption(1).value = True Then
        m_eSaveOption = saoAddAddress
    ElseIf optSaveOption(2).value = True Then
        m_eSaveOption = saoThisOrderOnly
    Else
        m_eSaveOption = saoCancel
    End If
    Me.Hide
End Sub


Private Sub optSaveOption_Click(Index As Integer)
    cmdOK.Enabled = True
End Sub


Public Function GetSaveOption(ByRef i_oAddr As Address, ByVal bCanOverwrite As Boolean) As SaveAddressOptions
    optSaveOption(0).Enabled = (i_oAddr.AddrKey <> 0 And bCanOverwrite)
    optSaveOption(0).value = False
    optSaveOption(1).value = False
    
    '***465 SMR 04-10-2006 - TOO is now an Address.AddrType ***
    'optSaveOption(2).value = i_oAddr.IsThisOrderOnly
    If i_oAddr.AddrType = TOO Then optSaveOption(2).value = True
    
    cmdOK.Enabled = optSaveOption(0).value _
                 Or optSaveOption(1).value _
                 Or optSaveOption(2).value
    
    Show vbModal
    
    GetSaveOption = m_eSaveOption
    Unload Me
End Function


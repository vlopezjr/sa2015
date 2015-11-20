VERSION 5.00
Begin VB.Form FResearchStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Research Status"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   320
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   320
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton optWaitCustomer 
      Caption         =   "Wait for Customer"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OptionButton optWaitFactory 
      Caption         =   "Wait for Factory"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OptionButton optContactCustomer 
      Caption         =   "Contact Customer"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton optContactFactory 
      Caption         =   "Contact Factory"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"FResearchStatus.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FResearchStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bResearchStatus As Boolean
Private m_eResearchStatus As ItemResearchStatus


Public Function LoadOrderResearchStatus(ByRef oOrder As Order) As Boolean
    Dim oItem As IItem
    Dim lCount As Long
    
    For Each oItem In oOrder.Items
        Select Case oItem.ResearchStatus
            Case ItemResearchStatus.irsContactFactory
                If Not optContactFactory.Enabled Then
                    lCount = lCount + 1
                    optContactFactory.Enabled = True
                    optContactFactory.Value = True
                    oOrder.ResearchStatus = irsContactFactory
                End If
            Case ItemResearchStatus.irsContactCustomer
                If Not optContactCustomer.Enabled Then
                    lCount = lCount + 1
                    optContactCustomer.Enabled = True
                    If Not optContactFactory.Enabled Then
                        optContactCustomer.Value = True
                        oOrder.ResearchStatus = irsContactCustomer
                    End If
                End If
            Case ItemResearchStatus.irsWaitFactory
                If Not optWaitFactory.Enabled Then
                    lCount = lCount + 1
                    optWaitFactory.Enabled = True
                    If Not optContactFactory.Enabled And Not optContactCustomer.Enabled Then
                        optWaitFactory.Value = True
                        oOrder.ResearchStatus = irsWaitFactory
                    End If
                End If
            Case ItemResearchStatus.irsWaitCustomer
                If Not optWaitCustomer.Enabled Then
                    lCount = lCount + 1
                    optWaitCustomer.Enabled = True
                    oOrder.ResearchStatus = irsWaitCustomer
                End If
        End Select
    Next
    If lCount < 2 Then
        LoadOrderResearchStatus = True
'10/12/05 added LR & DH
        Unload Me
    Else
        Me.Show vbModal
        'we return here when OK or Cancel is clicked
        LoadOrderResearchStatus = m_bResearchStatus
        If LoadOrderResearchStatus Then
            oOrder.ResearchStatus = m_eResearchStatus
        End If
'10/12/05 added LR & DH
        Unload Me
    End If
End Function


Private Sub cmdCancel_Click()
    m_bResearchStatus = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    m_bResearchStatus = True
    If optContactFactory.Enabled And optContactFactory.Value = True Then
        m_eResearchStatus = irsContactFactory
    ElseIf optContactCustomer.Enabled And optContactCustomer.Value = True Then
        m_eResearchStatus = irsContactCustomer
    ElseIf optWaitFactory.Enabled And optWaitFactory.Value = True Then
        m_eResearchStatus = irsWaitFactory
    Else
        m_eResearchStatus = irsWaitCustomer
    End If
    Me.Hide
End Sub

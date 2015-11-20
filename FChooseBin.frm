VERSION 5.00
Begin VB.Form FChooseBin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Bin"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.ListBox lstBins 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblNewBin 
      Caption         =   "Please choose new bin from the following list:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "FChooseBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lBin As Long


Public Function LoadNewBin(ByRef rstCurrBin As ADODB.Recordset) As Long
    Dim sSQL As String
    Dim rstBin As ADODB.Recordset
    
    sSQL = "Select WhseBinKey, WhseBinID from timWhseBin " _
            & "Where WhseKey = " & GetUserWhseKey
    
    Set rstBin = LoadDiscRst(sSQL)
    
    LoadList lstBins, rstBin, "WhseBinID", "WhseBinKey"
    With lstBins
        If .ListCount = 0 Then
            m_lBin = 0
        Else
            If rstCurrBin.RecordCount > 0 Then
                rstCurrBin.MoveFirst
                While Not rstCurrBin.EOF
                    ListRemoveItemByText lstBins, rstCurrBin.Fields("WhseBinID")
                    rstCurrBin.MoveNext
                Wend
                rstCurrBin.MoveFirst
            End If
            Me.Show vbModal
        End If
    End With
    LoadNewBin = m_lBin
    Unload Me
End Function


Private Sub cmdCancel_Click()
    m_lBin = 0
    Me.Hide
End Sub



Private Sub cmdOK_Click()
    If lstBins.ListIndex < 0 Then
        MsgBox "Please choose Bin from the list first.", vbExclamation + vbOKOnly, "Choose New Bin"
        Exit Sub
    End If
    
    m_lBin = lstBins.ItemData(lstBins.ListIndex)
    Me.Hide
End Sub


Private Sub lstBins_DblClick()
    cmdOK_Click
End Sub

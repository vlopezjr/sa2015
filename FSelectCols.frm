VERSION 5.00
Begin VB.Form FSelectCols 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available Columns:"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   2190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FSelectCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_chkCols As CheckBox
Dim m_gdxGrid As GridEX

Public Sub Init(gdxGrid As GridEX)
    Set m_gdxGrid = gdxGrid
    Me.Show 1
End Sub

Private Sub Form_Load()
    Dim oCol As JSColumn
    Dim i As Integer
    
    Me.Height = m_gdxGrid.Columns.Count * 400 + 1000
    
    For Each oCol In m_gdxGrid.Columns
        
        Set m_chkCols = Controls.Add("VB.checkbox", oCol.Key)
        If oCol.Caption = "" Then
            m_chkCols.Caption = oCol.DataField
        Else
            m_chkCols.Caption = oCol.Caption
        End If
        
        m_chkCols.Top = oCol.ColPosition * 400 - 400
        m_chkCols.Width = 2000
        m_chkCols.Left = 200
        m_chkCols.Visible = True
        
        If oCol.Visible = True Then m_chkCols.Value = vbChecked
    Next

    cmdSave.Top = m_gdxGrid.Columns.Count * 400 + 200
    cmdCancel.Top = m_gdxGrid.Columns.Count * 400 + 200
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim oCol As JSColumn
    Dim liCounter As Integer
    liCounter = 0
    
    'make sure at least one checkbox is checked
    For Each oCol In m_gdxGrid.Columns
        If Controls(oCol.Key).Value = vbChecked Then
            liCounter = liCounter + 1
        End If
    Next
    
    If liCounter = 0 Then
        MsgBox "Please select at least one column.", vbInformation, "Select Columns"
        Exit Sub
    End If
    
    For Each oCol In m_gdxGrid.Columns
        If Controls(oCol.Key).Value = vbChecked Then
            oCol.Visible = True
            liCounter = liCounter + 1
        Else
            oCol.Visible = False
        End If
    Next

    m_gdxGrid.LayoutString
    Unload Me
End Sub
















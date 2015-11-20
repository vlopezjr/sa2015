VERSION 5.00
Begin VB.Form FFontEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font"
   ClientHeight    =   660
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboFontName 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   1632
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   312
      Left            =   2820
      TabIndex        =   1
      Top             =   180
      Width           =   732
   End
   Begin VB.ComboBox cboFontSize 
      Height          =   315
      ItemData        =   "FFontEditor.frx":0000
      Left            =   1920
      List            =   "FFontEditor.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   672
   End
End
Attribute VB_Name = "FFontEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_FontName As String
Private m_FontSize As Integer


Public Property Get FName() As String
    FName = cboFontName.Text
End Property

Public Property Let FName(ByVal sNewValue As String)
    m_FontName = sNewValue
End Property

Public Property Get FSize() As Integer
    FSize = m_FontSize
End Property

Public Property Let FSize(ByVal iNewValue As Integer)
    m_FontSize = iNewValue
End Property

Private Sub cboFontSize_click()
    m_FontSize = CInt(cboFontSize.Text)
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    LoadFontNames
    SetComboByText cboFontName, m_FontName
    SetComboByText cboFontSize, CStr(m_FontSize)
End Sub

Private Sub LoadFontNames()
    cboFontName.AddItem "Arial"
    cboFontName.AddItem "Comic Sans MS"
    cboFontName.AddItem "MS Sans Serif"
    cboFontName.AddItem "Tahoma"
    cboFontName.AddItem "Times New Roman"
    cboFontName.AddItem "Verdana"
End Sub

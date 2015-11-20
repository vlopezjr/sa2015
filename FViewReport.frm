VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FViewReport 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   8970
   Begin SHDocVwCtl.WebBrowser webBrowser1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8745
      ExtentX         =   15425
      ExtentY         =   7223
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "FViewReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_lWindowID As Long

Private Sub Form_Resize()
    webBrowser1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Public Sub PopUp(ByRef i_sURL As String, Optional ByVal i_bModal As Boolean = False)
    webBrowser1.Navigate2 CStr(i_sURL)
    Show
End Sub

Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property


Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property

Public Sub SetCaption(ByRef i_sTitle As String)
    Me.Caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If m_lWindowID <> 0 Then
        MDIMain.UnloadTool m_lWindowID
    End If
    MDIMain.DoRefresh
End Sub

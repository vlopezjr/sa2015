VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FXSLViewer 
   Caption         =   "XSL Viewer"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   7545
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   285
      Left            =   5400
      TabIndex        =   1
      Top             =   5160
      Width           =   972
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   288
      Left            =   6480
      TabIndex        =   2
      Top             =   5160
      Width           =   972
   End
   Begin SHDocVwCtl.WebBrowser webBrowser1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7305
      ExtentX         =   12885
      ExtentY         =   8705
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "FXSLViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWindowID As Long
Private m_sXSLFile As String


Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property


Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property


Private Sub Form_Unload(Cancel As Integer)
    If m_lWindowID <> 0 Then
        MDIMain.UnloadTool m_lWindowID
    End If
    MDIMain.DoRefresh
End Sub


Private Sub Form_Resize()
    If Me.width - 300 > 0 Then
        WebBrowser1.width = Me.width - 300
    End If
    If Me.Height - 1050 > 0 Then
        WebBrowser1.Height = Me.Height - 1050
    End If
    cmdPrint.Top = WebBrowser1.Top + WebBrowser1.Height + 120
    cmdClose.Top = cmdPrint.Top
    cmdClose.Left = WebBrowser1.width - cmdClose.width + 120
    cmdPrint.Left = cmdClose.Left - cmdPrint.width - 120
End Sub



Public Sub ShowViewer(ByVal sCaption As String, ByVal sXSLFile As String, ByVal sGasketStatus As String)
    Me.caption = sCaption
    MDIMain.UpdateCaption Me
    LoadView sXSLFile, sGasketStatus
End Sub


Private Sub LoadView(ByVal sXSLFile As String, ByVal sStrXml As String)
    SaveToFile HtmlPath, XslHeader(sXSLFile) + vbCrLf + sStrXml
    WebBrowser1.Navigate2 XMLURL
    Me.SetFocus
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub


Private Function HtmlPath() As String
    HtmlPath = g_SnapshotPath & GetUserName & ".xml"
End Function














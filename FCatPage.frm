VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FCatPage 
   Caption         =   "Catalog PageBlock Viewer"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   9465
   Begin SHDocVwCtl.WebBrowser wbBrowser 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      ExtentX         =   5741
      ExtentY         =   3201
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
Attribute VB_Name = "FCatPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const FormWidth = 11000

Private m_sPartNo As String
Private m_iCustType As String

Private m_lWindowID As Long


Private Sub Form_Unload(Cancel As Integer)
   ' MDIMain.Toolbar1.Tools.Remove "Window" & Me.WindowID
    MDIMain.UnloadTool m_lWindowID
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

Public Property Get PartNo() As String
    PartNo = m_sPartNo
End Property


Public Property Let PartNo(ByVal sNewValue As String)
    m_sPartNo = sNewValue
End Property


Public Property Get CustType() As Integer
    CustType = m_iCustType
End Property


Public Property Let CustType(ByVal iNewValue As Integer)
    m_iCustType = iNewValue
End Property


Public Sub ShowPage()
    Dim URL As String
    SetCaption "View Part " & UCase(m_sPartNo)
    
    URL = g_ViewPageURL & "?partnumber=" & m_sPartNo
    
    wbBrowser.Navigate2 URL
End Sub


Private Sub Form_Load()
    Me.width = FormWidth
   wbBrowser.Navigate "about:blank"
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Resize()
    wbBrowser.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


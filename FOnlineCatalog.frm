VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FOnlineCatalog 
   Caption         =   "Catalog"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbBrowser 
      Height          =   2295
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   4048
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
Attribute VB_Name = "FOnlineCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private whseid As String

Public Sub SetCaption(ByRef i_sTitle As String)
    Me.Caption = i_sTitle
End Sub

Public Property Get WarehouseId() As String
    WarehouseId = whseid
End Property

Public Property Let WarehouseId(ByVal sNewValue As String)
    whseid = sNewValue
End Property

Private Sub Form_Load()
    Me.Caption = "Catalog Volume 9"
   wbBrowser.Navigate2 "http://www.caseparts.com/catalog/csrgateway.aspx"
End Sub

Private Sub Form_Resize()
    wbBrowser.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


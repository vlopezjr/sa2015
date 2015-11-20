VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FInventoryHistory 
   Caption         =   "Inventory History"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   8670
   Begin GridEX20.GridEX gdxInventoryHist 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8281
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   6
      Column(1)       =   "FInventoryHistory.frx":0000
      Column(2)       =   "FInventoryHistory.frx":01C4
      Column(3)       =   "FInventoryHistory.frx":0308
      Column(4)       =   "FInventoryHistory.frx":04C0
      Column(5)       =   "FInventoryHistory.frx":0604
      Column(6)       =   "FInventoryHistory.frx":0754
      SortKeysCount   =   1
      SortKey(1)      =   "FInventoryHistory.frx":0894
      FormatStylesCount=   5
      FormatStyle(1)  =   "FInventoryHistory.frx":08FC
      FormatStyle(2)  =   "FInventoryHistory.frx":0A34
      FormatStyle(3)  =   "FInventoryHistory.frx":0AE4
      FormatStyle(4)  =   "FInventoryHistory.frx":0B98
      FormatStyle(5)  =   "FInventoryHistory.frx":0C70
      ImageCount      =   0
      PrinterProperties=   "FInventoryHistory.frx":0D28
   End
End
Attribute VB_Name = "FInventoryHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lWindowID As Long


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

Public Function ShowHistory(ByVal WhseKey As Long, ByVal ItemKey As Long) As Boolean
    Dim oRst As ADODB.Recordset
    On Error GoTo EH
    SetWaitCursor True
    Set oRst = CallSP("cpoaGetInvHist_ForAnnualSalesAnalysis", _
        "@WhseKey", WhseKey, _
        "@ItemKey", ItemKey)
    SetWaitCursor False
    
    If oRst.EOF Then
        Msg "No transaction found.", vbInformation
        ShowHistory = False
    Else
        With gdxInventoryHist
            .HoldFields
            .HoldSortSettings = True
            Set .ADORecordset = oRst
        End With
        ShowHistory = True
    End If
    
    Exit Function
EH:
    MsgBox "Failed to retrieve Inventory History due to error " & Err.Number & " " & Err.Description, vbInformation
End Function

Private Sub Form_Load()
    Width = gdxInventoryHist.Width
    Height = gdxInventoryHist.Height
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDIMain.UnloadTool m_lWindowID
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        gdxInventoryHist.Width = Width - 100
        gdxInventoryHist.Height = Height - 400
    End If
End Sub


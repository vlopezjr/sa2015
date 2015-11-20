VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridEXWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oGrid As GridEX
Attribute m_oGrid.VB_VarHelpID = -1
Private m_lButton As Long
Private m_lShift As Long
Private m_lX As Long
Private m_lY As Long
Private m_sBKLayout As String

'events

Public Event RowChosen()

Public Event ColumnChosen(columnName As String)

'properties

Public Property Get Grid() As GridEX
    Set Grid = m_oGrid
End Property

Public Property Let Grid(ByRef i_oGrid As GridEX)
    Set m_oGrid = i_oGrid
End Property

Public Property Get Button() As Long
    Button = m_lButton
End Property

Public Property Get Shift() As Long
    Shift = m_lShift
End Property

Public Property Get X() As Long
    X = m_lX
End Property

Public Property Get Y() As Long
    Y = m_lY
End Property


'methods

'In bound mode, you can get the Bookmark of any row to position the Recordset clone
'this method is used only within this class (called by Value())

Public Sub SyncRecordset()
    Dim vBookmark As Variant
    Dim lRowIndex As Long
    Dim value1 As Variant
    Dim value2 As Variant
    
    On Error GoTo EH
    
    With m_oGrid
        lRowIndex = .RowIndex(.Row)
        If lRowIndex > 0 Then
            vBookmark = .RowBookmark(lRowIndex)
            .ADORecordset.Bookmark = vBookmark
        End If
    End With
    Exit Sub
EH:
    ThrowError "GridEXWrapper", "SyncRecordset", ""
End Sub

'Get the original index of the selected row
'I don't think this method is used anywhere

Public Property Get RowIndex() As Long
    With m_oGrid
        RowIndex = .RowIndex(.Row)
    End With
End Property

'i_vIndex is an index (int) or a key (string)

Public Function value(ByVal i_vIndex As Variant) As Variant
    On Error GoTo EH
    
    SyncRecordset

    With m_oGrid
        If .RowIndex(.Row) > 0 Then
            On Error Resume Next
            value = .ADORecordset.Fields(i_vIndex).value
        End If
    End With
    Exit Function
EH:
    ThrowError "GridEXWrapper", "Value", ""
End Function

'*** grab these events after the grid does

Private Sub m_oGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And RowIndex <> 0 Then
        KeyCode = 0
        RaiseEvent RowChosen
    'map Tab to ctrl-Tab
    ElseIf KeyCode = 9 And Shift = 0 Then
        SendKeys "^{TAB}"
    End If
End Sub


Private Sub m_oGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'cache this for use by DblClick
    m_lButton = Button
    m_lShift = Shift
    m_lX = X
    m_lY = Y
End Sub


Private Sub m_oGrid_DblClick()

    If RowIndex = 0 Then Exit Sub
    
    With m_oGrid
        Select Case .HitTest(m_lX, m_lY)
        
            Case jgexHTBackGround, jgexHTColumnHeader, jgexHTGroupByBox, jgexHTNewRow
                'Debug.Print "No Double-Click. HitTest = " & .HitTest(m_lX, m_lY) & " for (" & m_lX & ", " & m_lY & ")"
                Exit Sub
                   
            Case jgexHTCell
                Dim colClicked As JSColumn
                Set colClicked = .ColFromPoint(m_lX, m_lY)
                If Not colClicked Is Nothing Then
                    RaiseEvent ColumnChosen(colClicked.caption)
                End If
                RaiseEvent RowChosen
                    
            Case jgexHTNoWhere 'JJC: for some reason, gdxItems returns this code on valid hits
                'Debug.Print "GOT Double-Click. HitTest = " & .HitTest(m_lX, m_lY) & " for (" & m_lX & ", " & m_lY & ")"
                RaiseEvent RowChosen
                
            Case Else
                'Debug.Print "GOT Double-Click. HitTest = " & .HitTest(m_lX, m_lY) & " for (" & m_lX & ", " & m_lY & ")"
                RaiseEvent RowChosen
            
        End Select
    End With
End Sub


'********************************************************
'New public sub routines
'...will also add the same Error Handling
'********************************************************
Public Sub InitGridLayout(plUserKey As Long, piFormConst As Integer)
    Dim iGridRev As Integer
    Dim lors As ADODB.Recordset
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If grid's reg entry is empty, then fill it with the grid's const
    'If the grid's reg entry is equal to the grid's const because either
        'the grid's reg entry has not been set or
        'the grid's const has not been changed and still equals the grid's reg entry
        'in this case get the current users grid prefs
            'if they don't exist then don't set a layout, the grid's default layout will be used
    'If the grid's reg entry is not equal to the grid's const
        'delete the grid's layout entry from the database
        'put the new grid's reg entry into the registry
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'iGridRev = GetRegNumberValue(HKEY_LOCAL_MACHINE, g_GridRegKey, m_oGrid.Name, 0)
    'iGridRev = GetRegNumberValue(HKEY_CURRENT_USER, g_GridRegKey, m_oGrid.name, piFormConst)
    iGridRev = g_UserConfig.GetKeyValue("grids", m_oGrid.name)

    m_sBKLayout = m_oGrid.LayoutString

    'restores the layout for the current user, if the cached grid revision counter is the same
    If iGridRev = piFormConst Then
        Set lors = CallSP("spCPCGetUserPrefs", "@_iUserKey", plUserKey, "@_iGridName", m_oGrid.name)
                       
        'check if a layout exist for this user, if not save this layout for the user (so they can restore)
        If lors.BOF = True And lors.EOF = True Then Exit Sub
               
        'set grid properties
        m_oGrid.LoadLayoutString lors.Fields("LayoutString").value
        m_sBKLayout = m_oGrid.LayoutString
        m_oGrid.Refresh
    Else
        'Delete layout from tcpUserPrefs if the grid's layout const has changed
        CallSP "spCPCDeleteUserPrefs", "@_iUserKey", plUserKey, "@_iGridName", m_oGrid.name
            
        'set new regvalue
        'Call PutRegNumberValue(HKEY_CURRENT_USER, g_GridRegKey, m_oGrid.name, piFormConst)
        g_UserConfig.SetKeyValue "grids", m_oGrid.name, piFormConst
    End If
    
    Set lors = Nothing
End Sub


Public Sub RestoreGridLayout(plUserKey As Long)
    m_oGrid.LoadLayoutString m_sBKLayout
End Sub

Public Sub GridAutoFit()
    Dim oCol As JSColumn
    For Each oCol In m_oGrid.Columns
        oCol.AutoSize
        If oCol.width > 4000 Then oCol.width = 4000
    Next
End Sub


Public Sub GridSaveLayout(plUserKey As Long)
    m_sBKLayout = m_oGrid.LayoutString
    CallSP "spCPCDeleteUserPrefs", "@_iUserKey", plUserKey, "@_iGridName", m_oGrid.name
    CallSP "spCPCSaveUserPrefs", "@_iUserKey", plUserKey, "@_iGridName", m_oGrid.name, "@_iLayoutString", m_oGrid.LayoutString
End Sub



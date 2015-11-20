VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drawers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'12/3/08 LR
'renamed this from User.cls to Drawers.cls
'this is really a typed collection object
'it needs and indexer & enumerator

'This is part of Petty Cashier

Private m_sUserID As String
Private m_cDrawers As Collection
Private WithEvents m_oCurDrawer As Drawer
Attribute m_oCurDrawer.VB_VarHelpID = -1

Public Event CurDrawerStdTrxValidityChanged(IsValid As Boolean)
Public Event CurDrawerTransTrxValidityChanged(IsValid As Boolean)


Public Property Get UserID() As String
    UserID = m_sUserID
End Property


Public Property Get Drawers() As Collection
    Set Drawers = m_cDrawers
End Property


Public Property Get CurDrawer() As Drawer
    Set CurDrawer = m_oCurDrawer
End Property

Public Property Let CurDrawer(ByVal vNewValue As Drawer)
    Set m_oCurDrawer = vNewValue
End Property


Private Sub Class_Initialize()
    
    Set m_cDrawers = New Collection
    Set m_oCurDrawer = New Drawer

    m_sUserID = GetUserName
    
    Call GetDrawers
End Sub


Private Sub m_oCurDrawer_StdTrxValidityChanged(IsValid As Boolean)
    RaiseEvent CurDrawerStdTrxValidityChanged(IsValid)
End Sub


Private Sub m_oCurDrawer_TransTrxValidityChanged(IsValid As Boolean)
    RaiseEvent CurDrawerTransTrxValidityChanged(IsValid)
End Sub


Private Sub GetDrawers()
    Dim aDrawer As Drawer

    'Hit the DB (UserDrawer table) and get all the drawer keys for this user
    Dim lors As ADODB.Recordset
    Dim lsSql As String
    
    lsSql = "select DrawerKey, IsPref from tcpPCDrawer "
    lsSql = lsSql & "where UserID = '" & m_sUserID & "'"
    
    Set lors = New ADODB.Recordset
    lors.Source = lsSql
    Set lors.ActiveConnection = g_DB.Connection
    lors.Open
    
    While lors.EOF = False
        Set aDrawer = New Drawer
        aDrawer.Load lors!Drawerkey, True
        m_cDrawers.Add aDrawer
        
        If lors!IsPref = True Then
            Set m_oCurDrawer = aDrawer
        End If
        
        Set aDrawer = Nothing
        lors.MoveNext
    Wend
    
    'safeguard against NO preferred drawers
    If m_oCurDrawer Is Nothing Then
        If m_cDrawers.Count > 0 Then
            Set m_oCurDrawer = m_cDrawers(1)
        End If
    End If
        
    lors.Close
    Set lors = Nothing
End Sub



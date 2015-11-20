VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PRIVATE VARIABLES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_sDescr As String
Private m_iDrawerKey As Long
Private m_dBalance As Double

Private m_sPossDebitTrx As Collection
Private m_sPossCreditTrx As Collection

Private m_sTrxLog As Collection
Private m_oTranDrawer As Drawer

Private WithEvents m_oStdTrx As Transaction
Attribute m_oStdTrx.VB_VarHelpID = -1
Private WithEvents m_oTransTrx As Transaction
Attribute m_oTransTrx.VB_VarHelpID = -1

Public Event StdTrxValidityChanged(IsValid As Boolean)
Public Event TransTrxValidityChanged(IsValid As Boolean)

Private Sub m_oStdTrx_ValidityChanged(IsValid As Boolean)
    RaiseEvent StdTrxValidityChanged(IsValid)
End Sub

Private Sub m_oTransTrx_ValidityChanged(IsValid As Boolean)
    RaiseEvent TransTrxValidityChanged(IsValid)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PROPERITES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get StandardTrx() As Transaction
    Set StandardTrx = m_oStdTrx
End Property


Public Property Get TransTrx() As Transaction
    Set TransTrx = m_oTransTrx
End Property


Public Property Get TranDrawer() As Drawer
    Set TranDrawer = m_oTranDrawer
End Property


Public Property Get Descr() As String
    Descr = m_sDescr
End Property


Public Property Get Drawerkey() As String
    Drawerkey = m_iDrawerKey
End Property


Public Property Get Balance() As Double
    Balance = m_dBalance
End Property

Public Property Let Balance(value As Double)
    m_dBalance = value
End Property


Public Property Get PossDebitTrx() As Collection
    Set PossDebitTrx = m_sPossDebitTrx
End Property

Public Property Get PossCreditTrx() As Collection
    Set PossCreditTrx = m_sPossCreditTrx
End Property


Public Property Get TrxLog() As Collection
    Set TrxLog = m_sTrxLog
End Property



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PUBLIC SUB/FUNCTION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'LOAD
Public Function Load(Drawerkey As Long, AddTranDrawer As Boolean)
    Dim aTrxTypeDef As TransactionType
    Dim aTrx As Transaction
    
    m_iDrawerKey = Drawerkey
    
    'Hit the db and get info for this drawer
    Dim loRS As ADODB.Recordset
    Dim lsSql As String
    
    lsSql = "select * from tcpPCDrawerdef "
    lsSql = lsSql & "Where DrawerKey = " & m_iDrawerKey
    
    Set loRS = New ADODB.Recordset
    loRS.Source = lsSql
    Set loRS.ActiveConnection = g_DB.Connection
    loRS.Open
    
    If loRS.EOF = False Then
        m_sDescr = Trim("" & loRS!Descr)
        
        'Load TranDrawer if appropriate
        If AddTranDrawer Then
            If loRS!TranDrawerkey <> 0 Then
                Set m_oTranDrawer = New Drawer
                Call m_oTranDrawer.Load(loRS!TranDrawerkey, False)
            Else
                Set m_oTranDrawer = Nothing
            End If
        End If
    End If
    
'--------------------------------------------------
    'Get trx types
    lsSql = "Select * from tcpPCTrxType a inner join "
    lsSql = lsSql & "tcpPCtrxtypedef b on a.pctrxtypekey = b.pctrxtypekey "
    lsSql = lsSql & "Where DrawerKey = " & m_iDrawerKey
    
    Set loRS = New ADODB.Recordset
    loRS.Source = lsSql
    Set loRS.ActiveConnection = g_DB.Connection
    loRS.Open
    
    Do Until loRS.EOF = True
        Set aTrxTypeDef = New TransactionType
        aTrxTypeDef.Descr = loRS!PCTrxTypeDescr
        aTrxTypeDef.TrxTypeKey = loRS!PCTrxTypeKey

        If loRS!IsDebit Then
            aTrxTypeDef.IsDebit = True
            m_sPossDebitTrx.Add aTrxTypeDef
        Else
            aTrxTypeDef.IsDebit = False
            m_sPossCreditTrx.Add aTrxTypeDef
        End If

        Set aTrxTypeDef = Nothing
        loRS.MoveNext
    Loop
    loRS.Close
    
    'set initial trx log rowcount to 50 (form has not displayed yet)
    Call TrxLogData(50)
End Function


Public Sub TrxLogData(RowCount As Integer)
    Dim ocmd As ADODB.Command
    Dim loRS As ADODB.Recordset
    Dim aTrx As Transaction

    Set ocmd = New ADODB.Command
    Set loRS = New ADODB.Recordset
    
    With ocmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spCPC_PCGetTransactions"
        .Parameters.Append .CreateParameter("@InParamRowCount", adInteger, adParamInput)
            .Parameters("@InParamRowCount").value = RowCount
            
        .Parameters.Append .CreateParameter("@InParamDrawerKey", adInteger, adParamInput)
            .Parameters("@InParamDrawerKey").value = m_iDrawerKey
            
        Set loRS = .Execute
    End With
    
    'remove any records already in the grid
    While (m_sTrxLog.Count > 0)
        m_sTrxLog.Remove (m_sTrxLog.Count)
    Wend
    
    'save final balance
    m_dBalance = loRS!Balance
    
    Do Until loRS.EOF = True
        Set aTrx = New Transaction
 
        aTrx.TrxDate = loRS!TrxDate
        aTrx.TrxAmt = loRS!Amount
        aTrx.TrxDescr = loRS!Descr
        aTrx.TrxBalance = loRS!Balance
        
        m_sTrxLog.Add aTrx

        Set aTrx = Nothing
        loRS.MoveNext
    Loop
End Sub


Public Sub TransferBal(TranAmt As Double, Notes As String, UserID As String)
    Dim aTrx As Transaction
     
    Set aTrx = New Transaction
    
    'TRANSFEROUT
    'smr 12-01-2004 - added below line because the drawer is used by more than one person
    SetBalance
    aTrx.Init TRANSFEROUT, m_dBalance
    aTrx.TrxAmt = TranAmt
    aTrx.TrxDate = Now
    aTrx.TrxDescr = "Transfer Out"
    aTrx.TrxNotes = Notes
    Call aTrx.Save(m_iDrawerKey, UserID)
    m_dBalance = m_dBalance + aTrx.TrxAmt
    m_oTransTrx.TrxBalance = m_dBalance
        
        
    'TRANSFERIN
    'need to get the current tranDrawer balance from the database
    ''''aTrx.Init TRANSFERIN, m_oTranDrawer.Balance
    m_oTranDrawer.SetBalance
    aTrx.Init TRANSFERIN, m_oTranDrawer.Balance
    
    aTrx.TrxDescr = "Transfer In"
    aTrx.TrxAmt = aTrx.TrxAmt
    Call aTrx.Save(m_oTranDrawer.Drawerkey, UserID)
        
    'This updates the transfer drawer NOT the real Other drawer....
    m_oTranDrawer.Balance = m_oTranDrawer.Balance + aTrx.TrxAmt
    
    Set aTrx = Nothing
End Sub


Public Sub SetBalance()
    'hit the database and get the most recent balance
    Dim ocmd As ADODB.Command
    Dim loRS As ADODB.Recordset
    
    Set ocmd = New ADODB.Command
    Set loRS = New ADODB.Recordset
    
    With ocmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spCPC_PCGetTransactions"
        .Parameters.Append .CreateParameter("@InParamRowCount", adInteger, adParamInput)
            .Parameters("@InParamRowCount").value = 1
        .Parameters.Append .CreateParameter("@InParamDrawerKey", adInteger, adParamInput)
            .Parameters("@InParamDrawerKey").value = m_iDrawerKey
        Set loRS = .Execute
    End With
    
    'save final balance
    m_dBalance = loRS!Balance
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PRIVATE SUB
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    m_sDescr = ""
    m_iDrawerKey = 0
    m_dBalance = 0
    Set m_sPossDebitTrx = New Collection
    Set m_sPossCreditTrx = New Collection
    Set m_sTrxLog = New Collection
    
    Set m_oStdTrx = New Transaction
    Set m_oTransTrx = New Transaction
End Sub



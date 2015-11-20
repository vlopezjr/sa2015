VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PRIVATE VARIABLES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_dTrxDate As Date
Private m_dTrxAmt As Double
Private m_dTrxBalance As Double
Private m_sTrxDescr As String
Private m_sTrxNotes As String
Private m_iPCTrxKey As Long
Private m_oTrxType As TransactionType
Private m_dTransferTrxAmt As Double

'Broken Rules
Private m_colBrokenRules As Collection
Public Event ValidityChanged(IsValid As Boolean)

'Report tab
'Public Sub DeleteTransactionTemp()
'    Dim lsSql As String
'    lsSql = "Delete from ???"
'    g_DB.Connection.Execute (lsSql)
'End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PROPERITES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get TrxDate() As Date
    TrxDate = m_dTrxDate
End Property

Public Property Let TrxDate(Value As Date)
    m_dTrxDate = Value
End Property


Public Property Get TrxAmt() As Double
    TrxAmt = m_dTrxAmt
End Property

Public Property Let TrxAmt(Value As Double)
    Dim BrokenRule As Rule
    
    m_dTrxAmt = Value

    Set BrokenRule = New Rule
    BrokenRule.Key = 1

    BrokenRule.Descr = "Amount is a numeric value greater than zero."
    If m_dTrxAmt <= 0 Or Len(Trim(m_dTrxAmt)) = 0 Then
        AddBrokenRule BrokenRule
    Else
        RemoveBrokenRule BrokenRule
    End If
    
End Property


Public Property Get TransferTrxAmt() As Double
    TransferTrxAmt = m_dTransferTrxAmt
End Property

Public Property Let TransferTrxAmt(Value As Double)
    Dim BrokenRule As Rule
    
    m_dTransferTrxAmt = Value

    Set BrokenRule = New Rule
    BrokenRule.Descr = "Amount is a numeric value greater than zero."
    BrokenRule.Key = 8
    
    If m_dTransferTrxAmt >= 0 Or Len(Trim(m_dTransferTrxAmt)) = 0 Then
        AddBrokenRule BrokenRule
    Else
        RemoveBrokenRule BrokenRule
    End If
    
End Property


Public Property Get TrxBalance() As Double
    'The balance AFTER the trans is saved - for display in log
    TrxBalance = m_dTrxBalance
End Property

Public Property Let TrxBalance(Value As Double)
    'this should only be called when loading the transactions from the db
    m_dTrxBalance = Value
End Property


Public Property Get TrxDescr() As String
    TrxDescr = m_sTrxDescr
End Property

Public Property Let TrxDescr(Value As String)
    Dim BrokenRule As Rule
    
    m_sTrxDescr = Value

    Set BrokenRule = New Rule
    BrokenRule.Descr = "Transaction Descr cannot be empty."
    BrokenRule.Key = 2
    
    If Value = "" Then
        AddBrokenRule BrokenRule
    Else
        RemoveBrokenRule BrokenRule
    End If
    
End Property


Public Property Get PCTrxKey() As Long
    PCTrxKey = m_iPCTrxKey
End Property

Public Property Let PCTrxKey(Value As Long)
    m_iPCTrxKey = Value
End Property


Public Property Get TrxType() As TransactionType
    Set TrxType = m_oTrxType
End Property

Public Property Let TrxType(Value As TransactionType)
     Set m_oTrxType = Value
End Property


Public Property Get TrxNotes() As String
    TrxNotes = m_sTrxDescr
End Property

Public Property Let TrxNotes(Value As String)
    m_sTrxNotes = Value
End Property



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PUBLIC SUB/FUNCTION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(InitialTrxType As Long, CurBalance As Double)
    Set m_oTrxType = New TransactionType
    m_oTrxType.TrxTypeKey = InitialTrxType
    m_dTrxBalance = CurBalance
End Sub


Public Function Save(Drawerkey As Long, UserID As String)
    Dim ocmd As ADODB.Command
    Dim lngRetVal As Long
    
    'Cash Out (example: Cash Refund) - Credit to the customer
    If m_oTrxType.IsDebit = True Then
        m_dTrxAmt = m_dTrxAmt       'Cash In
    Else
        m_dTrxAmt = -m_dTrxAmt      'Cash Out
    End If
    
    'set the Trx balance of the trx to the NEW balance so the form can grab it
    m_dTrxBalance = m_dTrxAmt + m_dTrxBalance 'm_dCurBalance
    
    If m_sTrxNotes = "" Then m_sTrxNotes = "Note"
    
    Set ocmd = New ADODB.Command
    With ocmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spCPC_PCInsertTransaction"
        .Parameters.Append .CreateParameter("RETURN_VALUE", adVariant, adParamReturnValue, 0, lngRetVal)
        
        'smr - 12/01/2004 - instead we are using getdate() inside of the sproc
        '.Parameters.Append .CreateParameter("@InParamTrxDate", adDate, adParamInput)
        '    .Parameters("@InParamTrxDate").value = m_dTrxDate
            
        .Parameters.Append .CreateParameter("@InParamAmount", adCurrency, adParamInput, 11, m_dTrxAmt)
        .Parameters.Append .CreateParameter("@InParamDescr", adVarChar, adParamInput, Len(m_sTrxDescr), m_sTrxDescr)
        .Parameters.Append .CreateParameter("@InParamDrawerKey", adInteger, adParamInput)
            .Parameters("@InParamDrawerKey").Value = Drawerkey
        .Parameters.Append .CreateParameter("@InParamBalance", adCurrency, adParamInput, 11, m_dTrxBalance)
        .Parameters.Append .CreateParameter("@InParamPCTrxTypeKey", adInteger, adParamInput)
            .Parameters("@InParamPCTrxTypeKey").Value = m_oTrxType.TrxTypeKey
        .Parameters.Append .CreateParameter("@InParamUserID", adVarChar, adParamInput, Len(UserID), UserID)
        .Parameters.Append .CreateParameter("@InParamNotes", adVarChar, adParamInput, Len(m_sTrxNotes), m_sTrxNotes)
        .Execute
        'Add a PC Transaction Key to the Trx object
        m_iPCTrxKey = .Parameters("RETURN_VALUE").Value
    End With
End Function


Public Sub Reset(InitialTrxTypeKey As Long, CurBalance As Double)
    Call ClearVariables
    Call Init(InitialTrxTypeKey, CurBalance)
End Sub


Public Function GetBrokenRules() As Collection
    Set GetBrokenRules = m_colBrokenRules
End Function


Public Property Get IsValid() As Boolean
    IsValid = (m_colBrokenRules.Count = 0)
End Property


'To Support Applications
'-----------------------
Public Function GetInvAppls(CustID As String) As Recordset
    Dim lors As ADODB.Recordset
    Dim lsSql As String
    
    '***CustID is known
    lsSql = "SELECT dbo.tarInvoice.InvcKey as OPKey, 'InvNbr' as TranType, "
    lsSql = lsSql & "dbo.tarInvoice.TranID as DocNbr, dbo.tarInvoice.TranDate, "
    lsSql = lsSql & "dbo.tarCustomer.CustID, dbo.tarCustomer.CustName, "
    lsSql = lsSql & "dbo.tarInvoice.TranAmt, 0 "  'added 0
    
    lsSql = lsSql & "FROM dbo.tarCustomer INNER JOIN "
    lsSql = lsSql & "dbo.tarInvoice ON dbo.tarCustomer.CustKey = dbo.tarInvoice.CustKey "
    lsSql = lsSql & "WHERE (dbo.tarInvoice.Status = 1) "
    If Len(Trim(CustID)) > 0 Then
        lsSql = lsSql & "AND (dbo.tarCustomer.CustID like '%" & CustID & "%')"
    End If
    
    Set lors = New ADODB.Recordset
    lors.Source = lsSql
    Set lors.ActiveConnection = g_DB.Connection
    lors.Open
    
    Set GetInvAppls = lors
    Set lors = Nothing
End Function


Public Function GetInvApplsPO(PONumber As Long) As Recordset
    Dim lors As ADODB.Recordset
    Dim lsSql As String
    
    'if ***OP number*** is know
    lsSql = "SELECT dbo.tcpSO.OPKey as OPKey, 'OPNbr' as TranType, "
    lsSql = lsSql & "convert(varchar(50),dbo.tcpSO.OPKey) as DocNbr, "
    lsSql = lsSql & "dbo.tcpSO.CreateDate as TranDate, dbo.tarCustomer.CustID, "
    lsSql = lsSql & "dbo.tarCustomer.CustName, "
    lsSql = lsSql & "dbo.tcpSO.TotalPrice + dbo.tcpSO.TotalTax as TranAmt, "
    
    'smr for inbound freight (0 = no inbound freight;   1 = inbound freight)
    lsSql = lsSql & "(case when tcpso.flags&0x100=0x100 then 1 else 0 end) as IBF "
    
    lsSql = lsSql & "FROM dbo.tcpSO INNER JOIN "
    lsSql = lsSql & "dbo.tarCustomer ON dbo.tcpSO.CustKey = dbo.tarCustomer.CustKey "
    lsSql = lsSql & "Where (dbo.tcpSO.OPKey = " & PONumber & ")"
    
    Set lors = New ADODB.Recordset
    lors.Open lsSql, g_DB.Connection
    Set lors.ActiveConnection = Nothing
    Set GetInvApplsPO = lors
    
    'Set loRS = Nothing
    
End Function


Public Function GetInvApplsUnion(OPNumber As Long, CustID As String) As Recordset
    Dim lors As ADODB.Recordset
    Dim lsSql As String
    
    'if ***OP number*** is know and
    lsSql = "SELECT dbo.tcpSO.OPKey as OPKey, 'OPNbr' as TranType, "
    lsSql = lsSql & "convert(varchar(50),dbo.tcpSO.OPKey) as DocNbr, "
    lsSql = lsSql & "dbo.tcpSO.CreateDate as TranDate, dbo.tarCustomer.CustID, "
    lsSql = lsSql & "dbo.tarCustomer.CustName, "
    lsSql = lsSql & "dbo.tcpSO.TotalPrice + dbo.tcpSO.TotalTax as TranAmt, "
    
    'smr for inbound freight (0 = no inbound freight;   1 = inbound freight)
    lsSql = lsSql & "(case when tcpso.flags&0x100=0x100 then 1 else 0 end) as IBF "
    
    lsSql = lsSql & "FROM dbo.tcpSO INNER JOIN "
    lsSql = lsSql & "dbo.tarCustomer ON dbo.tcpSO.CustKey = dbo.tarCustomer.CustKey "
    lsSql = lsSql & "Where (dbo.tcpSO.OPKey = " & OPNumber & ")"
    
    lsSql = lsSql & " Union "
    
    '***CustID is known
    lsSql = lsSql & "SELECT dbo.tarInvoice.InvcKey as OPKey, 'InvNbr' as TranType, "
    lsSql = lsSql & "dbo.tarInvoice.TranID as DocNbr, dbo.tarInvoice.TranDate, "
    lsSql = lsSql & "dbo.tarCustomer.CustID, dbo.tarCustomer.CustName, "
    lsSql = lsSql & "dbo.tarInvoice.TranAmt, 0 " 'added 0 for IBF
    lsSql = lsSql & "FROM dbo.tarCustomer INNER JOIN "
    lsSql = lsSql & "dbo.tarInvoice ON dbo.tarCustomer.CustKey = dbo.tarInvoice.CustKey "
    lsSql = lsSql & "WHERE (dbo.tarInvoice.Status = 1) "
    If Len(Trim(CustID)) > 0 Then
        'lsSql = lsSql & "AND (dbo.tarCustomer.CustID = '" & CustID & "')"
        lsSql = lsSql & "AND (dbo.tarCustomer.CustID like '%" & CustID & "%')"
    End If
    
    Set lors = New ADODB.Recordset
    lors.Source = lsSql
    Set lors.ActiveConnection = g_DB.Connection
    lors.Open
    
    Set GetInvApplsUnion = lors
    Set lors = Nothing
End Function


Public Sub SetApplications(loSelItems As Collection)
    Dim aSelItem As Application

    For Each aSelItem In loSelItems
        Call aSelItem.Insert
    Next
End Sub


Public Function GetTransResearch(RowCount As Integer, CustName As String, StartDate As String, EndDate As String, _
    DocNbr As String, Amount As String) As Recordset

    Dim ocmd As ADODB.Command
    Dim lors As ADODB.Recordset
    Dim aTrx As Transaction
    Dim lsAmount As String
    
    Set lors = New ADODB.Recordset
    
    CustName = "%" & CustName & "%"
    DocNbr = "%" & DocNbr & "%"
    
    lsAmount = "%" & Format(Amount, "#0") & ".%"
    
    Set ocmd = New ADODB.Command
    With ocmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spCPC_PCGetResearchRecs"
        .Parameters.Append .CreateParameter("@InParamRowCount", adInteger, adParamInput)
            .Parameters("@InParamRowCount").Value = RowCount
        .Parameters.Append .CreateParameter("@InParamCustName", adVarChar, adParamInput, Len(CustName), CustName)
        .Parameters.Append .CreateParameter("@InParamDescr", adDate, adParamInput, , StartDate)
        .Parameters.Append .CreateParameter("@InParamDescr", adDate, adParamInput, , EndDate)
        
        .Parameters.Append .CreateParameter("@InParamDescr", adVarChar, adParamInput, Len(DocNbr), DocNbr)
        .Parameters.Append .CreateParameter("@InParamAmount", adVarChar, adParamInput, Len(lsAmount), lsAmount)
        Set lors = .Execute
    End With
    
    Set GetTransResearch = lors
    Set lors = Nothing
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PRIVATE SUB/FUNCTION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_colBrokenRules = New Collection
    Call ClearVariables
End Sub


Private Sub ClearVariables()
    m_dTrxDate = Now
    m_dTrxBalance = 0
    m_dTrxAmt = 0
    TrxDescr = ""
    Set m_oTrxType = New TransactionType
    m_sTrxNotes = ""
    m_iPCTrxKey = 0
    m_dTransferTrxAmt = 0
End Sub


'Routines to support broken rules
Private Sub AddBrokenRule(BrokenRule As Rule)
    Dim aRule As Rule
    For Each aRule In m_colBrokenRules
        If aRule.Key = BrokenRule.Key Then
            Exit Sub
        End If
    Next
    
    m_colBrokenRules.Add BrokenRule, CStr(BrokenRule.Key)
    If m_colBrokenRules.Count = 1 Then
        RaiseEvent ValidityChanged(False)
    End If
End Sub


Private Sub RemoveBrokenRule(BrokenRule As Rule)
    Dim aRule As Rule
    Dim i As Integer
    
    For i = 0 To m_colBrokenRules.Count - 1
        Set aRule = m_colBrokenRules(i + 1)
        If aRule.Key = BrokenRule.Key Then
            m_colBrokenRules.Remove (i + 1)
            If m_colBrokenRules.Count = 0 Then
                RaiseEvent ValidityChanged(True)
            End If
            Exit Sub
        End If
    Next
End Sub



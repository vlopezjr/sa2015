VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PRIVATE VARIABLES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_lAppKey As Long
Private m_sAppTranType As String
Private m_sAppDocNbr As String
Private m_dAppTranDate As Date
Private m_sAppCustID As String
Private m_sAppCustName As String
Private m_dAppAmount As Double
Private m_iIBF As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PROPERITES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get AppKey() As Long
    AppKey = m_lAppKey
End Property

Public Property Let AppKey(value As Long)
    m_lAppKey = value
End Property


Public Property Get AppTranType() As String
    AppTranType = m_sAppTranType
End Property

Public Property Let AppTranType(value As String)
    m_sAppTranType = value
End Property


Public Property Get AppDocNbr() As String
    AppDocNbr = m_sAppDocNbr
End Property

Public Property Let AppDocNbr(value As String)
    m_sAppDocNbr = value
End Property


Public Property Get AppTranDate() As Date
    AppTranDate = m_dAppTranDate
End Property

Public Property Let AppTranDate(value As Date)
    m_dAppTranDate = value
End Property


Public Property Get AppCustID() As String
    AppCustID = m_sAppCustID
End Property

Public Property Let AppCustID(value As String)
    m_sAppCustID = value
End Property


Public Property Get AppCustName() As String
    AppCustName = m_sAppCustName
End Property

Public Property Let AppCustName(value As String)
    m_sAppCustName = value
End Property


Public Property Get AppAmt() As Double
    AppAmt = m_dAppAmount
End Property

Public Property Let AppAmt(value As Double)
    m_dAppAmount = value
End Property


Public Property Get AppIBF() As Integer
    AppIBF = m_iIBF
End Property

Public Property Let AppIBF(value As Integer)
    m_iIBF = value
End Property



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PUBLIC SUB/FUNCTION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Insert()
    'called from transaction class
    Dim lsSql As String
    Dim liCounter As Integer
    
    lsSql = "Insert into tcpPCApplication "
    lsSql = lsSql & "(PCTrxKey, TranType, DocNbr, CustName, TranAmt) "
    lsSql = lsSql & "Values "
    lsSql = lsSql & "(" & m_lAppKey & ", '" & m_sAppTranType & "', '" & m_sAppDocNbr & "', '"
    lsSql = lsSql & Replace(m_sAppCustName, "'", "''") & "', " & m_dAppAmount & ")"

    g_DB.Connection.Execute (lsSql)
End Function


'Report tab
Public Sub DeleteApplTemp()
    Dim lsSql As String
    lsSql = "Delete from tcpPCApplTemp"
    g_DB.Connection.Execute (lsSql)
End Sub


'Report tab
Public Sub InsertApplTemp(StartDate As Date, EndDate As Date)
    Dim lsSql As String
    Dim liCounter As Integer
    Dim lors As ADODB.Recordset
    Dim lsNotes As String
    
    lsSql = "Select tcpPCApplication.*, tcpPCTransaction.trxdate, tcpPCTransaction.Notes, tcpPCDrawerDef.Descr "
    lsSql = lsSql & "From tcpPCTransaction inner join tcpPCApplication "
    lsSql = lsSql & "ON tcpPCTransaction.PCTrxKey = tcpPCApplication.PCTrxKey "
    lsSql = lsSql & "INNER JOIN tcpPCDrawerDef ON tcpPCTransaction.Drawerkey = tcpPCDrawerDef.Drawerkey "
    lsSql = lsSql & "Where TrxDate >= '" & StartDate & "' "
    lsSql = lsSql & "And TrxDate <= '" & EndDate & "'"
    
    
    lsSql = "SELECT dbo.tcpPCApplication.*, dbo.tcpPCTransaction.TrxDate AS TranDate, " & _
            "dbo.tcpPCTransaction.Notes AS Notes, dbo.tcpPCDrawerDef.Descr AS Descr, " & _
            "ISNULL(tarCustomer_1.CustID, tarCustomer_2.CustID) As CustID " & _
            "FROM dbo.tarCustomer tarCustomer_1 INNER JOIN " & _
            "dbo.tarInvoice ON tarCustomer_1.CustKey = dbo.tarInvoice.CustKey RIGHT OUTER JOIN " & _
            "dbo.tcpPCTransaction INNER JOIN " & _
            "dbo.tcpPCApplication ON dbo.tcpPCTransaction.PCTrxKey = dbo.tcpPCApplication.PCTrxKey INNER JOIN " & _
            "dbo.tcpPCDrawerDef ON dbo.tcpPCTransaction.DrawerKey = dbo.tcpPCDrawerDef.DrawerKey ON " & _
            "dbo.tarInvoice.TranID = dbo.tcpPCApplication.DocNbr LEFT OUTER JOIN " & _
            "dbo.tcpSO INNER JOIN " & _
            " dbo.tarCustomer tarCustomer_2 ON dbo.tcpSO.CustKey = tarCustomer_2.CustKey ON " & _
                "dbo.tcpPCApplication.DocNbr = CONVERT(varchar, dbo.tcpSO.OPKey) " & _
            "WHERE (dbo.tcpPCTransaction.TrxDate) >= '" & StartDate & "' " & _
            "AND (dbo.tcpPCTransaction.TrxDate) <= '" & EndDate & "'"
    
    Set lors = New ADODB.Recordset
    lors.Source = lsSql
    Set lors.ActiveConnection = g_DB.Connection
    lors.Open
    
    Do Until lors.EOF = True
        lsNotes = IIf(IsNull(lors!Notes), "", lors!Notes)
    
        lsSql = "Insert into tcpPCApplTemp "
        lsSql = lsSql & "(AppKey, PCTrxKey, TranType, DocNbr, CustName, TranAmt, "
        lsSql = lsSql & "Descr, TranDate, Notes, CustID) "
        lsSql = lsSql & "Values "
        lsSql = lsSql & "(" & lors!AppKey & ", "
        lsSql = lsSql & lors!PCTrxKey & ", '"
        lsSql = lsSql & lors!TranType & "', '"
        lsSql = lsSql & lors!DocNbr & "', '"
        lsSql = lsSql & Replace(lors!CustName, "'", "''") & "', "
        lsSql = lsSql & Format(lors!TranAmt, "#.00") & ", '"
        lsSql = lsSql & Replace(Trim(lors!Descr), "'", "''") & "', '"
        lsSql = lsSql & Format(lors!TranDate, "MM-DD-YYYY") & "', '"
        lsSql = lsSql & Replace(lsNotes, "'", "''") & "', '"
        lsSql = lsSql & Trim(lors!CustID) & "')"

        g_DB.Connection.Execute (lsSql)
        lors.MoveNext
    Loop
End Sub

Public Sub InsertApplWorking(StartDate As Date, EndDate As Date)
    Dim oCmd As ADODB.Command
    
    Set oCmd = New ADODB.Command
    
    With oCmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spCPC_PCApplWorking"
        .Parameters.Append .CreateParameter("@InParamStartDate", adDate, adParamInput, , StartDate)
        .Parameters.Append .CreateParameter("@InParamEndDate", adDate, adParamInput, , EndDate)
        .Execute
    End With
    Set oCmd = Nothing
End Sub










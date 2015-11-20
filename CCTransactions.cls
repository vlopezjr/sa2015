VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_bHasBeenAuthorized As Boolean
Private m_bHasBeenCharged As Boolean
Private m_bHasErrors As Boolean

Private m_colTransactions As Collection

Private m_sAuthorizationNo As String
Private m_sAuthorizationPNREF As String


Public Property Get HasBeenAuthorized() As Boolean
'iterate through the Transactions collection and see if there's an authorization record
    HasBeenAuthorized = m_bHasBeenAuthorized
End Property

'There is a HasBeenCharged field in tcpSOCreditCard. How is this maintained?

Public Property Get HasBeenCharged() As Boolean
    HasBeenCharged = m_bHasBeenCharged
End Property


Public Property Get HasErrors() As Boolean
    HasErrors = m_bHasErrors
End Property


Public Property Get AuthorizationNo() As String
    AuthorizationNo = m_sAuthorizationNo
End Property


Public Property Get AuthorizationPNREF() As String
    AuthorizationPNREF = m_sAuthorizationPNREF
End Property


Private Sub Class_Initialize()
    Set m_colTransactions = New Collection
End Sub


Private Sub Class_Terminate()
    Set m_colTransactions = Nothing
End Sub


Public Sub Add(ByRef oTran As CCTransaction)
    oTran.Save
    m_colTransactions.Add oTran
End Sub


Public Property Get Count() As Integer
    Count = m_colTransactions.Count
End Property


Public Function Item(vIndex As Variant) As CCTransaction
    Item = m_colTransactions.Item(vIndex)
End Function


'This is a hidden member. It supports For Each Next.

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colTransactions.[_NewEnum]
End Function


'Load all existing transactions for an order
'fetch a recordset of all transactions for this order
'build a collection of transaction objects
'check transaction types and set these flags accordingly
'    m_bHasBeenAuthorized
'    m_bIsCharged
'    m_colTransactions

Public Sub Load(CCKey As Long, OPKey As Long)
    Dim orst As ADODB.Recordset
    Dim oTran As CCTransaction
    
'    Set orst = New ADODB.Recordset
    Set orst = CallSP("spcpcCCTransactionsLoad", "@OPKey", OPKey, "@CCKey", CCKey)
    
    With orst
        If Not .EOF Then
            Do While .EOF = False
                Set oTran = New CCTransaction
                oTran.OPKey = OPKey
                oTran.CCKey = CCKey
                
                oTran.TranKey = .Fields("TranKey").value
                oTran.TranType = .Fields("TranType").value
                oTran.TimeStamp = .Fields("CreateDate").value
                oTran.UserID = .Fields("UserID").value
                oTran.Amount = .Fields("Amount").value
                oTran.InvcKey = .Fields("InvcKey").value
                
                oTran.Response = .Fields("Response").value
                
                m_colTransactions.Add oTran, CStr(oTran.TranKey)
                
                If oTran.TranType = "A" And oTran.RESPMSG = "Approved" Then
                    m_bHasBeenAuthorized = True
                    m_sAuthorizationNo = oTran.AUTHCODE
                    m_sAuthorizationPNREF = oTran.PNREF
                End If
                If (oTran.TranType = "D" Or oTran.TranType = "S") And oTran.RESPMSG = "Approved" Then
                    m_bHasBeenCharged = True
                End If
                If oTran.Result <> 0 Then
                    m_bHasErrors = True
                End If
                orst.MoveNext
            Loop
        End If
    End With
End Sub


Public Sub Backup()

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreditLimitList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_col As Collection

Private Sub Class_Initialize()
    Set m_col = New Collection
End Sub


Private Sub Class_Terminate()
    Set m_col = Nothing
End Sub


Public Sub Add(ByRef i_oCreditLimit As CreditLimitRecord)
    m_col.Add i_oCreditLimit
End Sub


Public Sub Remove(vIndex As Variant)
    On Error Resume Next
    m_col.Remove vIndex
End Sub


Public Property Get Count() As Long
    Count = m_col.Count
End Property


Public Function Item(vIndex As Variant) As CreditLimitRecord
Attribute Item.VB_UserMemId = 0
    Set Item = m_col(vIndex)
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_col.[_NewEnum]
End Function


Public Sub LoadData( _
    lTerritory As Long, dtpUpdate As Date, _
    lCurrCredit As Long, lNewCredit As Long, _
    lRowCount As Long, bUpdate As Boolean _
)
    Dim rstCreditLimit As ADODB.Recordset

    SetWaitCursor True
    Dim creditlimit As CreditLimitRecord

    Set rstCreditLimit = CallSP("spCPCGetCreditLimit", _
         "@_iSalesTerritoryKey", lTerritory, _
         "@_iLastUpdated", dtpUpdate, _
         "@_iMonthsCredit", lCurrCredit, _
         "@_iNewMonthsCredit", lNewCredit, _
         "@_iRowCount", lRowCount, _
         "@_bUpdate", bUpdate)
         
    If rstCreditLimit Is Nothing Then Exit Sub
    With rstCreditLimit
        While Not .EOF
            Set creditlimit = New CreditLimitRecord
            creditlimit.AgingDate = .Fields("AgingDate")
            creditlimit.ColRems = .Fields("RemCnt")
            creditlimit.CurAmount = .Fields("CurAmt")
            creditlimit.CurLimit = .Fields("CreditLimit")
            creditlimit.CustID = .Fields("CustID")
            creditlimit.CustName = .Fields("CustName")
            creditlimit.l30DaysAmount = .Fields("30days")
            creditlimit.l45DaysAmount = .Fields("45days")
            creditlimit.NewLimit = .Fields("NewLimit")
            creditlimit.Terms = .Fields("PmtTermsID")
            creditlimit.TerritoryKey = .Fields("SalesTerritoryKey")
            creditlimit.Update = .Fields("Upd")
            creditlimit.YearRec = .Fields("TotalYrlyRec")
            creditlimit.CustKey = .Fields("CustKey")
            If IsNull(.Fields("LastChanged")) Then
                creditlimit.LastChanged = Null
            Else
                creditlimit.LastChanged = .Fields("LastChanged")
            End If
            m_col.Add creditlimit
            .MoveNext
            Set creditlimit = Nothing
        Wend
    End With
    SetWaitCursor False
End Sub


Public Sub UpdateDatabase()
    Dim obj As CreditLimitRecord
    Dim lIndex As Long
    lIndex = 0
    
    For Each obj In m_col
        If obj.Update Then
            CallSP "spCPCUpdateCreditLimit", _
                    "@_iCustKey", obj.CustKey, _
                    "@_iOldLimit", obj.CurLimit, _
                    "@_iNewLimit", obj.NewLimit, _
                    "@_iUserID", GetUserID
        End If
    Next

    For Each obj In m_col
        If obj.Update Then
            m_col.Remove lIndex
        End If
        lIndex = lIndex + 1
    Next
End Sub

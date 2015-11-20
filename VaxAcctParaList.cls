VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UPSAcctList"
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


Public Sub Add(ByRef i_oUPSAcct As UPSAcct)
    m_col.Add i_oUPSAcct
End Sub


Public Sub Remove(vIndex As Variant)
    On Error Resume Next
    m_col.Remove vIndex
End Sub


Public Property Get Count() As Long
    Count = m_col.Count
End Property


Public Function Item(vIndex As Variant) As UPSAcct
    Set Item = m_col(vIndex)
End Function


Public Function NewEnum() As IUnknown
    Set NewEnum = m_col.[_NewEnum]
End Function


Public Sub LoadUpsAcct(lCustKey As Long)
    On Error GoTo ErrorHandler
    Dim rstUPSAcct As ADODB.Recordset
    Dim oUPSAcct As UPSAcct
    Dim sSQL As String
    
    SetWaitCursor True
    
    sSQL = "SELECT DISTINCT " _
                & "tciAddress.AddrKey , tciAddress.AddrName, tciAddress.AddrLine1, " _
                & "tciAddress.AddrLine2, rtrim(tciAddress.City) as City, tciAddress.StateID, " _
                & " tciAddress.PostalCode, tcpUPSAcct.UPSAcct, tarCustomer.DfltShipToAddrKey " _
                & " , AddrType = CASE " _
                & " WHEN tciAddress.AddrKey = tarCustomer.DfltBillToAddrKey AND tciAddress.AddrKey = tarCustomer.DfltShipToAddrKey THEN 'B&S' " _
                & " WHEN tciAddress.AddrKey = tarCustomer.DfltBillToAddrKey THEN 'Bill' " _
                & " WHEN tciAddress.AddrKey = tarCustomer.DfltShipToAddrKey THEN 'Ship' " _
                & " Else ' ' " _
                & " End " _
                & " from tciAddress inner join " _
                & "tarCustAddr on tciAddress.Addrkey = tarCustAddr.AddrKey inner join " _
                & "tarCustomer on tarCustomer.CustKey = tarCustAddr.CustKey " _
                & "left outer join tcpUPSAcct on tcpUPSAcct.CustAddrKey = tciAddress.AddrKey " _
                & "where tarCustAddr.CustKey = " & lCustKey
                
    
    Set rstUPSAcct = LoadDiscRst(sSQL)
    
    With rstUPSAcct
        While Not .EOF
            Set oUPSAcct = New UPSAcct
            oUPSAcct.AddrKey = .Fields("AddrKey").value
            oUPSAcct.Selected = False
            On Error Resume Next
            oUPSAcct.AddrLine1 = ""
            oUPSAcct.AddrLine2 = ""
            oUPSAcct.AddrName = ""
            oUPSAcct.City = ""
            oUPSAcct.StateID = ""
            oUPSAcct.PostalCode = ""
            oUPSAcct.UPSAcct = ""
            
            oUPSAcct.AddrLine1 = .Fields("AddrLine1").value
            oUPSAcct.AddrLine2 = .Fields("AddrLine2").value
            oUPSAcct.AddrName = .Fields("AddrName").value
            oUPSAcct.City = .Fields("City").value
            oUPSAcct.StateID = .Fields("StateID").value
            oUPSAcct.PostalCode = .Fields("PostalCode").value

            oUPSAcct.UPSAcct = .Fields("UPSAcct").value
            oUPSAcct.AddrType = .Fields("AddrType").value
            If .Fields("AddrKey").value = .Fields("DfltShipToAddrKey").value Then
                If m_col.Count = 0 Then
                     m_col.Add oUPSAcct
                Else
                    m_col.Add oUPSAcct, , 1
                End If
            Else
                m_col.Add oUPSAcct
            End If
            Set oUPSAcct = Nothing
            .MoveNext
        Wend
    End With
    SetWaitCursor False
    Exit Sub
ErrorHandler:
    msg Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Err.Source
    ClearWaitCursor
End Sub



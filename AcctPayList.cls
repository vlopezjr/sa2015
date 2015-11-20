VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AcctPayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum AcctPayType
    VendorID = 1
    PartNbrType = 2
    POType = 3
    BatchType = 4
End Enum

Private m_col As Collection


Private Sub Class_Initialize()
    Set m_col = New Collection
End Sub


Private Sub Class_Terminate()
    Set m_col = Nothing
End Sub


Public Sub Add(ByRef i_oAcctPay As AcctPay)
    m_col.Add i_oAcctPay
End Sub


Public Sub Remove(vIndex As Variant)
    On Error Resume Next
    m_col.Remove vIndex
End Sub


Public Property Get Count() As Long
    Count = m_col.Count
End Property


Public Function Item(vIndex As Variant) As AcctPay
Attribute Item.VB_UserMemId = 0
    Set Item = m_col(vIndex)
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_col.[_NewEnum]
End Function


Private Function VendorExists(ByVal ItemKey As Long, ByVal VendKey As Long) As Boolean
    Dim oRst As ADODB.Recordset
    
    Set oRst = CallSP("spCPCExistVendItem", "@_iItemKey", ItemKey, "@_iVendKey", VendKey)
    
    If Not oRst.EOF Then VendorExists = True
    
    Set oRst = Nothing
End Function


'Loads inventory MaxMin data from spreadsheet files (12/18/2002 TeddyX)

Public Function LoadSpreadSheet(ByVal sFilePath As String) As Boolean
    Dim oConn As ADODB.Connection
    Dim oRst As ADODB.Recordset
    Dim oRC As MemoMeister.RemarkContext
    Dim rstAcctPay As ADODB.Recordset
    Dim oAcctPay As AcctPay

    SetWaitCursor True
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & sFilePath & ";" & _
               "Extended Properties=""Excel 8.0;HDR=yes;"""
    Debug.Print oConn.ConnectionString
    'Set orst = LoadDiscRst("Select * from [sea max min$] where [New Min] is not null or [New Max] is not null", oConn)
    Set oRst = LoadDiscRst("Select * from [qryGetAllData$] where [MaxStockQty] is not null or [MinStockQty] is not null", oConn)
    If oRst.EOF Then
        LoadSpreadSheet = False
    Else
        While Not oRst.EOF
             Set rstAcctPay = CallSP("spCPCUpdateRobertMaxMin", "@WhseKey", 24, "@UserInput", oRst.Fields("ItemID"))
             If Not rstAcctPay.EOF Then
                With rstAcctPay
                    'Add the record to the AcctPay list only when it's qualified for updating.
                    
                    Set oAcctPay = New AcctPay
                    oAcctPay.MaxQty = .Fields("MaxQty").Value
                    oAcctPay.MinQty = .Fields("MinQty").Value
                    If Not IsNull(oRst.Fields("MaxStockQty")) And Trim(oRst.Fields("MaxStockQty")) <> "" Then
                        oAcctPay.NewMax = CLng(oRst.Fields("MaxStockQty").Value)
                    Else
                        oAcctPay.NewMax = oAcctPay.MaxQty
                    End If
                    
                    If Not IsNull(oRst.Fields("MinStockQty")) And Trim(oRst.Fields("MinStockQty")) <> "" Then
                        oAcctPay.NewMin = CLng(oRst.Fields("MinStockQty").Value)
                    Else
                        oAcctPay.NewMin = oAcctPay.MinQty
                    End If
                    
                    oAcctPay.bUpdate = (oAcctPay.NewMax <> oAcctPay.MaxQty Or oAcctPay.NewMin <> oAcctPay.MinQty)
                    
                    'Add the record to the list only when it's qualified for updating
                    If oAcctPay.bUpdate Then
                        oAcctPay.VendPartNbr = Trim(.Fields("PartNbr").Value)
                        oAcctPay.Descr = Trim(.Fields("Descr").Value)
                        oAcctPay.ItemKey = .Fields("ItemKey").Value
                        oAcctPay.VendKey = .Fields("VendKey").Value
                        oAcctPay.VendID = Trim(.Fields("VendID").Value)
                        oAcctPay.Vendor = Trim(.Fields("VendName").Value)
                          
                        m_col.Add oAcctPay
                    End If
                    
                    Set oAcctPay = Nothing
                End With
             End If
            oRst.MoveNext
        Wend
        LoadSpreadSheet = True
        Set rstAcctPay = Nothing
    End If
    
    SetWaitCursor False
    Set oRst = Nothing
    oConn.Close
    Set oConn = Nothing
End Function


Public Function LoadMMVendItem(sInput As String, lWhseKey As Long) As Boolean
    Dim rstAcctPay As ADODB.Recordset
    Dim oAcctPay As AcctPay
    
    SetWaitCursor True
    Set rstAcctPay = CallSP("spCPCUpdateMaxMin", "@WhseKey", lWhseKey, "@UserInput", sInput)
    If rstAcctPay.EOF Then
        SetWaitCursor False
        Exit Function
    End If
    
    With rstAcctPay
        While Not rstAcctPay.EOF
            Set oAcctPay = New AcctPay
            oAcctPay.VendPartNbr = Trim(.Fields("PartNbr").Value)
            oAcctPay.Descr = Trim(.Fields("Descr").Value)
            oAcctPay.ItemKey = .Fields("ItemKey").Value
            oAcctPay.VendKey = .Fields("VendKey").Value
            oAcctPay.VendID = Trim(.Fields("VendID").Value)
            oAcctPay.Vendor = Trim(.Fields("VendName").Value)
            oAcctPay.MaxQty = .Fields("MaxQty").Value
            oAcctPay.MinQty = .Fields("MinQty").Value
            oAcctPay.NewMax = oAcctPay.MaxQty
            oAcctPay.NewMin = oAcctPay.MinQty
            oAcctPay.bUpdate = False
            
            m_col.Add oAcctPay
            .MoveNext
            Set oAcctPay = Nothing
        Wend
    End With
    Set rstAcctPay = Nothing
    LoadMMVendItem = True
    SetWaitCursor False
End Function


'Used by PurchAssistant to load collection

Public Function LoadVendItem(sInput As String, lVendKey As Long, lUpdateVendKey As Long) As Boolean
    Dim rstAcctPay As ADODB.Recordset
    Dim oAcctPay As AcctPay
    Dim sSQL As String
    Dim lItemKey As Long
    Dim bTemp As Boolean

    SetWaitCursor True
    sSQL = PartNbrSQL(sInput, lVendKey)
    Set rstAcctPay = LoadDiscRst(sSQL)
    If rstAcctPay Is Nothing Then Exit Function
    With rstAcctPay
        While Not .EOF
            Set oAcctPay = New AcctPay
            On Error Resume Next
            oAcctPay.NewReplCost = .Fields("NewReplCost")
            oAcctPay.VendPartNbr = ""
            oAcctPay.VendPartNbr = .Fields("VendPartNbr")
            oAcctPay.VendID = .Fields("VendID")
            oAcctPay.Vendor = .Fields("VendName")
            oAcctPay.CPCPartNbr = .Fields("CPCPartNbr")
            oAcctPay.Descr = .Fields("Descr")
            oAcctPay.VendKey = .Fields("VendKey")
            oAcctPay.ItemKey = .Fields("ItemKey")
            
            oAcctPay.BreakType = .Fields("BreakType")
            oAcctPay.DiscMeth = .Fields("DiscMeth")
            oAcctPay.ItemAliasKey = .Fields("ItemAliasKey")
            oAcctPay.ListPrice = .Fields("ListPrice")
            oAcctPay.OrigCountry = .Fields("OrigCountry")
            oAcctPay.PurchUnitMeasKey = .Fields("PurchUnitMeasKey")
            oAcctPay.SubjToVendFamDisc = .Fields("SubjToVendFamDisc")
            oAcctPay.VendFamilyKey = .Fields("VendFamilyKey")
            'oAcctPay.VendFamilyKey = Null
            
            If lUpdateVendKey = 0 Then
                oAcctPay.bUpdate = (lItemKey <> .Fields("ItemKey"))
                If oAcctPay.bUpdate And Not bTemp Then
                    bTemp = True
                End If
            Else
                If Not VendorExists(.Fields("ItemKey"), lUpdateVendKey) Then
                    oAcctPay.bUpdate = (lItemKey <> .Fields("ItemKey"))
                    If oAcctPay.bUpdate And Not bTemp Then
                        bTemp = True
                    End If
                Else
                    oAcctPay.bUpdate = False
                End If
            End If
            lItemKey = .Fields("ItemKey")
            m_col.Add oAcctPay
            .MoveNext
            Set oAcctPay = Nothing
        Wend
    End With
    SetWaitCursor False
    LoadVendItem = bTemp
End Function


'Used by AcctPayable to load collection

Public Sub LoadData(ByVal APType As AcctPayType, sInput As String)
    Dim rstAcctPay As ADODB.Recordset
    Dim oAcctPay As AcctPay
    Dim sSQL As String

    SetWaitCursor True
    Dim creditlimit As CreditLimitRecord

    Select Case APType
        Case VendorID:
            sSQL = VendorIDSQL(sInput)
        Case PartNbrType:
            sSQL = PartNbrSQL(sInput)
        Case POType:
            sSQL = POSQL(sInput)
        Case BatchType:
            sSQL = BatchSQL(sInput)
        Case Else
            SetWaitCursor False
            Exit Sub
    End Select

    Set rstAcctPay = LoadDiscRst(sSQL)
    
    If rstAcctPay Is Nothing Then Exit Sub
    With rstAcctPay
        While Not .EOF
            Set oAcctPay = New AcctPay
            On Error Resume Next
            oAcctPay.bUpdate = False
            oAcctPay.NewReplCost = .Fields("NewReplCost")
            oAcctPay.VendPartNbr = ""
            oAcctPay.VendPartNbr = .Fields("VendPartNbr")
            oAcctPay.VendID = .Fields("VendID")
            oAcctPay.Vendor = .Fields("VendName")
            oAcctPay.CPCPartNbr = .Fields("CPCPartNbr")
            oAcctPay.Descr = .Fields("Descr")
            oAcctPay.VendKey = .Fields("VendKey")
            oAcctPay.ItemKey = .Fields("ItemKey")
            If APType = POType Or APType = BatchType Or APType = VendorID Then
                oAcctPay.CurReplCost = .Fields("CurReplCost")
                oAcctPay.CreateDate = .Fields("CreateDate")
                oAcctPay.bColor = (oAcctPay.CurReplCost <> oAcctPay.NewReplCost)
                'oAcctPay.bUpdate = (oAcctPay.CurReplCost <> oAcctPay.NewReplCost)
            End If
            
            If APType = BatchType Then
                oAcctPay.BatchID = .Fields("BatchID")
                oAcctPay.OrigUserID = .Fields("OrigUserID")
            End If
            
            m_col.Add oAcctPay
            .MoveNext
            Set oAcctPay = Nothing
        Wend
    End With
    SetWaitCursor False
End Sub


Private Function VendorIDSQL(sVendID As String) As String
     'Dim sTemp As String
    'sTemp = sVendID
    'sTemp = Replace(sTemp, "*", "_")
    
'    VendorIDSQL = "SELECT DISTINCT rtrim(timItem.ItemID) as CPCPartNbr, rtrim(timItemDescription.ShortDesc) as Descr, timVendItem.RplcmntUnitCost AS CurReplCost, " & _
'                "rtrim(timVendItem.VendItemID) AS VendPartNbr, rtrim(tapVendor.VendName) as VendName, timVendItem.VendKey, " & _
'                "timVendItem.ItemKey, tpoPurchOrder.TranID, tpoPurchOrder.CreateDate, tpoPOLine.UnitCost as NewReplCost, rtrim(tapVendor.VendID) as VendID " & _
'                "FROM timItem INNER JOIN timVendItem ON timItem.ItemKey = timVendItem.ItemKey INNER JOIN " & _
'                "tapVendor ON timVendItem.VendKey = tapVendor.VendKey INNER JOIN timItemDescription ON " & _
'                "timItem.ItemKey = timItemDescription.ItemKey INNER JOIN tpoPOLine ON timItem.ItemKey = " & _
'                "tpoPOLine.ItemKey INNER JOIN tpoPurchOrder ON timVendItem.VendKey = tpoPurchOrder.VendKey AND " & _
'                "tapVendor.VendKey = tpoPurchOrder.VendKey And tpoPOLine.POKey = tpoPurchOrder.POKey " & _
'                "WHERE (tpoPurchOrder.CompanyID = 'CPC') AND " & _
'                "tapVendor.VendID = '" & Trim(sVendID) & "'"

'    VendorIDSQL = "SELECT DISTINCT rtrim(timItem.ItemID) as CPCPartNbr, " _
'        & "rtrim(timItemDescription.ShortDesc) as Descr, timVendItem.RplcmntUnitCost AS CurReplCost, " _
'        & "timVendItem.RplcmntUnitCost AS NewReplCost, rtrim(timVendItem.VendItemID) AS VendPartNbr, " _
'        & "rtrim(tapVendor.VendName) as VendName, rtrim(tapVendor.VendID) as VendID " _
'        & "FROM timItem INNER JOIN timVendItem ON timItem.ItemKey = timVendItem.ItemKey INNER JOIN " _
'        & "tapVendor ON timVendItem.VendKey = tapVendor.VendKey INNER JOIN timItemDescription ON " _
'        & "timItem.ItemKey = timItemDescription.ItemKey INNER JOIN tpoPOLine ON timItem.ItemKey = " _
'        & "tpoPOLine.ItemKey INNER JOIN tpoPurchOrder ON timVendItem.VendKey = tpoPurchOrder.VendKey AND " _
'        & "tapVendor.VendKey = tpoPurchOrder.VendKey And tpoPOLine.POKey = tpoPurchOrder.POKey " _
'        & "WHERE (tpoPurchOrder.CompanyID = 'CPC') AND tapVendor.VendID = '" & Trim(sVendID) & "'"

'PRN 411 -- Above select statement doesn't include VendKey or ItemKey which makes DB updates impossible
    VendorIDSQL = "SELECT DISTINCT rtrim(timItem.ItemID) as CPCPartNbr, " _
        & "rtrim(timItemDescription.ShortDesc) as Descr, timVendItem.RplcmntUnitCost AS CurReplCost, " _
        & "timVendItem.RplcmntUnitCost AS NewReplCost, rtrim(timVendItem.VendItemID) AS VendPartNbr, " _
        & "rtrim(tapVendor.VendName) as VendName, rtrim(tapVendor.VendID) as VendID, dbo.timItem.ItemKey, dbo.tapVendor.VendKey " _
        & "FROM timItem INNER JOIN timVendItem ON timItem.ItemKey = timVendItem.ItemKey INNER JOIN " _
        & "tapVendor ON timVendItem.VendKey = tapVendor.VendKey INNER JOIN timItemDescription ON " _
        & "timItem.ItemKey = timItemDescription.ItemKey INNER JOIN tpoPOLine ON timItem.ItemKey = " _
        & "tpoPOLine.ItemKey INNER JOIN tpoPurchOrder ON timVendItem.VendKey = tpoPurchOrder.VendKey AND " _
        & "tapVendor.VendKey = tpoPurchOrder.VendKey And tpoPOLine.POKey = tpoPurchOrder.POKey " _
        & "WHERE (tpoPurchOrder.CompanyID = 'CPC') AND tapVendor.VendID = '" & Trim(sVendID) & "'"
End Function


Private Function POSQL(sTranID As String) As String
    POSQL = "SET ROWCOUNT 100 SELECT rtrim(timItem.ItemID) as CPCPartNbr, rtrim(timItemDescription.ShortDesc) as Descr, timVendItem.RplcmntUnitCost AS CurReplCost, " & _
                "rtrim(timVendItem.VendItemID) AS VendPartNbr, rtrim(tapVendor.VendName) as VendName, timVendItem.VendKey, " & _
                "timVendItem.ItemKey, tpoPurchOrder.TranID, tpoPurchOrder.CreateDate, tpoPOLine.UnitCost as NewReplCost, rtrim(tapVendor.VendID) as VendID " & _
                "FROM timItem INNER JOIN timVendItem ON timItem.ItemKey = timVendItem.ItemKey INNER JOIN " & _
                "tapVendor ON timVendItem.VendKey = tapVendor.VendKey INNER JOIN timItemDescription ON " & _
                "timItem.ItemKey = timItemDescription.ItemKey INNER JOIN tpoPOLine ON timItem.ItemKey = " & _
                "tpoPOLine.ItemKey INNER JOIN tpoPurchOrder ON timVendItem.VendKey = tpoPurchOrder.VendKey AND " & _
                "tapVendor.VendKey = tpoPurchOrder.VendKey And tpoPOLine.POKey = tpoPurchOrder.POKey " & _
                "WHERE (tpoPurchOrder.CompanyID = 'CPC') AND " & _
                "(tpoPurchOrder.TranID = '" & FormatBatch(Trim(sTranID), "PO-0000000000") & "') " & _
                " SET ROWCOUNT 0 "
End Function


Private Function BatchSQL(sBatchID As String) As String
    BatchSQL = "SET ROWCOUNT 100 SELECT rtrim(timItem.ItemID) as CPCPartNbr, rtrim(timItemDescription.ShortDesc) as Descr, timVendItem.RplcmntUnitCost AS CurReplCost, " & _
            "rtrim(timVendItem.VendItemID) AS VendPartNbr, rtrim(tapVendor.VendID) as VendID, rtrim(tapVendor.VendName) as VendName, timVendItem.VendKey, " & _
            "timVendItem.ItemKey , tpoPurchOrder.TranID, tpoPurchOrder.CreateDate, tpoPOLine.UnitCost as NewReplCost, " & _
            "tciBatchLog.BatchID, tciBatchLog.OrigUserID, tciBatchLog.CreateDate " & _
            "FROM tapVoucher INNER JOIN tpoPurchOrdVouch ON " & _
            "tapVoucher.VoucherKey = tpoPurchOrdVouch.VoucherKey INNER Join timItem INNER JOIN " & _
            "timVendItem ON timItem.ItemKey = timVendItem.ItemKey INNER JOIN tapVendor ON " & _
            "timVendItem.VendKey = tapVendor.VendKey INNER JOIN timItemDescription ON " & _
            "timItem.ItemKey = timItemDescription.ItemKey INNER JOIN tpoPOLine ON " & _
            "timItem.ItemKey = tpoPOLine.ItemKey INNER JOIN tpoPurchOrder ON " & _
            "timVendItem.VendKey = tpoPurchOrder.VendKey AND tapVendor.VendKey = tpoPurchOrder.VendKey AND " & _
            "tpoPOLine.POKey = tpoPurchOrder.POKey ON tpoPurchOrdVouch.POKey = tpoPurchOrder.POKey INNER JOIN " & _
            "tciBatchLog ON tapVoucher.BatchKey = tciBatchLog.BatchKey " & _
            "WHERE (tpoPurchOrder.CompanyID = 'CPC') AND (tciBatchLog.BatchID = '" & FormatBatch(Trim(sBatchID), "POVO-0000000") & "') AND " & _
            "(tpoPurchOrder.CreateDate > '1/1/00') " & _
            " SET ROWCOUNT 0 "
End Function


' Add timVendorItem.ListPrice to Grid (5/23/03 LR)

Private Function PartNbrSQL(sItemID As String, Optional bVendKey As Long = 0) As String
    'Dim sTemp As String
    'sTemp = sItemID
    'sTemp = Replace(sTemp, "*", "_")
    Dim sTemp As String
    
    If bVendKey <> 0 Then
        sTemp = " and timVendItem.VendKey = " & bVendKey
    End If

'SELECT rtrim(timItem.ItemID) as CPCPartNbr,
'rtrim(timItemDescription.ShortDesc) as Descr,
'timVendItem.RplcmntUnitCost as NewReplCost,
'rtrim(timVendItem.VendItemID) as VendPartNbr,
'rtrim(tapVendor.VendName) as VendName,
'timVendItem.VendKey,
'timVendItem.ItemKey,
'rtrim(tapvendor.vendid) as VendID,
'timVendItem.ItemAliasKey,
'timVendItem.BreakType,
'timVendItem.DiscMeth,
'timVendItem.ListPrice,
'timVendItem.OrigCountry,
'timVendItem.PurchUnitMeasKey,
'timVendItem.SubjToVendFamDisc,
'timVendItem.VendFamilyKey
'FROM timItem INNER JOIN timVendItem ON timItem.ItemKey = timVendItem.ItemKey
'INNER JOIN tapVendor ON timVendItem.VendKey = tapVendor.VendKey
'INNER JOIN timItemDescription ON timItem.ItemKey = timItemDescription.ItemKey
'WHERE (tapVendor.CompanyID = 'CPC') AND tapVendor.DfltPurchAcctKey = 3088

    PartNbrSQL = "SELECT rtrim(timItem.ItemID) as CPCPartNbr, rtrim(timItemDescription.ShortDesc) as Descr, " & _
                "timVendItem.RplcmntUnitCost as NewReplCost, rtrim(timVendItem.VendItemID) as VendPartNbr, " & _
                "rtrim(tapVendor.VendName) as VendName, timVendItem.VendKey, timVendItem.ItemKey, rtrim(tapvendor.vendid) as VendID, " & _
                "timVendItem.ItemAliasKey, timVendItem.BreakType, timVendItem.DiscMeth, timVendItem.ListPrice, " & _
                "timVendItem.OrigCountry, timVendItem.PurchUnitMeasKey, timVendItem.SubjToVendFamDisc, timVendItem.VendFamilyKey " & _
                "FROM timItem INNER JOIN timVendItem ON " & _
                "timItem.ItemKey = timVendItem.ItemKey INNER JOIN tapVendor ON " & _
                "timVendItem.VendKey = tapVendor.VendKey INNER JOIN timItemDescription ON " & _
                "timItem.ItemKey = timItemDescription.ItemKey " & _
                "WHERE (tapVendor.CompanyID = 'CPC') AND tapVendor.DfltPurchAcctKey = 3088 "
                
    If Trim(sItemID) <> "" Then
        PartNbrSQL = PartNbrSQL & " and timitem.itemid like '" & FormatSearch(Trim(sItemID)) & "'" & sTemp
    ElseIf bVendKey <> 0 Then
        PartNbrSQL = PartNbrSQL & " and timVendItem.VendKey = " & bVendKey
    End If
    
    PartNbrSQL = PartNbrSQL & " order by timVendItem.ItemKey "
End Function


Private Function FormatSearch(sInput As String) As String
    Dim sTemp As String

    sTemp = sInput
    sTemp = Replace(sTemp, "?", "_")
    sTemp = Replace(sTemp, "*", "%")
    FormatSearch = sTemp
End Function


Private Function FormatBatch(sInput As String, sMask As String) As String
    Dim sTemp As String
    
    sTemp = sMask
    sTemp = Left$(sTemp, Len(sTemp) - Len(sInput)) & sInput
    'sTemp = Replace(sTemp, "?", "_")
   ' sTemp = Replace(sTemp, "*", "%")
    
    FormatBatch = sTemp
End Function

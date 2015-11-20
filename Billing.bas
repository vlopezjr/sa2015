Attribute VB_Name = "Billing"
Public Sub ApproveRMAItem(ByVal RmaLineKey As Long, ByVal Approved As Boolean, ByVal CreditFreight As Boolean, ByVal QtyReceived As Integer)
    Set cmd = CreateCommandSP("spcpcRMAApproveItemUpdate")
    cmd.Parameters("@_iRMALineKey").value = RmaLineKey
    cmd.Parameters("@_iApproved").value = Approved
    cmd.Parameters("@_iCreditFreight").value = CreditFreight
    cmd.Parameters("@_iQtyRcvd").value = QtyReceived
    cmd.Execute
    Set cmd = Nothing
End Sub

Public Sub ApproveRMAAdjustmentForQtyReceived(ByVal RmaLineKey As Long, ByVal QtyReceived As Integer)
    Set cmd = CreateCommandSP("spcpcRMAApproveAdjustQtyRcvd")
    cmd.Parameters("@_iRMALineKey").value = RmaLineKey
    cmd.Parameters("@_iQtyRcvd").value = QtyReceived
    cmd.Execute
    Set cmd = Nothing
End Sub

Public Function LoadRMACredit(lWhseKey As Long) As RMAList
    Set LoadRMACredit = New RMAList
    
    Dim orst As ADODB.Recordset
    Dim oRMAOrder As RMAOrder
    
    On Error GoTo ErrorHandler
    
    If lWhseKey = 0 Then
        Set orst = CallSP("spcpcRMAGetApproved")
    Else
        Set orst = CallSP("spcpcRMAGetApproved", "@_iWhseKey", lWhseKey)
    End If
    
    With orst
        While Not .EOF
            Set oRMAOrder = New RMAOrder
            
            oRMAOrder.RmaLineKey = .Fields("RMALineKey").value
            oRMAOrder.RMAInfo = .Fields("RMAInfo").value
            '06/12/03 AVH PRN#5 Provide an option to sort the grid by RMA# or by CustID
            oRMAOrder.RMAInfo_ByCustID = .Fields("RMAInfo_ByCustID").value
            oRMAOrder.RMAKey = .Fields("RMAKey").value
            oRMAOrder.OPKey = .Fields("OPKey").value
            oRMAOrder.OPLineKey = .Fields("OPLineKey").value
            oRMAOrder.SOLineKey = .Fields("SOLineKey").value
            oRMAOrder.Credited = False
            oRMAOrder.Reason = .Fields("ReasonCode").value
            oRMAOrder.Restock = .Fields("Restock").value
            oRMAOrder.ItemID = .Fields("ItemID").value
            oRMAOrder.Descr = .Fields("Descr").value
            oRMAOrder.Cost = .Fields("Cost").value
            oRMAOrder.Price = .Fields("Price").value
            oRMAOrder.QtyAuthorized = .Fields("QtyAuth").value
            oRMAOrder.QtyRcvd = .Fields("QtyRcvd").value
            oRMAOrder.QtyCred = 0
            oRMAOrder.QtyPreCred = .Fields("QtyCred").value '.Fields("QtyPreCred").Value
            oRMAOrder.lMaxQtyCred = oRMAOrder.QtyCred
            oRMAOrder.CreditFreight = .Fields("CreditFreight").value
            oRMAOrder.RcvdWhseID = .Fields("RcvdWhseID").value
            oRMAOrder.ApproveDate = .Fields("ApproveDate").value
            oRMAOrder.CMNbr = Format(.Fields("CM#").value)
            
            LoadRMACredit.Add oRMAOrder
            Set oRMAOrder = Nothing
            .MoveNext
        Wend
    End With
    
    Set orst = Nothing
    
    Exit Function
    
ErrorHandler:
    msg Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Err.Source
    ClearWaitCursor
End Function

Public Function IsItemInInventory(ByVal iItemId As String, ByVal iWhseKey As Integer) As Boolean
    Dim ocmd As ADODB.Command
    Dim retval As Integer
    
    Set ocmd = CreateCommandSP("spcpcIsItemIdInInventory")
    With ocmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)

        .Parameters.Append .CreateParameter("@ItemId", adVarChar, adParamInput, 30, iItemId)
        .Parameters.Append .CreateParameter("@WhseKey", adSmallInt, adParamInput, , iWhseKey)
        
        .Execute
       
        IsItemInInventory = IIf(.Parameters("RETURN_VALUE").value = 1, True, False)
    End With
    
End Function

Public Sub UpdateRMAItemCredit(ByVal RmaLineKey As Long, ByVal QtyCred As Integer, ByVal CMNbr As String)
    Dim oCurrentTime As Date
    
    Set cmd = CreateCommandSP("spcpcRMACreditItemUpdate")
    cmd.Parameters("@_iRMALineKey").value = RmaLineKey
    cmd.Parameters("@_iQtyCred").value = QtyCred
    cmd.Parameters("@_iCredDate").value = oCurrentTime
    cmd.Parameters("@_iUserID").value = GetUserName
    cmd.Parameters("@_iCM#").value = CMNbr
    cmd.Execute
    Set cmd = Nothing
End Sub

Public Function GetRMAReasons() As Collection
    Dim orst As ADODB.Recordset
    Set GetRMAReasons = New Collection

    Set orst = LoadDiscRst("SELECT * FROM tcpRMAReason ORDER BY RMAReasonID")

    Do While Not orst.EOF
        GetRMAReasons.Add orst.Fields("RMAReasonID").value, CStr(orst.Fields("RMAReasonKey").value)
        orst.MoveNext
    Loop
End Function

' setup the DispositionID collection for the RMA grid
Public Function GetRMADispositions() As Collection
    Dim orst As ADODB.Recordset
    Set GetRMADispositions = New Collection

    Set orst = LoadDiscRst("SELECT * FROM tcpRMADisposition")
    
    GetRMADispositions.Add "-none-", CStr(0)
    Do While Not orst.EOF
        GetRMADispositions.Add orst.Fields("RMADispID").value, CStr(orst.Fields("RMADispKey").value)
        orst.MoveNext
    Loop
End Function


Public Function GetNexusStates() As ADODB.Recordset
    Dim sSQL As String
    Dim rstStates As ADODB.Recordset
    
    sSQL = "SELECT tciAddress.StateID AS State FROM timWarehouse INNER JOIN tciAddress " & _
            "ON timWarehouse.ShipAddrKey = AddrKey WHERE CompanyID='CPC' AND Transit = 0 " & _
            "ORDER BY tciAddress.StateID"

    Set GetNexusStates = LoadDiscRst(sSQL)
End Function


Public Function GetTerritory(sWhseID As String) As String
    Dim rst As ADODB.Recordset

    With g_rstWhses
        .Filter = "WhseID = '" & sWhseID & "'"
        If Not .EOF Then
            'TODO: use an output parameter rather than a recordset  8/26/03 LR
            Set rst = LoadDiscRst("Select SalesTerritoryID from tarSalesTerritory WHERE CompanyID = 'CPC' and SalesTerritoryKey = " & .Fields("SalesTerritoryKey"))
            GetTerritory = rst.Fields("SalesTerritoryID").value
            Set rst = Nothing
        End If
        .Filter = adFilterNone
    End With
End Function

Public Function GetCustKey(i_sCustID As String) As Long
    Dim cmd As ADODB.Command
    
    Set cmd = CreateCommandSP("spcpcCustIDtoKey")
    
    With cmd
        .Parameters("@_iCustID") = Left(i_sCustID, 12)
        .Execute
        GetCustKey = .Parameters(0).value
    End With
End Function

Public Function GetSoIdByOpId(ByVal OPKey) As Long
    Set cmd = CreateCommandSP("spcpcGetSOIDbyOPID")
    cmd.Parameters("@_iOPKey").value = OPKey
    cmd.Execute
    GetSoIdByOpId = IIf(IsNull(cmd.Parameters("@_oRetVal").value), 0, cmd.Parameters("@_oRetVal").value)
End Function

Public Function LoadReceivedRMAsToApprove(lWhseKey As Long) As RMAList
    Set LoadReceivedRMAsToApprove = New RMAList
    
    Dim orst As ADODB.Recordset
    Dim oRMAOrder As RMAOrder
    
    On Error GoTo ErrorHandler
    
    If lWhseKey = 0 Then
        Set orst = CallSP("spcpcRMAGetReceived")
    Else
        Set orst = CallSP("spcpcRMAGetReceived", "@_iWhseKey", lWhseKey)
    End If
    
    With orst
        While Not .EOF
            Set oRMAOrder = New RMAOrder
            
            oRMAOrder.RmaLineKey = .Fields("RMALineKey").value
            oRMAOrder.RMAInfo = .Fields("RMAInfo").value
            oRMAOrder.RMAKey = .Fields("RMAKey").value
            oRMAOrder.OPKey = .Fields("OPKey").value
            oRMAOrder.OPLineKey = .Fields("OPLineKey").value
            oRMAOrder.SOLineKey = .Fields("SOLineKey").value
            oRMAOrder.Approved = .Fields("Approved").value
            oRMAOrder.Disposition = .Fields("DispositionCode").value
            oRMAOrder.Reason = .Fields("ReasonCode").value
            oRMAOrder.Restock = .Fields("Restock").value
            oRMAOrder.ItemID = .Fields("ItemID").value
            oRMAOrder.Descr = .Fields("Descr").value
            oRMAOrder.Cost = .Fields("Cost").value
            oRMAOrder.Price = .Fields("Price").value
            oRMAOrder.QtyAuthorized = .Fields("QtyAuth").value
            oRMAOrder.QtyRcvd = .Fields("QtyRcvd").value
            '06/12/03 AVH PRN#103
            oRMAOrder.QtyRcvdOriginal = .Fields("QtyRcvd").value
            oRMAOrder.QtyPreCred = .Fields("QtyCred").value '.Fields("QtyPreCred").Value
            oRMAOrder.lMaxQtyCred = oRMAOrder.QtyCred
            oRMAOrder.CreditFreight = .Fields("CreditFreight").value
            oRMAOrder.RcvdWhseID = .Fields("RcvdWhseID").value
            oRMAOrder.ReceiveDate = .Fields("RcvdDate").value
            '08/08/14 VL
            oRMAOrder.OPItemType = .Fields("OPItemType").value
            
            LoadReceivedRMAsToApprove.Add oRMAOrder
            Set oRMAOrder = Nothing
            .MoveNext
        Wend
    End With
    
    Set orst = Nothing
    Exit Function
    
ErrorHandler:
    msg Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Err.Source
    ClearWaitCursor
End Function

Public Function LoadNewRMA(ByVal lOPKey As Long, Optional bAddItem As Boolean = False) As RMAList
    Set LoadNewRMA = New RMAList
    
    Dim oRstHdr As ADODB.Recordset
    Dim oRstLines As ADODB.Recordset
    Dim oRMAOrder As RMAOrder
    
    On Error GoTo ErrorHandler
    
    If bAddItem Then
        Set oRstHdr = LoadDiscRst("Select * from tcpRMA where OPKey = " & lOPKey)
        Set oRstLines = CallSP("spcpcRMAGetRemainingSOLines", "@RMAKey", oRstHdr.Fields("RMAKey").value, "@OPKey", lOPKey)
        Set oRstHdr = Nothing
    Else
        Set oRstLines = CallSP("spcpcRMAGetAllSOLines", "@OPKey", lOPKey)
    End If

    With oRstLines
        While Not .EOF
            Set oRMAOrder = New RMAOrder
            
            oRMAOrder.OPKey = .Fields("OPKey").value
            oRMAOrder.OPLineKey = .Fields("OPLineKey").value
            oRMAOrder.SOLineKey = .Fields("SOLineKey").value
            oRMAOrder.Authorized = False
            oRMAOrder.Approved = False
            oRMAOrder.Credited = False
            oRMAOrder.Reason = 0
            oRMAOrder.Restock = 0
            oRMAOrder.Disposition = 0
            oRMAOrder.CreditFreight = False
            oRMAOrder.ItemID = .Fields("ItemID").value
            oRMAOrder.OPItemType = .Fields("OPItemType").value
            oRMAOrder.Cost = .Fields("Cost").value
            oRMAOrder.Price = .Fields("Price").value
            oRMAOrder.QtyAuthorized = .Fields("QtyAuth").value 'this is really qty remaining to authorize
            oRMAOrder.lMaxQtyAuth = .Fields("QtyAuth").value
            oRMAOrder.ExtPrice = oRMAOrder.Price * oRMAOrder.QtyAuthorized
            
            LoadNewRMA.Add oRMAOrder
            Set oRMAOrder = Nothing
            .MoveNext
        Wend
    End With
    
    Set oRstLines = Nothing
    
    Exit Function
    
ErrorHandler:
    msg Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Err.Source
    ClearWaitCursor
    
    
End Function


Public Function LoadRMALine(ByVal lRMAKey As Long) As RMAList
    Set LoadRMALine = New RMAList
    Dim oRstLines As ADODB.Recordset
    Dim oRMAOrder As RMAOrder
    
    On Error GoTo ErrorHandler
    
    Set oRstLines = CallSP("spcpcRMALines", "@RMAKey", lRMAKey)

    With oRstLines
        
            While Not .EOF
                Set oRMAOrder = New RMAOrder
                oRMAOrder.RmaLineKey = .Fields("RMALineKey").value
                oRMAOrder.OPLineKey = .Fields("OPLineKey").value
                oRMAOrder.SOLineKey = .Fields("SOLineKey").value
                oRMAOrder.ItemID = .Fields("ItemID").value
                oRMAOrder.QtyAuthorized = .Fields("QtyAuth").value
                oRMAOrder.AuthBy = .Fields("AuthBy").value
                oRMAOrder.AuthDate = .Fields("AuthDate").value
                oRMAOrder.Cost = .Fields("Cost").value
                oRMAOrder.Price = .Fields("Price").value
                oRMAOrder.Reason = .Fields("ReasonCode").value
                oRMAOrder.Disposition = .Fields("DispositionCode").value
                oRMAOrder.Restock = .Fields("Restock").value
                oRMAOrder.QtyPreRcvd = .Fields("QtyPreRcvd").value
                oRMAOrder.QtyPreCred = .Fields("QtyPreCred").value
                oRMAOrder.CreditFreight = .Fields("CreditFreight").value
                oRMAOrder.lMaxQtyAuth = .Fields("MaxAuthQty").value
                oRMAOrder.ReturnToVendor = .Fields("ReturnToVendor").value
                oRMAOrder.DaysNoPenalty = .Fields("DaysNoPenalty").value
                oRMAOrder.VendorRMANumber = .Fields("VendorRMANumber").value
                oRMAOrder.VendKey = .Fields("VendKey").value
                LoadRMALine.Add oRMAOrder
                Set oRMAOrder = Nothing
                .MoveNext
            Wend
      
    End With
    
    Set oRstLines = Nothing
    Exit Function
    
ErrorHandler:
    msg Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Err.Source
    ClearWaitCursor
End Function

'Called by Warehouse Assistant to load all open RMALines that have not been fully received.
Public Function LoadOpenRMAByRMANumber(ByVal lRMANumber As Long, Optional bRMAKey As Boolean = False) As RMAList
    Set LoadOpenRMAByRMANumber = New RMAList
    
    Dim orst As ADODB.Recordset
    Dim oRMAOrder As RMAOrder
    Dim lOldRMAKey As Long
    Dim lNewRMAKey As Long
    Dim lCount As Long
    
    On Error GoTo ErrorHandler
    
    If Not bRMAKey Then
        Set orst = CallSP("spcpcRMAGetByRMANumber", "@_iRMANumber", lRMANumber)
    Else
        Set orst = CallSP("spCPCRMAGetByRMAKey", "@_iRMAKey", lRMANumber)
    End If
    
    With orst
        While Not .EOF
            If .Fields("QtyAuth") > .Fields("QtyPreRcvd") Then
                lNewRMAKey = .Fields("RMAKey").value
                
                Set oRMAOrder = New RMAOrder
                oRMAOrder.RmaLineKey = .Fields("RMALineKey").value
                oRMAOrder.RMAInfo = .Fields("RMAInfo").value
                oRMAOrder.RMAKey = .Fields("RMAKey").value
                oRMAOrder.OPKey = .Fields("OPKey").value
                oRMAOrder.OPLineKey = .Fields("OPLineKey").value
                oRMAOrder.SOLineKey = .Fields("SOLineKey").value
                oRMAOrder.Reason = .Fields("ReasonCode")
                oRMAOrder.Disposition = .Fields("Disposition")
                oRMAOrder.Restock = .Fields("Restock")
                oRMAOrder.AuthBy = .Fields("AuthBy").value
                oRMAOrder.AuthDate = .Fields("AuthDate").value
                oRMAOrder.ItemID = .Fields("ItemID").value
                oRMAOrder.Descr = .Fields("Descr").value
                oRMAOrder.Cost = .Fields("Cost").value
                oRMAOrder.Price = .Fields("Price").value
                oRMAOrder.QtyAuthorized = .Fields("QtyAuth").value
                oRMAOrder.QtyRcvd = 0
                oRMAOrder.QtyPreRcvd = .Fields("QtyPreRcvd").value
                oRMAOrder.QtyPreCred = .Fields("QtyPreCred").value
                oRMAOrder.lMaxQtyRcvd = oRMAOrder.QtyRcvd
                oRMAOrder.CustID = .Fields("CustID").value
                oRMAOrder.SOID = .Fields("SOID").value
                oRMAOrder.CustName = .Fields("CustName").value
                oRMAOrder.CreditFreight = .Fields("CreditFreight").value
                
                LoadOpenRMAByRMANumber.Add oRMAOrder
                Set oRMAOrder = Nothing
            End If
            
            'NOTE: I'm not sure if distinguishing between records and unique RMALines is necessary
            'The calling routines are only looking for 0 or >0.
            If lOldRMAKey <> lNewRMAKey Then
                lCount = lCount + 1
                lOldRMAKey = lNewRMAKey
            End If
            
            .MoveNext
        Wend
    End With
    
    Set orst = Nothing
   
    Exit Function
    
ErrorHandler:
    msg Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Err.Source
    ClearWaitCursor
End Function

'Called by Warehouse Assistant to load all open RMALines that have not been fully received.
Public Function LoadOpenRMA(Optional lCustKey As Long, Optional sCustID As String, Optional sItemID As String, Optional sItemDescr As String) As RMAList
    Set LoadOpenRMA = New RMAList
    
    Dim orst As ADODB.Recordset
    Dim oRMAOrder As RMAOrder
    Dim lOldRMAKey As Long
    Dim lNewRMAKey As Long
    Dim lCount As Long
    
    On Error GoTo ErrorHandler
    
    If lCustKey > 0 Then
        Set orst = CallSP("spcpcRMAGetByCustKey", "@_iCustKey", lCustKey)
    ElseIf sCustID <> "" And sItemID <> "" Then
        Set orst = CallSP("spCPCRMAGetByCustandPart", "@_iCustInput", sCustID, "@_iPartInput", sItemID)
    Else
        If sItemID <> "" And sItemDescr <> "" Then
            Set orst = CallSP("spCPCRMAGetByItem", "@_iItemID", sItemID, "@_iItemDescr", sItemDescr)
        ElseIf sItemID <> "" Then
            Set orst = CallSP("spCPCRMAGetByItem", "@_iItemID", sItemID)
        ElseIf sItemDescr <> "" Then
            Set orst = CallSP("spCPCRMAGetByItem", "@_iItemDescr", sItemDescr)
        End If
    End If
    
    With orst
        While Not .EOF
            If .Fields("QtyAuth") > .Fields("QtyPreRcvd") Then
                lNewRMAKey = .Fields("RMAKey").value
                
                Set oRMAOrder = New RMAOrder
                oRMAOrder.RmaLineKey = .Fields("RMALineKey").value
                oRMAOrder.RMAInfo = .Fields("RMAInfo").value
                oRMAOrder.RMAKey = .Fields("RMAKey").value
                oRMAOrder.OPKey = .Fields("OPKey").value
                oRMAOrder.OPLineKey = .Fields("OPLineKey").value
                oRMAOrder.SOLineKey = .Fields("SOLineKey").value
                oRMAOrder.Reason = .Fields("ReasonCode")
                oRMAOrder.Disposition = .Fields("Disposition")
                oRMAOrder.Restock = .Fields("Restock")
                oRMAOrder.AuthBy = .Fields("AuthBy").value
                oRMAOrder.AuthDate = .Fields("AuthDate").value
                oRMAOrder.ItemID = .Fields("ItemID").value
                oRMAOrder.Descr = .Fields("Descr").value
                oRMAOrder.Cost = .Fields("Cost").value
                oRMAOrder.Price = .Fields("Price").value
                oRMAOrder.QtyAuthorized = .Fields("QtyAuth").value
                oRMAOrder.QtyRcvd = 0
                oRMAOrder.QtyPreRcvd = .Fields("QtyPreRcvd").value
                oRMAOrder.QtyPreCred = .Fields("QtyPreCred").value
                oRMAOrder.lMaxQtyRcvd = oRMAOrder.QtyRcvd
                oRMAOrder.CustID = .Fields("CustID").value
                oRMAOrder.SOID = .Fields("SOID").value
                oRMAOrder.CustName = .Fields("CustName").value
                oRMAOrder.CreditFreight = .Fields("CreditFreight").value
                
                LoadOpenRMA.Add oRMAOrder
                Set oRMAOrder = Nothing
            End If
            
            'NOTE: I'm not sure if distinguishing between records and unique RMALines is necessary
            'The calling routines are only looking for 0 or >0.
            If lOldRMAKey <> lNewRMAKey Then
                lCount = lCount + 1
                lOldRMAKey = lNewRMAKey
            End If
            
            .MoveNext
        Wend
    End With
    
    Set orst = Nothing
  
    Exit Function
    
ErrorHandler:
    msg Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Err.Source
    ClearWaitCursor
End Function

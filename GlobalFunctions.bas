Attribute VB_Name = "GlobalFunctions"
Option Explicit
Option Compare Text

'Public Sub CloneRemarks(i_src As RemarkContext, o_dst As RemarkContext)
'Public Function GetRemarkTypeIndex(ByRef SRContext As RemarkContext, sTypeID As String) As Long
'Public Function GetLineStatus(oValue As LineStatus) As Strin
'Public Sub DisplayInfo(ByVal sCaption As String, ByVal sBody As String, ByVal iWidth As Integer, ByVal iHeight As Integer, Optional ByVal bFixedWidth = False)
'Public Function GetQuoteXmlName() As String
'Public Sub SaveAsTextFile(filename As String, buffer As String)
'Public Function msg( _
'        ByRef i_sPrompt As String, _
'        Optional ByVal i_eMsgBoxStyle As VbMsgBoxStyle = vbOKOnly, _
'        Optional ByRef i_sTitle As String = "Office Assistant" _
') As VbMsgBoxResult
'Public Function Version(Optional bIncludeName As Boolean = False) As String
'Public Function ReadFile(i_sPath As String) As String
'Public Sub PrepareToCommitItem(o_oBaseItem As Item, i_sPrefix As String, i_lWhseKey As Long)
'Public Function ShipCharge(ByRef oOrder As Order) As Double
'Public Sub UpdateExemptCerts(CustKey As Long, ExmptNo As String, StateID As String)
'Public Function ConvertSageItemType(ByVal i_lSageItemType As Long) As ItemTypeCode
'Public Function StatusCodeString(i_eStatusCode As ItemStatusCode) As String
'Public Function ResearchStatusString(i_eResearchStatus As ItemResearchStatus) As String
'Public Function GetWhseID(StateID As String, CountryID As String) As String
'Public Function WhseKeyToID(ByVal i_lWhseKey As Long) As String
'Public Sub SetUpWarehouses(cboWarehouse As ComboBox, _
'                            ByRef rstWhses As ADODB.Recordset, _
'                            ByVal WhseKey As Long)
'Public Function VendKeyToID(ByVal i_lVendKey As Long) As String
'Public Function WhoSetFreeFreight(ByVal OPKey As Long) As String
'Public Sub SendNotification(ByVal sSubject As String, ByVal sMessage As String, ByVal vsaRecipientsArray As Variant)

'Public Function GetVendorAddrKey(i_lVendKey As Long) As Long
'Public Function GetGLAcctKey(i_lUserKey As Long) As Long
'Public Function CreateUPSProxy() As MSSOAPLib30.SoapClient30


Private Declare Sub ProcessIdToSessionId Lib "Kernel32.dll" (ByVal lngPID As Long, ByRef lngSID As Long)
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function RedrawWindow Lib "user32" _
    (ByVal hwnd As Long, _
    lprcUpdate As RECT, _
    ByVal hrgnUpdate As Long, _
    ByVal fuRedraw As Long) As Long

Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_INVALIDATE = &H1
Private Const RDW_UPDATENOW = &H100

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Function GetProcessInfo() As String
    Dim SID As Long
    Dim PID As Long
    PID = GetCurrentProcessId()
    ProcessIdToSessionId PID, SID
    GetProcessInfo = "SessionId " & CStr(SID) & ", ProcessId " & CStr(PID)
End Function


Public Sub KillExistingInstance(processname As String)
    Dim WMI As Object
    Dim process As Object
    Dim sMessage As String
    Dim PID As Long
    Dim SID As Long
    
    PID = GetCurrentProcessId()
    ProcessIdToSessionId PID, SID
    
    Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process")
    
    For Each process In WMI
        If process.Name = processname And process.ProcessId <> PID And process.SessionId = SID Then
            sMessage = sMessage & "Killing existing instance: SessionId " & SID & ", ProcessId " & process.ProcessId & vbCrLf
            process.Terminate
        End If
    Next
    
    If Len(sMessage) > 0 Then
        LogEvent "EntryPoint", "Main", sMessage
    End If
End Sub


Public Sub ForceRefresh(ByVal hwnd As Long)
    Dim lRes As Long
    Dim oRect As RECT

    lRes = GetWindowRect(hwnd, oRect)
    If lRes = 0 Then
        Err.Raise -1, "modRefresh:ForceRefresh", "GetWindowRect Failed"
    End If
    
    lRes = RedrawWindow(hwnd, oRect, 0, RDW_INVALIDATE + RDW_UPDATENOW + RDW_ALLCHILDREN)
    If lRes = 0 Then
        Err.Raise -1, "modRefresh:ForceRefresh", "RedrawWindow Failed"
    End If
End Sub


'Called By:
'
'   MDIMain
'        DoSplitOrder
'   FOrder
'        MorphBTOItems
'        MorphSPOToBTO
'        UndoMorphBTOItmes
'        MorphBTOtoSPO
'        txtItemPartNbr.LostFocus
'        gdxItems.OLEDragDrop

Public Sub CloneRemarks(i_src As RemarkContext, o_dst As RemarkContext)
    Dim oRemark As remark
    Dim lTypeIndex As Long
    
    If i_src.RemarkList Is Nothing Then Exit Sub
    For Each oRemark In i_src.RemarkList
        'Sometimes, an error is produced while cloning remarks. The error is caused by
        'typeid mismatch. Using GetRemarkTypeIndex to guard in case this error happens.
        lTypeIndex = GetRemarkTypeIndex(i_src, oRemark.TypeID)
        If lTypeIndex > 0 Then
            o_dst.AddRemark lTypeIndex, oRemark.MemoText
        End If
    Next
End Sub


Public Function GetRemarkTypeIndex(ByRef SRContext As RemarkContext, sTypeID As String) As Long
    Dim lIndex As Long
    Dim oRemarkType As RemarkType
    
    For Each oRemarkType In SRContext
        lIndex = lIndex + 1
        If Trim(oRemarkType.TypeID) = Trim(sTypeID) Then
            GetRemarkTypeIndex = lIndex
            Exit Function
        End If
    Next
    
    GetRemarkTypeIndex = 0
End Function


Public Function GetLineStatus(oValue As LineStatus) As String
    Select Case oValue
        Case IsInvoiced:
            GetLineStatus = "Invoiced"
        Case IsShipComplete:
            GetLineStatus = "Shipped"
        Case IsShipBackorders:
            GetLineStatus = "Shipped with Backorders"
        Case IsReadyToShip:
            GetLineStatus = "Available to Pack"
        Case IsOnOrder:
            GetLineStatus = "On Order"
        Case IsNeedsToBeOrder:
            GetLineStatus = "Needs to Be Ordered"
        Case IsDropShipInActive:
            GetLineStatus = "DropShip In Active"
        Case IsDropShipCancelled:
            GetLineStatus = "DropShip Cancelled"
        Case IsDropShipClosed:
            GetLineStatus = "DropShip Closed"
        Case LineStatus.IsGskNew:
            GetLineStatus = "Not Yet Started"
        Case LineStatus.IsGskOutOfStock:
            GetLineStatus = "Out Of Stock"
        Case LineStatus.IsGskBegin:
            GetLineStatus = "Being Cut"
        Case LineStatus.IsGskCut:
            GetLineStatus = "Being Molded"
        Case LineStatus.IsGskMold:
            GetLineStatus = "Being Trimmed"
        Case LineStatus.IsGskTrim:
            GetLineStatus = "Complete"
        Case LineStatus.IsGskNotAvail:
            GetLineStatus = "Gsk Status Not Avail"
        Case LineStatus.IsBackOrderCancelled:
            GetLineStatus = "Back Order Cancelled"
        Case LineStatus.IsShipping:
            GetLineStatus = "Shipping"
        Case LineStatus.IsPacking:
            GetLineStatus = "Packing"
    End Select
End Function


'2/3/05 LR created
'to repurpose the FVendor form as a generic popup window
'factors code in FOrder.cmdRMAvendor_Click() and FOrder.cmdVendorSetails_Click()

'***466 LR added fixed-width-font parameter

Public Sub DisplayInfo(ByVal sCaption As String, ByVal sBody As String, ByVal iWidth As Integer, ByVal iHeight As Integer, Optional ByVal bFixedWidth = False)
    Dim oFrm As FDisplayInfo
    
    Set oFrm = New FDisplayInfo
    oFrm.width = iWidth
    oFrm.Height = iHeight
    oFrm.caption = sCaption
    
    If bFixedWidth Then
        oFrm.Font.Name = "Courier New"
    End If
    
    oFrm.body = sBody
    oFrm.Show vbModal
    Unload oFrm
    Set oFrm = Nothing
End Sub


Public Sub SaveAsTextFile(filename As String, buffer As String)
    Dim fso As FileSystemObject
    Dim ts As TextStream
    Set fso = New FileSystemObject
    Set ts = fso.CreateTextFile(filename, Overwrite:=True, Unicode:=True)
    ts.Write buffer
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub


Public Function msg( _
        ByRef i_sPrompt As String, _
        Optional ByVal i_eMsgBoxStyle As VbMsgBoxStyle = vbOKOnly, _
        Optional ByRef i_sTitle As String = "SageAssistant" _
) As VbMsgBoxResult
    SuppressWaitCursor True
    msg = MsgBox(i_sPrompt, i_eMsgBoxStyle, i_sTitle)
    SuppressWaitCursor False

End Function


Public Function Version(Optional bIncludeName As Boolean = False) As String
    Dim sOutput As String
    
    With App
        sOutput = .Major & "." & .Minor & "." & .Revision
        If bIncludeName Then
            sOutput = .ProductName & " " & sOutput
            sOutput = sOutput & " - (User: " & GetUserID(GetUserKey) & "   Whse: " & GetUserWhseID & "   DB: " & g_DB.server & "." & g_DB.database & ")"
        End If
    End With
    Version = sOutput
End Function


Public Function ReadFile(i_sPath As String) As String
    Dim fso As FileSystemObject
    Dim ts As TextStream

    Set fso = New FileSystemObject
    Set ts = fso.OpenTextFile(i_sPath, ForReading)
    ReadFile = ts.ReadAll

    Set ts = Nothing
    Set fso = Nothing
End Function


'Called by the IItem_Commit function in each of the specific item classes.

Public Sub PrepareToCommitItem(o_oBaseItem As Item, i_sPrefix As String, i_lWhseKey As Long)
     
    On Error GoTo ErrorHandler
    
    With o_oBaseItem
        .ItemID = i_sPrefix & "-" & GetWhseIDFromWhseKey(i_lWhseKey)
        g_rstPartIDs.Filter = "ItemID='" & .ItemID & "'"
        .ItemKey = g_rstPartIDs.Fields("ItemKey").value
        g_rstPartIDs.Filter = adFilterNone
    End With
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "PrepareToCommitItem", Err.Source
End Sub


'TODO: ShipMethod requires rstShipVia as a parameter. Ugh!
'What is Order.ShipMethod(g_rstShipVia) doing exactly?

Public Function ShipCharge(ByRef oOrder As Order) As Double
    With oOrder
        If InStr(1, .ShipMethod, "Red") <> 0 Then
            ShipCharge = 75#
        ElseIf InStr(1, .ShipMethod, "Blue") <> 0 Then
            ShipCharge = 40#
        ElseIf InStr(1, .ShipMethod, "3 day") <> 0 Then
            ShipCharge = 30#
        ElseIf InStr(1, .ShipMethod, "Ground") <> 0 Then
            ShipCharge = 20#
        ElseIf InStr(1, .ShipMethod, "STND") <> 0 Then
            ShipCharge = 100#
        End If
    End With
End Function


'************************************************************************
' SALESTAX stuff

'Called by
'   SalesTax.InheritExempt
'   FAcctRcv.cmdTaxUpdate_Click

'presumably a customer can have more than one certificate

Public Sub UpdateExemptCerts(CustKey As Long, ExmptNo As String, StateID As String)
    Dim sSQL As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrorHandler
       
    sSQL = "SELECT tarCustomer.CustID, tarCustSTaxExmpt.ExmptNo, tciAddress.StateID, " & _
                "tciSTaxSchdCodes.STaxCodeKey, tarCustAddr.AddrKey, tarCustomer.CustKey " & _
            "FROM tciSTaxSchdCodes INNER JOIN tarCustAddr INNER JOIN tciAddress ON tarCustAddr.AddrKey = tciAddress.AddrKey INNER JOIN " & _
                "tarCustomer ON tarCustAddr.CustKey = tarCustomer.CustKey ON tciSTaxSchdCodes.STaxSchdKey = tarCustAddr.STaxSchdKey LEFT Outer Join tciSTaxCode INNER JOIN " & _
                "tarCustSTaxExmpt ON tciSTaxCode.STaxCodeKey = tarCustSTaxExmpt.STaxCodeKey ON tciSTaxSchdCodes.STaxCodeKey = tciSTaxCode.STaxCodeKey AND " & _
                "tarCustAddr.AddrKey = tarCustSTaxExmpt.AddrKey  " & _
            "WHERE tarCustomer.CustKey = " & CStr(CustKey)

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdText
    End With

    Set rst = New ADODB.Recordset
    With rst
        .Open sSQL, g_DB.Connection, adOpenForwardOnly, adLockReadOnly
        Do While Not .EOF
            If Trim(.Fields("StateID").value) = Trim(StateID) Then
                If Not IsNull(.Fields("ExmptNo").value) Then
                    cmd.CommandText = "UPDATE tarCustSTaxExmpt SET ExmptNo = '" & Trim(ExmptNo) & "' WHERE AddrKey=" & CStr(.Fields("AddrKey").value) & " AND sTaxCodeKey=" & CStr(.Fields("sTaxCodeKey").value)
                Else
                    cmd.CommandText = "INSERT tarCustSTaxExmpt VALUES(" & CStr(.Fields("AddrKey").value) & ", " & CStr(.Fields("sTaxCodeKey").value) & ", '" & Trim(ExmptNo) & "')"
                End If
                cmd.Execute
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    Set rst = Nothing
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    ClearWaitCursor
    msg Err.Description, vbOKOnly + vbCritical, Err.Source
End Sub


'************************************************************************

'Called By
'   FItemSearch.Find
'   FItemSearch.ChooseFromGrid
'   FCrossRef.XRefSearch
'   FXRef.ItemType property

Public Function ConvertSageItemType(ByVal i_lSageItemType As Long) As ItemTypeCode
    Select Case i_lSageItemType
    Case 5
        ConvertSageItemType = itFinishedGood
    Case 7
        ConvertSageItemType = itBTOKit
    Case Else
        Err.Raise -1, "ConvertSageItemType", "Unexpected SageItemType: " & i_lSageItemType
    End Select
End Function


'Called By
'   FOrder.StatusCode
'   FOrder.ItemUpdateControls
'   Order.EventNewOrder
'   Order.EventOldOrder
'   Order.EventNewOrderLine
'   Order.EventOldOrderLine

Public Function StatusCodeString(i_eStatusCode As ItemStatusCode) As String
    Select Case i_eStatusCode
    Case ItemStatusCode.iscResearch
        StatusCodeString = "Need to Research"
    Case ItemStatusCode.iscQuote
        StatusCodeString = "Need to Quote"
    Case ItemStatusCode.iscAuthorize
        StatusCodeString = "Need Customer Authorization"
    Case ItemStatusCode.iscReadyToCommit
        StatusCodeString = "Ready to Commit"
    Case ItemStatusCode.iscEmpty
        StatusCodeString = "New Order"
    Case ItemStatusCode.iscCommitted
        StatusCodeString = "Committed"
    Case ItemStatusCode.iscHasRMA
        StatusCodeString = "Return Merch Auth"
    Case ItemStatusCode.iscDeleted
        StatusCodeString = "Deleted"
    Case ItemStatusCode.iscARHold
        StatusCodeString = "A/R Cust Hold"
    Case ItemStatusCode.iscPendingCommit
        StatusCodeString = "Pending Commit"
    Case Else
        StatusCodeString = "Unknown"
    End Select
End Function


Public Function ResearchStatusString(i_eResearchStatus As ItemResearchStatus) As String
    Select Case i_eResearchStatus
    Case ItemResearchStatus.irsNeedResearch
        ResearchStatusString = "Need to Research"
    Case ItemResearchStatus.irsContactFactory
        ResearchStatusString = "Contact Factory"
    Case ItemResearchStatus.irsContactCustomer
        ResearchStatusString = "Contact Customer"
    Case ItemResearchStatus.irsWaitFactory
        ResearchStatusString = "Waiting for Factory"
    Case ItemResearchStatus.irsWaitCustomer
        ResearchStatusString = "Waiting for Customer"
    Case Else
        ResearchStatusString = "Unknown research status"
    End Select
End Function


'*******************************************************************
' Warehouse stuff
' transit warehouses vs actual warehouses

Public Function GetWhseID(StateID As String, CountryID As String) As String
    Dim rst As ADODB.Recordset
    
    Set rst = LoadDiscRst("Select BranchID from tcpTerritory " _
                          & "Where CountryID = '" & CountryID & "' " _
                          & "and StateID = '" & StateID & "'")
    If rst.RecordCount = 0 Then
        GetWhseID = "MPK" 'HACK!
    Else
        GetWhseID = rst.Fields("BranchID").value
    End If
End Function


Public Sub SetUpWarehouses(cboWarehouse As ComboBox, _
                            ByRef rstWhses As ADODB.Recordset, _
                            ByVal WhseKey As Long)
    
    If Not rstWhses Is Nothing Then
        rstWhses.Filter = "transit = 0"
        Helpers.LoadCombo cboWarehouse, rstWhses, "WhseID", "WhseKey", WhseKey
        rstWhses.Filter = adFilterNone
    End If
    
End Sub

'*******************************************************************


Public Function WhoSetFreeFreight(ByVal OPKey As Long) As String
    Dim ocmd As ADODB.Command
    Set ocmd = CreateCommandSP("spcpcWhoSetFreeFreight")
    ocmd.Parameters("@_iOPKey").value = OPKey
    ocmd.Execute
    WhoSetFreeFreight = IIf(IsNull(ocmd.Parameters("@_oCSR").value), vbNullString, ocmd.Parameters("@_oCSR").value)
    Set ocmd = Nothing
End Function


'this is for sending an email to a list of recipients

Public Sub SendNotification(ByVal sSubject As String, ByVal sMessage As String, ByVal vsaRecipientsArray As Variant)
    Dim lvarRecipient As Variant
    Dim sRecipients As String
    
    On Error GoTo EH
    
    If IsArray(vsaRecipientsArray) Then
        For Each lvarRecipient In vsaRecipientsArray
            If Len(Trim$(lvarRecipient)) > 0 Then
                sRecipients = sRecipients & ";" & lvarRecipient
            End If
        Next lvarRecipient
    Else
        sRecipients = vsaRecipientsArray
    End If
    If Left$(sRecipients, 1) = ";" Then
        sRecipients = Mid(sRecipients, 2)
    End If

    EMail.Send GetUserID & "@caseparts.com", sRecipients, sSubject, sMessage, False
    
    Exit Sub
EH:
    'Handle error here to suppress error dialog
    LogDB.LogError "Failed to email notification due to error '", "GlobalFunctions", "SendNotification", Err.Source, Err.Number, Err.Description
End Sub


'10/24/08
'moved from Item.cls (& SOAPI.cls)

Public Function GetVendorAddrKey(i_lVendKey As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim orst As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT DfltPurchAddrKey FROM tapvendor WHERE VendKey = " & i_lVendKey
    Set orst = LoadDiscRst(sSQL)
    GetVendorAddrKey = 0
    If Not orst.EOF Then
        GetVendorAddrKey = orst.Fields("DfltPurchAddrKey").value
    End If
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "Item.GetVendorAddrKey", Err.Description
End Function

'10/24/08
'moved from Item.cls (& SOAPI.cls)

Public Function GetGLAcctKey(i_lUserKey As Long) As Long
    GetGLAcctKey = GetUserSalesAcctKey(i_lUserKey)
End Function

' Proxy for validating & classifying addresses through UPS.
' Called by FThisOrderOnlyAddress2

Public Function CreateUPSProxy() As MSSOAPLib30.SoapClient30
    Set CreateUPSProxy = New MSSOAPLib30.SoapClient30
    CreateUPSProxy.MSSoapInit g_UPSOnlineURL & "?WSDL"
End Function

Public Function OrderHasShipment(OPKey As Long) As Boolean
    Dim ocmd As ADODB.Command
    Set ocmd = CreateCommandSP("spcpcOrderHasShipment")
    With ocmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@OPKey", adInteger, adParamInput, , OPKey)
        .Execute
        OrderHasShipment = .Parameters("RETURN_VALUE")
    End With
End Function

'2-copy receipt printed status is maintained in the DB but not the order object.
'These functions are for CRUDing

Public Function OrderHasReceipt(OPKey As Long) As Boolean
    Dim ocmd As ADODB.Command
    Set ocmd = CreateCommandSP("spcpcOrderHasReceipt")
    With ocmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@OPKey", adInteger, adParamInput, , OPKey)
        .Execute
        OrderHasReceipt = .Parameters("RETURN_VALUE")
    End With
End Function


Public Sub SetReceiptPrinted(OPKey As Long)
    Dim ocmd As ADODB.Command
    Set ocmd = CreateCommandSP("spcpcSetOrderHasReceipt")
    With ocmd
        .Parameters.Append .CreateParameter("@OPKey", adInteger, adParamInput, , OPKey)
        .Execute
    End With
End Sub


Public Sub ClearReceiptPrinted(OPKey As Long)
    Dim ocmd As ADODB.Command
    Set ocmd = CreateCommandSP("spcpcClearOrderHasReceipt")
    With ocmd
        .Parameters.Append .CreateParameter("@OPKey", adInteger, adParamInput, , OPKey)
        .Execute
    End With
End Sub


'Save a string (typically XML or HTML) to a file

Public Sub SaveToFile(ByRef i_strPath As String, ByRef i_strText As String)
    Dim fs As FileSystemObject
    Dim ts As TextStream
    
    Set fs = New FileSystemObject
    Set ts = fs.CreateTextFile(i_strPath, Overwrite:=True, Unicode:=True)
    ts.Write i_strText
    ts.Close
    Set ts = Nothing
    Set fs = Nothing
End Sub


Public Function XMLURL() As String
    XMLURL = g_SnapshotPath & GetUserName & ".xml"
End Function


Public Function XslHeader(XslPath As String) As String
    XslHeader = "<?xml version=""1.0""?>" + vbCrLf _
              + "<!DOCTYPE xsl:stylesheet [<!ENTITY nbsp ""&#160;"">]>" + vbCrLf _
              + "<?xml-stylesheet type=""text/xsl"" href=""" _
              + XslPath + """?>" + vbCrLf
End Function



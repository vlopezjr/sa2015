VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const m_skSource = "Order"
    
Public Enum SOStatusCode
    sscUnacknowledged = 0
    sscOpen = 1
    sscInactive = 2
    sscCancelled = 3
    sscClosed = 4
    sscIncomplete = 5
    sscPendingApproval = 6
End Enum

Public Enum CustomerStatusCode
    cscActive = 1
    cscInactive = 2
    cscTemporary = 3
    cscDeleted = 4
End Enum


Private Type OrderProps
    
    lOPKey As Long
    lSOKey As Long
    lTranNo As Long
    
    lUserKey As Long
    sUserID As String

    lWhseKey As Long
    sWhseID As String
    
    lShipMethKey As Long
    lDropShipVendKey As Long
    lBillMethKey As Long

    dCreateDate As Date
    dUpdateDate As Date
   
    sPurchOrd As String * 15
    
    sInfo As String

    lFlagBitmap As Long
'        .bDropShip = ((.lFlagBitmap And &H1) = &H1)
'        .bShipComplete = ((.lFlagBitmap And &H2) = &H2)
'        .bPrintPickList = ((.lFlagBitmap And &H4) = &H4)
'        .bPricePackList = ((.lFlagBitmap And &H8) = &H8)
'        .bFreeFreight = ((.lFlagBitmap And &H10) = &H10)
'        .bReducedFreight = ((.lFlagBitmap And &H20) = &H20)
'        .bPartsNoCharge = ((.lFlagBitmap And &H40) = &H40)
'        .bInboundFreight = ((.lFlagBitmap And &H100) = &H100)
'        .bWalkup = ((.lFlagBitmap And &H200) = &H200)
'        .bMisc = ((.lFlagBitmap And &H400) = &H400)
'        .bDeposit = ((.lFlagBitmap And &H800) = &H800)
'        .bBillDifferentRate = ((.lFlagBitmap And &H4000) = &H4000)
    bDropShip As Boolean
    bShipComplete As Boolean
    bPrintPickList As Boolean
    bPricePackList As Boolean
    bFreeFreight As Boolean
    bReducedFreight As Boolean
    bPartsNoCharge As Boolean
    bInboundFreight As Boolean
    bWalkup As Boolean
    bMisc As Boolean
    bDeposit As Boolean
    bBillDifferentRate As Boolean
    
    bNewOrder As Boolean
    bARCustHold As Boolean
    bCommitting As Boolean      'this is "pending commit" or "marked for commit"
    bDeleted As Boolean
    bRMA As Boolean
    
    bUpdateCustomer As Boolean
    bRecommit As Boolean
    bPrintReceipt As Boolean

    lUpdateCounter As Long
    
    sOldSOID As String
    eResearchStatus  As ItemResearchStatus
    sUPSAcct As String * 6
    
    sShipToName As String * 40
    sShipToPhone As String * 17
    
    sShipToNote As String
End Type

Private m_udtOrder As OrderProps
Private m_udtOrderBackup As OrderProps

Private m_oBackUpStatusCode As ItemStatusCode

Private m_oPmtTerms As PaymentTerms

Private m_oContact As contact

Private m_oItems As Items
Private m_oBackUpItems As Items

Private m_oCustomer As Customer

Private WithEvents m_oSalesTax As SalesTax
Attribute m_oSalesTax.VB_VarHelpID = -1

Private m_oCreditCard As CreditCard

Private m_oRemarkContext As MemoMeister.RemarkContext

Private m_bLoading As Boolean


'****************************************************************************
'Default Class Methods
'****************************************************************************

Private Sub Class_Initialize()
    Set m_oCustomer = New Customer
    
    Set m_oItems = New Items
    Set m_oBackUpItems = New Items
    
    Set m_oPmtTerms = New PaymentTerms
    
    Set m_oSalesTax = New SalesTax
    
    Set m_oRemarkContext = New MemoMeister.RemarkContext
    m_oRemarkContext.EditBehavior = NoAutoSave
End Sub


'****************************************************************************
' Public Property Procedures
'****************************************************************************

Public Property Get Customer() As Customer
    Set Customer = m_oCustomer
End Property

Public Property Let Customer(ByVal oNewValue As Customer)
    Set m_oCustomer = oNewValue
End Property

Public Property Get SalesTax() As SalesTax
    Set SalesTax = m_oSalesTax
End Property

Public Property Get Items() As Items
    Set Items = m_oItems
End Property

Public Property Let Items(ByVal oNewValue As Items)
    Set m_oItems = oNewValue
End Property

Public Property Get CreditCard() As CreditCard
    Set CreditCard = m_oCreditCard
End Property

Public Property Let CreditCard(ByVal oNewValue As CreditCard)
    Set m_oCreditCard = oNewValue
End Property

Public Property Get OPKey() As Long
    OPKey = m_udtOrder.lOPKey
End Property

Public Property Let OPKey(ByVal vNewValue As Long)
    m_udtOrder.lOPKey = vNewValue
End Property


Public Property Get soKey() As Long
    soKey = m_udtOrder.lSOKey
End Property

Public Property Let soKey(ByVal vNewValue As Long)
    m_udtOrder.lSOKey = vNewValue
End Property


Public Property Get TranNo() As Long
    TranNo = m_udtOrder.lTranNo
End Property

Public Property Let TranNo(lNewValue As Long)
    m_udtOrder.lTranNo = lNewValue
End Property


Public Property Get OldSOID() As String
    OldSOID = m_udtOrder.sOldSOID
End Property

Public Property Let OldSOID(bNewValue As String)
    m_udtOrder.sOldSOID = bNewValue
End Property


Public Property Get bRMA() As Boolean
    bRMA = m_udtOrder.bRMA
End Property

Public Property Let bRMA(ByVal bNewValue As Boolean)
    m_udtOrder.bRMA = bNewValue
    m_udtOrder.bCommitting = Not m_udtOrder.bRMA
End Property


Public Property Get HasFreeFreight() As Boolean
    HasFreeFreight = m_udtOrder.bFreeFreight
End Property

Public Property Let HasFreeFreight(ByVal bNewValue As Boolean)
    m_udtOrder.bFreeFreight = bNewValue
End Property


Public Property Get ReducedFreight() As Boolean
    ReducedFreight = m_udtOrder.bReducedFreight
End Property

Public Property Let ReducedFreight(ByVal bNewValue As Boolean)
    m_udtOrder.bReducedFreight = bNewValue
End Property


Public Property Get BillDifferentRate() As Boolean
    BillDifferentRate = m_udtOrder.bBillDifferentRate
End Property

Public Property Let BillDifferentRate(ByVal bNewValue As Boolean)
    m_udtOrder.bBillDifferentRate = bNewValue
End Property


Public Property Get IsWalkup() As Boolean
    IsWalkup = m_udtOrder.bWalkup
End Property

Public Property Let IsWalkup(ByVal bNewValue As Boolean)
    m_udtOrder.bWalkup = bNewValue
End Property


Public Property Get IsMisc() As Boolean
    IsMisc = m_udtOrder.bMisc
End Property

Public Property Let IsMisc(ByVal bNewValue As Boolean)
    m_udtOrder.bMisc = bNewValue
End Property


Public Property Get HasInboundFreight() As Boolean
    HasInboundFreight = m_udtOrder.bInboundFreight
End Property

Public Property Let HasInboundFreight(ByVal bNewValue As Boolean)
    m_udtOrder.bInboundFreight = bNewValue
End Property


Public Property Get HasDeposit() As Boolean
    HasDeposit = m_udtOrder.bDeposit
End Property

Public Property Let HasDeposit(ByVal bNewValue As Boolean)
    m_udtOrder.bDeposit = bNewValue
End Property


Public Property Get HasPartsNoCharge() As Boolean
    HasPartsNoCharge = m_udtOrder.bPartsNoCharge
End Property

Public Property Let HasPartsNoCharge(ByVal bNewValue As Boolean)
    m_udtOrder.bPartsNoCharge = bNewValue
End Property


Public Property Get ShipComplete() As Boolean
    ShipComplete = m_udtOrder.bShipComplete
End Property

Public Property Let ShipComplete(ByVal bNewValue As Boolean)
    m_udtOrder.bShipComplete = bNewValue
End Property


Private Property Get FlagBitmap() As Long
    FlagBitmap = GetFlag
End Property

Private Property Let FlagBitmap(lNewValue As Long)
    m_udtOrder.lFlagBitmap = lNewValue
    ParseFlag
End Property


Public Property Get ShipMethod() As String
    Dim ShipMethId As String
    ShipMethId = ShipMethKeytoID(m_udtOrder.lShipMethKey)
    ShipMethod = Trim(Mid(ShipMethId, 5, 20))
End Property


Public Property Get ShipMethKey() As Long
    ShipMethKey = m_udtOrder.lShipMethKey
End Property

Public Property Let ShipMethKey(ByVal i_lShipMethKey As Long)
    Dim WasWillCall As Boolean
    Dim lShipMethKey As Long

    'Cache current values
    lShipMethKey = m_udtOrder.lShipMethKey
    WasWillCall = IsWillCall
    
    m_udtOrder.lShipMethKey = i_lShipMethKey

    'If the value has changed
    If i_lShipMethKey <> lShipMethKey Then
        'Load Sales Tax
        If IsWillCall Then
            m_oSalesTax.WillCallTaxOverride m_udtOrder.sWhseID
        ElseIf WasWillCall Then
            m_oSalesTax.Restore
        End If
    End If
End Property


Public Property Get PmtTerms() As PaymentTerms
    Set PmtTerms = m_oPmtTerms
End Property


Public Property Get ShipToName() As String
    ShipToName = Trim(m_udtOrder.sShipToName)
End Property

Public Property Let ShipToName(ByVal vNewValue As String)
    m_udtOrder.sShipToName = vNewValue
End Property

Public Property Get ShipToPhone() As String
    ShipToPhone = Trim(m_udtOrder.sShipToPhone)
End Property

Public Property Let ShipToPhone(ByVal vNewValue As String)
    m_udtOrder.sShipToPhone = vNewValue
End Property


Public Property Get ShipToNote() As String
    ShipToNote = Trim$(m_udtOrder.sShipToNote)
End Property
    
Public Property Let ShipToNote(ByVal vNewValue As String)
    m_udtOrder.sShipToNote = vNewValue
End Property


Public Property Get WhseKey() As Long
    WhseKey = m_udtOrder.lWhseKey
End Property

Public Property Let WhseKey(ByVal lNewValue As Long)
    If lNewValue <> m_udtOrder.lWhseKey Then
        m_udtOrder.lWhseKey = lNewValue
        m_udtOrder.sWhseID = WhseKeyToID(lNewValue)
    End If
End Property

Public Property Get whseid() As String
    whseid = m_udtOrder.sWhseID
End Property

Public Property Get PricePackList() As Boolean
    PricePackList = m_udtOrder.bPricePackList
End Property

Public Property Let PricePackList(bNewValue As Boolean)
    m_udtOrder.bPricePackList = bNewValue
End Property


Public Property Get Recommit() As Boolean
    Recommit = m_udtOrder.bRecommit
End Property

Public Property Let Recommit(bNewValue As Boolean)
    m_udtOrder.bRecommit = bNewValue
End Property


Public Property Get IsNewOrder() As Boolean
    IsNewOrder = m_udtOrder.bNewOrder
End Property


'This is called from FOrder.txtCustID_LostFocus
'It's used to indicate that a NewCustomer order has been converted into a Customer order.

Public Property Let UpdateCustomer(bNewValue As Boolean)
    m_udtOrder.bUpdateCustomer = bNewValue
End Property


Public Property Get UserKey() As Long
    UserKey = m_udtOrder.lUserKey
End Property


Public Property Let UserKey(ByVal lNewValue As Long)
    'If the UserKey is changed, change the UserID also.
    m_udtOrder.lUserKey = lNewValue
    m_udtOrder.sUserID = GetUserID(lNewValue)
End Property


Public Property Get UserID() As String
    UserID = m_udtOrder.sUserID
End Property


Public Property Get contact() As contact
    Set contact = m_oContact
End Property

Public Property Let contact(ByRef oContact As contact)
    Set m_oContact = oContact
End Property


Public Property Get UPSAcct() As String
    UPSAcct = Trim(m_udtOrder.sUPSAcct)
End Property

Public Property Let UPSAcct(ByVal sNewValue As String)
    m_udtOrder.sUPSAcct = sNewValue
End Property


Public Property Get PurchOrd() As String
    PurchOrd = Trim(m_udtOrder.sPurchOrd)
End Property

Public Property Let PurchOrd(ByVal vNewValue As String)
    m_udtOrder.sPurchOrd = vNewValue
End Property


Public Property Get Info() As String
    Info = Trim(m_udtOrder.sInfo)
End Property

Public Property Let Info(ByVal vNewValue As String)
    m_udtOrder.sInfo = vNewValue
End Property


Public Property Get IsWillCall() As Boolean
    IsWillCall = IIf((m_udtOrder.lShipMethKey = 27 Or m_udtOrder.lShipMethKey = 37 Or m_udtOrder.lShipMethKey = 32), True, False)
End Property


Public Property Get IsDropShip() As Boolean
    IsDropShip = m_udtOrder.bDropShip
End Property

Public Property Let IsDropShip(ByVal bNewValue As Boolean)
    m_udtOrder.bDropShip = bNewValue
End Property


Public Property Get bARCustHold() As Boolean
    bARCustHold = m_udtOrder.bARCustHold
End Property

Public Property Let bARCustHold(ByVal bNewValue As Boolean)
    m_udtOrder.bARCustHold = bNewValue
End Property


Public Property Get bCommitting() As Boolean
    bCommitting = m_udtOrder.bCommitting
End Property

Public Property Let bCommitting(ByVal bNewValue As Boolean)
    m_udtOrder.bCommitting = bNewValue
End Property


'Public Property Get PrintPickList() As Boolean
'    PrintPickList = m_udtOrder.bPrintPickList
'End Property
'
'Public Property Let PrintPickList(ByVal bNewValue As Boolean)
'    m_udtOrder.bPrintPickList = bNewValue
'End Property


Public Property Get PrintReceipt() As Boolean
    PrintReceipt = m_udtOrder.bPrintReceipt
End Property

Public Property Let PrintReceipt(ByVal bNewValue As Boolean)
    m_udtOrder.bPrintReceipt = bNewValue
End Property


Public Property Get DropShipVendKey() As Long
    DropShipVendKey = m_udtOrder.lDropShipVendKey
End Property

Public Property Let DropShipVendKey(ByVal lNewValue As Long)
    m_udtOrder.lDropShipVendKey = lNewValue
End Property


Public Property Get BillMethKey() As Long
    BillMethKey = m_udtOrder.lBillMethKey
End Property

Public Property Let BillMethKey(ByVal lNewValue As Long)
    m_udtOrder.lBillMethKey = lNewValue
End Property


Public Property Get HasSpecialHandling() As Boolean
    HasSpecialHandling = HasDeposit Or HasFreeFreight Or ReducedFreight Or HasInboundFreight Or HasPartsNoCharge
End Property


'Called internally by
'   Load()
'   LogSalesOrderEvent()
'   EventNewOrder()

Public Property Get StatusCode() As ItemStatusCode
    If m_udtOrder.bDeleted Then
        StatusCode = iscDeleted
    ElseIf m_udtOrder.bRMA Then
        StatusCode = ItemStatusCode.iscHasRMA
    ElseIf m_udtOrder.bARCustHold Then
        StatusCode = ItemStatusCode.iscARHold
    ElseIf m_udtOrder.bCommitting Then
        StatusCode = ItemStatusCode.iscPendingCommit
    ElseIf m_udtOrder.lSOKey > 0 Then
        StatusCode = ItemStatusCode.iscCommitted
    Else
        StatusCode = m_oItems.StatusCode
    End If
End Property


Public Property Get ResearchStatus() As ItemResearchStatus
    ResearchStatus = m_udtOrder.eResearchStatus
End Property

Public Property Let ResearchStatus(RHS As ItemResearchStatus)
    m_udtOrder.eResearchStatus = RHS
End Property
       

Public Property Get CreateDate() As Date
    CreateDate = m_udtOrder.dCreateDate
End Property


Public Property Get UpdateDate() As Date
    UpdateDate = m_udtOrder.dUpdateDate
End Property


Public Property Get RemarkContext() As MemoMeister.RemarkContext
    Set RemarkContext = m_oRemarkContext
End Property


'*****************************************************************************************
' Public Methods
'*****************************************************************************************

Public Sub Load(lOPKey As Long)
    Dim orstOrder As ADODB.Recordset
    Dim orstLine As ADODB.Recordset
    Dim oItem As Item
    Dim sSQL As String
    
    On Error GoTo EH

    If lOPKey <= 0 Then
        On Error GoTo 0
        Err.Raise -1, "Order.Load", _
            "Order.Load" & vbCrLf & "OPKey=" & lOPKey & " is not valid."
    End If
    
    Set orstOrder = LoadDiscRst("SELECT * FROM tcpSO WHERE OPKey=" & lOPKey)
    
    If orstOrder.EOF Then
        orstOrder.Close
        On Error GoTo 0
        Err.Raise -1, "Order.Load", _
            "Order.Load" & vbCrLf & "OPKey=" & lOPKey & " not found in database."
    End If
    
    Me.Clear ClearCust:=True
        
    LoadUdtFromRst orstOrder

    sSQL = "SELECT * FROM tcpSOLine WHERE OPKey=" & lOPKey & " ORDER BY [LineNo]"
    Set orstLine = LoadDiscRst(sSQL)
    
    'No LineItems
    If orstLine.EOF Then
        orstLine.Close
        m_oItems.Clear
        m_oBackUpItems.Clear
    Else
        Set m_oItems = New Items
        m_oItems.IsTaxable = orstOrder.Fields("TaxStatusCode")
        Set m_oBackUpItems = New Items
        Do While Not orstLine.EOF
            m_oItems.ImportItem ImportString(orstLine.Fields("LineXML")), m_udtOrder.lWhseKey
            m_oBackUpItems.ImportItem ImportString(orstLine.Fields("LineXML")), m_udtOrder.lWhseKey
            orstLine.MoveNext
        Loop
    End If

    LoadCustomerFromRst orstOrder
    
    m_oSalesTax.Init m_oCustomer
    If IsWillCall Then
        m_oSalesTax.WillCallTaxOverride m_udtOrder.sWhseID
    End If
    
    'if this order has a contact key, instantiate a contact object and load it
    'otherwise order.contact will be nothing.
    If Not IsNull(orstOrder.Fields("CntctKey")) And _
        orstOrder.Fields("CntctKey") > 0 Then
   
        'If the order's customer has an account
        If m_oCustomer.HasAccount Then
            'Find the contact in the collection by CntctKey
            'and set the order's contact reference to the found contact.
            'If the contact does not exist in the collection, just continue.
            On Error Resume Next
            Set m_oContact = m_oCustomer.Contacts.GetContactByKey(Trim(orstOrder.Fields("CntctKey")))
            On Error GoTo EH
        Else
            Set m_oContact = New contact
            'm_oContact.Connection = g_DB.Connection
            m_oContact.LoadByKey CLng(Trim(orstOrder.Fields("CntctKey")))
        End If
    End If
    
    m_oRemarkContext.Load "ViewOrder", lOPKey

    'Load the credit card and the transaction collection
    If Not IsNull(orstOrder.Fields("CCKey")) Then
        Set m_oCreditCard = New CreditCard
        m_oCreditCard.Load orstOrder.Fields("CCKey"), Me
    End If
    
    Backup
    m_oBackUpStatusCode = StatusCode
    
    Set orstOrder = Nothing
    Set orstLine = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, "Order.Load", _
        "Order.Load" & vbCrLf & Err.Description
End Sub


Private Sub LoadUdtFromRst(orstOrder As ADODB.Recordset)
    With m_udtOrder
        .lOPKey = orstOrder.Fields("OPKey")
        .lSOKey = orstOrder.Fields("SOKey")

        'NOTE: TranKey is an INT in tcpSO but it sources from tsoSalesOrder.TranNo which is a string
        .lTranNo = orstOrder.Fields("TranKey")
        
        .dCreateDate = orstOrder.Fields("CreateDate")
        .dUpdateDate = orstOrder.Fields("UpdateDate")
        
        .sUserID = orstOrder.Fields("UserID")
        .lUserKey = GetUserKey(orstOrder.Fields("UserID"))
        
        .lWhseKey = orstOrder.Fields("WhseKey")
        .sWhseID = WhseKeyToID(.lWhseKey)
        
        .lShipMethKey = orstOrder.Fields("ShipMethKey")

        .sPurchOrd = orstOrder.Fields("PurchOrd")
        
        If IsNull(orstOrder.Fields("Info")) Then
            .sInfo = ""
        Else
            .sInfo = orstOrder.Fields("Info")
        End If
 
        m_oPmtTerms.Load orstOrder.Fields("PmtTermsKey")
        
        .bDeleted = orstOrder.Fields("StatusCode") = ItemStatusCode.iscDeleted
        .bRMA = orstOrder.Fields("StatusCode") = ItemStatusCode.iscHasRMA
        .bARCustHold = orstOrder.Fields("StatusCode") = ItemStatusCode.iscARHold
        .bCommitting = orstOrder.Fields("StatusCode") = ItemStatusCode.iscPendingCommit
 
        'What's going on here?
        'Was the Flag field added at a later date?
        If IsNull(orstOrder.Fields("Flags")) Then
            .bDropShip = CBool(orstOrder.Fields("IsDropShip"))
            .bPrintPickList = CBool(orstOrder.Fields("PrintPickList"))
            .bPricePackList = CBool(orstOrder.Fields("PricePackList"))
            .bShipComplete = CBool(orstOrder.Fields("ShipComplete"))
            .bFreeFreight = False
            .bReducedFreight = False
            .bBillDifferentRate = False
            .bPartsNoCharge = False
            .bInboundFreight = False
            .bDeposit = False
            .bWalkup = False
            .bMisc = False
        Else
            FlagBitmap = orstOrder.Fields("Flags")
        End If
        
        If IsNull(orstOrder.Fields("DropShipVendKey")) Then
            .lDropShipVendKey = 0
        Else
            .lDropShipVendKey = orstOrder.Fields("DropShipVendKey")
        End If
        
        If IsNull(orstOrder.Fields("ResearchStatus")) Then
            .eResearchStatus = irsResearchEmpty
        Else
            .eResearchStatus = orstOrder.Fields("ResearchStatus")
        End If
        
        If IsNull(orstOrder.Fields("BillMethKey")) Then
            .lBillMethKey = 0
        Else
            .lBillMethKey = orstOrder.Fields("BillMethKey")
        End If
        
        If IsNull(orstOrder.Fields("UPSAcct")) Then
            .sUPSAcct = ""
        Else
            .sUPSAcct = Trim(orstOrder.Fields("UPSAcct"))
        End If
    
        If IsNull(orstOrder.Fields("ShipToName")) Then
            .sShipToName = ""
        Else
            .sShipToName = orstOrder.Fields("ShipToName")
        End If
        
        If IsNull(orstOrder.Fields("ShipToPhone")) Then
            .sShipToPhone = ""
        Else
            .sShipToPhone = orstOrder.Fields("ShipToPhone")
        End If
      
        .lUpdateCounter = orstOrder.Fields("UpdateCounter")
        .bUpdateCustomer = False
        .bNewOrder = False
        
        'value can be null
        If IsNull(orstOrder.Fields("ShipToNote")) Then
            .sShipToNote = ""
        Else
            .sShipToNote = orstOrder.Fields("ShipToNote")
        End If
    End With
End Sub


Private Sub LoadCustomerFromRst(orstOrder As ADODB.Recordset)
    If orstOrder.Fields("CustKey") = 0 Then
        'Temp/New customer
        m_oCustomer.Import ImportString(orstOrder.Fields("CustXML"))
    Else
        m_oCustomer.Load orstOrder.Fields("CustKey")
        
        'Load Billing Address
        If orstOrder.Fields("BillAddrKey") = 0 Then
            'this case should never occur, but is included for completeness
            m_oCustomer.BillAddr.Import ImportString(orstOrder.Fields("BillAddrXML"))
            'Add this line in case sometimes that the shipaddr is empty.
            'set it's custkey equals customer key will let it eligible for committing later
            m_oCustomer.BillAddr.CustKey = m_oCustomer.Key
        Else
            m_oCustomer.BillAddr.Load orstOrder.Fields("BillAddrKey")
        End If
        
        'Load Shipping Address
        If orstOrder.Fields("ShipAddrKey") = 0 Then
            m_oCustomer.ShipAddr.Import ImportString(orstOrder.Fields("ShipAddrXML"))
            'Add this line in case sometimes that the shipaddr is empty.
            'set it's custkey equals customer key will let it eligible for committing later
            m_oCustomer.ShipAddr.CustKey = m_oCustomer.Key
        Else
            m_oCustomer.ShipAddr.Load orstOrder.Fields("ShipAddrKey")
        End If
    End If
    
    'Override customer default values
    With m_oCustomer
        .Name = Format(orstOrder.Fields("CustName"))
        .CustType = Format(orstOrder.Fields("CustType"))
        .ReqPO = orstOrder.Fields("ReqPO")
    End With
End Sub


'Calling hierarchy:
'
'   Save
'       SaveOrder
'           SaveOrderInfo
'           SaveCreditCard
'           SaveOrderLineItems
'           SaveOrderAdditionalInfo
'       logOrderEvent
'       SaveComittedOrder
'       LogCommittedOrderEvent
'       LogSageMMEvent
'       OrderLogBackUp

'What is the purpose of each flag?
'   Optional i_bForcePending As Variant
'   Optional i_bDelRMA As Boolean = False
'   Optional i_bSage As Boolean = False
'   Optional i_bCommitOrder As Boolean = False

'Public Function Save(
'                Optional i_bForcePending As Variant, _
'                Optional i_bSage As Boolean = False, _
'                Optional i_bCommitOrder As Boolean = False) As Boolean

Public Function Save(Optional i_bForcePending As Variant, _
                Optional i_bSage As Boolean = False, _
                Optional i_bCommitOrder As Boolean = False) As Boolean

    On Error GoTo ErrorHandler
    
    If Not IsMissing(i_bForcePending) Then
        If CBool(i_bForcePending) = False Then
            m_udtOrder.bARCustHold = False
            m_udtOrder.bCommitting = False
        End If
    End If
    
    'Warning! The local variable bNew is required to cache m_udtOrder.bNewOrder because
    'SaveOrder alters the value of m_udtOrder.bNewOrder.
    Dim bNew As Boolean
    bNew = m_udtOrder.bNewOrder
    
    If Not i_bSage Then
        SaveOrder StatusCode, i_bCommitOrder

        'if m_udtOrder.bNewOrder was True, SaveOrder changed it to False
        'LogOrderEvent needs to know it's original value --
        'which is why we've cached it in bNew
        LogOrderEvent bNew
    Else
        SaveCommittedOrder
        LogCommittedOrderEvent
        LogSageMMEvent
    End If
    
    OrderLogBackUp
    
    Save = True
    Exit Function
    
ErrorHandler:
    'if this saving process is not in a committing process, show up error message to user
    'Otherwise, raise error and jump to Order.Commit ErrorHandler
    Save = False
    If i_bCommitOrder Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        '***TODO Raise an error. Don't throw up messageboxes within this object.
        msg "Error Number " & Err.Number & ": " & Err.Description, vbExclamation + vbOKOnly, "[Order.Save function] " & Err.Source
    End If
End Function


Public Function XmlCommit(ByRef rstWhses As ADODB.Recordset, _
                    ByRef rstVendor As ADODB.Recordset) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Dim oCCnode As JDMPDXML.XMLNode
    
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        .Tag = "order"
        .Add "opkey", m_udtOrder.lOPKey
        .Add "statuscode", Me.StatusCode
        .Add "commituserid", GetUserName
        .Add "userkey", m_udtOrder.lUserKey
        .Add "userid", m_udtOrder.sUserID
        .Add "whsekey", m_udtOrder.lWhseKey
        .Add "whseid", m_udtOrder.sWhseID
        .Add "pmttermskey", m_oPmtTerms.Key
        .Add "pmttermsid", m_oPmtTerms.ID
        .Add "shipmethkey", m_udtOrder.lShipMethKey
        .Add "shipmethid", ShipMethKeytoID(m_udtOrder.lShipMethKey)
        .Add "purchord", Trim$(m_udtOrder.sPurchOrd)
        .Add "isdropship", m_udtOrder.bDropShip
        .Add "pricepacklist", m_udtOrder.bPricePackList
        .Add "freefreight", m_udtOrder.bFreeFreight
        .Add "staxschedkey", m_oSalesTax.STaxSchdKey
        .Add "exmptno", m_oSalesTax.ExmptNo
        .Add "amount", m_oItems.TotalPrice + m_oItems.TotalTax(Me.SalesTax.TaxRate) + ShipCharge(Me)

        If Not m_oCustomer Is Nothing Then
            .AddChild m_oCustomer.XmlCommit
        End If
        
        If Not m_oContact Is Nothing Then
            .AddChild m_oContact.XmlCommit
        End If
        
        If Not m_oCreditCard Is Nothing Then
            .AddChild m_oCreditCard.XmlCommit(m_udtOrder.sWhseID)
        End If
        
        If Not m_oItems Is Nothing Then
            .AddChild m_oItems.XmlCommit(m_udtOrder.lWhseKey, rstWhses, rstVendor)
        End If
    End With

    Set XmlCommit = oXML
End Function


'Called by FViewOrder.GetStrXML

Public Function Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False, _
    Optional bNoPricing As Boolean = False, _
    Optional bWarehouse As Boolean = False, _
    Optional bManfNotes As Boolean = False _
) As JDMPDXML.XMLNode
    
    Dim oXML As JDMPDXML.XMLNode
    
    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "Order"
        .Add "OPKey", m_udtOrder.lOPKey
        .Add "SOKey", m_udtOrder.lSOKey
        .Add "OPID", m_udtOrder.lOPKey
        .Add "TranKey", TranNo
        
        .Add "CreateDate", m_udtOrder.dCreateDate
        .Add "UpdateDate", m_udtOrder.dUpdateDate
        
        .Add "UserKey", m_udtOrder.lUserKey
        .Add "UserID", m_udtOrder.sUserID
        
        .Add "WhseKey", m_udtOrder.lWhseKey
        .Add "ShipMethKey", m_udtOrder.lShipMethKey
        .Add "ShipMethod", ShipMethod
        .Add "ShipComplete", m_udtOrder.bShipComplete

        .Add "ShipToNote", Trim(m_udtOrder.sShipToNote)
            
        .Add "PaymentTerms", m_oPmtTerms.ID
        
        .Add "StatusCode", StatusCode
        .Add "PurchOrd", Trim(m_udtOrder.sPurchOrd)
        .Add "Info", Trim(m_udtOrder.sInfo)
        .Add "IsDropShip", m_udtOrder.bDropShip
        .Add "PrintPickList", m_udtOrder.bPrintPickList
        .Add "PricePackList", m_udtOrder.bPricePackList
        .Add "FreeFreight", m_udtOrder.bFreeFreight
        .Add "ReducedFreight", m_udtOrder.bReducedFreight
        .Add "BillDifferentRate", m_udtOrder.bBillDifferentRate
        .Add "InboundFreight", m_udtOrder.bInboundFreight
        .Add "WalkupOrder", m_udtOrder.bWalkup
        .Add "MiscOrder", m_udtOrder.bMisc
        .Add "Deposit", m_udtOrder.bDeposit
        .Add "PartsNoCharge", m_udtOrder.bPartsNoCharge
        .Add "TotalTax", m_oItems.TotalTax(m_oSalesTax.TaxRate)

        .AddChild m_oSalesTax.Export
        
        .Add "TotalPrice", m_oItems.TotalPrice

        'are these nodes required by the consumer of this XML?
        'if the order has a credit card assigned
        If Not m_oCreditCard Is Nothing Then
            .Add "CrCardType", m_oCreditCard.TypeID
            .Add "CrCardNo", m_oCreditCard.MaskedCCNo
            .Add "CrCardHolder", m_oCreditCard.CardHolderName
            .Add "CrCardStreetNbrZip", m_oCreditCard.StreetNbr
            .Add "CrCardExp", m_oCreditCard.ExpireDate
            .Add "CrCardAuth", vbNullString 'm_oCCTransactions.AuthorizationNo
        Else
            .Add "CrCardType", vbNullString
            .Add "CrCardNo", vbNullString
            .Add "CrCardHolder", vbNullString
            .Add "CrCardStreetNbrZip", vbNullString
            .Add "CrCardExp", vbNullString
            .Add "CrCardAuth", vbNullString
        End If
        
        .Add "UPSAcct", Trim(m_udtOrder.sUPSAcct)
        .Add "NoPricing", bNoPricing
        .Add "HasGasket", hasGasket

        If bPubNotes Then ExportRemarks Me.RemarkContext, "PubNotes", "Order.Public", oXML
        If bPrivNotes Then ExportRemarks Me.RemarkContext, "PrivNotes", "Order.Private", oXML
        If bWarehouse Then ExportRemarks Me.RemarkContext, "WarehouseNotes", "Order.Warehouse", oXML
        
        'NOTE: should this go into m_oItems.Export below?
        If bManfNotes Then ExportRemarksEx m_oItems, "ManfNotes", "OrderLine.Manf", oXML

        If Not m_oContact Is Nothing Then
            .AddChild m_oContact.Export
        End If
        
        If Not m_oCustomer Is Nothing Then
            '.Add "VaxAcct", GetVaxAcct(m_oCustomer.BillAddr.AddrKey)
            .AddChild m_oCustomer.Export
        End If
        
        If Not m_oItems Is Nothing Then
            .AddChild m_oItems.Export(bPubNotes, bPrivNotes, bPurchNotes)
        End If
        
    End With
    
    Set Export = oXML
End Function


'****************************************************************


Public Function Import(ByVal i_oXML As JDMPDXML.XMLNode) As Variant
    Dim oCustomerXML As JDMPDXML.XMLNode
    Dim oItemsXML As JDMPDXML.XMLNode
    
    With i_oXML
        If .Tag <> "Order" Then
            Err.Raise -1, "Order.Import", "Unexpected initial XML tag"
        End If
        
        m_udtOrder.lOPKey = .Item("OPKey").value
        m_udtOrder.lSOKey = .Item("SOKey").value
    
        '*** this is some residual refacoring crap that can be cleaned up LR 2/5/15
        On Error Resume Next    'HACK  need to fix this!!!  JJC
        'one of these is gonna fail
        m_udtOrder.lTranNo = 0
        m_udtOrder.lTranNo = .Item("TranKey").value
        m_udtOrder.lTranNo = CLng(Right$(.Item("TranID").value, 10))
        
        m_udtOrder.bPricePackList = .Item("PricePackList").value
        m_udtOrder.dUpdateDate = .Item("UpdateDate").value
        m_udtOrder.bShipComplete = .Item("ShipComplete").value
        
        On Error GoTo 0
        m_udtOrder.dCreateDate = .Item("CreateDate").value
        
        m_udtOrder.lUserKey = .Item("UserKey").value
        m_udtOrder.sUserID = .Item("UserID").value
        
        m_udtOrder.lWhseKey = .Item("WhseKey").value
        m_udtOrder.sWhseID = WhseKeyToID(m_udtOrder.lWhseKey)
        
        m_udtOrder.lShipMethKey = .Item("ShipMethKey").value
        m_udtOrder.sPurchOrd = .Item("PurchOrd").value
        
        On Error Resume Next
        m_udtOrder.sInfo = .Item("Info").value
        
        On Error Resume Next
        m_udtOrder.bDropShip = .Item("IsDropShip").value
        
        m_udtOrder.bFreeFreight = .Item("FreeFreight").value
        m_udtOrder.bReducedFreight = .Item("ReducedFreight").value
        m_udtOrder.bBillDifferentRate = .Item("BillDifferentRate").value
        m_udtOrder.bInboundFreight = .Item("InboundFreight").value
        m_udtOrder.bDeposit = .Item("Deposit").value
        m_udtOrder.bPartsNoCharge = .Item("PartsNoCharge").value
        
        m_udtOrder.bWalkup = .Item("WalkupOrder").value
        m_udtOrder.bMisc = .Item("MiscOrder").value
        m_udtOrder.sUPSAcct = .Item("UPSAcct").value
        
        Select Case CStr(.Item("PrintPickList").value)
            Case "Y", "True"
                m_udtOrder.bPrintPickList = True
            Case Else
                m_udtOrder.bPrintPickList = False
        End Select
        
        m_udtOrder.bARCustHold = False
        m_udtOrder.bCommitting = False
        m_udtOrder.bDeleted = False
        m_udtOrder.bRMA = False
    
        m_udtOrder.bARCustHold = .Item("StatusCode").value = ItemStatusCode.iscARHold
        m_udtOrder.bCommitting = .Item("StatusCode").value = ItemStatusCode.iscPendingCommit
        m_udtOrder.bDeleted = .Item("StatusCode").value = ItemStatusCode.iscDeleted
        m_udtOrder.bRMA = .Item("StatusCode").value = ItemStatusCode.iscHasRMA

        Set oCustomerXML = .Item("Customer")
        If oCustomerXML Is Nothing Then
            Set m_oCustomer = Nothing
        Else
            If m_oCustomer Is Nothing Then
                Set m_oCustomer = New Customer
            End If
            m_oCustomer.Import oCustomerXML
        End If
        
        Set oItemsXML = .Item("Items")
        If oItemsXML Is Nothing Then
            Set m_oItems = Nothing
        Else
            If m_oItems Is Nothing Then
                Set m_oItems = New Items
            End If
            m_oItems.Import oItemsXML, m_udtOrder.lWhseKey
        End If
    End With
End Function


'called only by Order.Load
Public Sub Backup()
    Dim oItem As IItem
    Dim oNewItem As IItem

    m_udtOrder.bUpdateCustomer = False

    m_udtOrderBackup = m_udtOrder
    m_oCustomer.Backup
    m_oItems.Backup
End Sub


Public Function Delete()
    m_udtOrder.bDeleted = True
End Function


'Not used
'Public Function Undelete()
'    m_udtOrder.bDeleted = False
'End Function


'called by FOrder.CancelButton
Public Sub Restore()
    m_udtOrder = m_udtOrderBackup
    m_oCustomer.Restore
    m_oItems.Restore m_udtOrder.lWhseKey
End Sub


'Guessing that the ClearCust Flag allows you to create a new order for the currently loaded Cust
'
'NOTE: The following three fields are not getting cleared.
'  bNewOrder
'  bRecommit
'  sOldSOID
'for a newly instantiated order, these will be False, False, 0 respectively,
'but for a recycled order object, they'll retain old values
'Is this intentional?
'ALSO NOTE: that Order.Create will set bNewOrder = True

Public Sub Clear(ClearCust As Boolean)

    With m_udtOrder
        .lOPKey = 0
        .lSOKey = 0
        .lTranNo = 0
        .dCreateDate = Now()
        .dUpdateDate = .dCreateDate
        .lUpdateCounter = 0
        
        m_bLoading = True
        'call the Let property (supressing its event) to keep key & id synced
        UserKey = GetUserKey
        m_bLoading = False
        
        .lWhseKey = GetUserWhseKey(.lUserKey)
        .sWhseID = GetUserWhseID(.lUserKey)
        
        .lShipMethKey = GetUserShipMethKey(.lUserKey)

        .lDropShipVendKey = 0
        .lBillMethKey = 0

        .sUPSAcct = ""
        .sPurchOrd = ""
        .sInfo = ""
        .eResearchStatus = irsResearchEmpty
        
        .bCommitting = False
        .bDropShip = False
        .bDeleted = False
        .bRMA = False
        .bARCustHold = False
        
        .bShipComplete = False
        .bPricePackList = False
        .bPrintPickList = False
        .bPrintReceipt = False
        .bFreeFreight = False
        .bReducedFreight = False
        .bBillDifferentRate = False
        .bInboundFreight = False
        .bDeposit = False
        .bPartsNoCharge = False
        .bWalkup = False
        .bMisc = False
        .bUpdateCustomer = False
        .lFlagBitmap = 0
        .sShipToName = ""
        .sShipToPhone = ""
        .sShipToNote = ""
    End With

    If ClearCust Then
        m_oCustomer.Clear
    End If

    m_oItems.Clear

    m_oPmtTerms.Clear

    m_oSalesTax.Clear
    
    Set m_oContact = Nothing
    Set m_oCreditCard = Nothing
    
End Sub


'Called By:
'   OfficeAssistant.FOrder.CommitOrder (commit from OP)
'   OfficeAssistant.FChoosePmtTerms.cmdOK_Click (commit from AR)
'
' SOAPClient can throw an error
' -2147024809   WSDLReader:XML Parser

Public Sub Commit()
    Dim oXMLNode As JDMPDXML.XMLNode
    Dim sOrderXML As String
    Dim oProxy As MSSOAPLib30.SoapClient30
    Dim sHeader As String
    
    'The local variable bNew is required to cache m_udtOrder.bNewOrder because
    'SaveOrder alters the value of m_udtOrder.bNewOrder
    'and LogOrderEvent (below) needs to original value.
    Dim bNew As Boolean
    bNew = m_udtOrder.bNewOrder
    
    'Set the order's status to PendingCOmmit before sending it to AutoCommit to avoid the plossibility of a race condition
    SaveOrder ItemStatusCode.iscPendingCommit
    
    LogDB.LogEvent "Order", "Commit", "OP" & OPKey & ": serializing order"
    'serialize the order
    Set oXMLNode = XmlCommit(g_rstWhses, g_rstVendors)
    oXMLNode.IndentWidth = 2
    sOrderXML = oXMLNode.ExportString

'commented out 6/29/15 LR. reimplement if necessary using confgi.xml
'    g_bWriteOrderXml = ConfigINI.Read("OrderCommit", "WriteFile")
'    If g_bWriteOrderXml Then
'        WriteXmlToFile sOrderXML
'    End If

    g_OrderCommitProxy.AddOrder sOrderXML
        
    LogOrderEvent bNew

    'When user clicks on commit directly without saving the order,
    'order remarks and item remarks will be lost.
    'Add the function in Commit to save order and line remarks.
    OrderLogBackUp

End Sub


Private Sub CheckAndNotifySTLGaskets()
    Dim oItem As IItem
    Dim hasGasket As Boolean
    
    For Each oItem In m_oItems
        If oItem.VendorKey = 1148 And m_udtOrder.sWhseID = "STL" Then
            hasGasket = True
        End If
    Next oItem
    
    If hasGasket Then
        EMail.Send "OP@caseparts.com", "joannar@caseparts.com", _
        "Alert: OP-" & m_udtOrder.lOPKey & " has a gasket from Greg's Gasket Works", False 'TextFormat
    End If
End Sub


Private Sub WriteXmlToFile(sOrderXML As String)
    Dim oXMLNode As JDMPDXML.XMLNode
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim sXMLPath As String
    
    sXMLPath = App.path & "\" & Me.OPKey & ".xml"

    Set fs = New FileSystemObject
    Set ts = fs.CreateTextFile(sXMLPath, Overwrite:=True, Unicode:=True)
    ts.Write sOrderXML
    ts.Close
End Sub


'NOTE: the bRecommit param isn't used

Public Function IsDirty(Optional bRecommit As Boolean = False) As Boolean
    IsDirty = True
    
    If m_udtOrder.bNewOrder Then Exit Function

    If FlagBitmap <> m_udtOrderBackup.lFlagBitmap Then Exit Function

    With m_udtOrder
        If .lShipMethKey <> m_udtOrderBackup.lShipMethKey Then Exit Function
        If .lUserKey <> m_udtOrderBackup.lUserKey Then Exit Function
        If .lWhseKey <> m_udtOrderBackup.lWhseKey Then Exit Function

        If .bBillDifferentRate Then
            If .lBillMethKey <> m_udtOrderBackup.lBillMethKey Then Exit Function
        End If

'TODO: look at this
        
        If Trim(.sPurchOrd) <> Trim(m_udtOrderBackup.sPurchOrd) Then Exit Function
        If Trim(.sInfo) <> Trim(m_udtOrderBackup.sInfo) Then Exit Function
        If Trim(.sUPSAcct) <> Trim(m_udtOrderBackup.sUPSAcct) Then Exit Function
        
        '10/02/02 TeddyX: Decide if the order is updated from new customer order to customer order
        If .bUpdateCustomer Then Exit Function
        
        If m_oCustomer.ShipAddr.IsDirty Then Exit Function
        
        If Trim(.sShipToName) <> Trim(m_udtOrderBackup.sShipToName) Then Exit Function
        If Trim(.sShipToPhone) <> Trim(m_udtOrderBackup.sShipToPhone) Then Exit Function
    End With
    
    If IsMemoMeisterDirty Then Exit Function
    
    If IsItemDirty Then Exit Function
    
    If m_oPmtTerms.IsDirty Then Exit Function
    
    IsDirty = False
End Function


Public Sub Create()
    
    Me.Clear ClearCust:=False
    
    'HACK: You would think that we'd want to use tcpSO with GetSurrogateKey.
    'And, ultimately, we do.
    'However, doing the transition stage, it's handy for OP .51 and .52
    'to use the same reference point for assigning OPKeys.
    'After the dust settles, we'll clean this up.
    'm_udtOrder.lOPKey = GetSurrogateKey("tcpSO")
    m_udtOrder.lOPKey = database.GetSurrogateKey("tcpOrder")
    
    m_udtOrder.bNewOrder = True
    
    m_oRemarkContext.Load "ViewOrder", CStr(m_udtOrder.lOPKey)
    
    LogOAEvent "Order", GetUserID, m_udtOrder.lOPKey, , StatusCode, "Created."
End Sub


Public Function Summary() As String
    Dim oItem As IItem
    Dim sOutput As String

    For Each oItem In m_oItems
        sOutput = sOutput & ItemSummary(oItem)
    Next
    Summary = sOutput
End Function


Public Function ItemSummary(oItem As IItem) As String
    Dim sType As String
    With oItem
        Select Case .OPItemType
        Case itFinishedGood, itBTOKit: sType = ""
        Case itMoldedGasket: sType = "GSKT:"
        Case itWarmerWire: sType = "WWIR:"
        Case itWireShelf: sType = "WSHF:"
        Case itSpecialOrder: sType = "SPO:"
        Case Else: sType = "???:"
        End Select
        ItemSummary = "(" & .Qty & ") " & sType & .ItemID & " " & .Descr & vbCrLf
    End With
End Function


' called by SaveOrder
' Should be create a StringBuilder?

Public Function Keywords() As String
    Dim oItem As IItem
    Dim sItemKeywords As String
    Dim sType As String
    Dim sOutput As String
    
    With m_udtOrder
        sOutput = ExtractRemarks(Me.RemarkContext) & Trim$(.sPurchOrd) & " "
        If Len(.sInfo) > 0 Then
            sOutput = sOutput & Trim$(.sInfo) & " "
        End If
        If Len(.sShipToNote) > 0 Then
            sOutput = sOutput & Trim$(.sShipToNote) & " "
        End If
    End With
    
    If Not m_oContact Is Nothing Then
        sOutput = sOutput & contact.Name & " " & contact.Phone & " " & contact.CellPhone & " "
    End If
    
    With m_oCustomer
        sOutput = sOutput & .ID & " " & .Name & " "
    End With
    
    With m_oCustomer.BillAddr
        sOutput = sOutput & .AddrName & " " & .Addr1 & " " & .Addr2 & " " & .City & " " & .State & " " & .Zip & " "
    End With
    
    With m_oCustomer.ShipAddr
        sOutput = sOutput & .AddrName & " " & .Addr1 & " " & .Addr2 & " " & .City & " " & .State & " " & .Zip & " "
    End With
    
    For Each oItem In m_oItems
        With oItem
            sItemKeywords = .ItemID & " " & .Descr & " " & .ModelNbr & " " & ExtractRemarks(oItem.RemarkContext) & .SerialNbr & " "
        End With
        sOutput = sOutput & sItemKeywords
    Next

    Keywords = sOutput
End Function



'************************************************************************************
'Private Functions
'************************************************************************************

'Called by:
'   Private Property Let FlagBitmap

Private Sub ParseFlag()
    With m_udtOrder
        .bDropShip = ((.lFlagBitmap And &H1) = &H1)
        .bShipComplete = ((.lFlagBitmap And &H2) = &H2)
        .bPrintPickList = ((.lFlagBitmap And &H4) = &H4)
        .bPricePackList = ((.lFlagBitmap And &H8) = &H8)
        .bFreeFreight = ((.lFlagBitmap And &H10) = &H10)
        .bReducedFreight = ((.lFlagBitmap And &H20) = &H20)
        .bPartsNoCharge = ((.lFlagBitmap And &H40) = &H40)
        .bInboundFreight = ((.lFlagBitmap And &H100) = &H100)
        .bWalkup = ((.lFlagBitmap And &H200) = &H200)
        .bMisc = ((.lFlagBitmap And &H400) = &H400)
        .bDeposit = ((.lFlagBitmap And &H800) = &H800)
        .bBillDifferentRate = ((.lFlagBitmap And &H4000) = &H4000)
    End With
End Sub

'Called by:
'   Private Property Get FlagBitmap() As Long

'Teddy: replace the sub routines with get/set for the FlagBitmap property.
Private Function GetFlag() As Long
    Dim lFlag As Long
    
    With m_udtOrder
        If .bDropShip Then
            lFlag = lFlag Or &H1
        End If
        
        If .bShipComplete Then
            lFlag = lFlag Or &H2
        End If
        
        If .bPrintPickList Then
            lFlag = lFlag Or &H4
        End If
        
        If .bPricePackList Then
            lFlag = lFlag Or &H8
        End If
        
        If .bFreeFreight Then
            lFlag = lFlag Or &H10
        End If
        
        If .bReducedFreight Then
            lFlag = lFlag Or &H20
        End If
        
        If .bBillDifferentRate Then
            lFlag = lFlag Or &H4000
        End If
        
        If .bPartsNoCharge Then
            lFlag = lFlag Or &H40
        End If
        
        If .bInboundFreight Then
            lFlag = lFlag Or &H100
        End If
        
        If .bWalkup Then
            lFlag = lFlag Or &H200
        End If
        
        If .bMisc Then
            lFlag = lFlag Or &H400
        End If
        
        If .bDeposit Then
            lFlag = lFlag Or &H800
        End If
        
        .lFlagBitmap = lFlag
        GetFlag = lFlag
    End With
End Function


'*******************************************************************
'The following functions are used to determine if the order is dirty
'*******************************************************************

'Called By IsDirty()

Private Function IsItemDirty() As Boolean
    Dim oItem As IItem
    Dim sLineChanges As String
    
    IsItemDirty = True
    For Each oItem In m_oItems
        sLineChanges = EventOldOrderLine(oItem)
        If Trim(sLineChanges) <> "" Then Exit Function
    Next
    
    If isItemAllDelete Then Exit Function
    If isItemPartialDelete Then Exit Function
    IsItemDirty = False
End Function


'Called By IsItemDirty()

Private Function isItemPartialDelete() As Boolean
    If m_oBackUpItems Is Nothing Then Exit Function
    If m_oBackUpItems.Count = 0 Then Exit Function

    Dim oItem As IItem
    Dim oNewItem As IItem
    Dim bItemDel As Boolean
    
    isItemPartialDelete = True
    For Each oItem In m_oBackUpItems
        bItemDel = True
        For Each oNewItem In m_oItems
            If oNewItem.OPLineKey = oItem.OPLineKey Then
                bItemDel = False
                Exit For
            End If
        Next
            
        If bItemDel Then Exit Function
    Next
    isItemPartialDelete = False
End Function


'Called By IsItemDirty()

Private Function isItemAllDelete() As Boolean
    If m_oBackUpItems Is Nothing Then Exit Function
    If m_oBackUpItems.Count = 0 Then Exit Function
    
    isItemAllDelete = True
    If m_oItems Is Nothing Then Exit Function
    If m_oItems.Count = 0 Then Exit Function
    isItemAllDelete = False
End Function


'Called By IsDirty()
'Doesn't MM save itself?
'We're allowed to create MM remarks when the order is ReadOnly

Private Function IsMemoMeisterDirty() As Boolean
    Dim oItem As IItem

    If m_oRemarkContext.RemarkList.Dirty Then
        IsMemoMeisterDirty = True
        Exit Function
    Else
        For Each oItem In m_oItems
            If oItem.RemarkContext.RemarkList.Dirty Then
                IsMemoMeisterDirty = True
                Exit For
            End If
        Next
    End If
End Function


'*******************************************************************
' Save logic
'*******************************************************************

'Called by
'   Save
'   Commit

Private Sub SaveOrder(ByVal oStatus As ItemStatusCode, _
                    Optional ByVal bCommitOrder As Boolean = False)

'    LogDB.LogEvent "Order", "SaveOrder", "OP" & OPKey & ": entering saveorder, ostatus=" & oStatus & ", commitorder=" & bCommitOrder
    
    On Error GoTo ErrorHandler
    Dim cmd As ADODB.Command
    
    If Not bCommitOrder Then
        'Execute everything within a transaction for saving order. If saving order
        'is executed in order.commit, transaction is not needed because order.commit
        'already has its own transaction.
        g_DB.Connection.BeginTrans
    End If
    
    '??? Why are we doing this?
    If m_udtOrder.bDeleted And m_oItems.Count = 0 Then
        'Permanently delete the order if it contains no items
        Set cmd = CreateCommandSP("Delete tcpSO where OPKey = " & m_udtOrder.lOPKey, adCmdText)
        cmd.Execute
        Set cmd = Nothing
        
    'This is the nominal case
    Else
    
    'TODO: combine these two subroutines
'        LogDB.LogEvent "Order", "SaveOrder", "OP" & OPKey & ": calling SaveOrderInfo(" & oStatus & ")"
        SaveOrderInfo oStatus
        
'        LogDB.LogEvent "Order", "SaveOrder", "OP" & OPKey & ": calling SaveOrderInfo"
        SaveOrderAdditionalInfo

        'Save order line items info
        If Not bRMA Then
'            LogDB.LogEvent "Order", "SaveOrder", "OP" & OPKey & ": calling SaveOrderLineItems"
            SaveOrderLineItems
        End If
        
'If you save more than once after loading, you need to keep our update count
'in sync with the database table. bNewOrder means the order has never been persisted.
        If Not m_udtOrder.bNewOrder Then
            LogDB.LogEvent "Order", "SaveOrder", "OP" & OPKey & ": incr update counter"
            m_udtOrder.lUpdateCounter = m_udtOrder.lUpdateCounter + 1
        End If
        
        m_udtOrder.bNewOrder = False
        
    End If
        
    Set cmd = Nothing
    
    If Not bCommitOrder Then
        g_DB.Connection.CommitTrans
    End If
    Exit Sub

ErrorHandler:

    LogDB.LogEvent "Order", "SaveOrder", "OP" & OPKey & ": Error" & Err.Number & " " & Err.Description
    
    If Not bCommitOrder Then
        LogDB.LogEvent "Order", "SaveOrder", "OP" & OPKey & ": rolling back commit transaction"
        g_DB.Connection.RollbackTrans
    End If
'    m_udtOrder.bNewOrder = bNewOrder  '?????
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


'Called by Save()
'Save 'lite-edits' to the Sage order

Private Sub SaveCommittedOrder()
    On Error GoTo ErrorHandler
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spCPCsoUpdateCommittedOrder")
    
    With cmd
        .Parameters("@_iOPKey") = m_udtOrder.lOPKey
        .Parameters("@_iSOKey") = m_udtOrder.lSOKey
        .Parameters("@_iPurchOrd") = m_udtOrder.sPurchOrd
        .Parameters("@_iShipMethKey") = m_udtOrder.lShipMethKey
        .Parameters("@_iBillMethKey") = m_udtOrder.lBillMethKey
        
        '6/12/15 LR needs a mod to spCPCsoUpdateCommittedOrder
        .Parameters("@_iPmtTermsKey") = m_oPmtTerms.Key
        
        .Parameters("@_iFlags") = FlagBitmap
        If m_udtOrder.sUPSAcct <> "" Then
            .Parameters("@_iUPSAcct") = m_udtOrder.sUPSAcct
        End If
        
        If m_udtOrder.sShipToName <> "" Then
            .Parameters("@_sShipToName") = m_udtOrder.sShipToName
        End If
        If m_udtOrder.sShipToPhone <> "" Then
            .Parameters("@_sShipToPhone") = m_udtOrder.sShipToPhone
        End If
        
        .Execute
    End With
    
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "SaveCommittedOrder", Err.Description
End Sub


'Called by SaveOrder()

Private Sub SaveOrderInfo(ByVal oStatus As ItemStatusCode)
    On Error GoTo ErrorHandler
    
    Dim cmd As ADODB.Command
    Dim iRetVal As Integer
    Dim bSPError As Boolean

'    LogDB.LogEvent "Order", "SaveOrderInfo", "OP" & OPKey & ": NewOrder=" & m_udtOrder.bNewOrder

    If m_udtOrder.bNewOrder Then
        Set cmd = CreateCommandSP("cpopOrderInsert")
    Else
        Set cmd = CreateCommandSP("cpopOrderUpdate")
    End If
    
    With cmd
        .Parameters("@_iOPKey").value = m_udtOrder.lOPKey
        .Parameters("@_iUpdateCounter").value = m_udtOrder.lUpdateCounter
        .Parameters("@_iSOKey").value = m_udtOrder.lSOKey
        .Parameters("@_iTranKey").value = m_udtOrder.lTranNo
        If m_oCreditCard Is Nothing Then
            .Parameters("@_iCCKey").value = Null
        Else
            .Parameters("@_iCCKey").value = m_oCreditCard.Key
        End If
        .Parameters("@_iStatusCode").value = oStatus
        
        .Parameters("@_iUserID").value = m_udtOrder.sUserID
        
        .Parameters("@_iWhseKey").value = m_udtOrder.lWhseKey
        .Parameters("@_iCustKey").value = m_oCustomer.Key

        .Parameters("@_iPmtTermsKey").value = m_oPmtTerms.Key
        
        .Parameters("@_iPurchOrd").value = m_udtOrder.sPurchOrd
        .Parameters("@_iInfo").value = m_udtOrder.sInfo

        If Not m_oContact Is Nothing Then
            .Parameters("@_iCntctKey").value = m_oContact.Key
        Else
            .Parameters("@_iCntctKey").value = Null
        End If

        .Parameters("@_iBillAddrKey").value = m_oCustomer.BillAddr.AddrKey
        .Parameters("@_iShipAddrKey").value = m_oCustomer.ShipAddr.AddrKey
        .Parameters("@_iCustName").value = m_oCustomer.Name
        .Parameters("@_iCustType").value = m_oCustomer.CustType
        
        'Temp/New cust order
        If m_oCustomer.Key = 0 Then
            .Parameters("@_iCustXML").value = m_oCustomer.Export.ExportString
            'Saving for the first time
            If m_udtOrder.bNewOrder Then
                'Send email notification to person who creates new accounts
                'This will be implemented in a later release.
                'SendEmailNotification
            End If
        Else
            .Parameters("@_iCustXML").value = ""
        End If
        
        If m_oCustomer.BillAddr.AddrKey = 0 And m_oCustomer.Key <> 0 Then
            .Parameters("@_iBillAddrXML").value = m_oCustomer.BillAddr.Export.ExportString
        Else
            .Parameters("@_iBillAddrXML").value = ""
        End If
        If m_oCustomer.ShipAddr.AddrKey = 0 And m_oCustomer.Key <> 0 Then
            .Parameters("@_iShipAddrXML").value = m_oCustomer.ShipAddr.Export.ExportString
        Else
            .Parameters("@_iShipAddrXML").value = ""
        End If
        
        .Parameters("@_iShipMethKey").value = m_udtOrder.lShipMethKey
        .Parameters("@_iTaxStatusCode").value = (m_oSalesTax.TaxRate <> 0)
        .Parameters("@_iReqPO").value = m_oCustomer.ReqPO
        
        .Parameters("@_iKeywords").value = PrepSQLText(Keywords)
        
        .Parameters("@_iSummary").value = Summary
        .Parameters("@_iDropShipVendKey").value = m_udtOrder.lDropShipVendKey
        
        .Execute

'!!! Is testing the RetVal against SageErrType appropriate?

'        LogDB.LogEvent "Order", "SaveOrderInfo", "OP" & OPKey & ": Insert/Update RetVal=" & iRetVal
        
        iRetVal = .Parameters("@_oRetVal").value
        If iRetVal <> SageErrType.aetNone Then
            'HACK: need better error handling
            'Teddy: We need to check for this condition much earlier in the processing of a Save or
            'Commit command.  Consider adding ConfirmFreshData method that will increment the version
            'number if it matches what we have in memory, then update both the disk and memory versions
            'so that they match.  Calls to ConfirmFreshData will be needed in higher level functions.
            'appropriate.
            bSPError = True
            GoTo ErrorHandler
        End If
    End With
    
'    LogDB.LogEvent "Order", "SaveOrderInfo", "OP" & OPKey & ": SaveOrderInfo exiting OK"
    
    Exit Sub
    
ErrorHandler:
    LogDB.LogEvent "Order", "SaveOrderInfo", "OP" & OPKey & ": SaveOrderInfo Error " & Err.Description
    
    If bSPError Then
        Err.Raise -1, "SaveOrderInfo", "Another user has updated this order since you loaded it. Reopen the order in another window and copy in your changes"
    Else
        Err.Raise Err.Number, "SaveOrderInfo", Err.Description
    End If
End Sub


'Called by SaveOrder()
'This subroutine save order line items information to database

Private Sub SaveOrderLineItems()
    On Error GoTo ErrorHandler
    
    Dim cmd As ADODB.Command
    Dim bSPError As Boolean
    Dim oItem As IItem
    Dim sLineInfo As String
    Dim lLineNo As Long
    Dim iRetVal As Integer
    
    Set cmd = CreateCommandSP("Delete tcpSOLine where OPKey = " & m_udtOrder.lOPKey, adCmdText)
    cmd.Execute
    
    Set cmd = CreateCommandSP("spCPCsoLineInsert1")
    With cmd
        For Each oItem In m_oItems
            lLineNo = lLineNo + 1
            .Parameters("@_iOPKey").value = m_udtOrder.lOPKey
            .Parameters("@_iLineNo").value = lLineNo
            
            'these keys already exist in memory
            .Parameters("@_iOPLineKey").value = oItem.OPLineKey
            .Parameters("@_iSOLineKey").value = oItem.LineKey
            
            'there's more to an item than the tcpSOLine columns support
            .Parameters("@_iLineXML").value = oItem.Export.ExportString
            
            .Parameters("@_iOPItemType").value = oItem.OPItemType
            .Parameters("@_iItemKey").value = oItem.ItemKey
            .Parameters("@_iItemID").value = oItem.ItemID
            .Parameters("@_iVendKey").value = oItem.VendorKey
            .Parameters("@_iMakeKey").value = oItem.MakeKey
            .Parameters("@_iCabModelNbr").value = oItem.ModelNbr
            .Parameters("@_iCabSerialNbr").value = oItem.SerialNbr
            .Parameters("@_iModModelNbr").value = ScrubText(oItem.ModelNbr)
            .Parameters("@_iModSerialNbr").value = ScrubText(oItem.SerialNbr)
            .Parameters("@_iCost").value = oItem.Cost
            .Parameters("@_iQty").value = oItem.Qty
            .Parameters("@_iQtyReturned").value = 0
            .Parameters("@_iDescr").value = oItem.Descr
            .Execute
            iRetVal = .Parameters("@_oRetVal").value
            ' if the insert failed
            If iRetVal = 0 Then
                bSPError = True
                sLineInfo = Trim(oItem.Descr)
                If Trim(oItem.ItemID) <> "" Then
                    sLineInfo = Trim(oItem.ItemID) & ": " & sLineInfo
                End If
                
                GoTo ErrorHandler
            End If
        Next
    End With
    
    Exit Sub
    
ErrorHandler:
    If bSPError Then
        Err.Raise -1, "SaveOrderLineItems", "Could not save line item " & sLineInfo & " of this order in spCPCsoLineInsert1."
    Else
        Err.Raise Err.Number, "SaveOrderLineItems", Err.Description
    End If
End Sub


'Called by SaveOrder()
'This subroutine saves order's additional information

Private Sub SaveOrderAdditionalInfo()
    On Error GoTo ErrorHandler
    
    Dim cmd As ADODB.Command
    
    Set cmd = CreateCommandSP("spCPCsoUpdateAdditionalInfo")
    With cmd
        .Parameters("@_iOPKey") = m_udtOrder.lOPKey
        .Parameters("@_iTotalPrice") = m_oItems.TotalPrice
        .Parameters("@_iTotalTax") = m_oItems.TotalTax(m_oSalesTax.TaxRate)
        .Parameters("@_iResearchStatus") = ResearchStatus
        .Parameters("@_iFlags") = FlagBitmap
        
        If m_udtOrder.bBillDifferentRate Then
            .Parameters("@_iBillMethKey") = m_udtOrder.lBillMethKey
        End If
        
        If m_udtOrder.sUPSAcct <> "" Then
            .Parameters("@_iUPSAcct") = m_udtOrder.sUPSAcct
        End If

        If m_udtOrder.sShipToName <> "" Then
            .Parameters("@_sShipToName") = m_udtOrder.sShipToName
        End If
        If m_udtOrder.sShipToPhone <> "" Then
            .Parameters("@_sShipToPhone") = m_udtOrder.sShipToPhone
        End If
        
        If m_udtOrder.sShipToNote <> "" Then
            .Parameters("@_sShipToNote") = m_udtOrder.sShipToNote
        End If
        
        .Execute
    End With
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "SaveOrderAdditionalInfo", Err.Description
End Sub


Private Function ImportString(sXML As String) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    
    Set oXML = New JDMPDXML.XMLNode
    oXML.ImportString sXML
    
    Set ImportString = oXML
End Function


Private Function hasGasket() As Boolean
    Dim oItem As IItem
    
    For Each oItem In m_oItems
        If oItem.OPItemType = itMoldedGasket Then
            hasGasket = True
            Exit For
        End If
    Next
End Function


Private Sub ExportRemarks(RC As RemarkContext, sXmlTag As String, sRemarkTypeID As String, oXML As JDMPDXML.XMLNode)
    Dim sRemarks As String
    Dim oRemark As remark
    
    If RC Is Nothing Then Exit Sub
    If RC.RemarkList Is Nothing Then Exit Sub
    If RC.RemarkList.Count = 0 Then Exit Sub
    
    For Each oRemark In RC.RemarkList
        If oRemark.RemarkType.TypeID = sRemarkTypeID Then
            sRemarks = sRemarks & oRemark.MemoText & vbCrLf
        End If
    Next
    
    If Len(sRemarks) > 0 Then
        oXML.Add sXmlTag, sRemarks
    End If
End Sub

Private Sub ExportRemarksEx(voItems As Items, sXmlTag As String, sRemarkTypeID As String, oXML As JDMPDXML.XMLNode)
    Dim oItem As IItem
    Dim sRemarks As String
    Dim oRemark As remark

    For Each oItem In voItems
        If oItem.RemarkContext Is Nothing Then Exit Sub
        If oItem.RemarkContext.RemarkList Is Nothing Then Exit Sub
        If oItem.RemarkContext.RemarkList.Count = 0 Then Exit Sub
        For Each oRemark In oItem.RemarkContext.RemarkList
            If oRemark.RemarkType.TypeID = sRemarkTypeID Then
                sRemarks = sRemarks & oRemark.MemoText & vbCrLf
            End If
        Next
    Next

    If Len(sRemarks) > 0 Then
        oXML.Add sXmlTag, sRemarks
    End If

End Sub


Private Sub ItemCopyGasket(ByRef oNewItem As IItem, ByRef oItem As IItem)
     Dim oTempItem1 As ItemGasket
     Dim otempItem2 As ItemGasket
     Set oTempItem1 = oNewItem
     Set otempItem2 = oItem
     
     With oTempItem1
        .Height = otempItem2.Height
        .width = otempItem2.width
        .Options = otempItem2.Options
        .materialId = otempItem2.materialId
        .IsMagnetic = otempItem2.IsMagnetic
    End With
End Sub


Private Sub ItemCopyShelf(ByRef oNewItem As IItem, ByRef oItem As IItem)
     Dim oTempItem1 As ItemShelf
     Dim otempItem2 As ItemShelf
     Set oTempItem1 = oNewItem
     Set otempItem2 = oItem
     
     With oTempItem1
        .Depth = otempItem2.Depth
        .width = otempItem2.width
        .Options = otempItem2.Options
        .FrameID = otempItem2.FrameID
        .FinishID = otempItem2.FinishID
    End With
End Sub


Private Sub ItemCopyWire(ByRef oNewItem As IItem, ByRef oItem As IItem)
     Dim oTempItem1 As ItemWWire
     Dim otempItem2 As ItemWWire
     Set oTempItem1 = oNewItem
     Set otempItem2 = oItem
     
     With oTempItem1
        .Inches = otempItem2.Inches
        .Feet = otempItem2.Feet
        .Voltage = otempItem2.Voltage
        .DoorHeight = otempItem2.DoorHeight
        .DoorWidth = otempItem2.DoorWidth
        .IsSinglePass = otempItem2.IsSinglePass
        .IsThreeSided = otempItem2.IsThreeSided
        .TotalInches = otempItem2.TotalInches
        .OhmsPerFoot = otempItem2.OhmsPerFoot
    End With
End Sub


Private Sub ItemCopy(ByRef oNewItem As IItem, ByRef oItem As IItem)
    If oNewItem.OPItemType = itMoldedGasket Then
        ItemCopyGasket oNewItem, oItem
    ElseIf oNewItem.OPItemType = itWireShelf Then
        ItemCopyShelf oNewItem, oItem
    ElseIf oNewItem.OPItemType = itWarmerWire Then
        ItemCopyWire oNewItem, oItem
    End If
            
    With oNewItem
        .Cost = oItem.Cost
        .CustType = oItem.CustType
        .DealerPrice = oItem.DealerPrice
        .IsCGMPN = oItem.IsCGMPN
        .IsTaxable = oItem.IsTaxable
        .ItemID = oItem.ItemID
        .ItemKey = oItem.ItemKey
        .LineKey = oItem.LineKey
        .ListPrice = oItem.ListPrice
        .MakeKey = oItem.MakeKey
        .ModelNbr = oItem.ModelNbr
        .NegotiatedPrice = oItem.NegotiatedPrice
        .OPLineKey = oItem.OPLineKey
        .OPKey = oItem.OPKey
        .Qty = oItem.Qty
        .SerialNbr = oItem.SerialNbr
        .VendorKey = oItem.VendorKey
        .WholesalePrice = oItem.WholesalePrice
        .BackNegotiatedPrice = oItem.BackNegotiatedPrice
        .SageItemType = oItem.SageItemType
        .Descr = oItem.Descr
        .StatusCode = oItem.StatusCode
    End With
End Sub


Private Function ExtractRemarks(i_oContext As RemarkContext) As String
    Dim oRemark As remark
    Dim str As String
    
    'We have some errors in error log that passed in i_oContext is nothing
    If i_oContext Is Nothing Then Exit Function
    If i_oContext.RemarkList Is Nothing Then Exit Function
    If i_oContext.RemarkList.Count = 0 Then Exit Function
    
    For Each oRemark In i_oContext.RemarkList
        str = str + oRemark.MemoText & " "
    Next
    ExtractRemarks = str
End Function


'Private Function GetOAErrMsg()
'    GetOAErrMsg = "* Order information in SageAssistant" & vbCrLf & _
'                vbTab & "Order:" & vbTab & vbTab & m_udtOrder.lOPKey & vbCrLf & _
'                vbTab & "CSR:" & vbTab & vbTab & m_udtOrder.sUserID & vbCrLf & _
'                vbTab & "Customer:" & vbTab & "[" & m_oCustomer.ID & "] " & m_oCustomer.Name & vbCrLf
'
'
'    If m_udtOrder.bDropShip Then
'        GetOAErrMsg = GetOAErrMsg & vbCrLf & "* This is a dropship order" & vbCrLf
'    End If
'
'    GetOAErrMsg = GetOAErrMsg & vbCrLf
'
'End Function

Private Sub m_oSalesTax_Changed()
    m_oItems.IsTaxable = (m_oSalesTax.TaxRate <> 0)
End Sub


'Private Function GetVaxAcct(lAddrKey As Long) As String
'    Dim rstVaxAcct As ADODB.Recordset
'    Dim sSQL As String
'
'    sSQL = "Select VaxAcct from tcpVaxAcct where addrkey = " & lAddrKey
'    Set rstVaxAcct = LoadDiscRst(sSQL)
'
'    If rstVaxAcct.RecordCount > 0 Then
'        GetVaxAcct = Trim(rstVaxAcct.Fields("VaxAcct").value)
'    End If
'
'    Set rstVaxAcct = Nothing
'End Function


'****************************************************************************
' Here's all the Event Logging stuff
'****************************************************************************

'LogOAEvent is defined in LogEvent.bas

'Called by
'   Save()
'   Commit()

Private Sub LogOrderEvent(bNewOrder As Boolean)
    LogSalesOrderEvent bNewOrder
    LogOrderLineEvent bNewOrder
End Sub


'Called by
'   LogOrderEvent()
'
'All of the subroutines prefaced by "Event" are string builders.

Private Sub LogSalesOrderEvent(bNewOrder As Boolean)
    Dim sSalesOrderEvent As String
    Dim sEventID As String
    Dim sNewStrEvent As String
    Dim sNewMMEvent As String
    Dim sNewBooleanEvent As String
    
    If StatusCode = ItemStatusCode.iscHasRMA Then
        sEventID = "RMA"
    Else
        sEventID = "Order"
    End If
    
    'PRN#96
    If bNewOrder Then
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventNewOrder
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventOrderUpdateCustomer
'        sSalesOrderEvent = EventNewOrder & EventOrderUpdateCustomer
    Else
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventOldOrder
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventOrderUpdateCustomer
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventShipMethod
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventCSR
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventWarehouse
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventPmtTerms
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventUPSAcct
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventPurchOrd
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventInfo
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventOrderMM
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventOrderLineMM
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventDropShip
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventShipComplete
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventPrintPickList
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventPricePackList
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventFreeFreight
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventReducedFreight
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventBillDifferentRate
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventInboundFreight
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventPartsNoCharge
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventDeposit
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventShipToName
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventShipToPhone
        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventShipToNote

'***CREDITCARD
'This would be handled in the CreditCard Editor
'        LogOAEvent sEventID, GetUserID, m_udtOrder.lOPKey, , StatusCode, EventCrCard
    
    End If
End Sub


'Called by
'   LogOrderEvent()

Private Sub LogOrderLineEvent(bNewOrder As Boolean)
    Dim oItem As IItem
    Dim sEventOrderLine As String
    
    For Each oItem In m_oItems
        If bNewOrder Then
            sEventOrderLine = EventNewOrderLine(oItem)
        Else
            sEventOrderLine = EventOldOrderLine(oItem)
            'PRN#96
            LogOrderLineDelEvent
        End If
        
        If Trim(sEventOrderLine) <> "" Then
            'PRN#96
            LogOAEvent "OrderLine", GetUserID, oItem.OPLineKey, m_udtOrder.lOPKey, oItem.StatusCode, sEventOrderLine
        End If
    Next
End Sub


Private Sub LogOrderLineDelEvent()
    Dim oItem As IItem
    Dim oNewItem As IItem
    Dim sTemp As String
    Dim bItemDel As Boolean
    
    'If m_bIsNew Then Exit Sub
    
    If m_oBackUpItems Is Nothing Then
        Exit Sub
    ElseIf m_oBackUpItems.Count = 0 Then
        Exit Sub
    End If
    
    If m_oItems Is Nothing Then
        For Each oItem In m_oBackUpItems
                sTemp = "Delete Item " & EventItemID(oItem) & " from order."
                LogOAEvent "OrderLine", GetUserID, oItem.OPLineKey, m_udtOrder.lOPKey, iscDeleted, sTemp
        Next
    ElseIf m_oItems.Count = 0 Then
        For Each oItem In m_oBackUpItems
                sTemp = "Delete Item " & EventItemID(oItem) & " from order."
                LogOAEvent "OrderLine", GetUserID, oItem.OPLineKey, m_udtOrder.lOPKey, iscDeleted, sTemp
        Next
    Else
        For Each oItem In m_oBackUpItems
            bItemDel = True
            For Each oNewItem In m_oItems
                If oNewItem.OPLineKey = oItem.OPLineKey Then
                    bItemDel = False
                    Exit For
                End If
            Next
            
            If bItemDel Then
                sTemp = "Delete Item " & EventItemID(oItem) & " from order."
                LogOAEvent "OrderLine", GetUserID, oItem.OPLineKey, m_udtOrder.lOPKey, iscDeleted, sTemp
            End If
        Next
    End If
End Sub


'Called by
'   Save()
Private Sub LogCommittedOrderEvent()
    LogOAEvent "Order", GetUserID, m_udtOrder.lOPKey, , StatusCode, EventShipMethod
    LogOAEvent "Order", GetUserID, m_udtOrder.lOPKey, , StatusCode, EventPurchOrd
    LogOAEvent "Order", GetUserID, m_udtOrder.lOPKey, , StatusCode, EventShipComplete
    LogOAEvent "Order", GetUserID, m_udtOrder.lOPKey, , StatusCode, EventUPSAcct
End Sub


'Called By:
'   Save()
'   Commit()

Private Sub OrderLogBackUp()
    Dim rstline As ADODB.Recordset
    Dim oItem As IItem
    
'    LogDB.LogEvent "Order", "OrderLogBackUp", "OP" & OPKey
        
    Set rstline = LoadDiscRst("SELECT * FROM tcpSOLine WHERE OPKey=" & m_udtOrder.lOPKey & " ORDER BY [LineNo]")

    'Since everything worked, go ahead an save the remarks.
    'This is not part of the transaction because MM is using a different connection object.
    m_oRemarkContext.Save True
    For Each oItem In m_oItems
        oItem.RemarkContext.Save True
    Next
    
    m_oBackUpItems.Clear
    m_oItems.Clear
    If Not rstline.EOF Then
        Do While Not rstline.EOF
            m_oItems.ImportItem ImportString(rstline.Fields("LineXML").value), m_udtOrder.lWhseKey
            m_oBackUpItems.ImportItem ImportString(rstline.Fields("LineXML").value), m_udtOrder.lWhseKey
            rstline.MoveNext
        Loop
    End If
    Backup
    m_oBackUpStatusCode = StatusCode
    
    Set rstline = Nothing
End Sub


'Called By:
'   Save()

Private Sub LogSageMMEvent()
    Dim sNewMMEvent As String
    
    sNewMMEvent = EventOrderMM & EventOrderLineMM
    If sNewMMEvent <> "" Then
        sNewMMEvent = Right(sNewMMEvent, Len(sNewMMEvent) - 2)
        LogOAEvent "Order", GetUserID, m_udtOrder.lOPKey, , StatusCode, sNewMMEvent
    End If
End Sub


'***************************************************************************************
'Event "string builder" functions
'***************************************************************************************

Private Function EventShipMethod() As String
    Dim sNewShipMethID As String
    Dim sOldShipMethID As String
    
    If m_udtOrder.lShipMethKey <> m_udtOrderBackup.lShipMethKey Then
        sNewShipMethID = ShipMethKeytoID(m_udtOrder.lShipMethKey)
        sOldShipMethID = ShipMethKeytoID(m_udtOrderBackup.lShipMethKey)
        
        If sOldShipMethID <> "" Then
            EventShipMethod = "Changed ShipMethod from '" & sOldShipMethID & "' to '" & sNewShipMethID & "'"
        Else
            EventShipMethod = "Ship Method is '" & sNewShipMethID & "'"
        End If
    End If
End Function


'Called By LogSalesOrderEvent()

Private Function EventPmtTerms() As String

    If m_oPmtTerms.ID <> m_oPmtTerms.PreviousID Then
        If m_oPmtTerms.PreviousID <> vbNullString Then
            EventPmtTerms = "Changed PmtTerms from " & m_oPmtTerms.PreviousID & " to " & m_oPmtTerms.ID
        Else
            EventPmtTerms = "PmtTerms are " & m_oPmtTerms.ID
        End If
    End If
    
End Function


'Called By
'    Save or Commit
'       LogOrderEvent
'           LogSalesOrderEvent
'
Private Function EventCSR() As String
    Dim sNewCSR As String
    Dim sOldCSR As String
    Dim lCSRWhseKey As Long
    Dim lOldCSRWhseKey As Long

'If the order's UserKey has changed
    If m_udtOrder.lUserKey <> m_udtOrderBackup.lUserKey Then
        sNewCSR = m_udtOrder.sUserID
        sOldCSR = m_udtOrderBackup.sUserID
        EventCSR = "Changed CSR from '" & sOldCSR & "' to '" & sNewCSR & "'"
        
        lCSRWhseKey = GetUserWhseKey(m_udtOrder.lUserKey)
        lOldCSRWhseKey = GetUserWhseKey(m_udtOrderBackup.lUserKey)
        
'if either the original or new CSR is in STL, inform Bob
'***TODO: pull mailing list and msg out to config file
        If ((lCSRWhseKey = g_STLWhseKey) Or (lOldCSRWhseKey = g_STLWhseKey)) Then

            EMail.Send "OP@caseparts.com", "bobg@caseparts.com", _
                "Alert: On OP-" & m_udtOrder.lOPKey & " CSR was changed from " & sOldCSR & " to " & sNewCSR, _
                vbNullString, False 'TextFormat

        End If
        
    End If
End Function


Private Function EventWarehouse() As String
    Dim sNewWhse As String
    Dim sOldWhse As String
    
    If m_udtOrder.lWhseKey <> m_udtOrderBackup.lWhseKey Then
        With g_rstWhses
            .Filter = "WhseKey = " & m_udtOrder.lWhseKey
                sNewWhse = Trim(.Fields("WhseID").value)
            .Filter = adFilterNone
            
            If m_udtOrderBackup.lWhseKey <> 0 Then
                .Filter = "WhseKey = " & m_udtOrderBackup.lWhseKey
                    sOldWhse = Trim(.Fields("WhseID").value)
                .Filter = adFilterNone
            End If
        End With
        
        If sOldWhse <> "" Then
            EventWarehouse = "Changed Warehouse from '" & sOldWhse & "' to '" & sNewWhse & "'"
        Else
            EventWarehouse = "Warehouse is '" & sNewWhse & "'"
        End If
    End If
End Function


'Add new event of UPS Acct
Private Function EventUPSAcct() As String
    If Trim(m_udtOrder.sUPSAcct) <> Trim(m_udtOrderBackup.sUPSAcct) Then
        EventUPSAcct = "Changed UPSAcct from '" & Trim(m_udtOrderBackup.sUPSAcct) & "' to '" & Trim(m_udtOrder.sUPSAcct) & "'"
    End If
End Function


Private Function EventPurchOrd() As String
    If Trim(m_udtOrder.sPurchOrd) <> Trim(m_udtOrderBackup.sPurchOrd) Then
        EventPurchOrd = "Changed PurchOrd from '" & Trim(m_udtOrderBackup.sPurchOrd) & "' to '" & Trim(m_udtOrder.sPurchOrd) & "'"
    End If
End Function


Private Function EventInfo() As String
    If Trim(m_udtOrder.sInfo) <> Trim(m_udtOrderBackup.sInfo) Then
        EventInfo = "Changed Info from '" & Trim(m_udtOrderBackup.sInfo) & "' to '" & Trim(m_udtOrder.sInfo) & "'"
    End If
End Function


Private Function EventDropShip() As String
    If m_udtOrder.bDropShip <> m_udtOrderBackup.bDropShip Then
        EventDropShip = "Dropship changed to '" & CStr(m_udtOrder.bDropShip) & "'"
    End If
End Function


Private Function EventShipComplete() As String
    If m_udtOrder.bShipComplete <> m_udtOrderBackup.bShipComplete Then
        EventShipComplete = "ShipComplete changed to '" & CStr(m_udtOrder.bShipComplete) & "'"
    End If
End Function


Private Function EventPrintPickList() As String
    If m_udtOrder.bPrintPickList <> m_udtOrderBackup.bPrintPickList Then
        EventPrintPickList = "PrintPickList changed to '" & CStr(m_udtOrder.bPrintPickList) & "'"
    End If
End Function


Private Function EventPricePackList() As String
    If m_udtOrder.bPricePackList <> m_udtOrderBackup.bPricePackList Then
        EventPricePackList = "PricePackList changed to '" & CStr(m_udtOrder.bPricePackList) & "'"
    End If
End Function


Private Function EventFreeFreight() As String
    If m_udtOrder.bFreeFreight <> m_udtOrderBackup.bFreeFreight Then
        EventFreeFreight = "FreeFreight changed to '" & CStr(m_udtOrder.bFreeFreight) & "'"
    End If
End Function


Private Function EventReducedFreight() As String
    If m_udtOrder.bReducedFreight <> m_udtOrderBackup.bReducedFreight Then
        EventReducedFreight = "ReducedFreight changed to '" & CStr(m_udtOrder.bReducedFreight) & "'"
    End If
End Function


Private Function EventBillDifferentRate() As String
    Dim sNewBillMethID As String
    Dim sOldBillMethID As String
    
    If m_udtOrder.bBillDifferentRate <> m_udtOrderBackup.bBillDifferentRate Then
        EventBillDifferentRate = "BillDifferentRate changed to '" & CStr(m_udtOrder.bBillDifferentRate) & "'"
        If m_udtOrder.bBillDifferentRate Then
            sNewBillMethID = ShipMethKeytoID(m_udtOrder.lBillMethKey)
            EventBillDifferentRate = EventBillDifferentRate & ". The BillMethod is " & sNewBillMethID
        End If
    Else
        If m_udtOrder.bBillDifferentRate And m_udtOrder.lBillMethKey <> m_udtOrderBackup.lBillMethKey Then
            sNewBillMethID = ShipMethKeytoID(m_udtOrder.lBillMethKey)
            sOldBillMethID = ShipMethKeytoID(m_udtOrderBackup.lBillMethKey)
            EventBillDifferentRate = "The BillMethod changed from " & sOldBillMethID & " to " & sNewBillMethID
        End If
    End If
End Function


Private Function EventInboundFreight() As String
    If m_udtOrder.bInboundFreight <> m_udtOrderBackup.bInboundFreight Then
        EventInboundFreight = "InboundFreight changed to '" & CStr(m_udtOrder.bInboundFreight) & "'"
    End If
End Function


Private Function EventItemID(ByRef oItem As IItem) As String
    Dim sItemID As String
    
    If oItem.ItemID = "" Then
        Select Case oItem.OPItemType
            Case itMoldedGasket
                sItemID = oItem.Descr
            Case itWarmerWire
                sItemID = oItem.Descr
            Case itSpecialOrder
                sItemID = "SpecialOrder " & oItem.Descr
            Case itBTOKit
                sItemID = "BTOKit " & oItem.Descr
            Case itFinishedGood
                sItemID = "FinishedGood " & oItem.Descr
            Case itWireShelf
                sItemID = oItem.Descr
        End Select
    Else
        sItemID = oItem.ItemID
    End If
    
    EventItemID = sItemID
End Function


Private Function EventDeposit() As String
    If m_udtOrder.bDeposit <> m_udtOrderBackup.bDeposit Then
        EventDeposit = "Deposit flag changed to '" & CStr(m_udtOrder.bDeposit) & "'"
    End If
End Function


Private Function EventPartsNoCharge() As String
    If m_udtOrder.bPartsNoCharge <> m_udtOrderBackup.bPartsNoCharge Then
        EventPartsNoCharge = "PartsNoCharge changed to '" & CStr(m_udtOrder.bPartsNoCharge) & "'"
    End If
End Function


Private Function EventOrderMM() As String
    If m_oRemarkContext.RemarkList.Dirty Then
        EventOrderMM = "Order Remarks were changed"
    End If
End Function


Private Function EventOrderUpdateCustomer() As String
    If m_udtOrder.bUpdateCustomer Then
        EventOrderUpdateCustomer = "New customer assigned. " & m_oCustomer.Name & ", Customer type is '" & m_oCustomer.CustType & "'"
    End If
End Function


Private Function EventOrderLineMM() As String
    Dim sTemp As String
    Dim sItemID As String
    Dim oItem As IItem
    
    sTemp = "OrderLine Remarks were changed for:"
    EventOrderLineMM = sTemp
    
    For Each oItem In m_oItems
        If oItem.RemarkContext.RemarkList.Dirty Then
                sItemID = sItemID & " " & oItem.ItemID & ","
                Exit For
        End If
    Next
    
    sTemp = sTemp & sItemID
    
    If sTemp = EventOrderLineMM Then
        EventOrderLineMM = ""
    Else
        EventOrderLineMM = Left(sTemp, Len(sTemp) - 1)
    End If
End Function


Private Function EventNewOrderLine(ByRef oItem As IItem) As String
    Dim sStatus As String

    'If this is a new order and line item's general status is need research,
    'log the special research status to the log event.  9/27/02 TX
    If oItem.StatusCode = iscResearch Then
        sStatus = ResearchStatusString(oItem.ResearchStatus)
    Else
        sStatus = StatusCodeString(oItem.StatusCode)
    End If

    If Trim(oItem.ItemID) <> "" Then

        EventNewOrderLine = "Added " & oItem.Qty & " " & oItem.ItemID & ". Price = $" _
                & oItem.EffectivePrice & " and cost is $" & oItem.Cost _
                & ". Item status = " & sStatus & " <oplinekey=" & oItem.OPLineKey & ">"
    Else
        EventNewOrderLine = "Added " & oItem.Qty & " " & oItem.Descr & ". Price = $" _
                & oItem.EffectivePrice & " and cost is $" & oItem.Cost _
                & ". Item status = " & sStatus & " <oplinekey=" & oItem.OPLineKey & ">"
    End If
End Function


Private Function EventOldOrderLine(ByRef oItem As IItem) As String
    Dim sTemp As String
    Dim oTempItem As IItem
    Dim bNewItem As Boolean
    Dim bItemChange As Boolean
    Dim sItemID As String
    
    sItemID = EventItemID(oItem)
    
    bNewItem = True
    If Not m_oBackUpItems Is Nothing Then
        If Not (m_oBackUpItems.Count = 0) Then
            For Each oTempItem In m_oBackUpItems
                If oTempItem.OPLineKey = oItem.OPLineKey Then
                    If oTempItem.EffectivePrice <> oItem.EffectivePrice Then
                        sTemp = "Price changed from $" & oTempItem.EffectivePrice & " to $" _
                                & oItem.EffectivePrice & ". "
                        bItemChange = True
                    End If
                    
                    If oTempItem.Cost <> oItem.Cost Then
                        sTemp = sTemp & vbCrLf & "Cost changed from $" & oTempItem.Cost & " to $" _
                                & oItem.Cost & ". "
                        bItemChange = True
                    End If
                    
                     If oTempItem.Qty <> oItem.Qty Then
                        sTemp = sTemp & vbCrLf & "Qty changed from " & oTempItem.Qty & " to " _
                                & oItem.Qty & ". "
                        bItemChange = True
                    End If
                    
                    If Trim(oTempItem.ItemID) <> Trim(oItem.ItemID) Then
                        sTemp = sTemp & vbCrLf & "PartNbr changed from " & Trim(oTempItem.ItemID) & " to " _
                                & Trim(oItem.ItemID) & "."
                        bItemChange = True
                    Else
                        If Trim(oTempItem.Descr) <> Trim(oItem.Descr) Then
                            sTemp = sTemp & vbCrLf & "Changed from " & Trim(oTempItem.Descr) & " to " _
                                    & Trim(oItem.Descr) & "."
                            bItemChange = True
                        End If
                    End If
                    
                    If oTempItem.StatusCode <> oItem.StatusCode Then
                        sTemp = sTemp & vbCrLf & "Item status changed from " & StatusCodeString(oTempItem.StatusCode) & " to " _
                                & StatusCodeString(oItem.StatusCode) & "."
                        bItemChange = True
                    ElseIf oTempItem.StatusCode = iscResearch Then
                        'If the line general status is need research, we have to check if the
                        'special research statuses changes.  9/27/02 TX
                        If oTempItem.ResearchStatus <> oItem.ResearchStatus Then
                            sTemp = sTemp & vbCrLf & "Item status changed from " & ResearchStatusString(oTempItem.ResearchStatus) & " to " _
                                & ResearchStatusString(oItem.ResearchStatus) & "."
                            bItemChange = True
                        End If
                    End If
                    
                    If bItemChange = True Then
                        sTemp = "Modified " & sItemID & ". " & sTemp
                    Else
                        sTemp = ""
                    End If
                    
                    
                    bNewItem = False
                    Exit For
                End If
            Next
        End If
        
        If bNewItem Then
            Dim sStatus As String
            
            'If this is a new order and line item's general status is need research,
            'log the special research status to the log event.
            If oItem.StatusCode = iscResearch Then
                sStatus = ResearchStatusString(oItem.ResearchStatus)
            Else
                sStatus = StatusCodeString(oItem.StatusCode)
            End If
            
            sTemp = "Add " & oItem.Qty & " " & sItemID & " to order. Price is $" _
                    & oItem.EffectivePrice & " and cost is $" & oItem.Cost _
                    & ". The new item status is '" & sStatus & "'"
                    
            If Trim(oItem.OriginalItemID) <> "" Then
                sTemp = sTemp & " This is a cross referenced item. Original Part Nbr is " & _
                    oItem.OriginalItemID & ". Reference Source is " & oItem.RefSource
            End If
        End If
    End If
    EventOldOrderLine = sTemp
End Function


'Called by LogSalesOrderEvent()

Private Function EventNewOrder() As String
    If StatusCode = ItemStatusCode.iscCommitted Then
        EventNewOrder = "The new order is committed (SO " & m_udtOrder.lTranNo & ")"
    ElseIf StatusCode = ItemStatusCode.iscARHold Then
        EventNewOrder = "Placed on AR Hold."
    Else
        EventNewOrder = "Saved. Status = " & StatusCodeString(StatusCode)
    End If
End Function


'Called by LogSalesOrderEvent()

Private Function EventOldOrder() As String
    
    If Recommit Then
        EventOldOrder = "Reopened for editing. SO " & OldSOID & " was cancelled."
        'EventOldOrder = "Change order status from 'committed' back to 'ready to commit' for recommitting"
    End If
     
    If m_oBackUpStatusCode <> StatusCode Then
        If StatusCode = ItemStatusCode.iscCommitted Then
            EventOldOrder = "Order Committed (SO " & m_udtOrder.lTranNo & ")"
        ElseIf StatusCode = ItemStatusCode.iscARHold Then
            EventOldOrder = "Placed on AR Hold"
        ElseIf StatusCode = ItemStatusCode.iscReadyToCommit And m_oBackUpStatusCode = ItemStatusCode.iscARHold Then
            EventOldOrder = "Returned by AR to " & m_udtOrder.sUserID
        ElseIf StatusCode = ItemStatusCode.iscHasRMA Then
            EventOldOrder = "Created RMA."
        Else
            EventOldOrder = "Change order status from '" & StatusCodeString(m_oBackUpStatusCode) & "' to '" & StatusCodeString(StatusCode) & "'"
        End If
    End If
End Function


Private Function EventShipToName() As String
    If Trim(m_udtOrder.sShipToName) <> Trim(m_udtOrderBackup.sShipToName) Then
        EventShipToName = "Changed ShipToName from '" & Trim(m_udtOrderBackup.sShipToName) & "' to '" & Trim(m_udtOrder.sShipToName) & "'"
    End If
End Function

Private Function EventShipToPhone() As String
    If Trim(m_udtOrder.sShipToPhone) <> Trim(m_udtOrderBackup.sShipToPhone) Then
        EventShipToPhone = "Changed ShipToPhone from '" & Trim(m_udtOrderBackup.sShipToPhone) & "' to '" & Trim(m_udtOrder.sShipToPhone) & "'"
    End If
End Function


Private Function EventShipToNote() As String
    If Trim(m_udtOrder.sShipToNote) <> Trim(m_udtOrderBackup.sShipToNote) Then
        EventShipToNote = "Changed ShipToNote from '" & Trim(m_udtOrderBackup.sShipToNote) & "' to '" & Trim(m_udtOrder.sShipToNote) & "'"
    End If
End Function


'************ End of String Builders ****************************

'This is intended for future use:

Private Sub SendEmailNotification()
    Dim sBody As String
    Dim sEmailList As String
    Dim sCSR As String
    Dim sWhseID As String
    
    'Create To: email address list
    Select Case m_udtOrder.lWhseKey
        Case g_MPKWhseKey
            sWhseID = g_MPKWhseID
            sEmailList = "lorir@caseparts.com; erickc@caseparts.com"
        Case g_SEAWhseKey
            sWhseID = g_SEAWhseID
            sEmailList = "elainer@caseparts.com; georget@caseparts.com"
        Case g_STLWhseKey
            sWhseID = g_STLWhseID
            sEmailList = "joannar@caseparts.com; danar@caseparts.com"
    End Select
    
    sCSR = GetUserName

    'Begin Create the Body of the email
    sBody = "This is an automated request for a new customer account." & vbCrLf & vbCrLf _
        & "OP Number: " & Me.OPKey & vbCrLf _
        & "Requesting CSR: " & sCSR _
        & "Warehouse: " & sWhseID _
        & vbCrLf & vbCrLf _
        & "Customer Information:" & vbCrLf _
        & m_oCustomer.Name & vbCrLf _
        & m_oCustomer.ShipAddr.Addr1 & vbCrLf
        
    If Len(m_oCustomer.ShipAddr.Addr2) > 0 Then
        sBody = sBody & m_oCustomer.ShipAddr.Addr2 & vbCrLf
    End If
    
    sBody = sBody & m_oCustomer.ShipAddr.City & ", " _
        & m_oCustomer.ShipAddr.State & ". " _
        & m_oCustomer.ShipAddr.Zip & vbCrLf _
        & m_oCustomer.ShipAddr.CountryID & vbCrLf & vbCrLf
        
    If Not Me.contact Is Nothing Then
        sBody = sBody & "Order Contact:" & vbCrLf _
            & Me.contact.Name & vbCrLf _
            & Format(Me.contact.Phone, "(###)###-####") & vbCrLf & vbCrLf
    End If
    
    sBody = sBody & "Please reply to this email with the new customer account number."
    
    'Create From: email address list
    sCSR = sCSR & "@caseparts.com"
    
    'This is enabled for testing
    EMail.Send sCSR, "dannyh@caseparts.com;jakew@caseparts.com", "Request for new customer account.", sBody, False 'TextFormat

End Sub


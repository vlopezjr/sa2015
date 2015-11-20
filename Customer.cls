VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const k_sDfltTempCustType = "EndUser"

Private Type CustProps
    sName As String * 40
    lKey As Long
    sID As String * 12
    bReqPO As Boolean
    bShipComplete As Boolean
    sCustType As String * 12    'Sage CustClassID ("Dealer", "EndUser", "Wholesale")
    sCollector As String
    bIsTemp As Boolean
    bIsWalkup As Boolean
    bIsMiscCustomer As Boolean  'CAUTION: this property only has a meaningful value
                                'when the order is first being created (it's not persisted)
    bPricePackList As Boolean
    lStatus As Long
    'sPassword As String
    sGiftRecipient As String
    lNatAcctKey As Long
    lBillKey As Long
    bIsGovt As Boolean
End Type


Private m_udtCust As CustProps
Private m_udtCustBackup As CustProps

Private m_oBillAddr As Address
Private m_oShipAddr As Address

Private m_oContacts As Contacts


Private Sub Class_Initialize()
    Set m_oBillAddr = New Address
    Set m_oShipAddr = New Address
    Set m_oContacts = New Contacts
    Clear
End Sub


Private Sub Class_Terminate()
    Set m_oBillAddr = Nothing
    Set m_oShipAddr = Nothing
    Set m_oContacts = Nothing
End Sub


'************************************************************************
' Public Properties
'************************************************************************

Public Property Get Name() As String
    Name = Trim(m_udtCust.sName)
End Property

Public Property Let Name(ByVal vNewValue As String)
    m_udtCust.sName = vNewValue
End Property


Public Property Get ID() As String
    ID = Trim(m_udtCust.sID)
End Property

Public Property Let ID(ByVal vNewValue As String)
    m_udtCust.sID = vNewValue
End Property


Public Property Get Key() As Long
    Key = m_udtCust.lKey
End Property

'6/13/05 JEJ
Public Property Get NatAcctKey() As Long
    NatAcctKey = m_udtCust.lNatAcctKey
End Property

Public Property Get BillKey() As Long
    BillKey = m_udtCust.lBillKey
End Property


'Public Property Get Password() As String
'    Password = Trim(m_udtCust.sPassword)
'End Property
'
'
'Public Property Let Password(ByVal vNewValue As String)
'    m_udtCust.sPassword = vNewValue
'End Property


Public Property Get Collector() As String
    Collector = Trim$(m_udtCust.sCollector)
End Property

Public Property Let Collector(ByVal vNewValue As String)
    m_udtCust.sCollector = vNewValue
End Property


Public Property Get CustType() As String
    CustType = Trim$(m_udtCust.sCustType)
End Property


Public Property Let CustType(ByVal vNewValue As String)
    m_udtCust.sCustType = vNewValue
End Property


Public Property Get CustTypeKey() As Byte
    Select Case Trim$(m_udtCust.sCustType)
        Case "EndUser"
            CustTypeKey = 1
        Case Is = "Dealer"
            CustTypeKey = 2
        Case Is = "Wholesale"
            CustTypeKey = 3
    End Select
End Property


Public Property Get Hold() As Boolean
    Dim rst As Recordset
    Dim sSQL As String

    With m_udtCust
        'No Hold for Temp Customers or Walk-Ups

        If Not HasAccount Then Exit Property
        
            sSQL = "SELECT Hold FROM tcpCustHold ch INNER JOIN tcpHoldStatus hs " _
                 & "ON ch.HoldStatusKey = hs.HoldStatusKey " _
                 & "WHERE Hold = 1 and CustKey = " & .lBillKey

    End With
    
    Set rst = LoadDiscRst(sSQL)
    Hold = (Not rst.EOF) 'Customer is on hold if we find a matching row
    Set rst = Nothing
End Property


'This property determines whether or not a Customer should be offered a catalog.
'It encapsultates all of our business logic.
'Do not record requests for
'   1. MISC customers (NOTE: m_oCustomer.IsMiscCustomer property is meaningless)
'   2. End Users
'   3. Shipping addresses whose first three characters don't match
'      the first three characters of the billing address
'
'Check to see if the custkey+addrkey has an entry in tcpCatalogRequest. If so, the customer
'(at this location) has already been asked.

Public Property Get QueryForCatalog() As Boolean
    'if this is a MISC customer, return false
    If (InStr(1, Me.ID, "-MISC") > 0) Then
        QueryForCatalog = False
    ElseIf Me.CustType = "EndUser" Then
        QueryForCatalog = False
    ElseIf Me.ShipAddr.AddrType <> Default Then
        QueryForCatalog = False
        
    Else
        Dim cmd As ADODB.Command
        Set cmd = CreateCommandSP("spcpcCheckForCatalogRequest")
        With cmd
            .Parameters("@_iCustKey").value = Me.Key
            .Parameters("@_iAddrKey").value = Me.ShipAddr.AddrKey
            .Execute
            'if there's an entry, return false
            QueryForCatalog = IIf(.Parameters("@_oRetVal").value > 0, False, True)
        End With
    End If
End Property


'Added 11/29/04 LR (PRN 499)
'Returns True if there's an unfulfilled entry (ResponseCode = 0)
'in tcpSLXmasGifts2004 for the CustKey/ShipAddr pair.

'ResponseCode values:
'  0 - default (no answer)
'  1 - yes
'  2 - no

Public Property Get QueryForGift() As Boolean
    Dim oCmd As ADODB.Command
    Set oCmd = CreateCommandSP("spcpcCheckForXmasGift")
    With oCmd
        .Parameters("@_iCustKey").value = Me.Key
        .Parameters("@_iAddrKey").value = Me.ShipAddr.AddrKey
        .Execute
        If IsNull(.Parameters("@_oRespCode").value) Then
            'the customer is not in the table
            QueryForGift = False
        ElseIf .Parameters("@_oRespCode").value > 0 Then
            'customer is "fulfilled"
            QueryForGift = False
        ElseIf .Parameters("@_oRespCode").value = 0 Then
            'customer is "unfulfilled"
            QueryForGift = True
            m_udtCust.sGiftRecipient = Trim(.Parameters("@_oRecipient").value)
        Else
            'catch all
            QueryForGift = False
        End If
    End With
    Set oCmd = Nothing
End Property

'added 11/29/04 LR (PRN 499)

Public Property Get GiftRecipient() As String
    GiftRecipient = m_udtCust.sGiftRecipient
End Property


Public Property Get ReqPO() As Boolean
    ReqPO = m_udtCust.bReqPO
End Property

Public Property Let ReqPO(ByVal vNewValue As Boolean)
    m_udtCust.bReqPO = vNewValue
End Property


Public Property Get ShipComplete() As Boolean
    ShipComplete = m_udtCust.bShipComplete
End Property


Public Property Let ShipComplete(ByVal vNewValue As Boolean)
    m_udtCust.bShipComplete = vNewValue
End Property


Public Property Get PricePackList() As Boolean
    PricePackList = m_udtCust.bPricePackList
End Property


Public Property Get IsTemp() As Boolean
    IsTemp = m_udtCust.bIsTemp
End Property


Public Property Let IsTemp(ByVal vNewValue As Boolean)
    m_udtCust.bIsTemp = vNewValue
End Property


Public Property Get IsWalkup() As Boolean
    IsWalkup = m_udtCust.bIsWalkup
End Property


'CAUTION: this property only has a meaningful value when the order is first being created (it's not persisted)
'In the code we test for MISC is a couple of ways:
'   Right(m_oOrder.Customer.ID, 5) <> "-MISC"
'   (InStr(1, Me.ID, "-MISC") > 0)

Public Property Get IsMiscCustomer() As Boolean
    IsMiscCustomer = m_udtCust.bIsMiscCustomer
End Property

Public Property Let IsMiscCustomer(ByVal vNewValue As Boolean)
    m_udtCust.bIsMiscCustomer = vNewValue
End Property

'***DH 12/11/08
Public Property Get IsGovernment() As Boolean
    IsGovernment = m_udtCust.bIsGovt
End Property

Public Property Get lStatusKey() As Long
    lStatusKey = m_udtCust.lStatus
End Property

Public Property Get BillAddr() As Address
    Set BillAddr = m_oBillAddr
End Property


Public Property Get ShipAddr() As Address
    Set ShipAddr = m_oShipAddr
End Property


Public Property Let BillAddr(oNewValue As Address)
    Set m_oBillAddr = oNewValue
End Property


Public Property Let ShipAddr(oNewValue As Address)
    Set m_oShipAddr = oNewValue
End Property


'indicate if Customer on hold is eligible for new orders
'What's the business rule in plain english?

Public Property Get CanBuy() As Boolean
    Dim rst As Recordset
    Dim sSQL As String
    
    CanBuy = True
    
    With m_udtCust

        If Not HasAccount Then Exit Property
        
        Set rst = CallSP("spcpcCustCanBuy", "@_iCustKey", .lBillKey)
    End With
    
    If Not rst.EOF Then
        If rst.Fields("HoldStatusID") <> "VIP" And rst.Fields("Over90").value > 0 Then
            CanBuy = False
        End If
    End If
    
    Set rst = Nothing
End Property


Public Property Get IsCOD() As Boolean
    IsCOD = m_oBillAddr.DefaultPmtTerms.IsCash
End Property


Public Property Get Contacts() As Contacts
    Set Contacts = m_oContacts
End Property


'Used by MDIMain.DoSplitOrder
Public Property Let Contacts(oNewValue As Contacts)
    Set m_oContacts = oNewValue
End Property


Private Function isMisc() As Boolean
    isMisc = (InStr(1, Me.ID, "-MISC") > 0)
End Function


Public Function HasAccount() As Boolean
    HasAccount = True
    With m_udtCust
        If .bIsTemp Or .bIsWalkup Or (InStr(1, .sID, "-MISC") > 0) Then
            HasAccount = False
        End If
    End With
End Function


'****************************************************************************
' Public Methods
'****************************************************************************

Public Function Export() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Dim oXMLAddr As JDMPDXML.XMLNode
    
    
    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "Customer"
        .Add "Name", Trim(m_udtCust.sName)
        .Add "Key", m_udtCust.lKey

'Added these 2 fields
'symptom: split orders were failing to carry over cust hold status
        .Add "NatAcctKey", m_udtCust.lNatAcctKey
        .Add "BillKey", m_udtCust.lBillKey
    
        .Add "ID", Trim(m_udtCust.sID)
        .Add "ReqPO", m_udtCust.bReqPO
        .Add "ShipComplete", m_udtCust.bShipComplete
        .Add "CustType", Trim(m_udtCust.sCustType)
        .Add "IsTemp", m_udtCust.bIsTemp
        .Add "lStatus", m_udtCust.lStatus
        .Add "Hold", Hold
        .Add "IsGovt", m_udtCust.bIsGovt
        
        Set oXMLAddr = New JDMPDXML.XMLNode
        oXMLAddr.Tag = "BillAddr"
        oXMLAddr.AddChild m_oBillAddr.Export
        .AddChild oXMLAddr
    
        Set oXMLAddr = New JDMPDXML.XMLNode
        oXMLAddr.Tag = "ShipAddr"
        oXMLAddr.AddChild m_oShipAddr.Export
        .AddChild oXMLAddr
    End With

    Set Export = oXML
End Function



Public Function XmlCommit() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    
    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "customer"
        .Add "key", m_udtCust.lKey
        .Add "id", Trim$(m_udtCust.sID)
        .Add "name", Trim$(m_udtCust.sName)
        .AddChild m_oBillAddr.XmlCommit("billaddr")
        .AddChild m_oShipAddr.XmlCommit("shipaddr")
    End With

    Set XmlCommit = oXML
End Function



Public Sub Import(ByVal i_oXML As JDMPDXML.XMLNode)
    With i_oXML
        If .Tag <> "Customer" Then
            Err.Raise -1, "Customer.ImportXML", "Unexpected initial XML tag"
        End If
        m_udtCust.sName = .Item("Name").value
        'm_udtCust.sVaxAcct = .Item("VaxAcct").Value
        m_udtCust.lKey = .Item("Key").value

'Added these 2 fields
'symptom: split orders were failing to carry over cust hold status
        m_udtCust.lNatAcctKey = .Item("NatAcctKey").value
        m_udtCust.lBillKey = .Item("BillKey").value
        
        m_udtCust.sID = .Item("ID").value
        m_udtCust.bReqPO = .Item("ReqPO").value
        m_udtCust.bShipComplete = .Item("ShipComplete").value
        m_udtCust.sCustType = .Item("CustType").value
        m_udtCust.bIsTemp = .Item("IsTemp").value
        On Error Resume Next
        m_udtCust.lStatus = 0
        m_udtCust.bIsGovt = .Item("IsGovt").value
        m_udtCust.lStatus = .Item("lStatus").value
        m_oBillAddr.Import .Item("BillAddr").Item("Address")
        m_oShipAddr.Import .Item("ShipAddr").Item("Address")
    End With

End Sub


Public Sub InitMiscCustomer(sCustID As String)
    Dim rst As Recordset
    'get keys for the warehouse's MISC customer
    Set rst = LoadDiscRst("Select CustKey, DfltBillToAddrKey, DfltShipToAddrKey from tarCustomer Where CustID = '" & sCustID & "'")
    Load rst.Fields("CustKey")
    BillAddr.Load rst.Fields("DfltBillToAddrKey")
    ShipAddr.Load rst.Fields("DfltShipToAddrKey")
    m_udtCust.bIsMiscCustomer = True
    Set rst = Nothing
End Sub


Public Sub InitWalkup(sCustID As String)
    Dim rst As Recordset
    Set rst = LoadDiscRst("Select CustKey, DfltBillToAddrKey, DfltShipToAddrKey from tarCustomer Where CustID = '" & sCustID & "'")
    Load rst.Fields("CustKey")
    BillAddr.Load rst.Fields("DfltBillToAddrKey")
    ShipAddr.Load rst.Fields("DfltShipToAddrKey")
    m_udtCust.bIsWalkup = True
    Set rst = Nothing
End Sub


Public Sub Load(ByVal i_lKey As Long)
    Dim oCmd As ADODB.Command
    Dim orst As ADODB.Recordset
    
    On Error GoTo EH

    Set orst = DATABASE.CallSP("spcpcLoadCustomer", "@_iCustKey", i_lKey)
    If orst.EOF Then
        On Error GoTo 0
        Err.Raise -1, "Customer.Load", _
            "Customer.Load" & vbCrLf & "CustKey=" & i_lKey & " does not exist in the database."
    End If
    
    With orst
        m_udtCust.lKey = .Fields("CustKey")
        m_udtCust.sID = .Fields("CustID")
        m_udtCust.sName = .Fields("CustName")

        If Not IsNull(.Fields("Collector")) Then
            m_udtCust.sCollector = .Fields("Collector")
        Else
            m_udtCust.sCollector = ""
        End If
        
        m_udtCust.bReqPO = .Fields("ReqPO").value <> 0
        m_udtCust.sCustType = .Fields("CustClassID")
        m_udtCust.lStatus = .Fields("Status")
        If Not IsNull(.Fields("PricePackList")) Then
            m_udtCust.bPricePackList = (Trim(.Fields("PricePackList")) = "Yes")
        Else
            m_udtCust.bPricePackList = False
        End If
        'm_udtCust.sPassword = Format(.Fields("Password"))
        
        'Populate lBillKey and lNatAcctKey as appropriate
        m_udtCust.lNatAcctKey = .Fields("NatAcctKey")
        If (.Fields("BillKey")) = 0 Then
            m_udtCust.lBillKey = m_udtCust.lKey
        Else
            m_udtCust.lBillKey = .Fields("BillKey")
        End If

        m_udtCust.bIsGovt = .Fields("isGovt")
        
        'The outer condition is needed to account for the bad behavior of FOrder which
        'loads the customer object several times during order creation instead of once.
        If m_oContacts.Count = 0 Then
            'Build the customer's contact collection (if not MISC, WALKUP, or TEMP)
            If HasAccount Then
                'm_oContacts.Connection = g_DB.Connection
                m_oContacts.OwnerKey = m_udtCust.lKey
                m_oContacts.LoadContacts
            End If
        End If

    End With

    Exit Sub
EH:
    Err.Raise Err.Number, "Customer.Load", _
        "Customer.Load" & vbCrLf & Err.Description
        
End Sub


Public Sub Backup()
    m_udtCustBackup = m_udtCust
    m_oBillAddr.Backup
    m_oShipAddr.Backup
End Sub


Public Sub Restore()
    m_udtCust = m_udtCustBackup
    m_oBillAddr.Restore
    m_oShipAddr.Restore
End Sub


Public Sub Clear()
    With m_udtCust
        .lKey = 0
        .sName = ""
        .sID = ""
        .sCollector = ""
        .bReqPO = False
        .sCustType = ""
        .bIsTemp = False
        .bShipComplete = False
        .bIsWalkup = False
        .bIsMiscCustomer = False
        .bPricePackList = False
        .lStatus = 0
        '.sPassword = ""
        .lBillKey = 0
        .lNatAcctKey = 0
        .bIsGovt = False
    End With
    
    m_oBillAddr.Clear
    m_oShipAddr.Clear

    m_oContacts.Clear
    
End Sub









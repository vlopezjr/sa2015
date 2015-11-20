VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreditCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type CardProps
    lKey As Long
    bPreferred As Boolean
    lTypeKey As Long
    sTypeID As String
'go through code and make sure this value gets properly trimmed on load and
'remove all of the trimming going on after the fact.
    sCardNo As String
    
'    sCID As String         '9/20/06 removed

    sMaskedCCNo As String
    
    sExpireDate As String
    sCardHolderName As String
    sStreetNbr As String
    sZipCode As String
    iStatus As Integer '1=Active  2=Delete
End Type

'used by Order.GetSalesOrdDflts (part of SO API)
Private m_rstCCrdTendTypes As ADODB.Recordset

Private WithEvents m_oState As BitMap
Attribute m_oState.VB_VarHelpID = -1
Private m_oHoldState As BitMap

Private m_udtCard As CardProps
Private m_udtCardBackup As CardProps

Private m_oCCTransactions As CCTransactions

'if this card is loaded with an order, this is a reference to the parent order
Private m_oOrder As Order


'************************************************************************
'PUBLIC EVENTS
'************************************************************************

Public Event StateChange(newstate As BitMap)


'************************************************************************
'CLASS EVENTS
'************************************************************************

Private Sub Class_Initialize()
    Dim sSQL As String
    
    sSQL = "SELECT TenderTypeKey, TenderTypeID " _
         & "FROM tcmTenderType " _
         & "WHERE CompanyID = 'CPC' and TenderTypeID like '%-CCRD'"
    Set m_rstCCrdTendTypes = LoadDiscRst(sSQL)
    
    Set m_oState = New BitMap
    Set m_oHoldState = New BitMap
End Sub


Private Sub Class_Terminate()
    Set m_oOrder = Nothing
    Set m_rstCCrdTendTypes = Nothing
    Set m_oState = Nothing
    Set m_oHoldState = Nothing
    Set m_oCCTransactions = Nothing
End Sub


'************************************************************************
' Public Properties
'************************************************************************

Public Property Get State() As BitMap
    Set State = m_oState
End Property


Public Property Get Key() As Long
    Key = m_udtCard.lKey
End Property

Public Property Let Key(ByVal vNewValue As Long)
    m_udtCard.lKey = vNewValue
End Property


'Public Property Get Order() As Order
'    Set Order = m_oOrder
'End Property

Public Property Let Order(ByRef oNewOrder As Order)
    Set m_oOrder = oNewOrder
End Property


Public Property Get CardNo() As String
    CardNo = Trim$(m_udtCard.sCardNo)
End Property

Public Property Let CardNo(ByVal vNewValue As String)
    m_udtCard.sCardNo = vNewValue
    m_udtCard.sMaskedCCNo = mask(vNewValue)
End Property


Private Function mask(Number As String) As String
    Dim i As Integer
    Number = Trim$(Number)
    For i = 1 To Len(Number) - 4
        Mid$(Number, i, 1) = "*"
    Next i
    mask = Number
End Function


Public Property Get MaskedCCNo() As String
    MaskedCCNo = m_udtCard.sMaskedCCNo
End Property


Public Property Get Status() As String
    'Status = Trim$(m_udtCard.iStatus)
    If m_udtCard.iStatus = 1 Then
        Status = "Active"
    ElseIf m_udtCard.iStatus = 2 Then
        Status = "Delete"
    Else
        Status = "Error"
    End If
End Property


Public Property Get CardHolderName() As String
    CardHolderName = Trim(m_udtCard.sCardHolderName)
End Property

Public Property Let CardHolderName(ByVal vNewValue As String)
    m_udtCard.sCardHolderName = vNewValue
End Property


Public Property Get StreetNbr() As String
    StreetNbr = Trim(m_udtCard.sStreetNbr)
End Property

Public Property Let StreetNbr(ByVal vNewValue As String)
    m_udtCard.sStreetNbr = vNewValue
End Property


Public Property Get ZipCode() As String
    ZipCode = Trim(m_udtCard.sZipCode)
End Property

Public Property Let ZipCode(ByVal vNewValue As String)
    m_udtCard.sZipCode = vNewValue
End Property


Public Property Get ExpireDate() As String
    ExpireDate = Trim$(m_udtCard.sExpireDate)
End Property

Public Property Let ExpireDate(ByVal vNewValue As String)
    m_udtCard.sExpireDate = vNewValue
End Property

'The Credit Card editor enforces an MMYYYY format on expiry dates, including leading 0
'The Verisign querystring wants MMYY, including leading 0

Public Property Get ExpireMonth() As String
    ExpireMonth = Mid$(ExpireDate, 1, 2)
End Property

Public Property Get ExpireYear() As String
    ExpireYear = Right$(ExpireDate, 2)
End Property


Public Property Get TypeKey() As Long
    TypeKey = m_udtCard.lTypeKey
End Property

Public Property Let TypeKey(ByVal vNewValue As Long)
    m_udtCard.lTypeKey = vNewValue
End Property


Public Property Get TypeID() As String
    TypeID = m_udtCard.sTypeID
End Property

Public Property Let TypeID(ByVal sNewVal As String)
    m_udtCard.sTypeID = sNewVal
End Property


Public Property Get Transactions() As CCTransactions
    Set Transactions = m_oCCTransactions
End Property


Public Property Get HasBeenAuthorized() As Boolean
    HasBeenAuthorized = m_oCCTransactions.HasBeenAuthorized
End Property


Public Property Get AuthorizationPNREF() As String
    AuthorizationPNREF = m_oCCTransactions.AuthorizationPNREF
End Property


Public Property Get HasBeenCharged() As Boolean
    HasBeenCharged = m_oCCTransactions.HasBeenCharged
End Property


Public Property Get HasErrors() As Boolean
    HasErrors = m_oCCTransactions.HasErrors
End Property


'9/20/06 removed
'
'Public Property Get CID() As String
'    CID = m_udtCard.sCID
'End Property
'
'Public Property Let CID(ByVal vNewValue As String)
'    m_udtCard.sCID = vNewValue
'End Property


Public Property Get Preferred() As Boolean
    Preferred = m_udtCard.bPreferred
End Property

Public Property Let Preferred(bNewValue As Boolean)
    m_udtCard.bPreferred = bNewValue
End Property


'used by Order.GetSalesOrdDflts (part of SO API)

Public Property Get TenderType(ByVal whseid As String) As String
    m_rstCCrdTendTypes.Filter = "TenderTypeID like '" & whseid & "%'"
    TenderType = m_rstCCrdTendTypes.Fields("TenderTypeKey").value
    m_rstCCrdTendTypes.Filter = adFilterNone
End Property


Public Function IsNew() As Boolean
    IsNew = m_oState.TestBits(eMask.IsNew)
End Function

Public Function IsValid() As Boolean
    IsValid = m_oState.TestBits(eMask.IsValid)
End Function

Public Function IsDirty() As Boolean
    IsDirty = m_oState.TestBits(eMask.IsDirty)
End Function


'********************************************************************
' Public Methods
'********************************************************************

'This is called to initialize a NEW CreditCard

Public Sub Init()
    m_udtCard.lKey = 0
    TypeKey = 0
    CardNo = ""
    ExpireDate = ""
    CardHolderName = ""
    StreetNbr = ""
    ZipCode = ""
    Preferred = 0
    m_udtCard.iStatus = 1
   
    m_oState.SetBits eMask.IsNew
    m_oState.ClearBits eMask.IsDirty

    Backup
    
End Sub


'Add a comment explaining why the Order parameter is optional
'6/6/06 LR why is this needed at all?

Public Sub Load(ByVal CCKey As Long, Optional oOrder As Variant)
    Dim oRst As ADODB.Recordset
    Dim oCrypto As Crypto

    If Not IsMissing(oOrder) Then
        Set m_oOrder = oOrder
        Set m_oCCTransactions = New CCTransactions
        m_oCCTransactions.Load CCKey, m_oOrder.OPKey
    End If
   
    Set oRst = CallSP("spcpcCreditCardLoad", "@CCKey", CCKey)
    
    Set oCrypto = New Crypto
    
    If Not oRst.EOF Then
        m_udtCard.lKey = oRst!CCKey
        m_udtCard.lTypeKey = oRst!CrCardTypeKey
        m_udtCard.sTypeID = oRst!CrCardTypeName
        m_udtCard.sCardNo = oCrypto.decrypt(oRst!CodedCCNo)
        m_udtCard.sMaskedCCNo = oRst!CrCardNo
        m_udtCard.sCardHolderName = oRst!CardHolderName
        m_udtCard.sExpireDate = oRst!CrCardExp
        m_udtCard.sStreetNbr = oRst!CrCardStreetNbrZip
        m_udtCard.sZipCode = oRst!CrCardZipCode
        m_udtCard.bPreferred = oRst!Preferred
        m_udtCard.iStatus = oRst!Status
    End If

    m_oState.ClearBits eMask.IsNew + eMask.IsDirty
    m_oState.SetBits eMask.IsValid

    Backup
    
End Sub


Public Sub Restore()
    m_udtCard = m_udtCardBackup
    m_oState.value = m_oHoldState.value
End Sub


Public Sub Add(CustKey As Long)
    Dim oCmd As ADODB.Command
    Dim oCrypto As Crypto
    
    Set oCmd = New ADODB.Command
    
    Set oCrypto = New Crypto
        
    With oCmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spCPC_CCEditorInsert"
        .Parameters("@InParamCardTypeKey").value = m_udtCard.lTypeKey
        .Parameters("@InParamCardNo").value = m_udtCard.sMaskedCCNo
        .Parameters("@InParamCodedCCNo").value = oCrypto.encrypt(m_udtCard.sCardNo)
        .Parameters("@InParamCardExp").value = m_udtCard.sExpireDate
        .Parameters("@InParamCardHolderName").value = m_udtCard.sCardHolderName
        .Parameters("@InParamCardStreetNbrZip").value = m_udtCard.sStreetNbr
        .Parameters("@InParamCardZipCode").value = m_udtCard.sZipCode
        .Parameters("@InParamCardPreferred").value = m_udtCard.bPreferred
        
        If isMisc(CustKey) Then CustKey = 0
        'note: a new customer will already have CustKey = 0
        
        .Parameters("@InParamCustKey").value = CustKey
               
        .Execute
         m_udtCard.lKey = .Parameters("@OutParamCCKey").value
    End With
    
    Set oCmd = Nothing
    Set oCrypto = Nothing
    
    m_oState.ClearBits eMask.IsDirty + eMask.IsNew
    
    Call Backup
End Sub


'Lookup a Customer by Key and see if it's one of our Misc accounts

Private Function isMisc(CustKey As Long) As Boolean
    isMisc = False
    
    Dim CustID As String
    Dim cmd As ADODB.Command
    
    Set cmd = CreateCommandSP("spCPCIsMisc")
    
    With cmd
        .Parameters("@_iCustKey").value = Trim(CustKey)
        .Execute
        CustID = IIf(IsNull(.Parameters("@_oCustID").value), vbNullString, .Parameters("@_oCustID").value)
    End With
    
    If (InStr(1, CustID, "-MISC") > 0) Then isMisc = True

End Function


Public Sub Update()
    Dim oCmd As ADODB.Command
    Dim oCrypto As Crypto
    
    Set oCmd = New ADODB.Command
    Set oCrypto = New Crypto
    
    With oCmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spCPC_CCEditorUpdate"
        .Parameters("@InParamCardTypeKey").value = m_udtCard.lTypeKey
        .Parameters("@InParamCardNo").value = m_udtCard.sMaskedCCNo
        .Parameters("@InParamCodedCCNo").value = oCrypto.encrypt(m_udtCard.sCardNo)
        .Parameters("@InParamCardExp").value = m_udtCard.sExpireDate
        .Parameters("@InParamCardHolderName").value = m_udtCard.sCardHolderName
        .Parameters("@InParamCardStreetNbrZip").value = m_udtCard.sStreetNbr
        .Parameters("@InParamCardZipCode").value = m_udtCard.sZipCode
        .Parameters("@InParamCardPreferred").value = m_udtCard.bPreferred
        .Parameters("@InParamKey").value = m_udtCard.lKey
        .Execute
    End With
    
    Set oCmd = Nothing
    Set oCrypto = Nothing
    
    m_oState.ClearBits eMask.IsDirty
    
    Backup
End Sub


Public Sub Delete(CustKey As Long)
    Dim oCmd As ADODB.Command
    
    Set oCmd = New ADODB.Command
    
    With oCmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spCPC_CCEditorUpdateStatus"
        .Parameters.Append .CreateParameter("@InParamCustKey", adInteger, adParamInput, Len(CustKey), CustKey)
        .Parameters.Append .CreateParameter("@Key", adInteger, adParamInput, Len(m_udtCard.lKey), m_udtCardBackup.lKey)
        .Execute
    End With
    
    Set oCmd = Nothing
    
    m_udtCard.iStatus = 2
    
End Sub


Public Function Export() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode

    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        .Tag = "CreditCard"
        .Add "Key", m_udtCard.lKey
        .Add "Preferred", m_udtCard.bPreferred
        .Add "TypeKey", m_udtCard.lTypeKey
        .Add "TypeID", Trim(m_udtCard.sTypeID)
        .Add "CardNo", Trim(m_udtCard.sCardNo)
        .Add "ExpireDate", m_udtCard.sExpireDate
        .Add "CardHolderName", Trim(m_udtCard.sCardHolderName)
        .Add "StreetNbr", Trim(m_udtCard.sStreetNbr)
        .Add "ZipCode", Trim(m_udtCard.sZipCode)
        .Add "Status", m_udtCard.iStatus
        .Add "State", m_oState.value
    End With
    
    Set Export = oXML
End Function


Public Function Import(ByVal i_oXML As JDMPDXML.XMLNode)
    With i_oXML
        Dim i As Integer
        If .Tag = "CreditCard" Then
            m_udtCard.lKey = .Item("Key").value
            m_udtCard.bPreferred = .Item("Preferred").value
            m_udtCard.lTypeKey = .Item("TypeKey").value
            m_udtCard.sTypeID = .Item("TypeID").value
            m_udtCard.sCardNo = .Item("CardNo").value
            m_udtCard.sExpireDate = .Item("ExpireDate").value
            m_udtCard.sCardHolderName = .Item("CardHolderName").value
            m_udtCard.sStreetNbr = .Item("StreetNbr").value
            m_udtCard.sZipCode = .Item("ZipCode").value
            m_udtCard.iStatus = .Item("Status").value
            m_oState.value = .Item("State").value
        End If
    End With
End Function


'Event Handlers

Private Sub m_oState_StateChange(newstate As BitMap)
    RaiseEvent StateChange(newstate)
End Sub


Private Sub Backup()
    m_udtCardBackup = m_udtCard
    m_oHoldState.value = m_oState.value
End Sub


'**** NOTE ********************************************************
'Removed 2/3/2009 Now handled within SOAPIServer in C#
'Restored this 2/16/09 for the temp resurrection of AutoCommit

'3/28/11 LR changed this function to return boolean and pass back the value returned from oTran.Authorize()

Public Function Authorize() As Boolean
    Dim oTran As CCTransaction

    Set oTran = New CCTransaction

    Authorize = oTran.Authorize(m_oOrder)
    m_oCCTransactions.Add oTran
    'Authorize = oTran.Result
End Function


Public Function Charge(invID As String, invKey As Long, invAmt As Double)
    Dim oTran As CCTransaction
    
    Set oTran = New CCTransaction
    
    oTran.InvcKey = invKey
    oTran.Amount = invAmt
    oTran.Charge m_oOrder, invID
    m_oCCTransactions.Add oTran
    Charge = oTran.Result
End Function


Public Function XmlCommit(whseid As String) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "creditcard"
        .Add "key", m_udtCard.lKey
        
        If m_oCCTransactions Is Nothing Then    'new credit card
            .Add "hasbeenauthorized", "False"
        Else
            .Add "hasbeenauthorized", Me.HasBeenAuthorized
        End If
        
        .Add "number", Trim$(m_udtCard.sCardNo)    ' already decrypted
        .Add "maskednumber", Trim$(m_udtCard.sMaskedCCNo)
        .Add "expirymonth", Me.ExpireMonth
        .Add "expiryyear", Me.ExpireYear
        .Add "cardholder", Me.CardHolderName
        .Add "streetno", Me.StreetNbr
        .Add "zipcode", Me.ZipCode
        .Add "tendertypekey", Me.TenderType(whseid)
    End With
    Set XmlCommit = oXML
End Function


Public Function Serialize(whseid As String) As IXMLDOMElement
    Dim dom As New DOMDocument
    Dim element As IXMLDOMElement
    Dim property As IXMLDOMElement
    
    Set element = dom.createElement("creditcard")
    
    Set property = dom.createElement("key")
    element.appendChild property
    property.text = m_udtCard.lKey
    
    If m_oCCTransactions Is Nothing Then    'new credit card
        Set property = dom.createElement("hasbeenauthorized")
        element.appendChild property
        property.text = "False"
    Else
        Set property = dom.createElement("hasbeenauthorized")
        element.appendChild property
        property.text = Me.HasBeenAuthorized
    End If
    
    Set property = dom.createElement("key")
    element.appendChild property
    property.text = m_udtCard.lKey
    
    Set property = dom.createElement("number")
    element.appendChild property
    property.text = Trim$(m_udtCard.sCardNo)    ' already decrypted!!!
    
    Set property = dom.createElement("maskednumber")
    element.appendChild property
    property.text = Trim$(m_udtCard.sMaskedCCNo)
    
    Set property = dom.createElement("expirymonth")
    element.appendChild property
    property.text = Me.ExpireMonth
    
    Set property = dom.createElement("expiryyear")
    element.appendChild property
    property.text = Me.ExpireYear
    
    Set property = dom.createElement("cardholder")
    element.appendChild property
    property.text = Me.CardHolderName
    
    Set property = dom.createElement("streetno")
    element.appendChild property
    property.text = Me.StreetNbr
    
    Set property = dom.createElement("zipcode")
    element.appendChild property
    property.text = Me.ZipCode
    
    Set property = dom.createElement("tendertypekey")
    element.appendChild property
    property.text = Me.TenderType(whseid)
    
    Set Serialize = element
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Changed(ByVal i_sProperty As String)

' These specific enums have been written to xml
' & when the order is reloaded will be imported.
Public Enum eAddrType
    Undefined = 0
    Default = 2
    CSA = 3
    TOO = 4
End Enum

Private Type AddrProps
    lAddrKey As Long
    sAddrName As String
    sAddr1 As String
    sAddr2 As String
    sCity As String
    sState As String
    sZip As String
    sCountryID As String
    lSalesTerritoryKey As Long
    
    lOPKey As Long
    lCustKey  As Long
    lWhseKey As Long
    
    bShipComplete As Boolean
'    bThisOrderOnly As Boolean      'removed 9/5/08 (see repercussions in this class below)
    bResidential As Boolean
    
    bIsDirty As Boolean

    AddrType As eAddrType
End Type

'Store the customer's default terms in the address billing address.
Private m_oPmtTerms As PaymentTerms

Private m_udtAddr As AddrProps
Private m_udtAddrBackup As AddrProps


'***********************************************************************
' Class Events
'***********************************************************************

Private Sub Class_Initialize()
    Set m_oPmtTerms = New PaymentTerms
    Clear
End Sub


Private Sub Class_Terminate()
    Set m_oPmtTerms = Nothing
End Sub


'***********************************************************************
' Public Properties
'***********************************************************************

Public Property Get DefaultPmtTerms() As PaymentTerms
    Set DefaultPmtTerms = m_oPmtTerms
End Property


Public Property Get SalesTerritoryKey() As Long
    SalesTerritoryKey = m_udtAddr.lSalesTerritoryKey
End Property

Public Property Let SalesTerritoryKey(ByVal lNewValue As Long)
    m_udtAddr.lSalesTerritoryKey = lNewValue
End Property


Public Property Get WhseKey() As Long
    WhseKey = m_udtAddr.lWhseKey
End Property

Public Property Let WhseKey(ByVal lNewValue As Long)
    m_udtAddr.lWhseKey = lNewValue
End Property


Public Property Get CustKey() As Long
    CustKey = m_udtAddr.lCustKey
End Property

Public Property Let CustKey(lNewValue As Long)
    m_udtAddr.lCustKey = lNewValue
End Property

    
Public Property Get AddrKey() As Long
    AddrKey = m_udtAddr.lAddrKey
End Property

Public Property Let AddrKey(ByVal lNewValue As Long)
    If lNewValue <> m_udtAddr.lAddrKey Then
        m_udtAddr.lAddrKey = lNewValue
        SetDirty "AddrKey"
    End If
End Property


' Only if addrtype is undefined, we get addrtype from sproc based on addrkey

Public Property Get AddrType() As eAddrType
    If m_udtAddr.AddrType = Undefined Then
        Dim cmd As ADODB.Command
        Set cmd = CreateCommandSP("spcpcGetAddrType")
        With cmd
            .Parameters("@_iShipAddrKey").value = m_udtAddr.lAddrKey
            .Execute
            m_udtAddr.AddrType = .Parameters("@_oAddrType").value
        End With
    End If
    AddrType = m_udtAddr.AddrType
End Property

' AddrType Let is needed for setting AddrType to TOO
' Note: Setting this property does not make the Address dirty.

Public Property Let AddrType(ByVal lNewValue As eAddrType)
    m_udtAddr.AddrType = lNewValue
End Property


Public Property Get OPKey() As Long
    OPKey = m_udtAddr.lOPKey
End Property

'This is only used by TOO addresses for event logging purposes.
Public Property Let OPKey(ByVal lNewValue As Long)
    If lNewValue <> m_udtAddr.lOPKey Then
        m_udtAddr.lOPKey = lNewValue
        SetDirty "OPKey"
    End If
End Property


Public Property Get AddrName() As String
    AddrName = Trim(m_udtAddr.sAddrName)
End Property

Public Property Let AddrName(ByVal sNewValue As String)
    If Trim$(sNewValue) <> Trim$(m_udtAddr.sAddrName) Then
        m_udtAddr.sAddrName = Trim$(sNewValue)
        SetDirty "AddrName"
    End If
End Property


Public Property Get Addr1() As String
    Addr1 = Trim(m_udtAddr.sAddr1)
End Property

Public Property Let Addr1(ByVal sNewValue As String)
    If Trim$(sNewValue) <> Trim$(m_udtAddr.sAddr1) Then
        m_udtAddr.sAddr1 = Trim$(sNewValue)
        SetDirty "Addr1"
    End If
End Property


Public Property Get Addr2() As String
    Addr2 = Trim(m_udtAddr.sAddr2)
End Property

Public Property Let Addr2(ByVal sNewValue As String)
    If Trim$(sNewValue) <> Trim$(m_udtAddr.sAddr2) Then
        m_udtAddr.sAddr2 = Trim$(sNewValue)
        SetDirty "Addr2"
    End If
End Property


Public Property Get City() As String
    City = Trim(m_udtAddr.sCity)
End Property

Public Property Let City(ByVal sNewValue As String)
    If Trim$(sNewValue) <> Trim$(m_udtAddr.sCity) Then
        m_udtAddr.sCity = Trim$(sNewValue)
        SetDirty "City"
    End If
End Property


Public Property Get State() As String
    State = Trim(m_udtAddr.sState)
End Property

Public Property Let State(ByVal sNewValue As String)
    If Trim$(sNewValue) <> Trim$(m_udtAddr.sState) Then
        m_udtAddr.sState = Trim$(sNewValue)
        SetDirty "State"
    End If
End Property


Public Property Get Zip() As String
    Zip = Trim(m_udtAddr.sZip)
End Property

Public Property Let Zip(ByVal sNewValue As String)
    With m_udtAddr
        If Trim$(sNewValue) <> Trim$(.sZip) Then
            .sZip = Trim$(sNewValue)
            SetDirty "Zip"
        End If
    End With
    
End Property


Public Property Get CountryID() As String
    CountryID = Trim(m_udtAddr.sCountryID)
End Property

Public Property Let CountryID(ByVal sNewValue As String)
    If Trim$(sNewValue) <> Trim$(m_udtAddr.sCountryID) Then
        m_udtAddr.sCountryID = Trim$(sNewValue)
        SetDirty "CountryID"
    End If
End Property


Public Property Get Residential() As Boolean
    Residential = m_udtAddr.bResidential
End Property

Public Property Let Residential(ByVal bNewValue As Boolean)
    If bNewValue <> m_udtAddr.bResidential Then
        m_udtAddr.bResidential = bNewValue
        SetDirty "Residential"
    End If
End Property


Public Property Get ShipComplete() As Boolean
    ShipComplete = m_udtAddr.bShipComplete
End Property

Public Property Let ShipComplete(ByVal bNewValue As Boolean)
    If bNewValue <> m_udtAddr.bShipComplete Then
        m_udtAddr.bShipComplete = bNewValue
        SetDirty "ShipComplete"
    End If
End Property


Public Property Get IsDirty() As Boolean
    IsDirty = m_udtAddr.bIsDirty
End Property


Public Property Let IsDirty(ByVal i_bNewValue As Boolean)
    m_udtAddr.bIsDirty = i_bNewValue
End Property


'******************************************************************************
' Public Functions
'******************************************************************************

Public Function Export() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode

    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "Address"
        .Add "AddrName", Trim(m_udtAddr.sAddrName)
        .Add "Addr1", Trim(m_udtAddr.sAddr1)
        .Add "Addr2", Trim(m_udtAddr.sAddr2)
        .Add "City", Trim(m_udtAddr.sCity)
        .Add "State", Trim(m_udtAddr.sState)
        .Add "Zip", Trim(m_udtAddr.sZip)
        .Add "AddrKey", m_udtAddr.lAddrKey
        .Add "OPKey", m_udtAddr.lOPKey
        .Add "CountryID", Trim(m_udtAddr.sCountryID)
        .Add "ShipComplete", m_udtAddr.bShipComplete
        .Add "PmtTermsKey", m_oPmtTerms.Key
        .Add "CustKey", m_udtAddr.lCustKey
        .Add "Residential", m_udtAddr.bResidential
        .Add "Modified", m_udtAddr.bIsDirty             '???
        .Add "AddrType", m_udtAddr.AddrType
    End With

    Set Export = oXML
End Function


Public Sub Import(ByVal i_oXML As JDMPDXML.XMLNode)
    With i_oXML
        If .Tag <> "Address" Then
            Err.Raise -1, "Address.ImportXML", "Unexpected initial XML tag"
        End If
        
        Me.Clear 'initialize all fields on m_udtAddr
        
        m_udtAddr.lAddrKey = .Item("AddrKey").value
        m_udtAddr.sAddrName = .Item("AddrName").value
        m_udtAddr.sAddr1 = .Item("Addr1").value
        m_udtAddr.sAddr2 = .Item("Addr2").value
        m_udtAddr.sCity = .Item("City").value
        m_udtAddr.sState = .Item("State").value
        m_udtAddr.sZip = .Item("Zip").value
        m_udtAddr.lOPKey = .Item("OPKey").value
        m_udtAddr.sCountryID = .Item("CountryID").value
        m_udtAddr.bShipComplete = .Item("ShipComplete").value

        m_oPmtTerms.Key = .Item("PmtTermsKey").value
        
        m_udtAddr.lCustKey = .Item("CustKey").value
        m_udtAddr.bIsDirty = .Item("Modified").value        '???
                
        'Some item such as Email, Residential, and Dfltcntctkey are
        'properties for Address. Therefore, it may crash the OP if the user to import
        'these new properties that don't exist in old XML.
        On Error Resume Next
        
        m_udtAddr.AddrType = .Item("AddrType").value

        m_udtAddr.bResidential = .Item("Residential").value
        
        On Error GoTo 0
    End With
End Sub


Public Sub Load(ByVal i_lKey As Long)
    Dim orst As ADODB.Recordset
    Dim sSQL As String
    
   If i_lKey = 0 Then Exit Sub
    
    Set orst = CallSP("spcpcOPAddrLoad", "@_AddrKey", CStr(i_lKey))

    With m_udtAddr
        .bIsDirty = False
        .lAddrKey = orst.Fields("AddrKey").value
        .lOPKey = 0 'Not saved to OP Tables yet
        
        .sAddrName = Trim(orst.Fields("AddrName").value)
        .sAddr1 = Trim(orst.Fields("AddrLine1").value)
        .sAddr2 = Trim(orst.Fields("AddrLine2").value)
        .sCity = Trim(orst.Fields("City").value)
        .sState = Trim(orst.Fields("StateID").value)
        .sZip = Trim(orst.Fields("PostalCode").value)
        .sCountryID = Trim(orst.Fields("CountryID").value)
        .bResidential = CBool(Trim(orst.Fields("AddrFaxExt").value) = "1")
        .bShipComplete = (orst.Fields("ShipComplete").value <> 0)

'This key comes from tarCustAddr
        m_oPmtTerms.Key = orst.Fields("PmtTermsKey").value

        .lCustKey = orst.Fields("CustKey").value
        .lSalesTerritoryKey = orst.Fields("SalesTerritoryKey").value
        .lWhseKey = orst.Fields("WhseKey").value

        ' Set addrtype to an initial state of UnDefined.
        ' When the addrtype is needed it will get the addrtype from sproc.
        .AddrType = Undefined
    End With
    
End Sub


Public Sub Clear(Optional bClearPmtTerms As Boolean = True)
    With m_udtAddr
        .bIsDirty = False
        .lAddrKey = 0
        .lOPKey = 0
        .sAddr1 = ""
        .sAddr2 = ""
        .sAddrName = ""
        .sCity = ""
        .sCountryID = "USA"
        .sState = ""
        .sZip = ""
        .bShipComplete = False
        .bResidential = False
        .lSalesTerritoryKey = 0
        .lWhseKey = 0

        If bClearPmtTerms Then m_oPmtTerms.Clear
        
        .AddrType = Undefined
    End With
    
End Sub


Public Function CompleteAddr() As String
    With m_udtAddr
        CompleteAddr = CompAddr(.sAddrName, .sAddr1, .sAddr2, .sCity, .sState, .sZip, .sCountryID)
    End With
End Function


Public Sub Backup()
    m_udtAddrBackup = m_udtAddr
    m_oPmtTerms.Backup
End Sub


Public Sub Restore()
    m_udtAddr = m_udtAddrBackup
    m_oPmtTerms.Restore
End Sub


Public Function XmlCommit(i_sAddrType As String) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        If i_sAddrType = "billaddr" Then
            .Tag = "billaddr"
            .Add "key", m_udtAddr.lAddrKey
        Else
            .Tag = "shipaddr"
            .Add "key", m_udtAddr.lAddrKey
            .Add "addrtype", m_udtAddr.AddrType
            .Add "residential", m_udtAddr.bResidential
            .Add "addrname", Trim$(m_udtAddr.sAddrName)
            .Add "addrline1", Trim$(m_udtAddr.sAddr1)
            .Add "addrline2", Trim$(m_udtAddr.sAddr2)
            .Add "city", Trim$(m_udtAddr.sCity)
            .Add "state", Trim$(m_udtAddr.sState)
            .Add "countryid", Trim$(m_udtAddr.sCountryID)
            .Add "postalcode", Trim(m_udtAddr.sZip)
        End If
    End With
    
    Set XmlCommit = oXML
End Function


Public Function Serialize(i_sAddrType As String) As IXMLDOMElement
    Dim dom As New DOMDocument
    Dim element As IXMLDOMElement
    Dim property As IXMLDOMElement
    
    If i_sAddrType = "billaddr" Then
        Set element = dom.createElement("billaddr")
        
        Set property = dom.createElement("key")
        element.appendChild property
        property.text = m_udtAddr.lAddrKey
    Else
        Set element = dom.createElement("shipaddr")
        
        Set property = dom.createElement("key")
        element.appendChild property
        property.text = m_udtAddr.lAddrKey
        
        Set property = dom.createElement("addrtype")
        element.appendChild property
        property.text = m_udtAddr.AddrType
        
        Set property = dom.createElement("residential")
        element.appendChild property
        property.text = m_udtAddr.bResidential
        
        Set property = dom.createElement("addrname")
        element.appendChild property
        property.text = Trim$(m_udtAddr.sAddrName)
        
        Set property = dom.createElement("addrline1")
        element.appendChild property
        property.text = Trim$(m_udtAddr.sAddr1)
        
        Set property = dom.createElement("addrline2")
        element.appendChild property
        property.text = Trim$(m_udtAddr.sAddr2)
        
        Set property = dom.createElement("city")
        element.appendChild property
        property.text = Trim$(m_udtAddr.sCity)
        
        Set property = dom.createElement("state")
        element.appendChild property
        property.text = Trim$(m_udtAddr.sState)
        
        Set property = dom.createElement("countryid")
        element.appendChild property
        property.text = Trim$(m_udtAddr.sCountryID)
        
        Set property = dom.createElement("postalcode")
        element.appendChild property
        property.text = Trim(m_udtAddr.sZip)
    End If
    
    Set Serialize = element
End Function


'FEditAddress catches this event and updates controls

Private Sub SetDirty(ByVal i_sProperty As String)
    m_udtAddr.bIsDirty = True
    RaiseEvent Changed(i_sProperty)
End Sub


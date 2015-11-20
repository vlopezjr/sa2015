VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum CustType
    ctUnknown = 0
    ctEndUser = 1
    ctDealer = 2
    ctWholesale = 3
End Enum

Private Type ItemProps
    lOPLineKey As Long
    eStatusCode As ItemStatusCode
    eCustType As CustType
    dCost As Double
    dListPrice As Double
    dDealerPrice As Double
    dWholesalePrice As Double
    dNegotiatedPrice As Double
    dBackNegotiatedPrice As Double
    lSageItemType As Long
    lItemKey As Long
    lLineKey As Long
    lMakeKey As Long
    lOPKey As Long
    lQty As Long
    lVendorKey As Long
    sDescr As String * 40
    sItemID As String * 30
    sModelNbr As String * 20
    sSerialNbr As String * 20
    bIsTaxable As Boolean
    bCGMPN As Boolean
    lMorphBTOKey As Long
    eResearchStatus As ItemResearchStatus
    sOriginalItemID As String
    sRefSource As String

'4/7/05 LR move to ItemFinGood?
    sWhseBinID As String 'PRN#118
    
    eItemInventoryStatus As enumItemInventoryStatus
End Type

Private m_udtItem As ItemProps
Private m_udtItemBackup As ItemProps
Private m_oRemarkContext As MemoMeister.RemarkContext

'Private Const k_lResaleTaxSchedKey = 166
'Private Const k_lCPCResaleTaxClass = 10


'***************************************************************************
' Class Event Handlers
'***************************************************************************

Private Sub Class_Initialize()
    With m_udtItem
        .eCustType = ctEndUser
        .dCost = 0
        .dListPrice = 0
        .dDealerPrice = 0
        .dWholesalePrice = 0
        .dNegotiatedPrice = -1
        .dBackNegotiatedPrice = 0
        .bCGMPN = False
        .lMorphBTOKey = 0
    
        .sItemID = "" 'otherwise, the string is initially set to NULLs
        .sDescr = ""
        .sModelNbr = ""
        .sSerialNbr = ""
        .lQty = 1   'HACK - this should be a user option later on
    End With
    
    Set m_oRemarkContext = New MemoMeister.RemarkContext
    m_oRemarkContext.EditBehavior = NoAutoSave
    
End Sub


Private Sub Class_Terminate()
    If Not m_oRemarkContext Is Nothing Then
        m_oRemarkContext.Dispose
        Set m_oRemarkContext = Nothing
    End If
End Sub


'***************************************************************************
' Public Properties
'***************************************************************************

Public Property Get CustType() As String
    Select Case m_udtItem.eCustType
    Case ctEndUser
        CustType = "EndUser"
    Case ctDealer
        CustType = "Dealer"
    Case ctWholesale
        CustType = "Wholesale"
    Case Else
        CustType = "Unknown"
    End Select
End Property

Public Property Let CustType(ByVal sNewValue As String)
    sNewValue = Trim(sNewValue)
    
    Select Case sNewValue
    Case "", "Unknown"
        m_udtItem.eCustType = ctUnknown
    
    Case "EndUser"
        m_udtItem.eCustType = ctEndUser

    Case "Dealer"
        m_udtItem.eCustType = ctDealer

    Case "Wholesale"
        m_udtItem.eCustType = ctWholesale
    
    Case Else
        Err.Raise -1, "Item.CustType", "Unexpected value for CustType: " & sNewValue
    End Select
End Property


Public Property Get Cost() As Double
    Cost = m_udtItem.dCost
End Property

Public Property Let Cost(ByVal dNewValue As Double)
    m_udtItem.dCost = dNewValue
End Property


Public Property Get ListPrice() As Double
    ListPrice = m_udtItem.dListPrice
End Property

Public Property Let ListPrice(ByVal dNewValue As Double)
    With m_udtItem
        If .dListPrice <> dNewValue Then
            .dListPrice = dNewValue
            If .eCustType = ctEndUser Or .eCustType = ctUnknown Then
                .dNegotiatedPrice = -1 'auto-reset on re-calc
            End If
        End If
    End With
End Property


Public Property Get DealerPrice() As Double
    DealerPrice = m_udtItem.dDealerPrice
End Property

Public Property Let DealerPrice(ByVal dNewValue As Double)
    With m_udtItem
        If .dDealerPrice <> dNewValue Then
            .dDealerPrice = dNewValue
            If .eCustType = ctDealer Then
                .dNegotiatedPrice = -1 'auto-reset on re-calc
            End If
        End If
    End With
End Property


Public Property Get WholesalePrice() As Double
    WholesalePrice = m_udtItem.dWholesalePrice
End Property

Public Property Let WholesalePrice(ByVal dNewValue As Double)
    With m_udtItem
        If .dWholesalePrice <> dNewValue Then
            .dWholesalePrice = dNewValue
            If .eCustType = ctWholesale Then
                .dNegotiatedPrice = -1 'auto-reset on re-calc
            End If
        End If
    End With
End Property

'4/7/05 LR remove?
'PRN#118
Public Property Get WhseBinID() As String
    WhseBinID = m_udtItem.sWhseBinID
End Property

Public Property Let WhseBinID(ByVal dNewValue As String)
    m_udtItem.sWhseBinID = Trim$(dNewValue)
End Property


Public Property Get NegotiatedPrice() As Double
    If m_udtItem.dNegotiatedPrice < 0 Then
        NegotiatedPrice = StandardPrice
        'NegotiatedPrice = EffectivePrice
    Else
        NegotiatedPrice = m_udtItem.dNegotiatedPrice
    End If
End Property

Public Property Let NegotiatedPrice(ByVal dNewValue As Double)
    m_udtItem.dNegotiatedPrice = dNewValue
End Property


Public Property Get BackNegotiatedPrice() As Double
    BackNegotiatedPrice = m_udtItem.dBackNegotiatedPrice
End Property


Public Property Let BackNegotiatedPrice(ByVal dNewValue As Double)
    m_udtItem.dBackNegotiatedPrice = dNewValue
End Property


Public Property Get StandardPrice() As Double
    Select Case m_udtItem.eCustType
    Case ctEndUser
        StandardPrice = m_udtItem.dListPrice
    Case ctDealer
        StandardPrice = m_udtItem.dDealerPrice
    Case ctWholesale
        StandardPrice = m_udtItem.dWholesalePrice
    Case Else
        StandardPrice = m_udtItem.dListPrice
    End Select
End Property


Public Property Get EffectivePrice() As Double
    If NegotiatedPricing Then
        EffectivePrice = NegotiatedPrice
    Else
        EffectivePrice = StandardPrice
    End If
End Property

Public Property Get NegotiatedPricing() As Boolean
    With m_udtItem
        NegotiatedPricing = .dNegotiatedPrice >= 0 And .dNegotiatedPrice <> StandardPrice
        'NegotiatedPricing = .dNegotiatedPrice <> StandardPrice Or StandardPrice = 0
    End With
End Property

Public Property Get ExtendedPrice() As Double
    ExtendedPrice = EffectivePrice * Qty
End Property

Public Property Get SageItemType() As Long
    SageItemType = m_udtItem.lSageItemType
End Property

Public Property Let SageItemType(ByVal vNewValue As Long)
    m_udtItem.lSageItemType = vNewValue
End Property

Public Property Get ItemKey() As Long
    ItemKey = m_udtItem.lItemKey
End Property

Public Property Let ItemKey(ByVal vNewValue As Long)
    m_udtItem.lItemKey = vNewValue
End Property

Public Property Get LineKey() As Long
    LineKey = m_udtItem.lLineKey
End Property

Public Property Let LineKey(ByVal vNewValue As Long)
    m_udtItem.lLineKey = vNewValue
End Property

Public Property Get MakeKey() As Long
    MakeKey = m_udtItem.lMakeKey
End Property

Public Property Let MakeKey(ByVal vNewValue As Long)
    m_udtItem.lMakeKey = vNewValue
End Property

Public Property Get OPLineKey() As Long
    OPLineKey = m_udtItem.lOPLineKey
End Property

Public Property Let OPLineKey(ByVal lNewValue As Long)
    m_udtItem.lOPLineKey = lNewValue
End Property

Public Property Get OPKey() As Long
    OPKey = m_udtItem.lOPKey
End Property


Public Property Let OPKey(ByVal vNewValue As Long)
    m_udtItem.lOPKey = vNewValue
End Property


Public Property Get Qty() As Long
    Qty = m_udtItem.lQty
End Property

Public Property Let Qty(ByVal vNewValue As Long)
    m_udtItem.lQty = vNewValue
End Property


'09/24/02       TeddyX
'Add new Research Status to base item

Public Property Get ResearchStatus() As ItemResearchStatus
    ResearchStatus = m_udtItem.eResearchStatus
End Property

Public Property Let ResearchStatus(ByVal eNewValue As ItemResearchStatus)
    'If StatusCode = iscResearch Then
    m_udtItem.eResearchStatus = eNewValue
'
'    Else
'        m_udtItem.eResearchStatus = iscResearchEmpty
'    End If
End Property


Public Property Get StatusCode() As ItemStatusCode
    StatusCode = m_udtItem.eStatusCode
End Property

Public Property Let StatusCode(ByVal eNewValue As ItemStatusCode)
    m_udtItem.eStatusCode = eNewValue
    
    '11/27/2002     TeddyX
    'Set Research Status back to empty if the Item statuscode is not iscResearch.
    
    If eNewValue <> iscResearch Then
        ResearchStatus = irsResearchEmpty
    End If
End Property


'add two more properties of Original ItemID and RefSource here
Public Property Get OriginalItemID() As String
    OriginalItemID = m_udtItem.sOriginalItemID
End Property


Public Property Let OriginalItemID(ByVal sNewValue As String)
    m_udtItem.sOriginalItemID = sNewValue
End Property


Public Property Get RefSource() As String
    RefSource = m_udtItem.sRefSource
End Property


Public Property Let RefSource(ByVal sNewValue As String)
    m_udtItem.sRefSource = sNewValue
End Property


Public Property Get VendorKey() As Long
    VendorKey = m_udtItem.lVendorKey
End Property


Public Property Let VendorKey(ByVal vNewValue As Long)
    m_udtItem.lVendorKey = vNewValue
End Property


Public Property Get Descr() As String
    Descr = Trim(m_udtItem.sDescr)
End Property

Public Property Let Descr(ByVal vNewValue As String)
    m_udtItem.sDescr = vNewValue
End Property


Public Property Get ItemID() As String
    ItemID = Trim(m_udtItem.sItemID)
End Property

Public Property Let ItemID(ByVal strNewValue As String)
    m_udtItem.sItemID = strNewValue
End Property


Public Property Get ModelNbr() As String
    ModelNbr = Trim(m_udtItem.sModelNbr)
End Property

Public Property Let ModelNbr(ByVal vNewValue As String)
    m_udtItem.sModelNbr = vNewValue
End Property


Public Property Get RemarkContext() As MemoMeister.RemarkContext
    Set RemarkContext = m_oRemarkContext
End Property


Public Property Get SerialNbr() As String
    SerialNbr = m_udtItem.sSerialNbr
End Property

Public Property Let SerialNbr(ByVal vNewValue As String)
    m_udtItem.sSerialNbr = vNewValue
End Property

Public Property Get IsCGMPN() As Boolean
    IsCGMPN = m_udtItem.bCGMPN
End Property


Public Property Let IsCGMPN(ByVal bNewValue As Boolean)
    m_udtItem.bCGMPN = bNewValue
End Property


Public Property Get MorphBTOKey() As Long
    MorphBTOKey = m_udtItem.lMorphBTOKey
End Property

Public Property Let MorphBTOKey(ByVal lNewValue As Long)
    m_udtItem.lMorphBTOKey = lNewValue
End Property

Public Property Get IsTaxable() As Boolean
    IsTaxable = m_udtItem.bIsTaxable
End Property

Public Property Let IsTaxable(ByVal vNewValue As Boolean)
    m_udtItem.bIsTaxable = vNewValue
End Property

Public Property Get ItemInventoryStatus() As enumItemInventoryStatus
    ItemInventoryStatus = m_udtItem.eItemInventoryStatus
End Property

Public Property Let ItemInventoryStatus(ByVal vNewValue As enumItemInventoryStatus)
    m_udtItem.eItemInventoryStatus = vNewValue
End Property


'***************************************************************************
' Public Methods
'***************************************************************************

'4/25/06 LR added optional XMLNode parameter

Public Function Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False, _
    Optional oNode As JDMPDXML.XMLNode = Nothing _
) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
    
        .Tag = "Item"

        .Add "CustType", m_udtItem.eCustType
        .Add "ItemID", Trim(m_udtItem.sItemID)
        .Add "Descr", Trim(m_udtItem.sDescr)
        .Add "Cost", m_udtItem.dCost
        .Add "ListPrice", Format(m_udtItem.dListPrice, "0.00")
        .Add "DealerPrice", Format(m_udtItem.dDealerPrice, "0.00")
        .Add "WholesalePrice", Format(m_udtItem.dWholesalePrice, "0.00")
        .Add "NegotiatedPrice", Format(m_udtItem.dNegotiatedPrice, "0.00")
        .Add "AcuityItemType", Format(m_udtItem.lSageItemType, "0.00")
        .Add "EffectivePrice", Format(EffectivePrice, "0.00")
        .Add "ExtendedPrice", Format(ExtendedPrice, "0.00")
        .Add "ItemKey", m_udtItem.lItemKey
        .Add "LineKey", m_udtItem.lLineKey
        .Add "OPLineKey", m_udtItem.lOPLineKey
        .Add "MakeKey", m_udtItem.lMakeKey
        .Add "OPKey", m_udtItem.lOPKey
        .Add "Qty", m_udtItem.lQty
        .Add "StatusCode", m_udtItem.eStatusCode
        .Add "ResearchStatus", m_udtItem.eResearchStatus
        .Add "VendorKey", m_udtItem.lVendorKey
        .Add "ModelNbr", Trim(m_udtItem.sModelNbr)
        .Add "SerialNbr", Trim(m_udtItem.sSerialNbr)
        .Add "IsTaxable", m_udtItem.bIsTaxable
        .Add "IsCGMPN", m_udtItem.bCGMPN
        .Add "MorphBTOKey", m_udtItem.lMorphBTOKey
        .Add "ComboMakeModelSerial", ComboMakeModelSerialNbr
        .Add "BackNegotiatedPrice", BackNegotiatedPrice

'4/7/05 LR remove?
        .Add "WhseBinID", WhseBinID
'added 4/19/06 LR
        .Add "ItemType", m_udtItem.lSageItemType
        
        ExportRemarks bPubNotes, "PubNotes", "OrderLine.Public", oXML
        ExportRemarks bPrivNotes, "PrivNotes", "OrderLine.Private", oXML
        ExportRemarks bPurchNotes, "PurchNotes", "OrderLine.Purch", oXML
        
'added 4/25/06 LR
        If Not oNode Is Nothing Then
            .AddChild oNode
        End If
        
    End With
    
    Set Export = oXML
End Function


'******************************************************************

Public Function Import(ByVal i_oXML As JDMPDXML.XMLNode, Optional lWhseKey As Long)
    With i_oXML
    
        If .Tag <> "Item" Then
            Err.Raise -1, "Item.Import", "Unexpected initial XML tag"
        End If

        m_udtItem.dCost = .Item("Cost").value
        m_udtItem.eCustType = .Item("CustType").value
        m_udtItem.dListPrice = .Item("ListPrice").value
        m_udtItem.dDealerPrice = .Item("DealerPrice").value
        m_udtItem.dWholesalePrice = .Item("WholesalePrice").value
        m_udtItem.dNegotiatedPrice = .Item("NegotiatedPrice").value
        'm_udtItem.dBackNegotiatedPrice = m_udtItem.dNegotiatedPrice
        m_udtItem.lSageItemType = .Item("AcuityItemType").value
        m_udtItem.lItemKey = .Item("ItemKey").value
        m_udtItem.lLineKey = .Item("LineKey").value
        On Error Resume Next
        m_udtItem.lOPLineKey = .Item("OPLineKey").value
        On Error GoTo 0
        m_udtItem.lMakeKey = .Item("MakeKey").value
        m_udtItem.lOPKey = .Item("OPKey").value
        m_udtItem.lQty = .Item("Qty").value
        m_udtItem.eStatusCode = .Item("StatusCode").value
        m_udtItem.lVendorKey = .Item("VendorKey").value
        m_udtItem.sDescr = .Item("Descr").value
        m_udtItem.sItemID = .Item("ItemID").value
        m_udtItem.sModelNbr = .Item("ModelNbr").value
        m_udtItem.sSerialNbr = .Item("SerialNbr").value
        m_udtItem.bIsTaxable = .Item("IsTaxable").value
        On Error Resume Next
        m_udtItem.bCGMPN = False
        m_udtItem.bCGMPN = .Item("IsCGMPN").value
        m_udtItem.eResearchStatus = irsResearchEmpty
        m_udtItem.eResearchStatus = .Item("ResearchStatus").value
        m_udtItem.lMorphBTOKey = 0
        m_udtItem.lMorphBTOKey = .Item("MorphBTOKey").value
        m_udtItem.dBackNegotiatedPrice = 0
        m_udtItem.dBackNegotiatedPrice = .Item("BackNegotiatedPrice").value
    End With
End Function


Public Function XmlCommit(i_eItemType As ItemTypeCode, _
                    i_lWhseKey As Long, _
                    ByRef rstWhses As ADODB.Recordset, _
                    ByRef rstVendor As ADODB.Recordset) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        .Tag = "item"
        .Add "opitemtype", i_eItemType
        .Add "oplinekey", m_udtItem.lOPLineKey
        .Add "key", ItemKey
        '.Add "id", ItemID
        .Add "id", Trim(m_udtItem.sItemID)
        .Add "descr", Descr
        .Add "acuityitemtype", SageItemType
        .Add "vendorkey", VendorKey
        .Add "vendoraddrkey", GetVendorAddrKey(VendorKey)
        .Add "cost", Format(Cost, "0.00")
        .Add "effectiveprice", Format(EffectivePrice, "0.00")
        .Add "standardprice", Format(StandardPrice, "0.00")
        .Add "qty", Qty
        .Add "extcomment", BuildExtComment(i_lWhseKey, rstWhses, rstVendor)
    End With
    
    Set XmlCommit = oXML
End Function


Public Sub Backup()
    m_udtItemBackup = m_udtItem
    m_oRemarkContext.RemarkList.Backup
End Sub


Public Sub Restore()
    m_udtItem = m_udtItemBackup
    m_oRemarkContext.RemarkList.Restore
End Sub

'***LR 10/31/08 This, along with a change in the way we round tax amount
'at the line level, fixes the tax rounding issue we've had in Sage for a long time.
'Public Function TaxAmt(i_dTaxRate) As Double
Public Function TaxAmt(i_dTaxRate As Double) As Double
    If IsTaxable Then
        TaxAmt = ExtendedPrice * i_dTaxRate
    Else
        TaxAmt = 0
    End If
End Function


'***************************************************************************
' Private Functions
'***************************************************************************

Private Function ComboMakeModelSerialNbr() As String
    Dim sMakeText As String
    Dim rst As ADODB.Recordset
    
    If m_udtItem.lMakeKey > 0 Then
        Set rst = LoadDiscRst("SELECT * FROM tcpNewMake where MakeID =" & m_udtItem.lMakeKey)
        If Not rst.EOF Then sMakeText = Trim(rst.Fields("MakeText").value)
        Set rst = Nothing
    End If
    
    If sMakeText <> "" Or Trim(m_udtItem.sModelNbr) <> "" Or Trim(m_udtItem.sSerialNbr) <> "" Then
        ComboMakeModelSerialNbr = sMakeText & " / " & Trim(m_udtItem.sModelNbr) & " / " & Trim(m_udtItem.sSerialNbr)
    End If
End Function


Private Sub ExportRemarks(bInclude As Boolean, sXmlTag As String, sRemarkTypeID As String, oXML As JDMPDXML.XMLNode)
    Dim sRemarks As String
    Dim oRemark As remark
    
    If Not bInclude Then Exit Sub
    If m_oRemarkContext Is Nothing Then Exit Sub
    If m_oRemarkContext.RemarkList Is Nothing Then Exit Sub
    If m_oRemarkContext.RemarkList.Count = 0 Then Exit Sub
    
    For Each oRemark In m_oRemarkContext.RemarkList
        If oRemark.RemarkType.TypeID = sRemarkTypeID Then
            sRemarks = sRemarks & oRemark.MemoText & vbCrLf
        End If
    Next
    
    If Len(sRemarks) > 0 Then
        oXML.Add sXmlTag, sRemarks
    End If
End Sub


'There are 2 private copies of this function
'One here for XMLCommit
'One in the SOAPI class

'A HACK to save SPO-WAs and code to get comments out of MM

'7/14/08 LR altered parameters
' replaced i_oOrder As Order with WhseKey as Long
'Private Function BuildExtComment(i_oOrder As Order, _
'                                ByRef rstWhses As ADODB.Recordset, _
'                                ByRef rstVendor As ADODB.Recordset) As String
Private Function BuildExtComment(WhseKey As Long, _
                                ByRef rstWhses As ADODB.Recordset, _
                                ByRef rstVendor As ADODB.Recordset) As String

    On Error GoTo ErrorHandler

    Dim oRemark As remark
    Dim sPurchNotes As String

    For Each oRemark In Me.RemarkContext.RemarkList
        If oRemark.RemarkType.TypeID = "OrderLine.Purch" Then
            sPurchNotes = sPurchNotes + oRemark.MemoText & " "
        End If
    Next
    
    If m_udtItem.lSageItemType = 0 Then
        If rstWhses Is Nothing Then Exit Function
        BuildExtComment = GetSPOExtCmnt(m_udtItem.lVendorKey, sPurchNotes, rstVendor)
        
    ElseIf Left(m_udtItem.sItemID, 4) = "SHF-" Then
        If rstWhses Is Nothing Then Exit Function
'        rstWhses.Filter = "WhseKey=" & CStr(i_oOrder.WhseKey)
        rstWhses.Filter = "WhseKey=" & CStr(WhseKey)
        BuildExtComment = GetSPOExtCmnt(rstWhses.Fields("WireShelfVendKey").value, sPurchNotes, rstVendor)
        rstWhses.Filter = adFilterNone
    Else
        BuildExtComment = sPurchNotes
    End If
    
    Exit Function
    
ErrorHandler:
    If Err.Source <> "Item.GetSPOExtCmnt" Then Err.Source = "Item.BuildExtComment"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function GetSPOExtCmnt(i_lVendorKey As Long, _
                                sPurchRemarks As String, _
                                ByRef rstVendors As ADODB.Recordset) As String
    On Error GoTo ErrorHandler
    
    Dim sVendorName As String
    Dim sVendorID As String
    
    If rstVendors Is Nothing Then Exit Function
    
    rstVendors.Filter = "VendKey=" & CStr(i_lVendorKey)
    If Not rstVendors.EOF Then
        sVendorName = Trim(rstVendors.Fields("VendName").value)
        sVendorID = Trim(rstVendors.Fields("VendID").value)
    End If
    rstVendors.Filter = adFilterNone
    
    GetSPOExtCmnt = Left$(Trim(sVendorID & " " & sVendorName & " C: (" & Format$(m_udtItem.dCost, g_MoneyMask) & ") " & sPurchRemarks), 250)
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "Item.GetSPOExtCmnt", Err.Description
End Function


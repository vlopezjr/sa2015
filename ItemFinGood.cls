VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemFinGood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ItemFinGoodProps
    lCatPage As Integer
    sLongDescr As String * 255
'4/7/05 LR add?
'    sWhseBinID As String
End Type

Private m_oInv As Inventories
Private m_udtFinGood As ItemFinGoodProps
Private m_udtFinGoodBackup As ItemFinGoodProps
Private m_bLoading As Boolean

Implements IItem

Public Event Change()

Private m_oBaseItem As Item

Public Property Get BaseClass() As Item
    Set BaseClass = m_oBaseItem
End Property


Public Sub Class_Initialize()
    Set m_oBaseItem = New Item
    Set m_oInv = New Inventories
    m_oBaseItem.StatusCode = iscAuthorize
End Sub


Public Sub Class_Terminate()
    Set m_oBaseItem = Nothing
    Set m_oInv = Nothing
End Sub


'Begin IItem Interface procedures
'================================

Private Function IItem_XMLCommit(i_lWhseKey As Long, _
                        ByRef rstWhses As ADODB.Recordset, _
                        ByRef rstVendor As ADODB.Recordset) As JDMPDXML.XMLNode
    Dim sHoldDescr As String
    Dim lItemKey As Long
    Dim sItemId As String
    Dim eItemType As ItemTypeCode

    If m_oBaseItem.itemkey = 0 Or Left$(m_oBaseItem.ItemID, 4) = "SPO-" Then
        'Save the values before PrepareToCommitItem
        eItemType = itSpecialOrder
        sHoldDescr = m_oBaseItem.Descr
        lItemKey = m_oBaseItem.itemkey
        sItemId = m_oBaseItem.ItemID
        m_oBaseItem.Descr = Left$((Trim$(m_oBaseItem.ItemID) & "-" & m_oBaseItem.Descr), 40)
        PrepareToCommitItem m_oBaseItem, "SPO", i_lWhseKey
    Else
        eItemType = itFinishedGood
    End If
    
    Set IItem_XMLCommit = m_oBaseItem.XmlCommit(eItemType, i_lWhseKey, rstWhses, rstVendor)
    
    If Left$(m_oBaseItem.ItemID, 4) = "SPO-" Then
        'Restore the Values
        m_oBaseItem.Descr = sHoldDescr
        m_oBaseItem.ItemID = sItemId
        m_oBaseItem.itemkey = lItemKey
    End If
        
End Function


Private Function IItem_Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False _
) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        If IItem_OPItemType = itFinishedGood Then
            .Tag = "FinGood"
        Else
            .Tag = "SpecialOrder"
        End If
        .Add "CatPage", m_udtFinGood.lCatPage
        .Add "LongDescr", Trim(m_udtFinGood.sLongDescr)
        
'4/7/05 LR add?
'        .Add "WhseBinID", m_udtFinGood.sWhseBinID

        .AddChild m_oBaseItem.Export(bPubNotes, bPrivNotes, bPurchNotes)
    End With
    
    Set IItem_Export = oXML
End Function


Private Sub IItem_Import(ByVal i_oXML As JDMPDXML.XMLNode, Optional lWhseKey As Long)
    With i_oXML
        If .Tag <> "FinGood" And .Tag <> "SpecialOrder" Then
            Err.Raise -1, "ItemFinGood.Import", "Unexpected initial XML tag"
        End If
        m_udtFinGood.lCatPage = .Item("CatPage").Value
        m_udtFinGood.sLongDescr = .Item("LongDescr").Value
        
'4/7/05 LR add?
'        m_udtFinGood.sWhseBinID = .Item("WhseBinID").value

        m_oBaseItem.Import .Item("Item")
        Load m_oBaseItem.itemkey, lWhseKey
    End With
End Sub


Public Sub IItem_Backup()
    m_oBaseItem.Backup
    m_udtFinGoodBackup = m_udtFinGood
End Sub


Public Sub IItem_Restore()
    m_oBaseItem.Restore
    m_udtFinGood = m_udtFinGoodBackup
End Sub


Public Function IItem_IsValidStatusCode(ByVal eStatus As ItemStatusCode) As Boolean
    Select Case eStatus
        Case ItemStatusCode.iscResearch
            IItem_IsValidStatusCode = True
        
        Case ItemStatusCode.iscQuote
            'IItem_IsValidStatusCode = (IItem_Cost <= 0) And IItem_VendorKey <> 0
            If IItem_OPItemType = itFinishedGood Then
                IItem_IsValidStatusCode = (IItem_Cost <= 0) And IItem_VendorKey <> 0
            Else
                IItem_IsValidStatusCode = True
            End If
    
        Case ItemStatusCode.iscAuthorize, ItemStatusCode.iscReadyToCommit
            IItem_IsValidStatusCode = (IItem_Cost > 0) And IItem_VendorKey <> 0
            
        Case Else
            IItem_IsValidStatusCode = False
    End Select
End Function


'Property Procedures

Public Property Let IItem_CustType(ByVal RHS As String)
    m_oBaseItem.CustType = RHS
End Property

Public Property Get IItem_CustType() As String
    IItem_CustType = m_oBaseItem.CustType
End Property


Public Property Let IItem_Cost(ByVal RHS As Double)
    If m_oBaseItem.Cost <> RHS Then
        m_oBaseItem.Cost = RHS
        Refresh
    End If
End Property

Public Property Get IItem_Cost() As Double
    IItem_Cost = m_oBaseItem.Cost
End Property


Public Property Let IItem_ListPrice(ByVal RHS As Double)
    m_oBaseItem.ListPrice = RHS
End Property

Public Property Get IItem_ListPrice() As Double
    IItem_ListPrice = m_oBaseItem.ListPrice
End Property


Public Property Let IItem_DealerPrice(ByVal RHS As Double)
    m_oBaseItem.DealerPrice = RHS
End Property

Public Property Get IItem_DealerPrice() As Double
    IItem_DealerPrice = m_oBaseItem.DealerPrice
End Property


Private Property Let IItem_ResearchStatus(ByVal RHS As ItemResearchStatus)
    m_oBaseItem.ResearchStatus = RHS
End Property


Private Property Get IItem_ResearchStatus() As ItemResearchStatus
    IItem_ResearchStatus = m_oBaseItem.ResearchStatus
End Property


Private Property Get IItem_OriginalItemID() As String
    IItem_OriginalItemID = m_oBaseItem.OriginalItemID
End Property


Private Property Let IItem_OriginalItemID(ByVal sNewValue As String)
    m_oBaseItem.OriginalItemID = sNewValue
End Property


Private Property Get IItem_RefSource() As String
    IItem_RefSource = m_oBaseItem.RefSource
End Property


Private Property Let IItem_RefSource(ByVal sNewValue As String)
    m_oBaseItem.RefSource = sNewValue
End Property

Private Function IItem_TaxAmt(i_dTaxRate As Double) As Double
    IItem_TaxAmt = m_oBaseItem.TaxAmt(i_dTaxRate)
End Function

Public Property Let IItem_WholesalePrice(ByVal RHS As Double)
    m_oBaseItem.WholesalePrice = RHS
End Property

Public Property Get IItem_WholesalePrice() As Double
    IItem_WholesalePrice = m_oBaseItem.WholesalePrice
End Property


Public Property Let IItem_NegotiatedPrice(ByVal RHS As Double)
    m_oBaseItem.NegotiatedPrice = RHS
End Property

Public Property Get IItem_NegotiatedPrice() As Double
    IItem_NegotiatedPrice = m_oBaseItem.NegotiatedPrice
End Property


Public Property Let IItem_BackNegotiatedPrice(ByVal RHS As Double)
    m_oBaseItem.BackNegotiatedPrice = RHS
End Property


Public Property Get IItem_BackNegotiatedPrice() As Double
    IItem_BackNegotiatedPrice = m_oBaseItem.BackNegotiatedPrice
End Property

Public Property Get IItem_StandardPrice() As Double
    IItem_StandardPrice = m_oBaseItem.StandardPrice
End Property


Public Property Get IItem_EffectivePrice() As Double
    IItem_EffectivePrice = m_oBaseItem.EffectivePrice
End Property


Public Property Get IItem_NegotiatedPricing() As Boolean
    IItem_NegotiatedPricing = m_oBaseItem.NegotiatedPricing
End Property


Public Property Get IItem_ExtendedPrice() As Double
    IItem_ExtendedPrice = m_oBaseItem.ExtendedPrice
End Property


Public Property Get IItem_StatusCode() As ItemStatusCode
    IItem_StatusCode = m_oBaseItem.StatusCode
End Property


Public Property Let IItem_StatusCode(ByVal RHS As ItemStatusCode)
    If IItem_IsValidStatusCode(RHS) Then
        m_oBaseItem.StatusCode = RHS
    Else
        If IItem_VendorKey = 0 Then
            m_oBaseItem.StatusCode = iscResearch
        ElseIf IItem_Cost <= 0 Then
            m_oBaseItem.StatusCode = iscQuote
        Else
            m_oBaseItem.StatusCode = iscAuthorize
        End If
    End If
    
    'Set the default research status to be irsNeedResearch if the statuscode is need research
    If m_oBaseItem.StatusCode = iscResearch Then
        If m_oBaseItem.ResearchStatus < irsNeedResearch Then
            m_oBaseItem.ResearchStatus = irsNeedResearch
        End If
    End If
    
End Property


Public Property Let IItem_Descr(ByVal RHS As String)
    m_oBaseItem.Descr = RHS
End Property

Public Property Get IItem_Descr() As String
    IItem_Descr = m_oBaseItem.Descr
End Property


Public Property Let IItem_IsTaxable(ByVal RHS As Boolean)
    m_oBaseItem.IsTaxable = RHS
End Property

Public Property Get IItem_IsTaxable() As Boolean
    IItem_IsTaxable = m_oBaseItem.IsTaxable
End Property


Public Property Let IItem_ItemID(ByVal RHS As String)
    m_oBaseItem.ItemID = RHS
End Property


Public Property Get IItem_ItemID() As String
    IItem_ItemID = m_oBaseItem.ItemID
End Property


Public Property Let IItem_ItemKey(ByVal RHS As Long)
    m_oBaseItem.itemkey = RHS
End Property

Public Property Get IItem_ItemKey() As Long
    IItem_ItemKey = m_oBaseItem.itemkey
End Property


Public Property Let IItem_SageItemType(ByVal RHS As Long)
    m_oBaseItem.SageItemType = RHS
End Property

Public Property Get IItem_SageItemType() As Long
    IItem_SageItemType = m_oBaseItem.SageItemType
End Property


Public Property Get IItem_OPItemType() As ItemTypeCode
    If m_oBaseItem.itemkey > 0 Then
        IItem_OPItemType = itFinishedGood
    Else
        IItem_OPItemType = itSpecialOrder
    End If
End Property


Public Property Let IItem_LineKey(ByVal RHS As Long)
    m_oBaseItem.LineKey = RHS
End Property

Public Property Get IItem_LineKey() As Long
    IItem_LineKey = m_oBaseItem.LineKey
End Property

Public Property Let IItem_OPLineKey(ByVal RHS As Long)
    m_oBaseItem.OPLineKey = RHS
End Property


Public Property Get IItem_IsCGMPN() As Boolean
    IItem_IsCGMPN = m_oBaseItem.IsCGMPN
End Property

Public Property Let IItem_IsCGMPN(ByVal bNewValue As Boolean)
    m_oBaseItem.IsCGMPN = bNewValue
End Property



Public Property Get IItem_MorphBTOKey() As Long
    IItem_MorphBTOKey = m_oBaseItem.MorphBTOKey
End Property

Public Property Let IItem_MorphBTOKey(ByVal lNewValue As Long)
    m_oBaseItem.MorphBTOKey = lNewValue
End Property


Public Property Get IItem_OPLineKey() As Long
    IItem_OPLineKey = m_oBaseItem.OPLineKey
End Property

Public Property Let IItem_MakeKey(ByVal RHS As Long)
    m_oBaseItem.MakeKey = RHS
End Property

Public Property Get IItem_MakeKey() As Long
    IItem_MakeKey = m_oBaseItem.MakeKey
End Property


Public Property Let IItem_ModelNbr(ByVal RHS As String)
    m_oBaseItem.ModelNbr = RHS
End Property

Public Property Get IItem_ModelNbr() As String
    IItem_ModelNbr = m_oBaseItem.ModelNbr
End Property


Public Property Let IItem_OPKey(ByVal RHS As Long)
    m_oBaseItem.OPKey = RHS
End Property

Public Property Get IItem_OPKey() As Long
    IItem_OPKey = m_oBaseItem.OPKey
End Property


Public Property Get IItem_RemarkContext() As MemoMeister.RemarkContext
    Set IItem_RemarkContext = m_oBaseItem.RemarkContext
End Property


Public Property Let IItem_Qty(ByVal RHS As Long)
    m_oBaseItem.Qty = RHS
End Property

Public Property Get IItem_Qty() As Long
    IItem_Qty = m_oBaseItem.Qty
End Property

Public Property Get IItem_QtyAvail(lWhseKey As Long) As Variant
    On Error Resume Next
'    With m_oInv.WhseInventory(lWhseKey)
'        'PRN#139
'        'IItem_QtyAvail = .QtyOnHand - .QtyOnSO - .QtyOnBO
'        IItem_QtyAvail = .QtyOnHand - .QtyOnSO - .QtyOnBO - .QtyOnPS
'    End With
    IItem_QtyAvail = m_oInv.WhseInventory(lWhseKey).QtyAvail
End Property

Public Property Get IItem_QtyOnHand(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_QtyOnHand = m_oInv.WhseInventory(lWhseKey).QtyOnHand
End Property

Public Property Get IItem_QtyOnSO(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_QtyOnSO = m_oInv.WhseInventory(lWhseKey).QtyOnSO
End Property

Public Property Get IItem_QtyOnPO(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_QtyOnPO = m_oInv.WhseInventory(lWhseKey).QtyOnPO
End Property



Public Property Get IItem_QtyOnBO(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_QtyOnBO = m_oInv.WhseInventory(lWhseKey).QtyOnBO
End Property

Public Property Get IItem_PendQtyDecrease(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_PendQtyDecrease = m_oInv.WhseInventory(lWhseKey).PendQtyDecrease
End Property

Public Property Get IItem_PendQtyIncrease(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_PendQtyIncrease = m_oInv.WhseInventory(lWhseKey).PendQtyIncrease
End Property

Public Property Get IItem_InConflict(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_InConflict = m_oInv.WhseInventory(lWhseKey).InConflict
End Property

Public Property Let IItem_SerialNbr(ByVal RHS As String)
    m_oBaseItem.SerialNbr = RHS
End Property

Public Property Get IItem_SerialNbr() As String
    IItem_SerialNbr = m_oBaseItem.SerialNbr
End Property


Public Property Let IItem_VendorKey(ByVal RHS As Long)
    m_oBaseItem.VendorKey = RHS
End Property

Public Property Get IItem_VendorKey() As Long
    IItem_VendorKey = m_oBaseItem.VendorKey
End Property


'End IItem Interface procedures
'================================

Public Sub Load(i_lItemKey As Long, Optional lWhseKey As Long = 0)
    Dim orst As ADODB.Recordset
    Dim sSQL As String

    On Error GoTo EH

    If i_lItemKey <= 0 Then
        Debug.Print "ItemFinGood.Load - invalid key ignored: " & i_lItemKey
        Exit Sub
    End If

'NOTE: 3/22/05 LR this returns up to three records ordered by WhseKey
'NOTE: 9/11/15 VL moved this call to a new stored proc ([spOPItemLoadFinGood])
    'sSQL = "SELECT i.* " _
         & ", isnull(p.QtyPending,0) As QtyPending " _
         & ", (isnull(i.QtyOnHand, 0) - isnull(i.QtyOnBO, 0) - isnull(i.QtyOnSO, 0) - isnull(p.QtyPending, 0)) As QtyAvail " _
         & "FROM vwOPItemLoadFinGood i " _
         & "left outer join vwPendingShipment p on i.ItemKey = p.ItemKey and i.WhseKey = p.WhseKey " _
         & "WHERE i.ItemKey = " & i_lItemKey & " ORDER BY i.WhseKey"

    'Set orst = LoadDiscRst(sSQL)
    
    Set orst = CallSP("spOPItemLoadFinGood", "@ItemKey", i_lItemKey)
    
    Init orst, lWhseKey
    IItem_StatusCode = IItem_StatusCode
    Set orst = Nothing
    Exit Sub
EH:
    ThrowError "ItemFinGood", "Load", "i_lItemKey = " & i_lItemKey & "; lWhseKey = " & lWhseKey
End Sub

'***************************************************************************************

' 10/22/03 Changed scope from public to private as it's not called externally

Private Sub Init(ByRef i_orst As ADODB.Recordset, lWhseKey As Long)
    On Error GoTo EH
    Debug.Print "Init Finished Good " & i_orst.Fields("ItemID").Value
    With m_udtFinGood
        .lCatPage = IIf(IsNull(i_orst.Fields("CatPage").Value), 0, i_orst.Fields("CatPage").Value)
        .sLongDescr = i_orst.Fields("LongDesc").Value
    End With

'HACK HACK HACK  10/30/03 LR
'This function is executed both when adding a new FinGood and when
'loading an existing FinGood. In the later case the code below was
'overwriting nine properties (cost and prices the most important)
'which had been populated from XML at a higher point in the calling
'hierarchy.
'The kludge is to check the BaseItem's ListPrice property.
'If it's zero then we're adding a new FinGood else we're loading an existing one.
'When loading an existing item only four properties (CatPage, LongDescr, WhseBinID,
'ItemInventoryStatus) need to be loaded here (because they aren't persisted in the XML).

    If m_oBaseItem.ListPrice > 0 Then
        'loading existing item
        With m_oBaseItem
            '.WhseBinID = i_oRst.Fields("WhseBinID").value 'PRN#118
            .WhseBinID = GetWhseBinLoc(i_orst, lWhseKey)
            .ItemInventoryStatus = i_orst.Fields("Status").Value
        End With
    Else
        'adding new item
        With m_oBaseItem
            .Cost = i_orst.Fields("RplcmntUnitCost").Value
            .ListPrice = i_orst.Fields("ListPrice").Value
            .DealerPrice = i_orst.Fields("DealerPrice").Value
            .WholesalePrice = i_orst.Fields("WholesalePrice").Value
            .Descr = i_orst.Fields("ShortDesc").Value
            .ItemID = i_orst.Fields("ItemID").Value
            .itemkey = i_orst.Fields("ItemKey").Value
            .SageItemType = i_orst.Fields("ItemType").Value
            .VendorKey = i_orst.Fields("PrimaryVendKey").Value
            '.WhseBinID = i_oRst.Fields("WhseBinID").value 'PRN#118
            .WhseBinID = GetWhseBinLoc(i_orst, lWhseKey)
            .ItemInventoryStatus = i_orst.Fields("Status").Value
        End With

        'HACK - HACK - HACK JEJ
        'The the HIGHEST Cost from each warehouse
        'And the first Non-Zero VendorKey

        If lWhseKey > 0 Then
            With i_orst
                .Filter = "WhseKey = " & lWhseKey
                    If Not .EOF Then
                        m_oBaseItem.VendorKey = .Fields("PrimaryVendKey").Value
                        m_oBaseItem.Cost = .Fields("RplcmntUnitCost").Value
                    End If
                .Filter = adFilterNone

                If m_oBaseItem.VendorKey = 0 Then
                    .Filter = "WhseKey <> " & lWhseKey
                    If Not .EOF Then .MoveFirst

                    Do While Not .EOF
                        If .Fields("PrimaryVendKey").Value <> 0 Then
                            m_oBaseItem.VendorKey = .Fields("PrimaryVendKey").Value
                            m_oBaseItem.Cost = .Fields("RplcmntUnitCost").Value
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                    .Filter = adFilterNone
                End If
    ''' replaced JEJ
    '''        With i_orst
    '''            .MoveNext
    '''            Do While Not .EOF
    '''                If m_oBaseItem.Cost < i_orst.Fields("RplcmntUnitCost").Value Then
    '''                    m_oBaseItem.Cost = i_orst.Fields("RplcmntUnitCost").Value
    '''                End If
    '''
    '''                If m_oBaseItem.VendorKey = 0 And i_orst.Fields("PrimaryVendKey").Value <> 0 Then
    '''                    m_oBaseItem.VendorKey = i_orst.Fields("PrimaryVendKey").Value
    '''                End If
    '''                .MoveNext
    '''            Loop
            End With
            i_orst.MoveFirst
        End If
        'End Hack
    End If

'TODO: Is there a problem down this rabbit hole? Follow the recordset. 10/30/03 LR
'See Inventory.Init()
'   m_udtInv.dVendRplcmntUnitCost = .Fields("RplcmntUnitCost").Value

    m_oInv.Init i_orst, i_orst.Fields("ItemKey").Value
    Exit Sub
EH:
    ThrowError "ItemFinGood", "Init", ""
End Sub


'4/8/05 LR added
'
Private Function GetWhseBinLoc(ByRef orst As ADODB.Recordset, ByVal lWhseKey As Long) As String
    GetWhseBinLoc = vbNullString
    With orst
        orst.Filter = "WhseKey = " & lWhseKey
        If Not .EOF Then
            GetWhseBinLoc = .Fields("WhseBinID").Value
        End If
        .Filter = adFilterNone
    End With
End Function


Public Property Get CatPage() As Integer
    CatPage = m_udtFinGood.lCatPage
End Property


Public Property Get Inventories() As Inventories
    Set Inventories = m_oInv
End Property

Private Sub Refresh()
    If Not m_bLoading Then
        m_bLoading = True
        IItem_StatusCode = m_oBaseItem.StatusCode 'ensure valid status
        RaiseEvent Change
        m_bLoading = False
    End If
End Sub

Private Property Let IItem_ItemInventoryStatus(ByVal RHS As enumItemInventoryStatus)
    m_oBaseItem.ItemInventoryStatus = RHS
End Property

Private Property Get IItem_ItemInventoryStatus() As enumItemInventoryStatus
    IItem_ItemInventoryStatus = m_oBaseItem.ItemInventoryStatus
End Property


'4/7/05 LR add?
'
'Public Property Get WhseBinID() As String
'End Property
'
'
'Public Property Let WhseBinID(ByVal sNewValue As String)
'End Property





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemBTOKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IItem

Public Event Change()

Private Type ItemBTOKitProps
    lCatPage As Integer
    sLongDescr As String * 255
End Type

Private m_udtBTOKit As ItemBTOKitProps
Private m_udtBTOKitBackup As ItemBTOKitProps

Private m_Components As Components   'Tracks Constituent Parts
Private m_Inventories As Inventories   'Exposes theoritical Qty on Hand

Private m_bLoading As Boolean

Private m_oBaseItem As Item

Public Property Get BaseClass() As Item
    Set BaseClass = m_oBaseItem
End Property


Public Sub Class_Initialize()
    Set m_oBaseItem = New Item
    Set m_Components = New Components
    Set m_Inventories = New Inventories
    m_oBaseItem.StatusCode = iscAuthorize
End Sub


Public Sub Class_Terminate()
    Set m_oBaseItem = Nothing
    Set m_Components = Nothing
    Set m_Inventories = Nothing
End Sub


'Begin IItem Interface procedures
'================================


Private Function IItem_XMLCommit(i_lWhseKey As Long, _
                    ByRef rstWhses As ADODB.Recordset, _
                    ByRef rstVendor As ADODB.Recordset) As JDMPDXML.XMLNode
    
    Set IItem_XMLCommit = m_oBaseItem.XmlCommit(ItemTypeCode.itBTOKit, i_lWhseKey, rstWhses, rstVendor)
    
End Function


Private Function IItem_Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False _
) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        .Tag = "BTOKit"
        .Add "CatPage", m_udtBTOKit.lCatPage
        .Add "LongDescr", Trim(m_udtBTOKit.sLongDescr)
        .AddChild m_oBaseItem.Export(bPubNotes, bPrivNotes, bPurchNotes, m_Components.Export)
    End With

    Set IItem_Export = oXML
End Function


Private Sub IItem_Import(ByVal i_oXML As JDMPDXML.XMLNode, Optional lWhseKey As Long)
    With i_oXML
        If .Tag <> "BTOKit" Then
            Err.Raise -1, "BTOKit.Import", "Unexpected initial XML tag"
        End If
        m_udtBTOKit.lCatPage = .Item("CatPage").Value
        m_udtBTOKit.sLongDescr = .Item("LongDescr").Value
        m_oBaseItem.Import .Item("Item")
        Load m_oBaseItem.ItemKey, lWhseKey
    End With
End Sub


'why public?
Public Sub IItem_Backup()
    m_oBaseItem.Backup
    m_udtBTOKitBackup = m_udtBTOKit
End Sub


'why public?
Public Sub IItem_Restore()
    m_oBaseItem.Restore
    m_udtBTOKit = m_udtBTOKitBackup
End Sub


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
        If IItem_Cost <= 0 Then
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


Public Function IItem_IsValidStatusCode(ByVal eStatus As ItemStatusCode) As Boolean
    Select Case eStatus
        Case ItemStatusCode.iscResearch
            IItem_IsValidStatusCode = True
        
        Case ItemStatusCode.iscQuote
            IItem_IsValidStatusCode = (IItem_Cost <= 0)
    
        Case ItemStatusCode.iscAuthorize, ItemStatusCode.iscReadyToCommit
            IItem_IsValidStatusCode = (IItem_Cost > 0)
            
        Case Else
            IItem_IsValidStatusCode = False
    End Select
End Function

'================== REVIEW THESE BELOW

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


Public Property Let IItem_ItemID(ByVal RHS As String)
    m_oBaseItem.ItemID = RHS
End Property

Public Property Get IItem_ItemID() As String
    IItem_ItemID = m_oBaseItem.ItemID
End Property


Public Property Let IItem_ItemKey(ByVal RHS As Long)
    m_oBaseItem.ItemKey = RHS
End Property

Public Property Get IItem_ItemKey() As Long
    IItem_ItemKey = m_oBaseItem.ItemKey
End Property


Public Property Let IItem_SageItemType(ByVal RHS As Long)
    m_oBaseItem.SageItemType = RHS
End Property

Public Property Get IItem_SageItemType() As Long
    IItem_SageItemType = m_oBaseItem.SageItemType
End Property


Public Property Get IItem_OPItemType() As ItemTypeCode
    IItem_OPItemType = itBTOKit
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
    With m_Inventories.WhseInventory(lWhseKey)
        'IItem_QtyAvail = .QtyOnHand - .QtyOnSO - .QtyOnBO - .QtyOnPS
        IItem_QtyAvail = .QtyAvail
    End With
End Property



Public Property Get IItem_QtyOnHand(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_QtyOnHand = m_Inventories.WhseInventory(lWhseKey).QtyOnHand
End Property

Public Property Get IItem_QtyOnSO(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_QtyOnSO = m_Inventories.WhseInventory(lWhseKey).QtyOnSO
End Property

Public Property Get IItem_QtyOnPO(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_QtyOnPO = m_Inventories.WhseInventory(lWhseKey).QtyOnPO
End Property

Public Property Get IItem_QtyOnBO(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_QtyOnBO = m_Inventories.WhseInventory(lWhseKey).QtyOnBO
End Property

Public Property Get IItem_InConflict(lWhseKey As Long) As Variant
    On Error Resume Next
    IItem_InConflict = m_Components.InConflict(lWhseKey)
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


Public Sub Load(i_lItemKey As Long, Optional lWhseKey As Long)
    Dim cmd As ADODB.Command
    Dim orst As ADODB.Recordset

    On Error GoTo EH
    
    If i_lItemKey <= 0 Then
        Debug.Print "ItemBTOKit.Load - invalid key ignored: " & i_lItemKey
        Exit Sub
    End If
    
    Set cmd = CreateCommandSP("spOPItemLoadBTOKit")
    cmd.Parameters("@ItemKey").Value = i_lItemKey
    
    Set orst = New ADODB.Recordset
    orst.Open cmd
    
    If Not orst.EOF Then
    
        With m_udtBTOKit
            .lCatPage = orst.Fields("KitCatPage").Value
            .sLongDescr = orst.Fields("KitLongDesc").Value
        End With
    
        'See the comments in ItemFinGood.Init()
        'A similar fix is required here.
    
        If m_oBaseItem.ListPrice > 0 Then
            'loading existing item
            m_oBaseItem.ItemInventoryStatus = orst.Fields("Status").Value
        Else
            'adding new item
            With m_oBaseItem
                .Cost = orst.Fields("RplcmntUnitCost").Value
                'Cost = 0 if there is not a '$' Kit; in this case, roll-up cost
                If .Cost = 0 Then
                    .Cost = CalcCost(orst)
                End If
    
                If Not IsMissing(lWhseKey) Then
                    If lWhseKey > 0 Then
                        orst.Filter = "WhseKey = " & lWhseKey
                        If Not orst.EOF Then
                            .VendorKey = orst.Fields("PrimaryVendKey").Value
                            If .VendorKey = 0 Then
                                .VendorKey = orst.Fields("CompVendKey").Value
                            End If
                        End If
                        orst.Filter = adFilterNone
                    End If
                Else
                    .VendorKey = orst.Fields("PrimaryVendKey").Value
                    If .VendorKey = 0 Then
                        .VendorKey = orst.Fields("CompVendKey").Value
                    End If
                End If
                    
                .ListPrice = orst.Fields("KitListPrice").Value
                .DealerPrice = orst.Fields("KitDealerPrice").Value
                .WholesalePrice = orst.Fields("KitWholesalePrice").Value
                .Descr = orst.Fields("KitShortDesc").Value
                .ItemID = orst.Fields("KitItemID").Value
                .ItemKey = orst.Fields("KitItemKey").Value
                .SageItemType = orst.Fields("KitItemType").Value
                .ItemInventoryStatus = orst.Fields("Status").Value
            End With
        End If
        
        If IsMissing(lWhseKey) Then
            orst.Filter = "CompWhseKey=" & orst.Fields("CompWhseKey").Value
        Else
            orst.Filter = "CompWhseKey=" & CLng(lWhseKey)
        End If
        m_Components.Init orst
        orst.Filter = adFilterNone
        
        CalcInventories

    End If
    
    CloseRst orst
    Set cmd = Nothing
    Set orst = Nothing
    
    IItem_StatusCode = IItem_StatusCode
    
    Exit Sub
EH:
    ThrowError "ItemBTOKit", "Load", "i_lItemKey = " & i_lItemKey & "; lWhseKey = " & lWhseKey
End Sub


Private Function CalcCost(ByRef i_orst As ADODB.Recordset) As Double
    Dim dTemp As Double

    If i_orst.EOF Then Exit Function
    If IsNull(i_orst.Fields("CompWhseKey").Value) Then Exit Function

    'JEJ 6/7/01 To roll up component costs for **first** warehouse

    i_orst.Filter = "CompWhseKey=" & i_orst.Fields("CompWhseKey").Value
    Do While Not i_orst.EOF
        dTemp = dTemp + (i_orst.Fields("CompVendCost").Value * i_orst.Fields("CompItemQty").Value)
        i_orst.MoveNext
    Loop
    CalcCost = dTemp
    i_orst.Filter = adFilterNone
End Function


Public Property Get CatPage() As Integer
    CatPage = m_udtBTOKit.lCatPage
End Property


Public Property Get Components() As Components
    Set Components = m_Components
End Property


Public Property Get Inventories() As Inventories
    Set Inventories = m_Inventories
End Property


Private Sub CalcInventories()
    Dim i As Long
    Dim oInv As Inventory
        
    'FYI: these constants come out of the Config.ini file
    'they are the WhseKeys. Choice & order matters.
    For i = g_MPKWhseKey To g_STLWhseKey
        Set oInv = New Inventory
        oInv.SetBTOKitInv i, _
            m_Components.CalcQty(i, "QtyOnHand"), _
            m_Components.CalcQty(i, "QtyOnSO"), _
            m_Components.CalcQty(i, "QtyOnBO"), _
            m_Components.CalcQty(i, "QtyOnPO"), _
            m_Components.CalcQty(i, "QtyAvail")
        m_Inventories.Items.Add oInv
        Set oInv = Nothing
    Next
End Sub


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




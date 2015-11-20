VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemShelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IItem

Private Type ItemShelfProps
    dDepth As Double
    dWidth As Double
    lFinishID As Long
    sFinishText As String * 15
    lFrameID As Long
    sFrameText As String * 10
    lOptions As Long
    sRemarks As String * 255
    sCustType As String * 12
End Type

Private Const k_lShf516 = 1
Private Const k_lShf38 = 2

Private Const k_lShfZinc = 1
Private Const k_lShfChrome = 2
Private Const k_lShfStainless = 3
Private Const k_lShfEpoxy = 4
Private Const k_lShfPlastic = 5
Private Const k_lShfNickle = 6

'ItemShelf
Private Const k_lShfCutOut = 1
Private Const k_lShfStraightLeg = 2
Private Const k_lShfBentLeg = 4
Private Const k_lShfProdStop = 8
Private Const k_lShfSupport = 16

Public Event Change()

Private m_udtShelf As ItemShelfProps
Private m_udtShelfBackup As ItemShelfProps

Private m_bLoading As Boolean

Private m_oBaseItem As Item

'***SOAPI 10/22/08 LR

Public Property Get BaseClass() As Item
    Set BaseClass = m_oBaseItem
End Property


'Begin IItem Interface procedures
'================================

Private Function IItem_XMLCommit(i_lWhseKey As Long, _
                    ByRef rstWhses As ADODB.Recordset, _
                    ByRef rstVendor As ADODB.Recordset) As JDMPDXML.XMLNode

    PrepareToCommitItem m_oBaseItem, "SHF", i_lWhseKey
    Set IItem_XMLCommit = m_oBaseItem.XmlCommit(itWireShelf, i_lWhseKey, rstWhses, rstVendor)
    
End Function


Private Function IItem_Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False _
) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        .Tag = "Shelf"
        .Add "Depth", m_udtShelf.dDepth
        .Add "Width", m_udtShelf.dWidth
        .Add "FinishID", m_udtShelf.lFinishID
        .Add "FinishText", Trim(m_udtShelf.sFinishText)
        .Add "FrameID", m_udtShelf.lFrameID
        .Add "FrameText", Trim(m_udtShelf.sFrameText)
        .Add "Options", m_udtShelf.lOptions
        .Add "Remarks", Trim(m_udtShelf.sRemarks)
        .Add "CustType", Trim(m_udtShelf.sCustType)
        .AddChild m_oBaseItem.Export(bPubNotes, bPrivNotes, bPurchNotes)
    End With
    
    Set IItem_Export = oXML
End Function


Private Sub IItem_Import(ByVal i_oXML As JDMPDXML.XMLNode, Optional lWhseKey As Long)
    With i_oXML
        If .Tag <> "Shelf" Then
            Err.Raise -1, "Shelf.Import", "Unexpected initial XML tag"
        End If
        
        m_udtShelf.dDepth = .Item("Depth").value
        m_udtShelf.dWidth = .Item("Width").value
        m_udtShelf.lFinishID = .Item("FinishID").value
        m_udtShelf.sFinishText = .Item("FinishText").value
        m_udtShelf.lFrameID = .Item("FrameID").value
        m_udtShelf.sFrameText = .Item("FrameText").value
        m_udtShelf.lOptions = .Item("Options").value
        m_udtShelf.sRemarks = .Item("Remarks").value
        m_udtShelf.sCustType = .Item("CustType").value
        m_oBaseItem.Import .Item("Item")
    End With
End Sub


Public Sub Class_Initialize()
    Set m_oBaseItem = New Item
    m_oBaseItem.SageItemType = 1 'Sage Non-Inventory Misc Item
    
    With m_udtShelf
        .sFinishText = "Zinc"
        .lFinishID = 1
        .sFrameText = "5/16" & """"
        .lFrameID = 1
        .lOptions = 0
    End With
    
    
    IItem_StatusCode = iscResearch
End Sub

Public Sub Class_Terminate()
    Set m_oBaseItem = Nothing
End Sub

Public Sub IItem_Backup()
    m_oBaseItem.Backup
    m_udtShelfBackup = m_udtShelf
End Sub


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

'09/24/02       TeddyX
'Add new property of research status
Private Property Let IItem_ResearchStatus(ByVal RHS As ItemResearchStatus)
'    If IItem_StatusCode = iscResearch Then
        m_oBaseItem.ResearchStatus = RHS
'    Else
'        m_oBaseItem.ResearchStatus = iscResearchEmpty
'    End If
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


Public Property Get IItem_BackNegotiatedPrice() As Double
    IItem_BackNegotiatedPrice = m_oBaseItem.BackNegotiatedPrice
End Property


Public Property Let IItem_BackNegotiatedPrice(ByVal RHS As Double)
    m_oBaseItem.BackNegotiatedPrice = RHS
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
    
    '09/25/02       TeddyX
    'Set the default research default status as irsNeedResearch
    'if item status code is iscResearch
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
            'IItem_IsValidStatusCode = (IItem_Cost <= 0)
            IItem_IsValidStatusCode = True
    
        Case ItemStatusCode.iscAuthorize, ItemStatusCode.iscReadyToCommit
            IItem_IsValidStatusCode = (IItem_Cost > 0)
            
        Case Else
            IItem_IsValidStatusCode = False
    End Select
End Function


'===== REVIEW BELOW

Public Property Let IItem_Descr(ByVal RHS As String)
                                                                    'Ignore - read only for Shelf
End Property

Public Property Get IItem_Descr() As String
    IItem_Descr = m_oBaseItem.Descr
End Property


Private Sub UpdateDescr()
    Dim sTemp As String
    
    With m_udtShelf
        sTemp = Trim(.sFrameText) & " " & Trim(.sFinishText) & " Shf " _
              & Helpers.FormatInches(m_udtShelf.dDepth, """") & "D X " _
              & Helpers.FormatInches(m_udtShelf.dWidth, """") & "W "

        If (.lOptions And k_lShfCutOut) Then
            sTemp = sTemp & "C"
        End If

        If (.lOptions And k_lShfStraightLeg) Then
            sTemp = sTemp & "S"
        End If

        If (.lOptions And k_lShfBentLeg) Then
            sTemp = sTemp & "B"
        End If

        If .lOptions And k_lShfProdStop Then
            sTemp = sTemp & "P"
        End If

        If .lOptions And k_lShfSupport Then
            sTemp = sTemp & "S"
        End If
    End With

    If Len(sTemp) > 40 Then
        sTemp = Replace(sTemp, "Stainless", "S/S")
        sTemp = Replace(sTemp, "Chrome", "Chr")
        sTemp = Replace(sTemp, "Plastic", "Plas")
        sTemp = Replace(sTemp, "Epoxy", "Epx")
    End If
    
    m_oBaseItem.Descr = Left$(sTemp, 40)
End Sub


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
    IItem_OPItemType = itWireShelf
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
                                                                    'Return empty because qty is n/a for this component type
End Property

Public Property Get IItem_QtyOnHand(lWhseKey As Long) As Variant
                                                                    'Return empty because qty is n/a for this component type
End Property

Public Property Get IItem_QtyOnSO(lWhseKey As Long) As Variant
                                                                    'Return empty because qty is n/a for this component type
End Property

Public Property Get IItem_QtyOnPO(lWhseKey As Long) As Variant
                                                                    'Return empty because qty is n/a for this component type
End Property

Public Property Get IItem_QtyOnBO(lWhseKey As Long) As Variant
                                                                    'Return empty because qty is n/a for this component type
End Property

Public Sub IItem_Restore()
    m_oBaseItem.Restore
    m_udtShelf = m_udtShelfBackup
End Sub

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


Public Function LoadFinishCombo(cboFinish As ComboBox)
    m_bLoading = True
    With cboFinish
        .AddItem "Zinc"
        .ItemData(.NewIndex) = k_lShfZinc
        .AddItem "Chrome"
        .ItemData(.NewIndex) = k_lShfChrome
        .AddItem "Stainless"
        .ItemData(.NewIndex) = k_lShfStainless
        .AddItem "Epoxy"
        .ItemData(.NewIndex) = k_lShfEpoxy
        .AddItem "Plastic"
        .ItemData(.NewIndex) = k_lShfPlastic
        .AddItem "Nickle"
        .ItemData(.NewIndex) = k_lShfNickle
        .ListIndex = 0
    End With
    m_bLoading = False
End Function


Public Function LoadFrameCombo(cboFrame As ComboBox)
    m_bLoading = True
    With cboFrame
        .AddItem "5/16"
        .ItemData(.NewIndex) = k_lShf516
        .AddItem "3/8"
        .ItemData(.NewIndex) = k_lShf38
        .ListIndex = 0
    End With
    m_bLoading = False
End Function


Public Property Get Depth() As Double
    Depth = m_udtShelf.dDepth
End Property

Public Property Let Depth(ByVal dNewValue As Double)
    If m_udtShelf.dDepth <> dNewValue Then
        m_udtShelf.dDepth = dNewValue
        CalcCostPrice
        Refresh
    End If
End Property


Public Property Get width() As Double
    width = m_udtShelf.dWidth
End Property

Public Property Let width(ByVal dNewValue As Double)
    If m_udtShelf.dWidth <> dNewValue Then
        m_udtShelf.dWidth = dNewValue
        CalcCostPrice
        Refresh
    End If
End Property


Public Property Get Options() As Long
    Options = m_udtShelf.lOptions
End Property

Public Property Let Options(ByVal lNewValue As Long)
    If m_udtShelf.lOptions <> lNewValue Then
        m_udtShelf.lOptions = lNewValue
        Refresh
    End If
End Property


Public Property Get FrameID() As Long
    FrameID = m_udtShelf.lFrameID
End Property

Public Property Let FrameID(ByVal lNewValue As Long)
    If Not m_bLoading Then
        If m_udtShelf.lFrameID <> lNewValue Then
            m_udtShelf.lFrameID = lNewValue
            Refresh
        End If
    End If
End Property


Public Property Get FinishID() As Long
    FrameID = m_udtShelf.lFrameID
End Property

Public Property Let FinishID(ByVal lNewValue As Long)
    If Not m_bLoading Then
        If m_udtShelf.lFinishID <> lNewValue Then
            m_udtShelf.lFinishID = lNewValue
            Refresh
        End If
    End If
End Property


Public Property Get remarks() As String
    remarks = Trim(m_udtShelf.sRemarks)
End Property

Public Property Let remarks(ByVal sNewValue As String)
    m_udtShelf.sRemarks = sNewValue
End Property


Public Property Get CustType() As String
    CustType = m_udtShelf.sCustType
End Property

Public Property Let CustType(ByVal vNewValue As String)
    m_udtShelf.sCustType = vNewValue
End Property


Public Property Get FinishText() As String
    FinishText = m_udtShelf.sFinishText
End Property

Public Property Let FinishText(ByVal sNewValue As String)
    m_udtShelf.sFinishText = sNewValue
End Property


Public Property Get FrameText() As String
    FrameText = m_udtShelf.sFrameText
End Property

Public Property Let FrameText(ByVal sNewValue As String)
    m_udtShelf.sFrameText = sNewValue
End Property


Private Sub CalcCostPrice()
                                                                'TODO JON: We need to discuss this.  My assumption is that we'll use
                                                                '          the NegotiatedPrice property for dealing with shelf pricing.
End Sub


Private Sub Refresh()
    If Not m_bLoading Then
        m_bLoading = True
        IItem_StatusCode = m_oBaseItem.StatusCode 'ensure valid status
        UpdateDescr
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




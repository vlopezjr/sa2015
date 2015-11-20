VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemGasket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IItem

Private Type ItemGasketProps
    dHeight As Double
    dWidth As Double
    lMaterialID As Long
    sGMPartNbr As String * 15
    lOptions As Long
    sRemarks As String * 255
    sCustType As String * 12
    bIsMagnetic As Boolean
    bIsDart As Boolean
    bIsLA As Boolean
    bIsSL As Boolean
End Type

Public Event Change()


'***GMOD
'Private Const k_lGasThreeSided = 1
'Private Const k_lGasDartToDart = 2
'Private Const k_lGasNoMagHinge = 4
'Private Const k_lGasInverted = 8
'***

Private m_udtGasket As ItemGasketProps
Private m_udtGasketBackup As ItemGasketProps

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

    PrepareToCommitItem m_oBaseItem, "GSK", i_lWhseKey
    Set IItem_XMLCommit = m_oBaseItem.XmlCommit(itMoldedGasket, i_lWhseKey, rstWhses, rstVendor)
    
End Function


Private Function IItem_Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False _
) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    
    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "Gasket"
        .Add "Height", m_udtGasket.dHeight
        .Add "Width", m_udtGasket.dWidth
        .Add "MaterialID", m_udtGasket.lMaterialID
        .Add "GMPartNbr", Trim(m_udtGasket.sGMPartNbr)
        .Add "Options", m_udtGasket.lOptions
        .Add "Remarks", Trim(m_udtGasket.sRemarks)
        .Add "CustType", Trim(m_udtGasket.sCustType)
        .Add "IsMagnetic", m_udtGasket.bIsMagnetic
        .Add "IsDart", m_udtGasket.bIsDart
        .Add "IsLA", m_udtGasket.bIsLA
        .Add "IsSL", m_udtGasket.bIsSL
        .AddChild m_oBaseItem.Export(bPubNotes, bPrivNotes, bPurchNotes)
    End With
    
    Set IItem_Export = oXML
End Function


Private Sub IItem_Import(ByVal i_oXML As JDMPDXML.XMLNode, Optional lWhseKey As Long)
    With i_oXML
        If .Tag <> "Gasket" Then
            Err.Raise -1, "Gasket.Import", "Unexpected initial XML tag"
        End If
        
        m_udtGasket.dHeight = .Item("Height").value
        m_udtGasket.dWidth = .Item("Width").value
        m_udtGasket.lMaterialID = .Item("MaterialID").value
        m_udtGasket.sGMPartNbr = .Item("GMPartNbr").value
        m_udtGasket.lOptions = .Item("Options").value
        m_udtGasket.sRemarks = .Item("Remarks").value
        m_udtGasket.sCustType = .Item("CustType").value
        m_udtGasket.bIsMagnetic = .Item("IsMagnetic").value
        m_udtGasket.bIsDart = .Item("IsDart").value
        m_udtGasket.bIsLA = .Item("IsLA").value
        m_udtGasket.bIsSL = .Item("IsSL").value
        m_oBaseItem.Import .Item("Item")
    End With
End Sub


Public Sub Class_Initialize()
    Set m_oBaseItem = New Item
    m_oBaseItem.SageItemType = 1 'Sage Non-Inventory Misc Item
    
    With m_udtGasket
        .lMaterialID = 0
        .bIsMagnetic = True
    End With
    
    IItem_StatusCode = iscQuote
End Sub


Public Sub Class_Terminate()
    Set m_oBaseItem = Nothing
End Sub


Public Sub IItem_Backup()
    m_oBaseItem.Backup
    m_udtGasketBackup = m_udtGasket
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
    'Set the default research status to be irsNeedResearch if
    'the statuscode is need research
    If m_oBaseItem.StatusCode = iscResearch Then
        If m_oBaseItem.ResearchStatus < irsNeedResearch Then
            m_oBaseItem.ResearchStatus = irsNeedResearch
        End If
    End If
End Property


Public Function IItem_IsValidStatusCode(ByVal eStatus As ItemStatusCode) As Boolean
    Select Case eStatus
        Case ItemStatusCode.iscResearch
            IItem_IsValidStatusCode = False
        
        Case ItemStatusCode.iscQuote
            IItem_IsValidStatusCode = (IItem_Cost <= 0)
    
        Case ItemStatusCode.iscAuthorize, ItemStatusCode.iscReadyToCommit
            IItem_IsValidStatusCode = (IItem_Cost > 0)
            
        Case Else
            IItem_IsValidStatusCode = False
    End Select
End Function


'===== REVIEW BELOW

Public Property Let IItem_Descr(ByVal RHS As String)
                                                                ' ignore - read only for Gasket
End Property

Public Property Get IItem_Descr() As String
    IItem_Descr = m_oBaseItem.Descr
End Property


Private Sub UpdateDescr()
    Dim sTemp As String
    
    With m_udtGasket
        sTemp = Trim(.sGMPartNbr) & " " _
              & Helpers.FormatInches(.dHeight, """") & "x" _
              & Helpers.FormatInches(.dWidth, """") & " "
    
        If (.lOptions And k_lGasketDartToDart) Then
            sTemp = sTemp & "/D"
        End If

        If (.lOptions And k_lGasketInverted) Then
            sTemp = sTemp & "/I"
        End If

'***GMOD
        If (.lOptions And k_lGasketNoMagLHS) Then
            sTemp = sTemp & "/LH"
        End If

        If (.lOptions And k_lGasketNoMagRHS) Then
            sTemp = sTemp & "/RH"
        End If
'***

        If .lOptions And k_lGasketThreeSided Then
            sTemp = sTemp & "/3"
'***GMOD
        Else
            sTemp = sTemp & "/4"
'***
        End If
    End With
    
    sTemp = sTemp & " Gasket"
    If Len(sTemp) > 40 Then
        sTemp = Replace(sTemp, "Gasket", "Gskt")
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
    IItem_OPItemType = itMoldedGasket
End Property


Public Property Let IItem_LineKey(ByVal RHS As Long)
    m_oBaseItem.LineKey = RHS
End Property

Public Property Get IItem_LineKey() As Long
    IItem_LineKey = m_oBaseItem.LineKey
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
    m_udtGasket = m_udtGasketBackup
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



Public Sub LoadMaterialCombo(cboControl As ComboBox)
   If m_udtGasket.bIsMagnetic Then
        g_rstGaskMats.Filter = "IsMagnetic=-1"
    Else
        g_rstGaskMats.Filter = "IsMagnetic=0"
    End If
    
    If Not m_bLoading Then
        m_bLoading = True
        Helpers.LoadCombo cboControl, g_rstGaskMats, "GMPartNbr", "GasketMaterialID", 0, True
        m_bLoading = False
    Else
        Helpers.LoadCombo cboControl, g_rstGaskMats, "GMPartNbr", "GasketMaterialID", 0, True
    End If

    g_rstGaskMats.Filter = adFilterNone
End Sub


Public Property Get Height() As Double
    Height = m_udtGasket.dHeight
End Property

Public Property Let Height(ByVal dNewValue As Double)
    If m_udtGasket.dHeight <> dNewValue Then
        m_udtGasket.dHeight = dNewValue
        CalcCostPrice
        Refresh
    End If
End Property


Public Property Get width() As Double
    width = m_udtGasket.dWidth
End Property

Public Property Let width(ByVal dNewValue As Double)
    If m_udtGasket.dWidth <> dNewValue Then
        m_udtGasket.dWidth = dNewValue
        CalcCostPrice
        Refresh
    End If
End Property


Public Property Get Options() As Long
    Options = m_udtGasket.lOptions
End Property

Public Property Let Options(ByVal lNewValue As Long)
    m_udtGasket.lOptions = lNewValue
    Refresh
End Property


Public Property Get materialId() As Long
    materialId = m_udtGasket.lMaterialID
End Property

Public Property Let materialId(ByVal lNewValue As Long)
    If Not m_bLoading Then
        If m_udtGasket.lMaterialID <> lNewValue Then
            m_udtGasket.lMaterialID = lNewValue
            SetMats
            CalcCostPrice
            Refresh
        End If
    End If
End Property


Public Property Get MoldedBy() As String
    With m_udtGasket
        If .bIsLA Then
            If .bIsSL Then
                MoldedBy = "Los Angeles and St. Louis"
            Else
                MoldedBy = "Los Angeles"
            End If
        Else
            If .bIsSL Then
                MoldedBy = "St. Louis"
            Else
                MoldedBy = ""
            End If
        End If
    End With
End Property


Public Property Get IsMagnetic() As Boolean
    IsMagnetic = m_udtGasket.bIsMagnetic
End Property

Public Property Let IsMagnetic(ByVal bNewValue As Boolean)
    m_udtGasket.bIsMagnetic = bNewValue
    Refresh
End Property


Public Property Get IsDart() As Boolean
    IsDart = m_udtGasket.bIsDart
End Property

Public Property Let IsDart(ByVal bNewValue As Boolean)
    m_udtGasket.bIsDart = bNewValue
End Property


Public Property Get IsLA() As Boolean
    IsLA = m_udtGasket.bIsLA
End Property

Public Property Let IsLA(ByVal bNewValue As Boolean)
    m_udtGasket.bIsLA = bNewValue
End Property


Public Property Get IsSL() As Boolean
    IsSL = m_udtGasket.bIsSL
End Property

Public Property Let IsSL(ByVal bNewValue As Boolean)
    m_udtGasket.bIsSL = bNewValue
End Property


Public Property Get remarks() As String
    remarks = Trim(m_udtGasket.sRemarks)
End Property

Public Property Let remarks(ByVal sNewValue As String)
    m_udtGasket.sRemarks = sNewValue
End Property


Public Property Get CustType() As String
    CustType = m_udtGasket.sCustType
End Property

Public Property Let CustType(ByVal sNewValue As String)
    m_udtGasket.sCustType = sNewValue
End Property


Public Sub SetMats()
    With m_udtGasket
        If .lMaterialID = 0 Then
            .bIsLA = True
            .bIsSL = True
            .bIsDart = True
            .sGMPartNbr = ""
        Else
             g_rstGaskMats.Filter = "GasketMaterialID=" & CStr(.lMaterialID)
            .bIsLA = g_rstGaskMats.Fields("IsLA").value
            .bIsSL = g_rstGaskMats.Fields("IsSL").value
            .bIsDart = g_rstGaskMats.Fields("IsDart").value
            .sGMPartNbr = g_rstGaskMats.Fields("GMPartNbr").value
            g_rstGaskMats.Filter = adFilterNone
        End If
    End With
End Sub



Private Sub CalcCostPrice()
    With m_oBaseItem
        If Not Valid Then
            .Cost = 0
            .ListPrice = 0
            .DealerPrice = 0
            .WholesalePrice = 0
            
        ElseIf m_udtGasket.dHeight < 31 And m_udtGasket.dWidth < 31 Then
            .Cost = 4.5
            .ListPrice = 33             '30.4
            .DealerPrice = 20.5         '19
            .WholesalePrice = 16.5      '15
            
        ElseIf m_udtGasket.dHeight < 51 And m_udtGasket.dWidth < 51 Then
            .Cost = 6
            .ListPrice = 43             '39.2
            .DealerPrice = 26.5         '24.5
            .WholesalePrice = 21        '19.5
            
        ElseIf m_udtGasket.dHeight < 71 And m_udtGasket.dWidth < 71 Then
            .Cost = 8.5
            .ListPrice = 51             '47.2
            .DealerPrice = 32           '29.5
            .WholesalePrice = 26.5      '24.5

        ElseIf m_udtGasket.dHeight < 81 And m_udtGasket.dWidth < 81 Then
            .Cost = 10
            .ListPrice = 59             '54.4
            .DealerPrice = 37           '34
            .WholesalePrice = 32        '29.5
            
        Else
            .Cost = 12
            .ListPrice = 67
            .DealerPrice = 46
            .WholesalePrice = 37.5
        End If

        'OEM pricing is high for this material
        If Trim$(m_udtGasket.sGMPartNbr) = "02-14160" Then
        .ListPrice = .ListPrice * 1.45              '1.65
        .DealerPrice = .DealerPrice * 1.33          '1.53
        .WholesalePrice = .WholesalePrice * 1.45    '1.62
        End If
        
    End With
End Sub


Private Function Valid() As Boolean
    Dim MaxHeight As Integer
    With m_udtGasket
        '8/14/13 MOD
        If .bIsMagnetic Then
            MaxHeight = g_MaxHeightMagnetic
        Else
            MaxHeight = g_MaxHeightCompression
        End If
        If .dHeight <= 0 Or .dHeight > MaxHeight Then Exit Function
        If .dWidth <= 0 Or .dWidth > MaxHeight Then Exit Function
        '---
        If .lMaterialID <= 0 Then Exit Function
    End With
    Valid = True
End Function


Private Sub Refresh()
    If Not m_bLoading Then
        m_bLoading = True
        IItem_StatusCode = m_oBaseItem.StatusCode 'in case of status change
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



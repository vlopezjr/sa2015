VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemWWire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IItem

Private Type ItemWWireProps
    lFeet As Long
    lInches As Long
    lTotalInches As Long
    lVoltage As Long
    lDoorHeight As Long
    lDoorWidth As Long
    bIsThreeSided As Boolean
    bIsSinglePass As Boolean
    dWattsPerFoot As Double
    dAmperage As Double
    dOhmsPerFoot As Double
    sCustType As String * 12
    lWhseKey As Long
End Type

Public Event Change()
Private m_bLoading As Boolean

Private m_udtWWire As ItemWWireProps
Private m_udtWWireBackup As ItemWWireProps

Private m_arWires() As Double
Private m_arSafeWires() As Double
Private m_lWireCount As Double

Private m_oBaseItem As Item

'***SOAPI 10/22/08 LR

Public Property Get BaseClass() As Item
    Set BaseClass = m_oBaseItem
End Property


'Private Const k_dFixedCostPerWire = 2.5
'Private Const k_dCostPerFoot = 0.09
'Private Const k_lWireCount = 20
'Private Const k_lMinSafeWPF = 2
'Private Const k_lMaxSafeWPF = 12
'Private Const k_lBestSingleWPF = 6
'Private Const k_lBestDoubleWPF = 3


'================================
'Begin IItem Interface procedures


Private Function IItem_XMLCommit(i_lWhseKey As Long, _
                    ByRef rstWhses As ADODB.Recordset, _
                    ByRef rstVendor As ADODB.Recordset) As JDMPDXML.XMLNode

    PrepareToCommitItem m_oBaseItem, "WWR", i_lWhseKey
    Set IItem_XMLCommit = m_oBaseItem.XmlCommit(itWarmerWire, i_lWhseKey, rstWhses, rstVendor)

End Function


Private Function IItem_Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False _
) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        .Tag = "WWire"
        .Add "Feet", m_udtWWire.lFeet
        .Add "Inches", m_udtWWire.lInches
        .Add "TotalInches", m_udtWWire.lTotalInches
        .Add "Voltage", m_udtWWire.lVoltage
        .Add "DoorHeight", m_udtWWire.lDoorHeight
        .Add "DoorWidth", m_udtWWire.lDoorWidth
        .Add "IsThreeSided", m_udtWWire.bIsThreeSided
        .Add "IsSinglePass", m_udtWWire.bIsSinglePass
        .Add "WattsPerFoot", m_udtWWire.dWattsPerFoot
        .Add "Amperage", m_udtWWire.dAmperage
        .Add "OhmsPerFoot", m_udtWWire.dOhmsPerFoot
        .Add "CustType", Trim(m_udtWWire.sCustType)
        .AddChild m_oBaseItem.Export(bPubNotes, bPrivNotes, bPurchNotes)
    End With
    
    Set IItem_Export = oXML
End Function


Private Sub IItem_Import(ByVal i_oXML As JDMPDXML.XMLNode, Optional lWhseKey As Long)
    With i_oXML
        If .Tag <> "WWire" Then
            Err.Raise -1, "WWire.Import", "Unexpected initial XML tag"
        End If
        
        m_udtWWire.lFeet = .Item("Feet").value
        m_udtWWire.lInches = .Item("Inches").value
        m_udtWWire.lTotalInches = .Item("TotalInches").value
        m_udtWWire.lVoltage = .Item("Voltage").value
        m_udtWWire.lDoorHeight = .Item("DoorHeight").value
        m_udtWWire.lDoorWidth = .Item("DoorWidth").value
        m_udtWWire.bIsThreeSided = .Item("IsThreeSided").value
        m_udtWWire.bIsSinglePass = .Item("IsSinglePass").value
        m_udtWWire.dWattsPerFoot = .Item("WattsPerFoot").value
        m_udtWWire.dAmperage = .Item("Amperage").value
        m_udtWWire.sCustType = .Item("CustType").value
        '10/30/2002         Teddy
        'Get and load the safe wires here
        m_udtWWire.lWhseKey = lWhseKey
        CalcSafeWires
        m_udtWWire.dOhmsPerFoot = .Item("OhmsPerFoot").value
        m_oBaseItem.Import .Item("Item"), lWhseKey
    End With
End Sub


Public Sub Class_Initialize()
    Set m_oBaseItem = New Item
    m_oBaseItem.SageItemType = 1 'Sage Non-Inventory Misc Item
    
    With m_udtWWire
        .lVoltage = 115
        .bIsThreeSided = False
        .dAmperage = 0
        .dOhmsPerFoot = 0
        .dWattsPerFoot = 0
        .bIsSinglePass = True
        .sCustType = ""
    End With
    'LoadWires
    IItem_StatusCode = iscResearch
End Sub


Public Sub Class_Terminate()
    Set m_oBaseItem = Nothing
End Sub


Public Sub IItem_Backup()
    m_oBaseItem.Backup
    m_udtWWireBackup = m_udtWWire
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


Public Property Get IItem_StandardPrice() As Double
    IItem_StandardPrice = m_oBaseItem.StandardPrice
End Property


Public Property Get IItem_EffectivePrice() As Double
    IItem_EffectivePrice = m_oBaseItem.EffectivePrice
End Property


Public Property Get IItem_NegotiatedPricing() As Boolean
    IItem_NegotiatedPricing = m_oBaseItem.NegotiatedPricing
End Property


Public Property Get IItem_BackNegotiatedPrice() As Double
    IItem_BackNegotiatedPrice = m_oBaseItem.BackNegotiatedPrice
End Property


Public Property Let IItem_BackNegotiatedPrice(ByVal RHS As Double)
    m_oBaseItem.BackNegotiatedPrice = RHS
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
                                                                    'Item is a read-only property for warmer wires.  Consider refining IItem to account for that
End Property

Public Property Get IItem_Descr() As String
    IItem_Descr = m_oBaseItem.Descr
End Property


' Called by Refresh

Private Sub UpdateDescr()
    Dim sPass As String
    
    With m_udtWWire
        If .bIsSinglePass Then
            sPass = "S"
        Else
            sPass = "D"
        End If
        m_oBaseItem.Descr = .lTotalInches & """" & " Length of " & .dOhmsPerFoot & " Ohms/ft. Wire (" & .lVoltage & "V)" & " " & sPass
    End With
End Sub


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
    IItem_OPItemType = itWarmerWire
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
    m_udtWWire = m_udtWWireBackup
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


Public Property Get lWhseKey() As Long
    lWhseKey = m_udtWWire.lWhseKey
End Property

Public Property Let lWhseKey(ByVal lNewValue As Long)
    m_udtWWire.lWhseKey = lNewValue
End Property

Public Property Get TotalInches() As Long
    TotalInches = m_udtWWire.lTotalInches
End Property

Public Property Let TotalInches(ByVal vNewValue As Long)
    If vNewValue <> m_udtWWire.lTotalInches Then
        m_udtWWire.lTotalInches = vNewValue
        CalcLength "TotalInches"
    End If
End Property


Public Property Get Inches() As Long
    Inches = m_udtWWire.lInches
End Property

Public Property Let Inches(ByVal vNewValue As Long)
    If vNewValue <> m_udtWWire.lInches Then
        m_udtWWire.lInches = vNewValue
        CalcLength "Inches"
    End If
End Property


Public Property Get Feet() As Long
    Feet = m_udtWWire.lFeet
End Property

Public Property Let Feet(ByVal vNewValue As Long)
    If vNewValue <> m_udtWWire.lFeet Then
        m_udtWWire.lFeet = vNewValue
        CalcLength "Feet"
    End If
End Property


Public Property Get Voltage() As Long
    Voltage = m_udtWWire.lVoltage
End Property

Public Property Let Voltage(ByVal vNewValue As Long)
    m_udtWWire.lVoltage = vNewValue
    CalcLength "Voltage"
End Property


Public Property Get DoorHeight() As Long
    DoorHeight = m_udtWWire.lDoorHeight
End Property

Public Property Let DoorHeight(ByVal vNewValue As Long)
    If vNewValue <> m_udtWWire.lDoorHeight Then
        m_udtWWire.lDoorHeight = vNewValue
        CalcLength "DoorHeight"
    End If
End Property


Public Property Get DoorWidth() As Long
    DoorWidth = m_udtWWire.lDoorWidth
End Property

Public Property Let DoorWidth(ByVal vNewValue As Long)
    If vNewValue <> m_udtWWire.lDoorWidth Then
        m_udtWWire.lDoorWidth = vNewValue
        CalcLength "DoorWidth"
    End If
End Property


Public Property Get IsThreeSided() As Boolean
    IsThreeSided = m_udtWWire.bIsThreeSided
End Property
    
Public Property Let IsThreeSided(ByVal vNewValue As Boolean)
    m_udtWWire.bIsThreeSided = vNewValue
    m_udtWWire.bIsSinglePass = Not vNewValue
    CalcLength "IsThreeSided"
End Property


Public Property Get IsSinglePass() As Boolean
    IsSinglePass = m_udtWWire.bIsSinglePass
End Property
    
Public Property Let IsSinglePass(ByVal vNewValue As Boolean)
    m_udtWWire.bIsSinglePass = vNewValue
    CalcLength "IsSinglePass"
End Property


Public Property Get Amperage() As Double
    Dim lWattage As Long
        
    With m_udtWWire
        If .dOhmsPerFoot = 0 Or .lFeet = 0 Or .lVoltage = 0 Then
            Amperage = 0
            Exit Property
        End If
        lWattage = (.lVoltage * .lVoltage) / (.dOhmsPerFoot * .lFeet)
        Amperage = lWattage / .lVoltage
    End With
End Property

Public Property Let Amperage(ByVal vNewValue As Double)
    m_udtWWire.dAmperage = vNewValue
End Property


Public Property Get WattsPerFoot() As Double
    'WattsPerFoot = m_udtWWire.dWattsPerFoot
    WattsPerFoot = CalcWattsPerFoot(m_udtWWire.dOhmsPerFoot)
End Property

Public Property Let WattsPerFoot(ByVal vNewValue As Double)
    m_udtWWire.dWattsPerFoot = vNewValue
End Property


Public Property Get OhmsPerFoot() As Double
    OhmsPerFoot = m_udtWWire.dOhmsPerFoot
End Property

Public Property Let OhmsPerFoot(ByVal vNewValue As Double)
    If m_udtWWire.dOhmsPerFoot <> vNewValue Then
        m_udtWWire.dOhmsPerFoot = vNewValue
        Refresh
    End If
End Property


Public Property Let CustType(ByVal vNewValue As String)
    m_udtWWire.sCustType = vNewValue
End Property


' Called by CalcSafeWires
' fill array m_arWires with the OhmsPerFoot values in stock for a specific warehouse

Private Sub LoadWires()
    If g_rstWarmerWire Is Nothing Or g_rstWarmerWire.EOF Then Exit Sub
    
    Dim i As Long
    g_rstWarmerWire.Filter = "WhseKey=" & m_udtWWire.lWhseKey & " and InStock <> 0 "
    If Not g_rstWarmerWire.EOF Then
        m_lWireCount = g_rstWarmerWire.RecordCount
        ReDim m_arWires(1 To m_lWireCount) As Double
        g_rstWarmerWire.MoveFirst
        i = 1
        While Not g_rstWarmerWire.EOF
            m_arWires(i) = g_rstWarmerWire.Fields("OhmsPerFoot").value
            i = i + 1
            g_rstWarmerWire.MoveNext
        Wend
    Else
        m_lWireCount = 0
    End If
    g_rstWarmerWire.Filter = adFilterNone
    'ReDim m_arWires(1 To g_WireCount) As Double
    
    'm_arWires(1) = 0.22
    'm_arWires(2) = 1.25
    'm_arWires(3) = 1.75
    'm_arWires(4) = 2.57
    'm_arWires(5) = 3
    'm_arWires(6) = 4
    'm_arWires(7) = 4.5
    'm_arWires(8) = 7
    'm_arWires(9) = 9
    'm_arWires(10) = 12
    'm_arWires(11) = 14
    'm_arWires(12) = 16.5
    'm_arWires(13) = 18
    'm_arWires(14) = 22
    'm_arWires(15) = 32
    'm_arWires(16) = 41
    'm_arWires(17) = 44
    'm_arWires(18) = 73
    'm_arWires(19) = 178
    'm_arWires(20) = 265
End Sub


Private Sub CalcLength(i_sPropName)
    With m_udtWWire
        Select Case i_sPropName
        Case "TotalInches"
            'use total inches to calc feet/inches
            'delete any data in door size
            .lFeet = Int(.lTotalInches / 12)
            .lInches = Int(.lTotalInches Mod 12)
            .lDoorHeight = 0
            .lDoorWidth = 0

        Case Is = "Inches", "Feet"
            'use feet/inches to calc total inches
            'delete any data in door size
            .lTotalInches = (.lFeet * 12) + .lInches
            .lDoorHeight = 0
            .lDoorWidth = 0
        
        Case Is = "IsSinglePass"
            'signal to the user that door size is irrelevant
            .lDoorHeight = 0
            .lDoorWidth = 0
        
        Case Is = "DoorHeight", "DoorWidth", "IsThreeSided"
            'Use door dimensions to calc totalinches and feet / inches
            If .lDoorHeight = 0 Or .lDoorWidth = 0 Then
                .lTotalInches = 0
                .lFeet = 0
                .lInches = 0
            Else
                If .bIsThreeSided Then
                    .lTotalInches = (4 * .lDoorHeight) + (2 * .lDoorWidth)
                Else
                    .lTotalInches = (2 * .lDoorHeight) + (2 * .lDoorWidth)
                End If
                .lFeet = Int(.lTotalInches / 12)
                .lInches = Int(.lTotalInches Mod 12)
            End If
        Case Is = "Voltage"
        End Select
    End With
    
    'Calc Cost and Price
    CalcCostPrice
    
    'Calc available wires
    CalcSafeWires
        
    Refresh
    'Spec change loads the lengths, price/cost info and the list box on the client

    'set wire to BEST available - calling property to raise change wire event
    OhmsPerFoot = m_udtWWire.dOhmsPerFoot
End Sub


Private Sub CalcCostPrice()
    With m_oBaseItem
        .Cost = g_FixedCostPerWire + (m_udtWWire.lFeet * g_CostPerFoot)

        Select Case (m_udtWWire.lTotalInches / 12)
        Case Is <= 0
            .Cost = 0
            .ListPrice = 0
            .DealerPrice = 0
            .WholesalePrice = 0

        Case Is < 11
            .ListPrice = 17.5       '14.4
            .DealerPrice = 11       '9
            .WholesalePrice = 10    '8
   
        Case Is < 21
            .ListPrice = 21         '17.6
            .DealerPrice = 13       '11
            .WholesalePrice = 12    '10
   
        Case Is < 31
            .ListPrice = 25         '20.8
            .DealerPrice = 15       '13
            .WholesalePrice = 14    '12
   
        Case Is < 41
            .ListPrice = 28         '23.2
            .DealerPrice = 18       '15
            .WholesalePrice = 17    '14
   
        Case Is < 51
            .ListPrice = 31         '25.6
            .DealerPrice = 20       '17
            .WholesalePrice = 19    '16
   
        Case Else
            .ListPrice = (m_udtWWire.lTotalInches / 12) * 0.62          '0.54
            .DealerPrice = (m_udtWWire.lTotalInches / 12) * 0.4         '0.34
            .WholesalePrice = (m_udtWWire.lTotalInches / 12) * 0.38     '0.32
        End Select
    End With
End Sub


Private Function CalcWattsPerFoot(OhmsPerFoot As Double) As Double
    Dim dTemp As Double
    Dim dFeet As Double
    
    On Error GoTo ErrorHandler

    With m_udtWWire

'PRN 615
        dFeet = .lTotalInches / 12
        
        If OhmsPerFoot = 0 Or .lTotalInches = 0 Then
            WattsPerFoot = 0
            Exit Function
        End If
        
'PRN 615
'        dTemp = ((.lVoltage * .lVoltage) / (.lFeet * OhmsPerFoot)) / .lFeet
        dTemp = ((.lVoltage * .lVoltage) / (dFeet * OhmsPerFoot)) / dFeet
        CalcWattsPerFoot = dTemp
    End With
    Exit Function

ErrorHandler:
    'Ignore error and return zero WPF upon divide by zero or other error
End Function


Public Property Get SafeWires() As Variant
    SafeWires = m_arSafeWires
End Property


' Called by CalcLength
' Using the array m_arWires
' Populate m_arSafeWires with the wires that will work; array m_arSafeWires

Private Sub CalcSafeWires()
    Dim lWPF As Long
    Dim i
    Dim lBestWPF As Long
    Dim lDiffFromBest As Long
    
    LoadWires
    If m_udtWWire.bIsSinglePass Then
        lBestWPF = g_BestSingleWPF
    Else
        lBestWPF = g_BestDoubleWPF
    End If
        
    lDiffFromBest = g_MaxSafeWPF
    m_udtWWire.dOhmsPerFoot = 0
    
    ReDim m_arSafeWires(0)
    'For i = 1 To g_WireCount
    For i = 1 To m_lWireCount
        lWPF = CalcWattsPerFoot(m_arWires(i))
        If lWPF > g_MinSafeWPF And lWPF < g_MaxSafeWPF Then
            ReDim Preserve m_arSafeWires(UBound(m_arSafeWires) + 1)
            m_arSafeWires(UBound(m_arSafeWires)) = m_arWires(i)
            'set the wire closest too lbestwpf
            If Abs(lWPF - lBestWPF) < lDiffFromBest Then
                lDiffFromBest = Abs(lWPF - lBestWPF)
                m_udtWWire.dOhmsPerFoot = m_arWires(i)
            End If
        End If
    Next
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




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colItems As Collection
Private m_lSelectedIndex As Long
Private m_sBackup As String

Private m_eIsTaxable As Boolean

Public Enum ItemTypeCode
    itFinishedGood = 5
    itBTOKit = 7
    itMoldedGasket = 10
    itWireShelf = 20
    itWarmerWire = 30
    itSpecialOrder = 40
End Enum


'*** 12/19/08 LR modified
Public Enum ItemStatusCode
    iscEmpty = 0
    iscResearch = 1
    iscQuote = 2
    iscAuthorize = 3
    iscReadyToCommit = 4
    iscARHold = 5
    iscPendingCommit = 6
    iscHasRMA = 128
    iscCommitted = 256
    iscDeleted = 512
End Enum

Public Enum ItemResearchStatus
    irsResearchEmpty = 0
    irsNeedResearch = 1
    irsContactFactory = 2
    irsContactCustomer = 3
    irsWaitFactory = 4
    irsWaitCustomer = 5
End Enum

Public Enum enumItemInventoryStatus
    iisActive = 1
    iisInactive
    iisDiscontinued
    iisDeleted
End Enum


Public Event Add(ByRef oItem As IItem)
Public Event Remove(ByRef oItem As IItem)


Public Function Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False _
) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Dim oItem As IItem
    
    Set oXML = New JDMPDXML.XMLNode
    oXML.Tag = "Items"
    oXML.Add "IsTaxable", m_eIsTaxable
    For Each oItem In m_colItems
        oXML.AddChild oItem.Export(bPubNotes, bPrivNotes, bPurchNotes)
    Next
    Set Export = oXML
End Function


Public Sub Import(ByVal i_oXML As JDMPDXML.XMLNode, lWhseKey As Long)
    Dim oChild As JDMPDXML.XMLNode
    Dim oItem As IItem

    Clear

    For Each oChild In i_oXML
        ImportItem oChild, lWhseKey
    Next

    SelectedIndex = 0
End Sub


Public Sub ImportItem(ByVal i_oXML As JDMPDXML.XMLNode, lWhseKey As Long)
    Dim oItem As IItem

    Select Case i_oXML.Tag
    Case "FinGood", "SpecialOrder"
        Set oItem = New ItemFinGood
    Case "BTOKit"
        Set oItem = New ItemBTOKit
    Case "Gasket"
        Set oItem = New ItemGasket
    Case "Shelf"
        Set oItem = New ItemShelf
    Case "WWire"
        Set oItem = New ItemWWire
    Case "IsTaxable"
        m_eIsTaxable = i_oXML.value
        Exit Sub
    Case Else
        Err.Raise -1, "Items.ImportItem", "Unexpected Item tag: " & i_oXML.Tag
    End Select
    
    oItem.Import i_oXML, lWhseKey
    
    Add oItem
End Sub

Public Sub Backup()
    Dim oXML As JDMPDXML.XMLNode

    Set oXML = Export
    m_sBackup = oXML.ExportString
End Sub

Public Sub Restore(lWhseKey As Long)
    Dim oXML As JDMPDXML.XMLNode

    If Len(m_sBackup) = 0 Then
        Clear
    Else
        Set oXML = New JDMPDXML.XMLNode
        oXML.ImportString m_sBackup
        Import oXML, lWhseKey
    End If
End Sub

Private Sub Class_Initialize()
    Set m_colItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colItems = Nothing
End Sub


Public Property Get SelectedIndex() As Long
    SelectedIndex = m_lSelectedIndex
End Property

Public Property Let SelectedIndex(lNewValue As Long)
    If lNewValue < 0 Or lNewValue > m_colItems.Count Then
        Err.Raise -1, "Items.SelectedIndex", "Invalid value for SelectedIndex: " & lNewValue
    End If

    If lNewValue = 1 And m_lSelectedIndex = 1 Then
        Debug.Print "Where's Waldo?"
    End If
    m_lSelectedIndex = lNewValue
    Debug.Print "Item Index set to " & m_lSelectedIndex & " of " & m_colItems.Count
End Property


Public Property Get SelectedItem() As IItem
    If m_lSelectedIndex > 0 And m_lSelectedIndex <= m_colItems.Count Then
        Set SelectedItem = m_colItems(m_lSelectedIndex)
    End If
End Property


Public Property Get TotalPrice() As Double
    Dim oItem As IItem
    For Each oItem In m_colItems
        TotalPrice = TotalPrice + oItem.ExtendedPrice
    Next
End Property



Public Function AuthorizeAll() As Long
    Dim oItem As IItem
    Dim lCount As Long
    For Each oItem In m_colItems
        If oItem.StatusCode = iscAuthorize Then
            oItem.StatusCode = ItemStatusCode.iscReadyToCommit
            lCount = lCount + 1
        End If
    Next
    AuthorizeAll = lCount
End Function


Public Function UnAuthorizeAll() As Long
    Dim oItem As IItem
    Dim lCount As Long
    For Each oItem In m_colItems
        If oItem.StatusCode = iscReadyToCommit Then
            oItem.StatusCode = ItemStatusCode.iscAuthorize
            lCount = lCount + 1
        End If
    Next
    UnAuthorizeAll = lCount
End Function


Public Property Get ResearchStatusCode() As ItemResearchStatus
    Dim oItem As IItem
    Dim eResearchStatus As ItemResearchStatus
    
    eResearchStatus = irsResearchEmpty
    If m_colItems.Count = 0 Then
        Exit Property
    Else
        For Each oItem In m_colItems
            If oItem.StatusCode = iscResearch Then
                If oItem.ResearchStatus < eResearchStatus Then
                    eResearchStatus = oItem.ResearchStatus
                End If
            End If
        Next
    End If
    
    ResearchStatusCode = eResearchStatus
End Property
    

Public Property Get StatusCode() As ItemStatusCode
    Dim oItem As IItem
    Dim eOrderStatus As ItemStatusCode

    If m_colItems.Count = 0 Then
        eOrderStatus = iscEmpty
        Exit Property
    Else
        eOrderStatus = m_colItems(1).StatusCode
    End If

    For Each oItem In m_colItems
        If oItem.StatusCode < eOrderStatus Then
            eOrderStatus = oItem.StatusCode
        End If
    Next

    StatusCode = eOrderStatus
End Property

Public Property Get IsTaxable() As Boolean
    IsTaxable = m_eIsTaxable
End Property


Public Property Let IsTaxable(ByVal i_eNewValue As Boolean)
    Dim oItem As IItem

    m_eIsTaxable = i_eNewValue
    For Each oItem In m_colItems
        oItem.IsTaxable = i_eNewValue
    Next

End Property


Public Sub Clear()
    m_lSelectedIndex = 0
    While m_colItems.Count > 0
        Remove 1
    Wend
    Set m_colItems = New Collection
End Sub


Public Function CreateItem(i_eItemType As ItemTypeCode) As IItem
    Dim oItem As IItem
    
    Select Case i_eItemType
    Case itWarmerWire
        Set oItem = New ItemWWire
        
    Case itWireShelf
        Set oItem = New ItemShelf
    
    Case itMoldedGasket
        Set oItem = New ItemGasket

    Case itSpecialOrder
        Set oItem = New ItemFinGood

    Case itFinishedGood
        Set oItem = New ItemFinGood
    
    Case itBTOKit
        Set oItem = New ItemBTOKit
    
    Case Else
        Err.Raise -1, "Items.CreateItem", "Unexpected item type: " & i_eItemType
        Exit Function
    End Select
    
    Set CreateItem = oItem
End Function


Public Sub Add(ByRef i_oItem As IItem)
    RaiseEvent Add(i_oItem)
    If Not i_oItem Is Nothing Then
        With i_oItem
            If .OPLineKey = 0 Then .OPLineKey = DATABASE.GetSurrogateKey("tcpSOLine")
            .RemarkContext.Load "ViewOrderLine", CStr(.OPLineKey)
            .IsTaxable = Me.IsTaxable
        End With
        m_colItems.Add i_oItem
        SelectedIndex = m_colItems.Count
    End If
End Sub


Public Sub Remove(vIndex As Variant)
    Dim oItem As IItem
    Set oItem = Item(vIndex)
    RaiseEvent Remove(oItem)
    
    If Not oItem Is Nothing Then
        m_colItems.Remove vIndex
        If m_lSelectedIndex = vIndex Then
            m_lSelectedIndex = 0
        End If
    End If
End Sub


Public Property Get Count() As Long
    Count = m_colItems.Count
End Property


Public Function Item(vIndex As Variant) As IItem
Attribute Item.VB_UserMemId = 0
    Dim lIndex As Long

    On Error GoTo ErrorHandler

    lIndex = CLng(vIndex)
    If lIndex > 0 And lIndex <= m_colItems.Count Then
        Set Item = m_colItems(lIndex)
    End If
    Exit Function

ErrorHandler:
    Err.Raise -1, "Items.Item", "Items.Item only supports numeric index values"
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colItems.[_NewEnum]
End Function



Public Function XmlCommit(i_lWhseKey As Long, _
                    ByRef rstWhses As ADODB.Recordset, _
                    ByRef rstVendor As ADODB.Recordset) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Dim oItem As IItem
    
    Set oXML = New JDMPDXML.XMLNode
    oXML.Tag = "items"

    For Each oItem In m_colItems
        oXML.AddChild oItem.XmlCommit(i_lWhseKey, rstWhses, rstVendor)
    Next
    Set XmlCommit = oXML
End Function


Public Function TotalTax(i_dTaxRate As Double) As Double
    Dim oItem As IItem
    For Each oItem In m_colItems
        If oItem.IsTaxable Then
'***DH 11/04/08 this solves the penny rounding discrepency between OA and Sage.
            'TotalTax = TotalTax + oItem.TaxAmt(i_dTaxRate)
            TotalTax = TotalTax + Round(CCur(oItem.TaxAmt(i_dTaxRate)), 2)
        End If
    Next
'***DH 10/31/08
'This was intended to truncate, but was not working properly
    'TotalTax = Round((TotalTax - 0.005), 2)
End Function




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Methods

Public Function Export( _
    Optional bPubNotes As Boolean = False, _
    Optional bPrivNotes As Boolean = False, _
    Optional bPurchNotes As Boolean = False _
) As JDMPDXML.XMLNode
End Function


Public Sub Import(ByVal i_oXML As JDMPDXML.XMLNode, Optional lWhseKey As Long)
End Sub


Public Sub Backup()
End Sub


Public Sub Restore()
End Sub

'Properties

Public Property Let CustType(ByVal sNewValue As String)
End Property


Public Property Get CustType() As String
End Property


Public Property Get Cost() As Double
End Property


Public Property Let Cost(ByVal dNewValue As Double)
End Property


Public Property Get ListPrice() As Double
End Property


Public Property Let ListPrice(ByVal dNewValue As Double)
End Property


Public Property Get DealerPrice() As Double
End Property


Public Property Let DealerPrice(ByVal dNewValue As Double)
End Property


Public Property Get WholesalePrice() As Double
End Property


Public Property Let WholesalePrice(ByVal dNewValue As Double)
End Property


Public Property Get NegotiatedPrice() As Double
End Property


Public Property Let NegotiatedPrice(ByVal dNewValue As Double)
End Property


Public Property Get StandardPrice() As Double
End Property


Public Property Get EffectivePrice() As Double
End Property


Public Property Get BackNegotiatedPrice() As Double
End Property


Public Property Let BackNegotiatedPrice(ByVal dNewValue As Double)
End Property


Public Property Get NegotiatedPricing() As Boolean
End Property


Public Property Get ExtendedPrice() As Double
End Property


'09/24/2002     TeddyX
'Add new Research Status for item
Public Property Get ResearchStatus() As ItemResearchStatus
End Property


Public Property Let ResearchStatus(ByVal eNewValue As ItemResearchStatus)
End Property


Public Property Get StatusCode() As ItemStatusCode
End Property


Public Property Let StatusCode(ByVal eNewValue As ItemStatusCode)
End Property


Public Function IsValidStatusCode(ByVal eStatus As ItemStatusCode) As Boolean
End Function


'TODO: REVIEW all below
Public Property Get SageItemType() As Long
End Property


Public Property Let SageItemType(ByVal vNewValue As Long)
End Property


Public Property Get OPItemType() As ItemTypeCode  'read-only and implied by subclass
End Property


Public Property Get ItemKey() As Long
End Property


Public Property Let ItemKey(ByVal vNewValue As Long)
End Property


Public Property Get LineKey() As Long
End Property


Public Property Let LineKey(ByVal vNewValue As Long)
End Property


Public Property Get MakeKey() As Long
End Property


Public Property Let MakeKey(ByVal vNewValue As Long)
End Property


Public Property Get OPKey() As Long
End Property


Public Property Let OPKey(ByVal vNewValue As Long)
End Property


Public Property Get Qty() As Long
End Property


Public Property Let Qty(ByVal vNewValue As Long)
End Property


Public Property Get QtyAvail(lWhseKey As Long) As Variant
End Property


Public Property Get QtyOnHand(lWhseKey As Long) As Variant
End Property


Public Property Get QtyOnSO(lWhseKey As Long) As Variant
End Property


Public Property Get QtyOnPO(lWhseKey As Long) As Variant
End Property


Public Property Get QtyOnBO(lWhseKey As Long) As Variant
End Property


Public Property Get VendorKey() As Long
End Property


Public Property Let VendorKey(ByVal vNewValue As Long)
End Property


Public Property Get Descr() As String
End Property


Public Property Let Descr(ByVal vNewValue As String)
End Property


Public Property Get ItemID() As String
End Property


Public Property Let ItemID(ByVal vNewValue As String)
End Property


Public Property Get ModelNbr() As String
End Property


Public Property Let ModelNbr(ByVal vNewValue As String)
End Property


Public Property Get RemarkContext() As MemoMeister.RemarkContext
End Property


Public Property Get SerialNbr() As String
End Property


Public Property Let SerialNbr(ByVal vNewValue As String)
End Property


Public Property Get IsCGMPN() As Boolean
End Property


Public Property Let IsCGMPN(ByVal bNewValue As Boolean)
End Property



Public Property Get MorphBTOKey() As Long
End Property


Public Property Let MorphBTOKey(ByVal lNewValue As Long)
End Property


Public Property Get IsTaxable() As Boolean
End Property


Public Property Let IsTaxable(ByVal vNewValue As Boolean)
End Property


Public Function XmlCommit(i_lWhseKey As Long, ByRef rstWhses As ADODB.Recordset, ByRef rstVendor As ADODB.Recordset) As JDMPDXML.XMLNode
End Function


Public Function TaxAmt(i_dTaxRate As Double) As Double
End Function


Public Property Get OPLineKey() As Long
End Property


Public Property Let OPLineKey(ByVal lNewValue As Long)
End Property


'add two more properties for cross reference information
Public Property Get OriginalItemID() As String
End Property


Public Property Let OriginalItemID(ByVal sNewValue As String)
End Property


Public Property Get RefSource() As String
End Property


Public Property Let RefSource(ByVal sNewValue As String)
End Property


Public Property Get ItemInventoryStatus() As enumItemInventoryStatus
End Property

Public Property Let ItemInventoryStatus(ByVal eNewValue As enumItemInventoryStatus)
End Property



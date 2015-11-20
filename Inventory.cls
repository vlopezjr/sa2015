VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'private backing store

Private m_lWhseKey As Long
Private m_lQtyOnBO As Long
Private m_lQtyOnPO As Long
Private m_lQtyOnSO As Long
Private m_lQtyOnHand As Long
Private m_lQtyAvail As Long
Private m_lVendKey As Long

'LR 11/16/15
Private m_sWhseID As String
Private m_lItemKey As Long
Private m_sVendID As String
Private m_sVendName As String

Private m_dVendRplcmntUnitCost As Double
Private m_dVendListPrice As Double
    
'VL 10/14/2015
Private m_lPendQtyDecrease As Long
Private m_lPendQtyIncrease As Long
Private m_bInConflict As Boolean
    
'VL 11/13/2015
Private m_lPicked As Long
Private m_lPacked As Long
Private m_lPendingAdjustments As Long
    
Private m_colBins As Collection


' public properties

Public Property Get WhseID() As String
    WhseID = m_sWhseID
End Property

Public Property Get WhseKey() As Long
    WhseKey = m_lWhseKey
End Property

Public Property Get QtyOnHand() As Long
    QtyOnHand = m_lQtyOnHand
End Property

Public Property Get QtyOnSO() As Long
    QtyOnSO = m_lQtyOnSO
End Property

Public Property Get QtyOnPO() As Long
    QtyOnPO = m_lQtyOnPO
End Property

Public Property Get QtyOnBO() As Long
    QtyOnBO = m_lQtyOnBO
End Property

Public Property Get QtyAvail() As Long
    QtyAvail = m_lQtyAvail
End Property

Public Property Get VendKey() As Long
    VendKey = m_lVendKey
End Property

Public Property Get VendID() As String
    VendID = m_sVendID
End Property

Public Property Get VendName() As String
    VendName = m_sVendName
End Property

Public Property Get VendRplcmntUnitCost() As Double
    VendRplcmntUnitCost = m_dVendRplcmntUnitCost
End Property

Public Property Get VendListPrice() As Double
    VendListPrice = m_dVendListPrice
End Property

Public Property Get PendQtyDecrease() As Long
    PendQtyDecrease = m_lPendQtyDecrease
End Property

Public Property Get PendQtyIncrease() As Long
    PendQtyIncrease = m_lPendQtyIncrease
End Property

Public Property Get InConflict() As Boolean
    InConflict = m_bInConflict
End Property

Public Property Get Picked() As Long
    Picked = m_lPicked
End Property

Public Property Get Packed() As Long
    Packed = m_lPacked
End Property

Public Property Get PendingAdjustments() As Long
    PendingAdjustments = m_lPendingAdjustments
End Property


Private Sub Class_Initialize()
    Set m_colBins = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colBins = Nothing
End Sub


Public Property Get Bins() As Collection
    Set Bins = m_colBins
End Property


Public Sub Init(i_rst As ADODB.Recordset)
    With i_rst
        m_lItemKey = .Fields("ItemKey").value
        m_sWhseID = .Fields("WhseID").value
        m_lWhseKey = .Fields("WhseKey").value
        m_lQtyOnBO = .Fields("QtyOnBO").value
        m_lQtyOnPO = .Fields("QtyOnPO").value
        m_lQtyOnSO = .Fields("QtyOnSO").value
        m_lQtyAvail = .Fields("QtyAvail").value
        m_lVendKey = .Fields("PrimaryVendKey").value
        m_sVendID = .Fields("VendID").value
        m_sVendName = .Fields("VendName").value
        m_dVendListPrice = .Fields("VendorListPrice").value
        m_dVendRplcmntUnitCost = .Fields("RplcmntUnitCost").value
        m_lPendQtyDecrease = .Fields("PendQtyDecrease").value
        m_lPendQtyIncrease = .Fields("PendQtyIncrease").value
        m_bInConflict = .Fields("InConflict").value
        m_lPicked = .Fields("Picked").value
        m_lPacked = .Fields("Packed").value
        m_lPendingAdjustments = .Fields("PendingAdjustments").value
    End With

'LR 11/16/15 would like to move it here
'    PopulateBins i_rst
End Sub


'LR 11/16/15 removed the itemkey parameter
' remains public becasue of the way it shares a recordset with its parent

'Public Sub PopulateBins(i_rst As ADODB.Recordset, i_lItemKey As Long)
Public Sub PopulateBins(i_rst As ADODB.Recordset)
    Dim o_Bin As Bin
    Dim lQtyOnHand
    
    With i_rst
        .Filter = "WhseKey=" & m_lWhseKey & " AND Itemkey=" & m_lItemKey 'i_lItemKey
        
        m_lQtyOnHand = .Fields("QtyOnHand").value
        
        Do While Not .EOF
            Set o_Bin = New Bin
            o_Bin.BinID = .Fields("WhseBinID").value
            o_Bin.QtyOnHand = .Fields("QtyOnHand").value
            o_Bin.PendQtyDecrease = .Fields("PendQtyDecrease").value
            o_Bin.PendQtyIncrease = .Fields("PendQtyIncrease").value
            m_colBins.Add o_Bin
            Set o_Bin = Nothing
            .MoveNext
        Loop
        .Filter = adFilterNone
    End With
End Sub


'this is used to initialize a BTO Kit's inventory object

Public Sub SetBTOKitInv(i_lWhseKey As Long, _
    i_lQtyOnHand As Long, _
    i_lQtyOnSO, _
    i_lQtyOnBO, _
    i_lQtyOnPO, _
    i_lQtyAvail _
)
    m_sWhseID = WhseKeyToID(i_lWhseKey)
    m_lWhseKey = i_lWhseKey
    m_lQtyOnHand = i_lQtyOnHand
    m_lQtyOnSO = i_lQtyOnSO
    m_lQtyOnBO = i_lQtyOnBO
    m_lQtyOnPO = i_lQtyOnPO
    m_lQtyAvail = i_lQtyAvail
End Sub


'LR 11/16/15
'remove the itemkey and whsekey parameters from all of these methods

'Public Function GetOrders(i_lItemKey As Long, i_lWhseKey) As ADODB.Recordset
'    Set GetOrders = CallSP("spCPCGetOrdersforSO", "@i_ItemKey", i_lItemKey, "@i_WhseKey", i_lWhseKey)
'End Function
Public Function GetOrders() As ADODB.Recordset
    Set GetOrders = CallSP("spCPCGetOrdersforSO", "@i_ItemKey", m_lItemKey, "@i_WhseKey", m_lWhseKey)
End Function

'Public Function GetPurchaseOrders(i_lItemKey As Long, i_lWhseKey) As ADODB.Recordset
'    If i_lWhseKey = 0 Then
'        Set GetPurchaseOrders = CallSP("spcpcGetOpenPOsByKey", "@_iItemKey", i_lItemKey)
'    Else
'        Set GetPurchaseOrders = CallSP("spcpcGetOpenPOsByKey", "@_iItemKey", i_lItemKey, "@_iWhseKey", i_lWhseKey)
'    End If
'End Function
Public Function GetPurchaseOrders() As ADODB.Recordset
    Set GetPurchaseOrders = CallSP("spcpcGetOpenPOsByKey", "@_iItemKey", m_lItemKey, "@_iWhseKey", m_lWhseKey)
End Function

'Public Function GetPickedOrders(i_lItemKey As Long, i_lWhseKey) As ADODB.Recordset
'    Set GetPickedOrders = CallSP("spcpcGetPickedOrders", "@_iItemKey", i_lItemKey, "@_iWhseKey", i_lWhseKey)
'End Function
Public Function GetPickedOrders() As ADODB.Recordset
    Set GetPickedOrders = CallSP("spcpcGetPickedOrders", "@_iItemKey", m_lItemKey, "@_iWhseKey", m_lWhseKey)
End Function

'Public Function GetPackedOrders(i_lItemKey As Long, i_lWhseKey) As ADODB.Recordset
'    Set GetPackedOrders = CallSP("spcpcGetPackedOrders", "@_iItemKey", i_lItemKey, "@_iWhseKey", i_lWhseKey)
'End Function
Public Function GetPackedOrders() As ADODB.Recordset
    Set GetPackedOrders = CallSP("spcpcGetPackedOrders", "@_iItemKey", m_lItemKey, "@_iWhseKey", m_lWhseKey)
End Function

'Public Function GetPendingIncreaseAdjustments(i_lItemKey As Long, i_lWhseKey) As ADODB.Recordset
'    Set GetPendingIncreaseAdjustments = CallSP("spcpcGetPendingIncreaseAdjustments", "@_iItemKey", i_lItemKey, "@_iWhseKey", i_lWhseKey)
'End Function
Public Function GetPendingIncreaseAdjustments() As ADODB.Recordset
    Set GetPendingIncreaseAdjustments = CallSP("spcpcGetPendingIncreaseAdjustments", "@_iItemKey", m_lItemKey, "@_iWhseKey", m_lWhseKey)
End Function

'Public Function GetPendingDecreaseAdjustments(i_lItemKey As Long, i_lWhseKey) As ADODB.Recordset
'    Set GetPendingDecreaseAdjustments = CallSP("spcpcGetPendingDecreaseAdjustments", "@_iItemKey", i_lItemKey, "@_iWhseKey", i_lWhseKey)
'End Function
Public Function GetPendingDecreaseAdjustments() As ADODB.Recordset
    Set GetPendingDecreaseAdjustments = CallSP("spcpcGetPendingDecreaseAdjustments", "@_iItemKey", m_lItemKey, "@_iWhseKey", m_lWhseKey)
End Function

'Public Function GetPendingInventoryAdjustments(i_lItemKey As Long, i_lWhseKey) As ADODB.Recordset
'    Set GetPendingInventoryAdjustments = CallSP("spcpcGetPendingInventoryAdjustments", "@_iItemKey", i_lItemKey, "@_iWhseKey", i_lWhseKey)
'End Function
Public Function GetPendingInventoryAdjustments() As ADODB.Recordset
    Set GetPendingInventoryAdjustments = CallSP("spcpcGetPendingInventoryAdjustments", "@_iItemKey", m_lItemKey, "@_iWhseKey", m_lWhseKey)
End Function

Public Function GetConflicts(i_lItemKey As Long, i_lWhseKey) As ADODB.Recordset
    Set GetConflicts = CallSP("spcpcGetConflicts", "@_iItemKey", i_lItemKey, "@_iWhseKey", i_lWhseKey)
End Function




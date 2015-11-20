VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private Type BinProps
'    sBinID As String * 10
'    sBinLoc As String * 10
'    lQtyOnHand As Long
'End Type
'
'Private m_udtBin As BinProps

Private sBinID As String
Private lQtyOnHand As Long
Private lPendQtyDecrease As Long
Private lPendQtyIncrease As Long

Public Property Get BinID() As String
    BinID = sBinID
End Property

Public Property Let BinID(ByVal vNewValue As String)
    sBinID = vNewValue
End Property


'Public Property Get BinLoc() As String
'    BinLoc = m_udtBin.sBinLoc
'End Property
'
'Public Property Let BinLoc(ByVal vNewValue As String)
'    m_udtBin.sBinLoc = vNewValue
'End Property


Public Property Get QtyOnHand() As Long
    QtyOnHand = lQtyOnHand
End Property

Public Property Let QtyOnHand(ByVal vNewValue As Long)
    lQtyOnHand = vNewValue
End Property

Public Property Get PendQtyDecrease() As Long
    PendQtyDecrease = lPendQtyDecrease
End Property

Public Property Let PendQtyDecrease(ByVal vNewValue As Long)
    lPendQtyDecrease = vNewValue
End Property

Public Property Get PendQtyIncrease() As Long
    PendQtyIncrease = lPendQtyIncrease
End Property

Public Property Let PendQtyIncrease(ByVal vNewValue As Long)
    lPendQtyIncrease = vNewValue
End Property


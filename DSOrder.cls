VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lPOID As Long
Private sTranNo As String
Private dFreight As Double
Private sVouchers As String
Private sTrackNo As String

Public Property Get POID() As Long
    POID = lPOID
End Property

Public Property Let POID(ByVal oNewValue As Long)
    lPOID = oNewValue
End Property

Public Property Get TranNo() As String
    TranNo = sTranNo
End Property

Public Property Let TranNo(ByVal oNewValue As String)
    sTranNo = oNewValue
End Property

Public Property Get Freight() As Double
    Freight = dFreight
End Property

Public Property Let Freight(ByVal oNewValue As Double)
    dFreight = oNewValue
End Property

Public Property Get Vouchers() As String
    Vouchers = sVouchers
End Property

Public Property Let Vouchers(ByVal oNewValue As String)
    sVouchers = oNewValue
End Property

Public Property Get TrackNo() As String
    TrackNo = sTrackNo
End Property

Public Property Let TrackNo(ByVal oNewValue As String)
    sTrackNo = oNewValue
End Property

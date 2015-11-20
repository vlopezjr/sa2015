VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OSItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Order Status Item

Public ItemKey As Long
Public Status As LineStatus
Public ItemID As String
Public Description As String
Public QtyOrdered As Long
Public UnitPrice As Double
Public QtyOpenToShip As Long
Public QtyInvcd As Long
Public SOLineKey As Long

Public QtyOnHand As Long
Public QtyOnSO As Long
Public QtyAvail As Long
Public Cost As Double
Public Pick As Boolean
Public ItemType As ItemTypeCode
Public Conflict As Boolean

Public Property Get TypeDesc() As String
    Select Case ItemType
        Case itFinishedGood: TypeDesc = "Component"
        Case itBTOKit: TypeDesc = "Kit"
        Case itMoldedGasket: TypeDesc = "Gasket"
        Case itWarmerWire: TypeDesc = "WWire"
        Case itWireShelf: TypeDesc = "WShelf"
        Case itSpecialOrder: TypeDesc = "SPO"
    End Select
End Property


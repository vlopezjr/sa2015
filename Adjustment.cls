VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Adjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sItemID As String
Private m_lQuantity As Integer
Private m_sTranDescr As String


Public Property Get ItemID() As String
    ItemID = m_sItemID
End Property
Public Property Let ItemID(value As String)
    m_sItemID = value
End Property


Public Property Get Quantity() As Integer
    Quantity = m_lQuantity
End Property
Public Property Let Quantity(value As Integer)
    m_lQuantity = value
End Property


Public Property Get TranDescr() As String
    TranDescr = m_sTranDescr
End Property
Public Property Let TranDescr(value As String)
    m_sTranDescr = value
End Property

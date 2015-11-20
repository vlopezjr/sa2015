VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RMAList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_col As Collection


Private Sub Class_Initialize()
    Set m_col = New Collection
End Sub


Private Sub Class_Terminate()
    Set m_col = Nothing
End Sub


Public Sub Add(ByRef i_oRMAOrder As RMAOrder)
    m_col.Add i_oRMAOrder
End Sub


Public Sub Remove(vIndex As Variant)
    On Error Resume Next
    m_col.Remove vIndex
End Sub


Public Property Get Count() As Long
    Count = m_col.Count
End Property


Public Function Item(vIndex As Variant) As RMAOrder
Attribute Item.VB_UserMemId = 0
    Set Item = m_col(vIndex)
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_col.[_NewEnum]
End Function






















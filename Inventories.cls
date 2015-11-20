VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inventories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colInventories As Collection

Private Sub Class_Initialize()
    Set m_colInventories = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colInventories = Nothing
End Sub

Public Function Count() As Long
    Count = m_colInventories.Count
End Function

Public Function Item(vIndex As Variant) As Inventory
    Item = m_colInventories.Item(vIndex)
End Function

Public Property Get Items() As Collection
    Set Items = m_colInventories
End Property

Public Function WhseInventory(ByVal i_lWhseKey As Long) As Inventory
    Dim oInv As Inventory

    For Each oInv In m_colInventories
        If oInv.WhseKey = i_lWhseKey Then
            Set WhseInventory = oInv
            Exit Function
        End If
    Next
End Function

Public Sub Init(i_rst As ADODB.Recordset, i_lPntItemKey As Long)
    Dim oInv As Inventory
    Dim lHoldWhseKey As Long
    
    On Error GoTo EH
       
    Set m_colInventories = New Collection
    i_rst.Filter = "ItemKey=" & i_lPntItemKey
    With i_rst
        .MoveFirst
        Do While Not .EOF
            If Not IsNull(.Fields("WhseKey").value) Then
                If lHoldWhseKey <> .Fields("WhseKey").value Then
                    lHoldWhseKey = .Fields("WhseKey").value
                    Set oInv = New Inventory
                    oInv.Init i_rst
                    m_colInventories.Add oInv
                    Set oInv = Nothing
                End If
            End If
            .MoveNext
        Loop
    End With
    
'LR 11/16/15 this is ugly
'the reason it is done out here instead of within the Inventory object
'is because it is using the same recordset

    For Each oInv In m_colInventories
        'oInv.PopulateBins i_rst, i_lPntItemKey
        oInv.PopulateBins i_rst
    Next
    Exit Sub
EH:
    ThrowError "Inventories", "Init", "i_lPntItemKey = " & i_lPntItemKey
End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colInventories.[_NewEnum]
End Function


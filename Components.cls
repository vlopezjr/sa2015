VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Components"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ksSource = "Components"

Private m_colItems As Collection
'Private m_KitInConflict As Boolean

Private Sub Class_Initialize()
    Set m_colItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colItems = Nothing
End Sub


'Public Property Get InConflict()
'    InConflict = m_KitInConflict
'End Property


Public Sub Init(i_rst As ADODB.Recordset)
    Dim oItem As ItemFinGood
    Dim lItemKey As Long

    On Error GoTo EH
    
    With i_rst
        Do While Not .EOF
            If .Fields("ItemKey").Value <> lItemKey Then
                lItemKey = .Fields("ItemKey").Value
                Set oItem = New ItemFinGood
                oItem.Load lItemKey, .Fields("WhseKey").Value
'                If oItem.IItem_InConflict Then
'                    m_KitInConflict = True
'                End If
                oItem.IItem_Qty = i_rst.Fields("CompItemQty").Value
                m_colItems.Add oItem
                Set oItem = Nothing
            End If
            .MoveNext
        Loop
    End With
    
    Exit Sub
EH:
    ThrowError "Components", "Init", ""
    
End Sub


Public Function Export() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Dim oItem As IItem

    Set oXML = New JDMPDXML.XMLNode
    oXML.Tag = "CompItems"
    For Each oItem In m_colItems
        oXML.AddChild oItem.Export(bPubNotes:=False, bPrivNotes:=False, bPurchNotes:=False)
    Next
    Set Export = oXML
    
End Function

    
Public Property Get Count() As Long
    Count = m_colItems.Count
End Property


Public Function Item(vIndex As Variant) As IItem
    Set Item = m_colItems(vIndex)
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colItems.[_NewEnum]
End Function


Public Function InConflict(i_lWhseKey As Long) As Boolean
    Dim oItem As ItemFinGood
    For Each oItem In m_colItems
        If oItem.IItem_InConflict(i_lWhseKey) Then
            InConflict = True
        End If
    Next
End Function


Public Function CalcQty(i_lWhseKey As Long, i_sQtyName As String) As Long
    Dim oItem As ItemFinGood
    Dim oInvs As Inventories
    Dim oInv As Inventory
    Dim lCompQtyWhse As Long
    Dim lTempQty As Long
    Dim bFirstPass As Boolean

    Dim l_ItemQty As Long
    Dim l_InvQty As Long
    
    lTempQty = 0
    bFirstPass = True
    
    For Each oItem In m_colItems
        lCompQtyWhse = 0
        
        Select Case i_sQtyName
            Case Is = "QtyOnHand"
                l_ItemQty = oItem.IItem_QtyOnHand(i_lWhseKey)
            Case Is = "QtyOnSO"
                l_ItemQty = oItem.IItem_QtyOnSO(i_lWhseKey)
            Case Is = "QtyOnBO"
                l_ItemQty = oItem.IItem_QtyOnBO(i_lWhseKey)
            Case Is = "QtyOnPO"
                l_ItemQty = oItem.IItem_QtyOnPO(i_lWhseKey)
            Case Is = "QtyAvail"
                l_ItemQty = oItem.IItem_QtyAvail(i_lWhseKey)
            Case Else
                Err.Raise -1, ksSource, "Invalid qty type in Components.CalcQty"
                Exit Function
        End Select

        For Each oInv In oItem.Inventories
            If oInv.WhseKey = i_lWhseKey Then
                Select Case i_sQtyName
                    Case Is = "QtyOnHand"
                        l_InvQty = oInv.QtyOnHand
                    Case Is = "QtyOnSO"
                        l_InvQty = oInv.QtyOnSO
                    Case Is = "QtyOnBO"
                        l_InvQty = oInv.QtyOnBO
                    Case Is = "QtyOnPO"
                        l_InvQty = oInv.QtyOnPO
                    Case Is = "QtyAvail"
                        l_InvQty = oInv.QtyAvail
                    Case Else
                        Err.Raise -1, ksSource, "Invalid qty type in Components.CalcQty"
                        Exit Function
                End Select
    
                lCompQtyWhse = l_InvQty
                Exit For
            End If
        Next

        'Truncate rather than round the Qty calcs
        If bFirstPass Then
            lTempQty = Int(lCompQtyWhse / oItem.IItem_Qty)
            bFirstPass = False
        Else
            If lTempQty > Int(lCompQtyWhse / oItem.IItem_Qty) Then
                lTempQty = Int(lCompQtyWhse / oItem.IItem_Qty)
            End If
        End If

    Next
    
    CalcQty = lTempQty
    
End Function

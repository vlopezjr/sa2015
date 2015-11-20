VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OSItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Order Status Item collections are used in the following forms:
'   FOrder
'   FBilling
'   FWarehouse
'   FChooseConflictItems


Public Enum LineStatus
    IsNeedsToBeOrder = 1
    IsOnOrder = 2
    IsReadyToShip = 3
    IsShipComplete = 4
    IsShipBackorders = 5
    IsInvoiced = 6
    IsDropShipInActive = 7
    IsDropShipCancelled = 8
    IsDropShipClosed = 9
    IsGskNew = 10
    IsGskBegin = 11
    IsGskCut = 12
    IsGskMold = 13
    IsGskTrim = 14
    IsGskNotAvail = 15
    IsBackOrderCancelled = 16
    IsPacking = 17
    IsShipping = 18
    IsGskOutOfStock = 19
End Enum

Private m_col As Collection

'Add new gasket line status the Office Assistant (10/8/02 TX)
'This is the date the Gasket Tracking went online in MPK.
'It is used in the status calculation below.
Private Const k_sGskDate = "11/07/02"


Private Sub Class_Initialize()
    Set m_col = New Collection
End Sub


Private Sub Class_Terminate()
    Set m_col = Nothing
End Sub


Public Sub Add(ByRef i_oOSItem As OSItem)
    m_col.Add i_oOSItem
End Sub


Public Sub Remove(vIndex As Variant)
    On Error Resume Next
    m_col.Remove vIndex
End Sub


Public Property Get Count() As Long
    Count = m_col.Count
End Property


Public Function Item(vIndex As Variant) As OSItem
    Set Item = m_col(vIndex)
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_col.[_NewEnum]
End Function


Public Sub Load(lOPKey As Long, bDropShip As Boolean)
    Dim oOSItem As OSItem
    Dim oBTOKit As ItemBTOKit
    Dim sql As String
    Dim lWhseKey As Long
    Dim rstLineItem As ADODB.Recordset
    Dim bShipping As Boolean
    Dim bPacking As Boolean
    
    'added this to avoid the expensive query if called with OPKey = 0
    If lOPKey = 0 Then Exit Sub
    
    Set rstLineItem = CallSP("spOPOrdStatGetLineItems", "@_iOPKey", lOPKey)
    
    If Not rstLineItem.EOF Then
        rstLineItem.MoveFirst
        With rstLineItem
            While Not .EOF
                Set oOSItem = New OSItem
                oOSItem.ItemID = .Fields("ItemID").Value
                oOSItem.Description = .Fields("Description").Value
                oOSItem.UnitPrice = .Fields("UnitPrice").Value
                oOSItem.QtyOrdered = .Fields("QtyOrd").Value
                oOSItem.ItemKey = .Fields("ItemKey").Value
                oOSItem.SOLineKey = .Fields("SOLineKey").Value
                If bDropShip Then
                    oOSItem.Status = GetDropShipItemsStatus(.Fields("SOLineKey").Value)
                Else
                    'set special gasket status for gasket items.
                    If Left(oOSItem.ItemID, 4) = "GSK-" Then
                        'If order committed date is before "11/07/2002", no gasket status
                        'info is available
                        If .Fields("CommitDate").Value <= CDate(k_sGskDate) Then
                            oOSItem.Status = IsGskNotAvail
                        Else
                            oOSItem.Status = GetGsktStatus(lOPKey, .Fields("SOLineKey").Value)
                        End If
                    Else
                        bShipping = (Not IsNull(.Fields("ShipKey"))) And Not (IsNull(.Fields("FreightAmt")) And IsNull(.Fields("ShipTrackNo")))
                        bPacking = (Not IsNull(.Fields("ShipKey"))) And IsNull(.Fields("FreightAmt")) And IsNull(.Fields("ShipTrackNo"))
                       
                        If Left(oOSItem.ItemID, 4) = "SPO-" Or Left(oOSItem.ItemID, 4) = "SHF-" Then
                            oOSItem.Status = GetPOStatus( _
                                        .Fields("QtyInvcd").Value, .Fields("QtyOrd").Value, _
                                        .Fields("QtyOpenToShip").Value, .Fields("PORecdComplete").Value, _
                                        .Fields("FreezeSOLineKey").Value, .Fields("Status").Value, bShipping, bPacking)
                        Else
                            oOSItem.Status = GetNonSPOLineStatus( _
                                    .Fields("QtyInvcd").Value, .Fields("QtyOrd").Value, _
                                    .Fields("QtyOpenToShip").Value, .Fields("QtyOnHand").Value, _
                                    .Fields("QtyOnPO").Value, .Fields("Status").Value, bShipping, bPacking)
                        End If
                    End If
                End If

                oOSItem.QtyOpenToShip = .Fields("QtyOpenToShip").Value
                oOSItem.QtyInvcd = .Fields("QtyInvcd").Value
                oOSItem.ItemType = .Fields("ItemType").Value
                
                If oOSItem.ItemType = itBTOKit Then
                    Set oBTOKit = New ItemBTOKit
                    lWhseKey = .Fields("WhseKey").Value
                    oBTOKit.Load oOSItem.ItemKey, lWhseKey
                    oOSItem.QtyOnHand = oBTOKit.IItem_QtyOnHand(lWhseKey)
                    oOSItem.QtyOnSO = oBTOKit.IItem_QtyOnSO(lWhseKey)
                    Set oBTOKit = Nothing
                Else
                    oOSItem.QtyOnHand = .Fields("QtyOnHand").Value
                    oOSItem.QtyOnSO = .Fields("QtyOnSO").Value
                End If
                
                If oOSItem.ItemType = itBTOKit Or oOSItem.ItemType = itFinishedGood Then
                    If oOSItem.QtyOnHand > 0 And oOSItem.QtyOpenToShip <> oOSItem.QtyOnSO And oOSItem.QtyOnSO > oOSItem.QtyOnHand Then
                        oOSItem.Conflict = True
                    Else
                        oOSItem.Conflict = False
                    End If
                Else
                    oOSItem.Conflict = False
                End If
                
                m_col.Add oOSItem
                .MoveNext
            Wend
        End With
    End If
    Set oOSItem = Nothing
End Sub


'fetch the gasket status from tcpGsktState

Private Function GetGsktStatus(ByVal lOPKey As Long, ByVal lSOLineKey As Long) As LineStatus
    On Error GoTo ErrorHandler
    Dim oCmd As ADODB.Command
    
    Set oCmd = CreateCommandSP("spcpcGsktGetLatestState")
    With oCmd
        '.Parameters("@_iOrderID").value = lOPKey
        .Parameters("@_iLineKey").Value = lSOLineKey
        .Execute
    
        Select Case .Parameters("@_oState").Value
            Case 0:
                GetGsktStatus = IsGskNew
            Case 1:
                GetGsktStatus = IsGskBegin
            Case 2:
                GetGsktStatus = IsGskCut
            Case 3:
                GetGsktStatus = IsGskMold
            Case 4:
                GetGsktStatus = IsGskTrim
            Case 5:
                GetGsktStatus = IsGskOutOfStock
            Case Else
                GetGsktStatus = IsGskNew
        End Select
    End With
    
    Set oCmd = Nothing
    Exit Function
    
ErrorHandler:
    msg Err.Description, vbOKOnly, Err.Source
End Function


Private Function GetDropShipItemsStatus(ByVal lSOLineKey As Long) As LineStatus
    Dim rst As ADODB.Recordset
    
    Set rst = CallSP("spCPCsoGetDropShipStatus", "@_SOLineKey", lSOLineKey)
    If rst.EOF Then
        GetDropShipItemsStatus = IsNeedsToBeOrder
    Else
        If IsNull(rst.Fields("Status")) Then
            GetDropShipItemsStatus = IsNeedsToBeOrder
        Else
            Select Case rst.Fields("Status").Value
            Case 0:
                GetDropShipItemsStatus = IsNeedsToBeOrder
            Case 1:
                GetDropShipItemsStatus = IsOnOrder
            Case 2:
                GetDropShipItemsStatus = IsDropShipInActive
            Case 3:
                GetDropShipItemsStatus = IsDropShipCancelled
            Case 4:
                 GetDropShipItemsStatus = IsDropShipClosed
            End Select
        End If
    End If
End Function


Private Function GetNonSPOLineStatus( _
        lQtyInvcd As Long, lQtyOrd As Long, _
        lQtyOpenToShip As Long, lQtyOnHand As Long, _
        lQtyOnPO As Long, lLineStatus As Long, bShipping As Boolean, bPacking As Boolean) As LineStatus
        
        If lQtyInvcd = lQtyOrd Then
            GetNonSPOLineStatus = IsInvoiced
        ElseIf lQtyOpenToShip = 0 Then
            GetNonSPOLineStatus = IsShipComplete
        ElseIf lQtyOpenToShip > 0 And lLineStatus = 2 Then
            GetNonSPOLineStatus = IsBackOrderCancelled
        ElseIf lQtyOpenToShip > 0 And lQtyOpenToShip <> lQtyOrd Then
            GetNonSPOLineStatus = IsShipBackorders
        ElseIf lQtyOpenToShip > 0 And bShipping Then
            GetNonSPOLineStatus = IsShipping
        ElseIf lQtyOpenToShip > 0 And bPacking Then
            GetNonSPOLineStatus = IsPacking
        ElseIf Not (lQtyOrd > lQtyOnHand) Then
            GetNonSPOLineStatus = IsReadyToShip
        ElseIf (lQtyOrd > (lQtyOnHand + lQtyOnPO)) Then
            GetNonSPOLineStatus = IsNeedsToBeOrder
        ElseIf Not (lQtyOrd > (lQtyOnHand + lQtyOnPO)) Then
            GetNonSPOLineStatus = IsOnOrder
        End If
End Function


Private Function GetPOStatus( _
                lQtyInvcd As Long, lQtyOrd As Long, _
                lQtyOpenToShip As Long, bPORecdComplete As Boolean, _
                lFreezeSOLineKey As Long, lLineStatus As Long, bShipping As Boolean, bPacking As Boolean) As LineStatus
                
    If lQtyInvcd = lQtyOrd Then
        GetPOStatus = IsInvoiced
    ElseIf lQtyOpenToShip = 0 Then
        GetPOStatus = IsShipComplete
    ElseIf lQtyOpenToShip > 0 And lLineStatus = 2 Then
        GetPOStatus = IsBackOrderCancelled
    ElseIf lQtyOpenToShip > 0 And lQtyOpenToShip <> lQtyOrd Then
        GetPOStatus = IsShipBackorders
    ElseIf lQtyOpenToShip > 0 And bShipping Then
        GetPOStatus = IsShipping
    ElseIf lQtyOpenToShip > 0 And bPacking Then
        GetPOStatus = IsPacking
    ElseIf bPORecdComplete = True Then
        GetPOStatus = IsReadyToShip
    ElseIf lFreezeSOLineKey = 0 Then
        GetPOStatus = IsNeedsToBeOrder
    ElseIf bPORecdComplete = False Then
        GetPOStatus = IsOnOrder
    End If
End Function


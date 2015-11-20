Attribute VB_Name = "ShipMethod"
Option Explicit

'what's the relationship between ShipVia and Warehouse?

Public Sub SetUpShipVia(cboShipVia As ComboBox, _
                        ByVal WhseKey As Long, _
                        ByVal ShipMethKey As Long)

    g_rstShipVia.Filter = "ShipMethID LIKE '" & GetWhseIDFromWhseKey(WhseKey) & "%'"
    LoadShipViaCombo cboShipVia, g_rstShipVia, "ShipMethID", "ShipMethKey", ShipMethKey
    g_rstShipVia.Filter = adFilterNone

End Sub


Public Function RecalcShipVia(cboShipVia As ComboBox, _
                            ByRef rstShipVia As ADODB.Recordset, _
                            ByRef rstWhses As ADODB.Recordset, _
                            ByVal WhseKey As Long) As Long
    Dim lIdx As Long
    Dim sCurText As String

    If rstShipVia Is Nothing Then Exit Function
    
    rstShipVia.Filter = adFilterNone
    rstShipVia.Filter = "ShipMethID LIKE '" & GetWhseIDFromWhseKey(WhseKey) & "%'"
    
    sCurText = Trim(cboShipVia.text)
    cboShipVia.Clear
    
    rstShipVia.MoveFirst
    lIdx = -1
    With rstShipVia
        Do While Not .EOF
            cboShipVia.AddItem Trim(Mid(.Fields("ShipMethID").value, 5, 20))
            cboShipVia.ItemData(cboShipVia.NewIndex) = .Fields("ShipMethKey").value
            .MoveNext
        Loop
    End With

    Dim i As Integer
    For i = 0 To cboShipVia.ListCount - 1
        If Trim(cboShipVia.List(i)) = sCurText Then
            lIdx = i
        End If
    Next

    If cboShipVia.ListCount > 0 Then
        If lIdx > -1 Then
            cboShipVia.ListIndex = lIdx
            RecalcShipVia = cboShipVia.ItemData(lIdx)
        Else
            cboShipVia.ListIndex = 0
            RecalcShipVia = cboShipVia.ItemData(0)
            msg "Warning - Previously selected shipping method not supported at this warehouse."
        End If
    Else
        msg "Error - There are no shipping methods for this warehouse - See Jon"
        RecalcShipVia = 0
    End If
    
    rstShipVia.Filter = adFilterNone

End Function


'Called by
'   SetUpShipVia() above

Private Sub LoadShipViaCombo(cboCombo As ComboBox, _
                                rst As ADODB.Recordset, _
                                sDisplayField As String, _
                                Optional vKeyField As Variant, _
                                Optional vDfltKeyValue As Variant)
    Dim lIdx As Long
    Dim sCurText As String
    Dim lIndex As Long
    
    sCurText = Trim(cboCombo.text)
    Debug.Print "*" & sCurText
    cboCombo.Clear
    
    rst.MoveFirst
    lIdx = -1
    With rst
        Do While Not .EOF
            'Strip Out 'MPK-'
            cboCombo.AddItem Trim(Mid(.Fields(sDisplayField).value, 5, 20))
            
            If Not IsMissing(vKeyField) Then
                cboCombo.ItemData(cboCombo.NewIndex) = .Fields(vKeyField).value
            End If
            
            .MoveNext
        Loop
    End With
    
    For lIndex = 0 To cboCombo.ListCount - 1
        If Not IsMissing(vDfltKeyValue) Then
            If Not IsMissing(vKeyField) Then
                If vDfltKeyValue = cboCombo.ItemData(lIndex) Then
                    lIdx = lIndex
                    Exit For
                End If
            Else
                If vDfltKeyValue = Trim(cboCombo.List(lIndex)) Then
                    lIdx = lIndex
                End If
            End If
        End If
    Next

    If lIdx = -1 Then
        Dim i As Integer
        For i = 0 To cboCombo.ListCount - 1
            If InStr(1, cboCombo.List(i), sCurText) > 0 Then
                lIdx = i
            End If
        Next
    End If
    
    If cboCombo.ListCount > 0 Then
        If lIdx > -1 Then
            cboCombo.ListIndex = lIdx
            'm_udtOrder.lShipMethKey = cboCombo.ItemData(lIdx)
        Else
            cboCombo.ListIndex = 0
            'm_udtOrder.lShipMethKey = cboCombo.ItemData(0)
            msg "Warning - Previously selected shipping method not supported at this warehouse."
        End If
    Else
        msg "Error - There are no shipping methods for this warehouse - See Jon"
        'm_udtOrder.lShipMethKey = 0
    End If
End Sub





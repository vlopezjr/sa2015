VERSION 5.00
Begin VB.Form FLoadOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load Order"
   ClientHeight    =   2145
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEdit 
      Caption         =   "I want to Edit this order and understand that this will cancel the associated Acuity order."
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   840
      Width           =   3552
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2100
      TabIndex        =   3
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "Do you want to View or Edit this committed order?"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "This order is committed to Acuity."
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4152
   End
End
Attribute VB_Name = "FLoadOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sResult As String
Private m_SOKey As Long
Private m_bLoad As Boolean


Public Property Get Result() As String
    Result = m_sResult
End Property


Private Sub chkEdit_Click()
    If m_bLoad = True Then Exit Sub
    
    m_bLoad = True
    If chkEdit.value = vbChecked Then
        If CheckSPOFrozen Then
            chkEdit.value = vbUnchecked
        End If
    End If
    
    cmdEdit.Enabled = (chkEdit.value = vbChecked)
    
    If cmdEdit.Enabled = True Then
        TryToSetFocus cmdEdit
    Else
        TryToSetFocus cmdView
    End If
    
    m_bLoad = False
End Sub


'09/16/02       TeddyX
'This function is used to check if dropship order has associated PO
'If it has, check the status of the PO. If the status is cancelled, this order can be editted
'If the status is 0 and 1, ask the user to contact purchase agent.
'Otherwise, throw out a message to user that this order can only be viewed.

Private Function CheckDropShipPO() As Boolean
    Dim sSQL As String
    Dim orst As ADODB.Recordset
    Dim sTemp As String
    
    sSQL = "select distinct tsoSalesOrder.TranNo as SOID, tpoPurchOrder.TranNo AS POID, timBuyer.BuyerID, tpoPurchOrder.Status " _
        & "from tsoSalesOrder inner join tsoSOLine on tsoSOLine.SOKey = tsoSalesOrder.SOKey " _
        & "inner join tpoPOLine on tpoPOLine.POLineKey = tsoSOLine.POLineKey " _
        & "inner join tpoPurchOrder on tpoPOLine.POKey = tpoPurchOrder.POKey " _
        & "inner join timBuyer on tpoPurchOrder.BuyerKey = timBuyer.BuyerKey " _
        & "where tsoSOLine.SOKey = " & m_SOKey
    
    Set orst = LoadDiscRst(sSQL)
    
    If Not orst.EOF Then
        If orst.Fields("Status").value = 0 Or orst.Fields("Status").value = 1 Then
            sTemp = "Dropship Order SO " & StripLeadingZeros(Trim(orst.Fields("SOID").value)) & " has items on PO " & StripLeadingZeros(Trim(orst.Fields("POID").value)) & _
                ". You need to contact " & vbCrLf & "Purchasing Agent " & Trim(orst.Fields("BuyerID").value) & " and resolve this before you can edit this order. " & vbCrLf & vbCrLf & _
                "You can only view this order for now."
            msg sTemp, vbOKOnly + vbExclamation
            CheckDropShipPO = True
        ElseIf orst.Fields("Status").value <> 3 Then
            sTemp = "Dropship Order SO " & StripLeadingZeros(Trim(orst.Fields("SOID").value)) & " has items on PO " & StripLeadingZeros(Trim(orst.Fields("POID").value)) & _
                ". The PO status indicates that" & vbCrLf & "this order is not eligible for editting. You can contact " & vbCrLf & "Purchasing Agent " & Trim(orst.Fields("BuyerID").value) & " for details. " & vbCrLf & vbCrLf & _
                "You can only view this order for now."
            msg sTemp, vbOKOnly + vbExclamation
            CheckDropShipPO = True
        End If
    End If
    
    Set orst = Nothing
End Function


Private Function CheckSPOFrozen() As Boolean
    Dim orst As ADODB.Recordset
    Dim oRstDropship As ADODB.Recordset
    Dim oRstDropShipPO As ADODB.Recordset
    Dim sTemp As String
    Dim iIndex As Long
    Dim sSQL As String
    
    '09/16/02       TeddyX
    'Check if this order is dropship order first.
    
    sSQL = "Select (CASE WHEN tcpSO.flags&0x1 = 0x1 THEN 1 ELSE 0 END) AS IsDropShip " _
        & "from tcpSO where SOKey = " & m_SOKey
    Set oRstDropship = LoadDiscRst(sSQL)
    
    If oRstDropship.EOF Then
        msg "This Sage order does not exist in SageAssistant. Please contact IT department."
    Else
        If oRstDropship.Fields("IsDropShip").value = 1 Then
            CheckSPOFrozen = CheckDropShipPO
        Else
            Set orst = CallSP("spcpcHasFrozenSPO", "@SOKey", m_SOKey)
            
            If Not orst.EOF Then
                If orst.RecordCount = 1 Then
                    sTemp = "SO " & Trim(orst.Fields("SOID").value) & " has speical order items on PO " & Trim(orst.Fields("POID").value) & _
                        ". You need to contact " & vbCrLf & "Purchasing Agent " & Trim(orst.Fields("BuyerID").value) & " and resolve this before this order can be edited. " & vbCrLf & vbCrLf & _
                        "At this point the order can only be opened for viewing."
                Else
                    orst.MoveFirst
                    iIndex = 1
                    sTemp = "SO " & Trim(orst.Fields("SOID").value) & " has special order items on" & vbCrLf & vbCrLf
                    While Not orst.EOF
                        sTemp = sTemp & iIndex & ". PO " & Trim(orst.Fields("POID").value) & ". Purchasing Agent is " & Trim(orst.Fields("BuyerID").value) & vbCrLf
                        iIndex = iIndex + 1
                        orst.MoveNext
                    Wend
                    
                    sTemp = sTemp & vbCrLf & "You need to contact Purchasing Agents and resolve this before this order can be edited." & vbCrLf & _
                                "At this point the order can only be opened for viewing."
                
                End If
                msg sTemp, vbOKOnly + vbExclamation, "Frozen SPOs"
                CheckSPOFrozen = True
            End If
        End If
    End If
    
    Set orst = Nothing
    Set oRstDropship = Nothing
End Function


Private Sub cmdCancel_Click()
    m_sResult = "Cancel"
    Me.Hide
End Sub

Private Sub cmdEdit_Click()
    m_sResult = "Edit"
    Me.Hide
End Sub


Public Sub LoadSageOrder(ByRef lSOKey As Long)
    m_SOKey = lSOKey
    Me.Show vbModal
End Sub


Private Sub cmdView_Click()
    m_sResult = "View"
    Me.Hide
End Sub

Private Sub Form_Activate()
    TryToSetFocus cmdView
End Sub

Private Sub lblEditWarning_Click()
    chkEdit_Click
End Sub

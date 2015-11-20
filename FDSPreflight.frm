VERSION 5.00
Begin VB.Form FDSPreflight 
   Caption         =   "Preflight"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtShipWarnings 
      Height          =   3495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   8532
   End
End
Attribute VB_Name = "FDSPreflight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sPSKeys As String
Private m_sWarnings As String
Private m_aDataSource As Variant
Private m_iRecordCount As Integer



Public Property Get DataSource() As Variant
    DataSource = m_aDataSource
End Property

Public Property Let DataSource(ByVal oNewValue As Variant)
    m_aDataSource = oNewValue
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    Dim remarks As String
   
    SetWaitCursor True
    
    Printer.Font = "Arial"
    Printer.FontSize = 14
    Printer.Print "Open Drop Ship Preflight/Order Report" & "     " & Date
    Printer.Print
    Printer.FontSize = 10
    
    Printer.Print m_sWarnings
    Printer.Print
           
    Printer.EndDoc
    
    SetWaitCursor False
    msg "Report printed on " & Printer.DeviceName
    
End Sub

Private Sub Form_Load()

    Dim sTemp As String
    Dim sTempWarnings As String
    Dim i As Integer
    
    m_iRecordCount = UBound(m_aDataSource, 2) + 1
    For i = 0 To m_iRecordCount - 1
         If CBool(m_aDataSource(0, i)) Then
             sTemp = sTemp & "SO#: " & CStr(m_aDataSource(9, i)) & " PO#: " & CStr(m_aDataSource(8, i)) & " Vendor: " & CStr(m_aDataSource(16, i)) & " Customer: " & CStr(m_aDataSource(17, i)) & vbCrLf _
             & Space(4) & "Tracking No: " & GetTrackingNumber(CStr(m_aDataSource(3, i))) & " Freight: " & GetAmount(CStr(m_aDataSource(4, i))) & " Tax: " & GetAmount(CStr(m_aDataSource(6, i))) _
             & " Handling: " & GetAmount(CStr(m_aDataSource(5, i))) & " Packing: " & GetAmount(m_aDataSource(7, i)) & vbCrLf _
             & GetItemQuantities(CInt(m_aDataSource(15, i))) _
             & Space(4) & "Created By: " & CStr(m_aDataSource(26, i)) & vbCrLf & vbCrLf
             
                          
             sTempWarnings = GetShipWarnings(CStr(m_aDataSource(15, i)), CStr(m_aDataSource(8, i)), CStr(m_aDataSource(4, i)), CStr(m_aDataSource(3, i)))
             
             If Len(sTempWarnings) > 0 Then
                sTemp = sTemp & sTempWarnings
             End If
            
             
         End If
    Next
    
    m_sWarnings = sTemp
    
    If Len(Trim(m_sWarnings)) = 0 Then
        txtShipWarnings.text = "No Exceptions Found!"
    Else
        txtShipWarnings.text = m_sWarnings
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Function GetShipWarnings(sVouchers As String, sPOID As String, sFreight As String, sTrackingNo As String) As String
    Dim rst As ADODB.Recordset
    Dim sPSKey As String
    Dim iOpKey As Long
    Dim sWarnings As String
    Dim sBuffer As String

    
    Set rst = CallSP("spcpcDropShipmentCheck", "@Shipments", sVouchers)

    If rst.EOF Then Exit Function

    
    With rst
        Do While Not .EOF
        
            If sPSKey <> CStr(.Fields("PSKey").value) Then
                If Len(sWarnings) > 0 Then
              
                    sBuffer = sBuffer & vbCrLf & sWarnings
                  
                    sWarnings = vbNullString
                End If
                
                sPSKey = CStr(.Fields("PSKey").value)
                iOpKey = CLng(.Fields("OPNbr").value)
            End If
            
            'Free Parts - No Charge
            If .Fields("PartsNoCharge").value = 1 And .Fields("OrderAmt").value > 0 Then
                sWarnings = sWarnings & "    Error - Order is marked No Charge but shipment has value. " & vbCrLf
            End If
            
            'International
            If Trim(.Fields("ShiptoCountryId").value) <> "USA" Then
                sWarnings = sWarnings & "    Warning - Order is shipped to " & Trim(.Fields("ShiptoCountryId").value) _
                & "." & vbCrLf
            End If
            
            'Sales Tax
            If .Fields("TaxError").value = 1 Then
                sWarnings = sWarnings & "    Error - Sage and Order Pad Sales Tax values don't match. " & vbCrLf
            End If

            'Interstate
            If (.Fields("ShipAddrSTaxSchdKey").value = g_lIStateDfltSchdKey) _
                And (.Fields("STaxAmt").value > 0) Then
                sWarnings = sWarnings & "    Error - Customer is Interstate and is being charged Sales Tax." & vbCrLf
            End If
            
            'Government
            If (.Fields("ShipAddrSTaxSchdKey").value = g_lGovtDfltSchdKey) _
                And (.Fields("STaxAmt").value > 0) Then
                sWarnings = sWarnings & "    Error - Customer is Government and is being charged Sales Tax." & vbCrLf
            End If
            
            'International
            If (.Fields("ShipAddrSTaxSchdKey").value = g_lIntlDfltSchdKey) _
                And (.Fields("STaxAmt").value > 0) Then
                sWarnings = sWarnings & "    Error - Customer is International and is being charged Sales Tax." & vbCrLf
            End If
            
            'Resale
            If (Len(Trim(.Fields("ShipAddrSTaxExemptNo").value)) > 0) And (.Fields("STaxAmt").value > 0) Then
                sWarnings = sWarnings & "    Warning - Customer has a Resale Certificate and is being charged Sales Tax." _
                & vbCrLf
            End If
            
            'No Charge
            If .Fields("PartsNoCharge").value = 0 And .Fields("OrderAmt").value = 0 Then
                sWarnings = sWarnings & "    Warning - Order is not marked No Charge but shipment has no value. " _
                & vbCrLf
            End If
            
            'Free Freight
            If .Fields("FreeFreight").value = 1 And .Fields("FreightAmt").value > 0 _
                And InStr(1, .Fields("ShipMethod").value, "Call") = 0 Then
                sWarnings = sWarnings & "    Error - Order is marked Free Freight but shipment has freight. " & vbCrLf
            End If
            
            'High Freight
            If .Fields("FreightAmt").value > 75 Then
                sWarnings = sWarnings & "    Warning - Order has high freight - " & _
                GetAmount(.Fields("FreightAmt").value) & "." & vbCrLf
            End If

            'Zero Freight and Not Free Freight or WillCall
            If .Fields("FreeFreight").value = 0 And .Fields("FreightAmt").value = 0 _
                And InStr(1, .Fields("ShipMethod").value, "Call") = 0 Then
                'And InStr(1, .Fields("ShipMethod").Value, "Call") = 0 And Trim(.Fields("UPSAcct").Value) = "" Then
                sWarnings = sWarnings & "    Warning - The Shipment has no freight but the order is not WillCall or marked Free Freight " _
                & vbCrLf
            End If
                        
            'Ship Complete
            If .Fields("ShipComplete").value = 1 And .Fields("BackOrders").value = 1 Then
                sWarnings = sWarnings & "    Error - Order is marked Ship Complete but shipment has backorders. " _
                & vbCrLf
            End If
            
            'Inbound Freight
            If .Fields("InboundFreight").value = 1 Then
                If .Fields("InboundFreightAmt").value > 0 Then
                    sWarnings = sWarnings & "    Warning - Order is marked Inbound Freight (" & GetAmount(.Fields("InboundFreightAmt").value) _
                    & ")" & vbCrLf
                Else
                    sWarnings = sWarnings & "    Warning - Order is marked Inbound Freight. " & vbCrLf
                End If
            End If
            
            'Reduced Freight
            If .Fields("ReducedFreight").value = 1 Then
                'Add the information of Reduced Freight Amt to the report
                If .Fields("BillMethKey").value > 0 Then
                    g_rstShipVia.Filter = "ShipMethKey = " & .Fields("BillMethKey").value
                    
                    sWarnings = sWarnings & "    Warning - Order is marked Reduced Freight (shipped " _
                    & .Fields("ShipMethod").value & ", billed " & g_rstShipVia.Fields("ShipMethID").value & ")" & vbCrLf
                    
                    g_rstShipVia.Filter = adFilterNone
                Else
                    sWarnings = sWarnings & "    Warning - Order is marked Reduced Freight. " & vbCrLf
                End If
            End If
            
            'Deposit
            If .Fields("Deposit").value = 1 Then
                If .Fields("DepositAmt").value > 0 Then
                    sWarnings = sWarnings & "    Warning - Order is marked Deposit (" & GetAmount(.Fields("DepositAmt").value) & ")" & vbCrLf
                Else
                    sWarnings = sWarnings & "    Warning - Order is marked Deposit. " & vbCrLf
                End If
            End If
            
            'No Tracking Number
            'If Trim(.Fields("ShipTrackNo").value) = "" And InStr(1, .Fields("ShipMethod").value, "Call") = 0 Then
            '    sWarnings = sWarnings & "    Warning - Order is missing a tracking number." & vbCrLf
            'End If
            
            'Bill Recipient - Freight more then handling charge.
            'Is marked Bill Recipient, is not WIllCall,and the
            'freight does not contain only Handling charges.
            If Len(Trim(.Fields("UPSAcct").value)) > 0 And Not IsHandlingCharge(.Fields("FreightAmt")) Then
                sWarnings = sWarnings & "    Error - Order is UPS Bill Recipient, " _
                            & "but has inappropriate handling. Freight: " & _
                            GetAmount(.Fields("FreightAmt").value) & vbCrLf
            End If
                        
            'Not Bill Recipient - Freight less then or egual to $2.00.
            If Len(Trim(.Fields("UPSAcct"))) = 0 And InStr(1, .Fields("ShipMethod").value, "Call") = 0 _
            And .Fields("FreightAmt") <= 2 And .Fields("FreightAmt") > 0 And .Fields("FreeFreight") = 0 Then
                sWarnings = sWarnings & "    Warning - Freight is too low. Freight: " & _
                            GetAmount(.Fields("FreightAmt").value) & vbCrLf
            End If

            If .Fields("HasTrueCompressor") > 0 Then
                sWarnings = sWarnings & "    Attention - Has a True Compressor" & vbCrLf
            End If
            
            'Comment back in when bill recpient issues are resolved
            'If default ship address is Bill Recipient and freight > $2.00
            'If (.Fields("CustUPSAcct") = 1 And .Fields("FreightAmt") > 2) Then
                'sWarnings = sWarnings & "    Warning - This customer is usually bill recipient, but freight is " & _
                            'Format$(.Fields("FreightAmt").value, "$###,###.##") & vbCrLf
            'End If
            
            
            If .Fields("FreeFreight").value = 1 Then
                sWarnings = sWarnings & "    Free Freight Memos" & vbCrLf
                sWarnings = sWarnings & "        " & FetchRemarks(iOpKey, "Order.FreeFreight")
            End If
            
            If .Fields("ReducedFreight").value = 1 Then
                sWarnings = sWarnings & "    Reduced Freight Memos" & vbCrLf
                sWarnings = sWarnings & "        " & FetchRemarks(iOpKey, "Order.ReducedFreight")
            End If
            
            If .Fields("InboundFreight").value = 1 Then
                sWarnings = sWarnings & "    Inbound Freight Memos" & vbCrLf
                sWarnings = sWarnings & "        " & FetchRemarks(iOpKey, "Order.InboundFreight")
            End If
            
            If .Fields("Deposit").value = 1 Then
                sWarnings = sWarnings & "    Deposit Memos" & vbCrLf
                sWarnings = sWarnings & "        " & FetchRemarks(iOpKey, "Order.Deposit")
            End If
            
            If .Fields("PartsNoCharge").value = 1 Then
                sWarnings = sWarnings & "    Parts No Charge Memos" & vbCrLf
                sWarnings = sWarnings & "        " & FetchRemarks(iOpKey, "Order.PartsNoCharge")
            End If
            
           
            
            'If Len(Trim((.Fields("OtherSHMemo")))) > 0 Then
            '    sWarnings = sWarnings & "    Other Shipment Memo - " & .Fields("OtherSHMemo") & vbCrLf
            'End If
            
            
            .MoveNext

            
        Loop
        
        If Len(sWarnings) > 0 Then
            sBuffer = sBuffer & vbCrLf & sWarnings & vbCrLf & vbCrLf
        End If
        
        GetShipWarnings = sBuffer
    End With
    
    CloseRst rst
End Function

Private Function IsHandlingCharge(amount As Double) As Boolean
    'Handling charges for MPK, STL, and SEA repectively.
    If amount = 1.5 Or amount = 2 Or amount = 1.75 Then
        IsHandlingCharge = True
    Else
        IsHandlingCharge = False
    End If
End Function

'fetch all dropship remarks for an order
'if there are none, a null string is returned
Private Function FetchRemarks(OPKey As Long, Addressee As String) As String
    Dim orst As ADODB.Recordset
    Dim s As String

    Set orst = New ADODB.Recordset
    With orst
        .Open "SELECT EffectiveDate, Sender, MemoText FROM tciMemo WHERE tciMemo.MemoOwnerKey=" & OPKey _
                & " AND (tciMemo.Addressee = '" & Addressee & "')", g_DB.Connection
        Do While Not .EOF
            s = s & Space(4) & .Fields("EffectiveDate") & ", " & .Fields("Sender") & ", " & .Fields("MemoText") & vbCrLf
            .MoveNext
        Loop
    End With
    FetchRemarks = s
End Function

Private Function GetAmount(ByVal amount As String) As String
    If Len(amount) = 0 Or amount = "0" Then
        GetAmount = "$0.00"
    Else
        GetAmount = Format$(amount, "$###,###.##")
    End If
End Function

Private Function GetTrackingNumber(ByVal trackingNo As String) As String
    If Len(trackingNo) > 0 Then
        GetTrackingNumber = trackingNo
    Else
        GetTrackingNumber = "    "
    End If
End Function

Private Function GetItemQuantities(ByVal pskey As Integer) As String
    Dim orst As ADODB.Recordset
    Dim s As String
    Dim sql As String
    
    sql = "select l.ItemId, d.QtyOrd, l.QtyShipped from tcpProvisionalShipment s " _
                & "join  tcpProvisionalShipLine l on s.PSKey=l.PSKey " _
                & "join tpoPOLine pl on l.PoLineKey=pl.POLineKey " _
                & "join tsoSOLineDist d on l.SoLineKey=d.SOLineKey " _
                & "and s.PSKey=" & pskey
                
    Set orst = New ADODB.Recordset
    With orst
        .Open sql, g_DB.Connection
                
        Do While Not .EOF
            s = s & Space(4) & .Fields("ItemId") & "    Shipping " & .Fields("QtyShipped") & " of " & .Fields("QtyOrd") & vbCrLf
            .MoveNext
        Loop
    End With
    
    GetItemQuantities = s
End Function

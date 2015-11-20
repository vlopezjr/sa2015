VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SalesTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'For Future use with Customer Maintenance
'This event will be used by FBilling to notify the
'AR Rep that SalesTax Schedule is not correct.

'Public Event DefaultTax(io_sErrorMsg As String)

Public Event Changed()

Private Type TaxProps
    lSTaxSchdKey As Long
    sSTaxSchdID As String
    sExmptNo As String
    dTaxRate As Double
    lCustKey As Long
End Type

Private Type EmailNotification
    sProcedure As String
    sSubject As String
    sMessage As String
End Type


Private m_udtTax As TaxProps
Private m_udtTaxBackup As TaxProps
Private m_udtNotify As EmailNotification


'***Initialization***

Private Sub Class_Initialize()
End Sub



'***Public Properties***

Public Property Get IsInterState() As Boolean
    'This string constant could be a problem (6/24/10 LR)
    IsInterState = IIf(LCase(Trim$(m_udtTax.sSTaxSchdID)) = "interstate", True, False)
End Property


Public Property Get STaxSchdKey() As Long
    STaxSchdKey = m_udtTax.lSTaxSchdKey
End Property


Public Property Get STaxSchdID() As String
    STaxSchdID = m_udtTax.sSTaxSchdID
End Property


Public Property Get ExmptNo() As String
    ExmptNo = Trim(m_udtTax.sExmptNo)
End Property

'Do we need this anymore?
Public Property Let ExmptNo(ByVal sNewValue As String)
    m_udtTax.sExmptNo = sNewValue
End Property


Public Property Get TaxRate() As Double
    TaxRate = m_udtTax.dTaxRate
End Property

Public Property Get FormatedTaxRate() As String
    FormatedTaxRate = Format$(m_udtTax.dTaxRate * 100, "#0.0000")
End Property



'***Private Methods***

'Called by
'   SalesTax.Init
Private Sub LoadWATax(ByRef i_oAddr As Address)
    Dim cmd As ADODB.Command
    Dim sLocCode As String
    Dim iResultCode As Integer

    iResultCode = GetWASTaxLocCode(i_oAddr.Addr1, i_oAddr.City, i_oAddr.Zip, sLocCode)
    If (iResultCode > 2) Then
        '3: The address, Zip+4, and Zip could not be found.
        '4: Invalid arguments.
        '5: Internal error.
            
        'Use a default tax rate.
        m_udtNotify.sProcedure = "LoadWATax"
        m_udtNotify.sSubject = "Error calculating Sales Tax. The default Sales Tax rate has been applied."
        m_udtNotify.sMessage = "Unable to retrieve a Location Code from the WebService for the address " _
                & vbCrLf & "provided. A default Sales Tax rate has been applied for this address." _
                & vbCrLf & "OP: " & i_oAddr.OPKey
        
        UseDefaultTaxRateWithNotify i_oAddr, False
    Else
        '0: The address was found.
        '1: The address was not found, but the Zip+4 was located.
        '2: Neither address or Zip+4 were found, but the 5-digit Zip was located.
        
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = g_DB.Connection
        cmd.CommandText = "spcpcLoadWASalesTax"
        cmd.CommandType = adCmdStoredProc
        
        cmd.Parameters("@_iLocCode").value = sLocCode
        cmd.Execute
        
        If Not IsNull(cmd.Parameters("@_oSchdKey").value) Then
            m_udtTax.sSTaxSchdID = Trim(cmd.Parameters("@_oSchdID").value)
            m_udtTax.lSTaxSchdKey = cmd.Parameters("@_oSchdKey").value
            m_udtTax.dTaxRate = cmd.Parameters("@_oRate").value
        Else
            'Use a default tax rate.
            m_udtNotify.sProcedure = "LoadWATax"
            m_udtNotify.sSubject = "Error calculating Sales Tax. The default Sales Tax rate has been applied."
            m_udtNotify.sMessage = "Unable to retrieve Sales Tax Schedule information from database for " _
                    & vbCrLf & "Location Code: " & sLocCode & ". A default Sales Tax rate has been applied " _
                    & vbCrLf & "for this address. Please check tciSTaxCode." _
                    & vbCrLf & "OP: " & i_oAddr.OPKey
            
            UseDefaultTaxRateWithNotify i_oAddr, True
        End If
    End If
End Sub


Private Sub LoadByZip(ByRef i_oAddr As Address)
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = g_DB.Connection
    cmd.CommandText = "spcpcLoadSalesTaxByZip"
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters("@_iZipCode").value = i_oAddr.Zip
    cmd.Execute
    
    If Not IsNull(cmd.Parameters("@_oSchdKey").value) Then
        m_udtTax.sSTaxSchdID = Trim(cmd.Parameters("@_oSchdID").value)
        m_udtTax.lSTaxSchdKey = cmd.Parameters("@_oSchdKey").value
        m_udtTax.dTaxRate = cmd.Parameters("@_oRate").value
    Else
        'Use a default tax rate.
        m_udtNotify.sProcedure = "LoadByZip"
        m_udtNotify.sSubject = "Error calculating Sales Tax. The default Sales Tax rate has been applied."
        m_udtNotify.sMessage = "Unable to retrieve Sales Tax Schedule information from database for " _
                & vbCrLf & "Zip Code: " & Left(i_oAddr.Zip, 5) & ". A default Sales Tax rate has been applied " _
                & vbCrLf & "for this address. Please check tcpZipTax and tsmPostalCode." _
                & vbCrLf & "OP: " & i_oAddr.OPKey
        
        UseDefaultTaxRateWithNotify i_oAddr, True
    End If
End Sub


'Called By:
'   SalesTax.Init
Private Sub LoadTax(i_lCustKey As Long, ByRef i_oAddr As Address)
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = g_DB.Connection
    cmd.CommandText = "spcpcLoadSalesTax"
    cmd.CommandType = adCmdStoredProc

    cmd.Parameters("@_iCustKey").value = i_lCustKey
    cmd.Parameters("@_iAddrKey").value = i_oAddr.AddrKey
    cmd.Execute
    
    If Not IsNull(cmd.Parameters("@_oSchdKey").value) Then
        m_udtTax.sSTaxSchdID = Trim(cmd.Parameters("@_oSchdID").value)
        m_udtTax.lSTaxSchdKey = cmd.Parameters("@_oSchdKey").value
        m_udtTax.dTaxRate = cmd.Parameters("@_oRate").value
    Else
        'Use a default tax rate.
        m_udtNotify.sProcedure = "LoadTax"
        m_udtNotify.sSubject = "Error calculating Sales Tax. The default Sales Tax rate has been applied."
        m_udtNotify.sMessage = "Unable to retrieve Sales Tax Schedule information from database for " _
                & vbCrLf & "CustKey: " & i_lCustKey & " and AddrKey: " & i_oAddr.AddrKey & "." _
                & vbCrLf & "A default Sales Tax rate has been applied for this address."
        
        UseDefaultTaxRateWithNotify i_oAddr, True
    End If
End Sub

Private Function GetWASTaxLocCode(i_sStreet As String, i_sCity As String, i_sZip As String, ByRef i_sLocCode As String) As Integer
    Dim oProxy As MSSOAPLib30.SoapClient30
    Set oProxy = New MSSOAPLib30.SoapClient30

    On Error GoTo EH
    
    oProxy.MSSoapInit g_WASalesTaxUri & "?WSDL"
    
    GetWASTaxLocCode = oProxy.GetLocCode(i_sStreet, i_sCity, i_sZip, i_sLocCode)
    Exit Function
    
EH:
    GetWASTaxLocCode = 5
End Function

Private Function HasNexus(i_sState As String) As Boolean
    HasNexus = (i_sState = "CA" Or i_sState = "MO" Or i_sState = "WA")
End Function

Private Sub GetExemptNo(i_sState As String)
    Dim cmd As ADODB.Command

    m_udtTax.sExmptNo = ""
    
    If m_udtTax.lCustKey <> 0 Then
        If HasNexus(i_sState) Then
            Set cmd = CreateCommandSP("spcpcGetExemptionNo")
        
            With cmd
                .Parameters("@_iCustKey") = m_udtTax.lCustKey
                .Parameters("@_sState") = i_sState
                .Execute
                
                'Returns an empty string if not found.
                m_udtTax.sExmptNo = Trim$(.Parameters("@_sExemptNo").value)
                
                If Len(m_udtTax.sExmptNo) > 0 Then
                    m_udtTax.dTaxRate = 0
                End If
            End With
        End If
    End If
End Sub



'***Public Methods***

'Called By
'   Order.Load
'   FOrder.cmdSelectCustomer_Click(2)
'   FOrder.cmdSelectCustomer_Click(3)
'   FOrder.cmdNewOrder_Click
'   FOrder.cmdEditAddr_Click(0)
'   FOrder.cmdEditAddr_Click(1)
'   FOrder.txtCustID_LostFocus
Public Sub Init(ByRef i_oCustomer As Customer)
    Dim orst As ADODB.Recordset
    Dim sSQL As String
    Dim iResultCode As Integer
    Dim sLocCode As String
    Dim oAddr As Address

    Set oAddr = i_oCustomer.ShipAddr

    m_udtTax.lCustKey = i_oCustomer.Key
 
    If oAddr.AddrType = TOO Then
        If i_oCustomer.IsGovernment Then
            m_udtTax.sSTaxSchdID = g_sGovtDfltSchdID
            m_udtTax.lSTaxSchdKey = g_lGovtDfltSchdKey
            m_udtTax.dTaxRate = g_dGovtDfltTaxRate
        Else
            If oAddr.CountryID = "USA" Then
                If HasNexus(oAddr.State) Then
                    'WA
                    If oAddr.State = "WA" Then
                        LoadWATax oAddr
                    Else
                    'MO, CA
                        LoadByZip oAddr
                    End If
                    
                    If i_oCustomer.HasAccount Then
                        GetExemptNo oAddr.State
                    End If
                Else
                    'Interstate
                    m_udtTax.sSTaxSchdID = g_sIStateDfltSchdID
                    m_udtTax.lSTaxSchdKey = g_lIStateDfltSchdKey
                    m_udtTax.dTaxRate = g_dIStateDfltTaxRate
                    m_udtTax.sExmptNo = ""
                End If
            Else
                'International
                m_udtTax.sSTaxSchdID = g_sIntlDfltSchdID
                m_udtTax.lSTaxSchdKey = g_lIntlDfltSchdKey
                m_udtTax.dTaxRate = g_dIntlDfltTaxRate
                m_udtTax.sExmptNo = ""
            End If
        End If
    Else 'Permanent Address
        'All permanent addresses have a tax schedule assigned to them
        'including Government, International, and Interstate.
        LoadTax i_oCustomer.Key, oAddr

        If i_oCustomer.HasAccount And Not i_oCustomer.IsGovernment Then
            GetExemptNo oAddr.State
        End If
    End If

    Backup
    
    RaiseEvent Changed

End Sub

Public Sub Clear()
    m_udtTax.sSTaxSchdID = ""
    m_udtTax.lSTaxSchdKey = 0
    m_udtTax.sExmptNo = ""
    m_udtTax.dTaxRate = 0
End Sub

Public Function Export() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode

    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "SalesTax"
        .Add "STaxSchdKey", m_udtTax.lSTaxSchdKey
        .Add "STaxSchdID", Trim$(m_udtTax.sSTaxSchdID)
        .Add "ExmptNo", Trim$(m_udtTax.sExmptNo)
        .Add "TaxRate", m_udtTax.dTaxRate
        .AddChild ExportBackup
    End With
    
    Set Export = oXML
End Function

'to support split order
Private Function ExportBackup() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode

    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "SalesTaxBackup"
        .Add "STaxSchdKey", m_udtTaxBackup.lSTaxSchdKey
        .Add "STaxSchdID", Trim$(m_udtTaxBackup.sSTaxSchdID)
        .Add "ExmptNo", Trim$(m_udtTaxBackup.sExmptNo)
        .Add "TaxRate", m_udtTaxBackup.dTaxRate
    End With
    Set ExportBackup = oXML
End Function


'to support split order
Public Sub Import(i_oXML As JDMPDXML.XMLNode)
    Dim oTaxBackupXML As JDMPDXML.XMLNode

    With i_oXML
        m_udtTax.lSTaxSchdKey = .Item("STaxSchdKey").value
        m_udtTax.sSTaxSchdID = .Item("STaxSchdID").value
        m_udtTax.sExmptNo = .Item("ExmptNo").value
        m_udtTax.dTaxRate = .Item("TaxRate").value
    End With

    Set oTaxBackupXML = i_oXML.Item("SalesTaxBackup")

    With oTaxBackupXML
        m_udtTaxBackup.lSTaxSchdKey = .Item("STaxSchdKey").value
        m_udtTaxBackup.sSTaxSchdID = .Item("STaxSchdID").value
        m_udtTaxBackup.sExmptNo = .Item("ExmptNo").value
        m_udtTaxBackup.dTaxRate = .Item("TaxRate").value
    End With

    RaiseEvent Changed
End Sub


Public Function XmlCommit() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    With oXML
        .Tag = "salestax"
        .Add "schdkey", m_udtTax.lSTaxSchdKey
        .Add "schdid", Trim$(m_udtTax.sSTaxSchdID)
        .Add "exmptno", Trim$(m_udtTax.sExmptNo)
        .Add "rate", m_udtTax.dTaxRate
    End With
    Set XmlCommit = oXML
End Function


Public Function Serialize() As IXMLDOMElement
    Dim dom As New DOMDocument
    Dim element As IXMLDOMElement
    Dim property As IXMLDOMElement

    Set element = dom.createElement("salestax")
    Set property = dom.createElement("schdkey")
    element.appendChild property
    property.text = m_udtTax.lSTaxSchdKey
    Set property = dom.createElement("schdid")
    element.appendChild property
    property.text = Trim$(m_udtTax.sSTaxSchdID)
    Set property = dom.createElement("exmptno")
    element.appendChild property
    property.text = Trim$(m_udtTax.sExmptNo)
    Set property = dom.createElement("rate")
    element.appendChild property
    property.text = m_udtTax.dTaxRate
    
    Set Serialize = element
End Function


Public Sub Backup()
    m_udtTaxBackup = m_udtTax
End Sub

'Used by
'   Order.ShipMethKey when switching from WillCall.
Public Sub Restore()
    m_udtTax = m_udtTaxBackup
    RaiseEvent Changed
End Sub

'Called By
'   Order.Load
'   Order.ShipMethKey (Let)
'   FOrder.cmdSelectCustomer_Click(2)
'   FOrder.cmdSelectCustomer_Click(3)
'   FOrder.cmdNewOrder_Click
'   FOrder.cmdEditAddr_Click(0)
'   FOrder.cmdEditAddr_Click(1)
'   FOrder.txtCustID_LostFocus
'***Default values are hard-coded in the modular level constants.
Public Sub WillCallTaxOverride(i_sWhseID As String)
    Dim sState As String

    If m_udtTax.sSTaxSchdID <> g_sGovtDfltSchdID Then
        Select Case i_sWhseID
            Case "MPK"
                m_udtTax.sSTaxSchdID = g_sMPKDfltSchdID
                m_udtTax.lSTaxSchdKey = g_lMPKDfltSchdKey
                m_udtTax.dTaxRate = g_dMPKDfltTaxRate
                sState = "CA"
            Case "SEA"
                m_udtTax.sSTaxSchdID = g_sSEADfltSchdID
                m_udtTax.lSTaxSchdKey = g_lSEADfltSchdKey
                m_udtTax.dTaxRate = g_dSEADfltTaxRate
                sState = "WA"
            Case "STL"
                m_udtTax.sSTaxSchdID = g_sSTLDfltSchdID
                m_udtTax.lSTaxSchdKey = g_lSTLDfltSchdKey
                m_udtTax.dTaxRate = g_dSTLDfltTaxRate
                sState = "MO"
        End Select

        GetExemptNo sState
        
        RaiseEvent Changed
    End If
End Sub

Public Sub PropogateExemptCert(i_lSOKey As Long, i_sState As String)
    Dim cmd As ADODB.Command
    
    If HasNexus(i_sState) And (Len(ExmptNo) > 0) Then
        Debug.Print "propogate exmpt"
        Set cmd = CreateCommandSP("spcpcPropagateExmptCert")
        cmd.Parameters("@i_ExmptNo") = ExmptNo
        cmd.Parameters("@i_SOKey") = i_lSOKey
        cmd.Execute
    End If
End Sub


'Called By
'   SalesTax.LoadWATax
'   SalesTax.LoadByZip
'   SalesTax.LoadTax

Private Sub UseDefaultTaxRateWithNotify(ByRef i_oAddr As Address, i_bNotify As Boolean)

    Select Case i_oAddr.State
        Case "CA"
            m_udtTax.sSTaxSchdID = g_sMPKDfltSchdID
            m_udtTax.lSTaxSchdKey = g_lMPKDfltSchdKey
            m_udtTax.dTaxRate = g_dMPKDfltTaxRate
        Case "WA"
            m_udtTax.sSTaxSchdID = g_sSEADfltSchdID
            m_udtTax.lSTaxSchdKey = g_lSEADfltSchdKey
            m_udtTax.dTaxRate = g_dSEADfltTaxRate
        Case "MO"
            m_udtTax.sSTaxSchdID = g_sSTLDfltSchdID
            m_udtTax.lSTaxSchdKey = g_lSTLDfltSchdKey
            m_udtTax.dTaxRate = g_dSTLDfltTaxRate
    End Select
    
    LogDB.LogEvent "SalesTax", m_udtNotify.sProcedure, m_udtNotify.sMessage
    
    If i_bNotify Then
        EMail.Send GetUserName & "@caseparts.com", "operations@caseparts.com", m_udtNotify.sSubject, _
                    m_udtNotify.sMessage, False 'TextFormat
    End If
    
'*** 3/12/09 For Future use with Customer Maintenance
    'RaiseEvent DefaultTax(m_udtNotify.sMessage)
End Sub


'*************************
'*** Used Only by FBilling
'*************************
' On entry, address is guaranteed to be in the US
'TODO: deal with int'l addresses
'Called By
'   FBilling.CreateAcct
Public Sub SetTax(i_sZipCode As String, i_sStateID As String, Optional i_sStreet As Variant, Optional i_sCity As Variant)
    Dim cmd As ADODB.Command
    Dim iResultCode As Integer
    Dim sLocCode As String
    Dim sErrMsg As String

    m_udtTax.sSTaxSchdID = g_sIStateDfltSchdID
    m_udtTax.lSTaxSchdKey = g_lIStateDfltSchdKey
    m_udtTax.dTaxRate = g_dIStateDfltTaxRate
    
    On Error GoTo EH

    If i_sStateID = "WA" Then
        If IsMissing(i_sStreet) Or IsMissing(i_sCity) Then
            'Raise an error
            On Error GoTo 0
            Err.Raise -1, "SalesTax.SetTax", _
                "SalesTax.SetTax" & vbCrLf & "Invalid address. Missing information. Please check the address."
        Else
            iResultCode = GetWASTaxLocCode(CStr(i_sStreet), CStr(i_sCity), i_sZipCode, sLocCode)
            If (iResultCode > 2) Then
                '3: The address, Zip+4, and Zip could not be found.
                '4: Invalid arguments.
                '5: Internal error.
                
                'Raise error
                On Error GoTo 0
                Err.Raise -1, "SalesTax.SetTax", _
                    "SalesTax.SetTax" & vbCrLf & "Unable to obtain a location code for the zip: " & i_sZipCode
            Else
                '0: The address was found.
                '1: The address was not found, but the Zip+4 was located.
                '2: Neither address or Zip+4 were found, but the 5-digit Zip was located.
                sErrMsg = "Unable to find Tax Schedule information for Location code: " & _
                            sLocCode & " in the database."
                
                Set cmd = New ADODB.Command
                Set cmd.ActiveConnection = g_DB.Connection
                cmd.CommandText = "spcpcLoadWASalesTax"
                cmd.CommandType = adCmdStoredProc
                
                cmd.Parameters("@_iLocCode").value = sLocCode
                cmd.Execute
                
                'If nothing is returned, this assignment will generate an "Invalid Use of Null" error.
                m_udtTax.sSTaxSchdID = Trim(cmd.Parameters("@_oSchdID").value)
                m_udtTax.lSTaxSchdKey = cmd.Parameters("@_oSchdKey").value
                m_udtTax.dTaxRate = cmd.Parameters("@_oRate").value
            
            End If
        End If
    Else
        If HasNexus(i_sStateID) Then
            'LoadByZip i_sZipCode
            sErrMsg = "Unable to obtain Tax information for the zip code: " & i_sZipCode
            
            Set cmd = New ADODB.Command
            Set cmd.ActiveConnection = g_DB.Connection
            
            cmd.CommandText = "spcpcLoadSalesTaxByZip"
            cmd.CommandType = adCmdStoredProc
            
            cmd.Parameters("@_iZipCode").value = i_sZipCode
            cmd.Execute
            
            'If nothing is returned, this assignment will generate an "Invalid Use of Null" error.
            m_udtTax.sSTaxSchdID = Trim(cmd.Parameters("@_oSchdID").value)
            m_udtTax.lSTaxSchdKey = cmd.Parameters("@_oSchdKey").value
            m_udtTax.dTaxRate = cmd.Parameters("@_oRate").value
        End If
    End If

    RaiseEvent Changed
    
    Exit Sub
EH:
    Err.Raise -1, "SalesTax.SetTax", _
        "SalesTax.SetTax" & vbCrLf & sErrMsg

End Sub




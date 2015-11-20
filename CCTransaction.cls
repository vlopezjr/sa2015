VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*********************************************************************************
' Created 7/2/04 Len Russell
'
' These are Authorization, Charge and Error objects
'
' Uses the PayPal PayFlowPro DLL
'
'*********************************************************************************


'NOTE: These constants are also declared in CreditCard.Service1.asmx
'Private Const HostPort As Long = 443
'Private Const TimeoutValue As Long = 45
'Private Const UserName As String = "caseparts"
'Private Const Password As String = "abcd1234"


Private m_sAcct As String           'max length 19
Private m_sComment1 As String       'max length 128
Private m_sComment2 As String       'max length 128
Private m_dAmount As Double

Private m_lTranKey As Long
Private m_sTranType As String
Private m_datTimeStamp As Date
Private m_lOPKey As Long
Private m_sTranNo As String
Private m_lCCKey As Long

Private m_sParamList As String

Private m_sResponse As String
Private m_iResult As Integer
Private m_sPNREF As String
Private m_sRespMsg As String
Private m_sAuthCode As String
Private m_sAVSAddr As String
Private m_sAVSZip As String
Private m_sIAVS As String
'***466 9/20/06 LR removed
'Private m_sCVV2Match As String
Private m_lInvcKey As Long
Private m_sUserID As String


Public Property Get TranKey() As Long
    TranKey = m_lTranKey
End Property

Public Property Let TranKey(ByVal lNewValue As Long)
    m_lTranKey = lNewValue
End Property


Public Property Get TranType() As String
    TranType = m_sTranType
End Property

Public Property Let TranType(ByVal sNewValue As String)
    m_sTranType = sNewValue
End Property


Public Property Get TimeStamp() As Date
    TimeStamp = m_datTimeStamp
End Property

Public Property Let TimeStamp(ByVal datNewValue As Date)
    m_datTimeStamp = datNewValue
End Property


Public Property Get OPKey() As Long
    OPKey = m_lOPKey
End Property

Public Property Let OPKey(ByVal lNewValue As Long)
    m_lOPKey = lNewValue
End Property


Public Property Get tranNo() As String
    tranNo = m_sTranNo
End Property

Public Property Let tranNo(ByVal sNewValue As String)
    m_sTranNo = sNewValue
End Property


Public Property Get CCKey() As Long
    CCKey = m_lCCKey
End Property

Public Property Let CCKey(ByVal lNewValue As Long)
    m_lCCKey = lNewValue
End Property


Public Property Get ParamList() As String
    ParamList = m_sParamList
End Property

Public Property Let ParamList(ByVal sNewValue As String)
    m_sParamList = sNewValue
End Property


Public Property Get Comment1() As String
    Comment1 = m_sComment1
End Property

'max length 128
Public Property Let Comment1(ByVal sNewValue As String)
    m_sComment1 = Mid$(sNewValue, 1, 128)
End Property


Public Property Get Comment2() As String
    Comment2 = m_sComment2
End Property

'max length 128
Public Property Let Comment2(ByVal sNewValue As String)
    m_sComment2 = Mid$(sNewValue, 1, 128)
End Property


Public Property Get Response() As String
    Response = m_sResponse
End Property

Public Property Let Response(ByVal sNewValue As String)
    m_sResponse = Trim$(sNewValue)
    If Len(m_sResponse) > 0 Then
        m_iResult = GetValue(sNewValue, "RESULT=")
        m_sPNREF = GetValue(sNewValue, "PNREF=")
        m_sRespMsg = GetValue(sNewValue, "RESPMSG=")
        m_sAVSAddr = GetValue(sNewValue, "AVSADDR=")
        m_sAVSZip = GetValue(sNewValue, "AVSZIP=")
        m_sIAVS = GetValue(sNewValue, "IAVS=")
'***466 9/20/06 LR removed
'        m_sCVV2Match = GetValue(sNewValue, "CVV2MATCH=")
        m_sAuthCode = GetValue(sNewValue, "AUTHCODE=")
    End If
End Property


Public Property Get Result() As Integer
    Result = m_iResult
End Property

Public Property Get PNREF() As String
    PNREF = m_sPNREF
End Property

Public Property Get RESPMSG() As String
    RESPMSG = m_sRespMsg
End Property

Public Property Get AVSADDR() As String
    AVSADDR = m_sAVSAddr
End Property

Public Property Get AVSZIP() As String
    AVSZIP = m_sAVSZip
End Property

Public Property Get IAVS() As String
    IAVS = m_sIAVS
End Property

'***466 9/20/06 LR removed
'Public Property Get CVV2MATCH() As String
'    CVV2MATCH = m_sCVV2Match
'End Property

Public Property Get AUTHCODE() As String
    AUTHCODE = m_sAuthCode
End Property


Public Property Get UserID() As String
    UserID = m_sUserID
End Property

Public Property Let UserID(ByVal sNewValue As String)
    m_sUserID = Trim$(sNewValue)
End Property


Public Property Get Amount() As Double
    Amount = m_dAmount
End Property

Public Property Let Amount(ByVal dNewValue As Double)
    m_dAmount = dNewValue
End Property


Public Property Get InvcKey() As Long
    InvcKey = m_lInvcKey
End Property

Public Property Let InvcKey(ByVal lNewValue As Long)
    m_lInvcKey = lNewValue
End Property

'2/16/09 created this method to support the temp resurrection of AutoCommit

Public Function Authorize(ByRef oOrder As Order) As Boolean
    Dim oProxy As MSSOAPLib30.SoapClient30
    Dim strErrorMessage As String
    
    On Error GoTo ErrorHandler

    'Initialize the Transaction object

    TranType = "A"
    TimeStamp = Now
    OPKey = oOrder.OPKey
    tranNo = oOrder.tranNo         'NOTE: this is 0 at this point
    Comment1 = Trim$(oOrder.Customer.ID)
    Comment2 = "OP-" & oOrder.OPKey
    UserID = GetUserID(oOrder.UserKey)
    CCKey = oOrder.CreditCard.Key
    
    ' Add $10 to ensure the authorized amount will exceed the bill amount to prevent duplicate authorizations
    Amount = oOrder.Items.TotalPrice + oOrder.Items.TotalTax(oOrder.SalesTax.TaxRate) + ShipCharge(oOrder) + 10#

    ParamList = ParamList & "&ACCT=" & oOrder.CreditCard.CardNo
    ParamList = ParamList & FormatParam("NAME", oOrder.CreditCard.CardHolderName)
    ParamList = ParamList & "&EXPDATE=" & oOrder.CreditCard.ExpireMonth & oOrder.CreditCard.ExpireYear
    ParamList = ParamList & "&AMT=" & FormatAsMoney(Amount)
    ParamList = ParamList & FormatParam("COMMENT1", Comment1)
    ParamList = ParamList & FormatParam("COMMENT2", Comment2)
    ParamList = ParamList & FormatParam("STREET", oOrder.CreditCard.StreetNbr)
    '6/10/10 LR
    ParamList = ParamList & FormatParam("ZIP", oOrder.CreditCard.ZipCode)

    ' create the SOAP Proxy
    Set oProxy = New MSSOAPLib30.SoapClient30
    oProxy.MSSoapInit g_PayPalUri & "?WSDL"
    
    'the Response set property parses out the response string (see above)
    Response = oProxy.Authorize(ParamList)

    'testing RESULT in the Repsonse isn't enough
    'also want to test AVSZIP
    If Result <> 0 Then GoTo ErrorHandler

    'set return value
    Authorize = True

    Exit Function

'11/3/10 LR
'This might be better moved to AutoCommit
'At the moment though, the RESPMSG property of the Transaction object is not
'readily available to AutoCommit

ErrorHandler:
    On Error Resume Next    '???

    'setup the Error CCTransaction here

    strErrorMessage = "Credit card authorization for OP " & oOrder.OPKey & " failed ("
    If RESPMSG = "Invalid account number" Then
        strErrorMessage = strErrorMessage & RESPMSG & " " & oOrder.CreditCard.CardNo & ")"
    Else
        strErrorMessage = strErrorMessage & RESPMSG & ")"
    End If

    strErrorMessage = strErrorMessage & vbCrLf & "The order was not committed."


    EMail.SendToList "01", "creditcard@caseparts.com", "AutoCommit.Transaction.Authorize (" & g_DB.Server & ")", strErrorMessage & vbCrLf & Response, False 'TextFormat
    EMail.Send "creditcard@caseparts.com", UserID & "@caseparts.com", "AutoCommit.Transaction.Authorize", strErrorMessage, False 'TextFormat

    LogDB.LogEvent "CCTransaction", "Authorize", strErrorMessage & "  " & Response

    LogOAEvent "Order", GetUserID, oOrder.OPKey, , , strErrorMessage
    
    Authorize = False
    
End Function


Public Function Save() As Long
    Dim oCmd As ADODB.Command
    
    Set oCmd = New ADODB.Command
    
    With oCmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "cpopCCTranInsert"
        .Parameters("@_iTranType").value = TranType
        .Parameters("@_iOPKey").value = OPKey
        .Parameters("@_iCCKey").value = CCKey
        .Parameters("@_iUserID").value = UserID
        .Parameters("@_iAmount").value = Amount
        .Parameters("@_iCreateDate").value = TimeStamp
        .Parameters("@_iPNRef").value = PNREF
        .Parameters("@_iResponse").value = Response
        .Parameters("@_iInvcKey").value = InvcKey
        .Parameters("@_iComment1").value = Comment1
        .Parameters("@_iComment2").value = Comment2
        .Parameters("@_iResult").value = Result
        .Parameters("@_iRespMsg").value = RESPMSG
        .Parameters("@_iAuthCode").value = AUTHCODE
'        .Parameters("@_iParamList").Value = ParamList
        .Parameters("@_iTranNo").value = tranNo
'        .Parameters("@_iOrigTranID").Value =
'        .Parameters("@_iOrigAmt").Value =
        .Execute
    End With
    
    Set oCmd = Nothing
End Function



'*************************************************************************
' Description:
'   Does Final Charge Transaction on previously authorized CC using Web Services.
'   This is used in place of the Charge method above.
'
' Return Value
'   0 for Success
'
' Parameters
'   AuthNumber - Output Parameter of the Auth Number
'   SOCreditCardKey - value needed to send email with right socreditcardkey
'   HasBeenCharged - boolean value that determines if authorization has already been captured previously
'
' Called By
'   FBilling.cmdChargeCC_Click
'
'************************************************************************


' Modidifed 2/3/2009 LR to use the new PayPal WebService

Public Function Charge(ByRef oOrder As Order, ByVal InvoiceNumber As String) As Boolean
    Dim oProxy As MSSOAPLib30.SoapClient30
    'Dim s As String
    'Dim sResponse As String
    'Dim strErrorMessage As String
    'Dim strWebServer As String
    'Dim code As String
    'Dim lngResult As Long
    
    'Initialize the transaction object (through property procedures)
    
    TranType = IIf(oOrder.CreditCard.HasBeenCharged, "S", "D")
    TimeStamp = Now
    OPKey = oOrder.OPKey
    tranNo = ""
    Comment1 = Trim$(oOrder.Customer.ID)
    Comment2 = InvoiceNumber
    UserID = GetUserName
    CCKey = oOrder.CreditCard.Key
    
    'Build the ParamList
    
'    ParamList = "USER=" & g_VerisignUserName
'    ParamList = ParamList & "&VENDOR=caseparts&PARTNER=Verisign"
'    ParamList = ParamList & "&PWD=" + g_VerisignPassword
'    ParamList = ParamList & "&TRXTYPE=" & TranType & "&TENDER=C"
    ParamList = "&AMT=" & FormatAsMoney(Amount)
    ParamList = ParamList & "&ORIGID=" & oOrder.CreditCard.AuthorizationPNREF
    ParamList = ParamList & FormatParam("COMMENT1", Comment1)
    ParamList = ParamList & FormatParam("COMMENT2", Comment2)

    ' create the SOAP Proxy
        
    'strWebServer = modRegistry.GetRegStringValue(modRegistry.HKEY_LOCAL_MACHINE, "Software\Case Parts Company", "WebServer", "cpWebPro")
    'If strWebServer = vbNullString Then
    '    Charge = -10000001
    '    Exit Function
    'End If

    On Error GoTo SOAPError
    
    Set oProxy = New MSSOAPLib30.SoapClient30
    'oProxy.MSSoapInit "http://cpwebpro/PayPal/CreditCard.asmx?WSDL"
    oProxy.MSSoapInit g_PayPalUri & "?WSDL"
        
    If TranType = "D" Then
        Response = oProxy.DelayedCapture(ParamList)
    Else
        Response = oProxy.Sale(ParamList)
    End If
    
    ' Process the response
    
'    If InStr(1, sResponse, "Error:") Then
'        s = Mid$(sResponse, 8)
'        If InStr(1, s, "Create") Then
'            code = "-1001"
'        ElseIf InStr(1, s, "Submit") Then
'            code = "-1002"
'        Else
'            code = "-1003"
'            s = "unknown"
'        End If
'        Response = "RESULT=" & code & "&RESPMSG=webservice: " & s
'    Else
'        Response = sResponse
'    End If

    If Result <> 0 Then
        Charge = False
    Else
        Charge = True
    End If
    Exit Function

SOAPError:
    MsgBox oProxy.FaultString + vbCrLf + oProxy.Detail, vbOKOnly, "SOAP Error"
    Charge = False
    
End Function


Private Function GetValue(string1 As String, string2 As String) As String
    Dim Found As Integer
    Dim Start As Integer
    Dim Delimiter As Integer
    Found = InStr(1, string1, string2)
    If Found Then
        Start = Found + Len(string2)
        Delimiter = InStr(Start, string1, "&")
        If Delimiter Then
            GetValue = Mid$(string1, Start, Delimiter - Start)
        Else
            GetValue = Mid$(string1, Start)
        End If
    Else
        GetValue = vbNullString
    End If
End Function


Private Function FormatAsMoney(ByVal cMoneyAmount As Double) As String
    Dim dblHolder As Double
    Dim strHolder As String
    
    On Error GoTo EH
    dblHolder = Round(cMoneyAmount, 2)
    strHolder = CStr(dblHolder)
    If InStr(1, strHolder, ".") = 0 Then
        strHolder = strHolder + ".00"
    ElseIf (Len(strHolder) - InStr(1, strHolder, ".")) = 1 Then
        strHolder = strHolder + "0"
    Else
        strHolder = Mid(strHolder, 1, InStr(1, strHolder, ".") + 2)
    End If
    FormatAsMoney = strHolder
    Exit Function
    
EH:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


'pg18 Payflow Pro Developer's Guide Using Special Characters in Values
'Because the ampersand (&) and the equal sign (=) characters have special meanings
'in the ParamList, name/value pairs like NAME=Ruff & Johnson are not valid.
'To use special charaters in the value of a name/value pair, use a length tag.

Private Function FormatParam(Name As String, value As String) As String
    If InStr(1, value, "=") Or InStr(1, value, "&") Then
        FormatParam = "&" & Name & "[" & Len(value) & "]=" & value
    Else
        FormatParam = "&" & Name & "=" & value
    End If
End Function


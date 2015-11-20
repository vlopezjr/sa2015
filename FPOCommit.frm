VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FPOCommit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Commit PO to Sage ERP"
   ClientHeight    =   2265
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   4200
      TabIndex        =   8
      Top             =   1680
      Width           =   1032
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   1032
   End
   Begin VB.ComboBox cboShipVia 
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "Commit"
      Default         =   -1  'True
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1032
   End
   Begin VB.TextBox txtComment 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker dtRequest 
      Height          =   312
      Left            =   1500
      TabIndex        =   3
      Top             =   600
      Width           =   1812
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   103940097
      CurrentDate     =   37718
   End
   Begin VB.Label lblPONum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   9
      Top             =   1620
      Width           =   3555
   End
   Begin VB.Label Label17 
      Caption         =   "Requested Date"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   1272
   End
   Begin VB.Label Label2 
      Caption         =   "Comment"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Ship Via"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   1272
   End
End
Attribute VB_Name = "FPOCommit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' Commit PO module-level constants & variables
'*****************************************************************

Const LINE_STATUS_OPEN = 1

Private m_rstLines As ADODB.Recordset
Private m_lVendKey As Long
Private m_sUserID As String
Private m_sWhseID As String

Private m_lPOKey As Long
Private m_sTranID As String
Private m_lLineKey As Long
Private m_lLineDistKey As Long
Private m_lMatchToleranceKey As Variant

Private m_bCancel As Boolean


Public Property Get Cancel() As Boolean
    Cancel = m_bCancel
End Property


Private Sub Form_Unload(Cancel As Integer)
    Set m_rstLines = Nothing
End Sub


'ofrm.Init(m_rstLines, m_lVendKey, m_sUserID, m_lWhseKey)

Public Sub Init(rstLines As ADODB.Recordset, VendKey As Long, userid As String, WhseKey As Long)
    Dim orst As ADODB.Recordset
    
    'cache these
    m_sWhseID = WhseKeyToID(WhseKey)
    m_sUserID = userid
    Set m_rstLines = rstLines
    m_lVendKey = VendKey
    
    g_rstShipVia.Filter = "ShipMethID LIKE '" & m_sWhseID & "%'"
    LoadCombo cboShipVia, g_rstShipVia, "ShipMethID", "ShipMethKey"
    cboShipVia.AddItem "-Select One-", 0
    cboShipVia.ListIndex = 0
    g_rstShipVia.Filter = adFilterNone

    Set orst = LoadDiscRst("SELECT MatchToleranceKey FROM tapVendor WHERE VendKey=" & VendKey)
    m_lMatchToleranceKey = orst.Fields("MatchToleranceKey").Value   'cache this for PO commit

    dtRequest.Value = DateAdd("d", 1, Now)
    
    cmdCommit.Visible = True
    cmdCancel.Visible = True
    cmdOK.Visible = False
    lblPONum.Visible = False
    lblPONum.Caption = vbNullString
    
    Me.Show vbModal
End Sub


Private Sub cmdCancel_Click()
    m_bCancel = True
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    m_bCancel = False
    Me.Hide
End Sub


Private Sub cmdCommit_Click()
    Dim CtrlIdx As Integer
    Dim lSpid As Long
    Dim sContext As String
    
    On Error GoTo ErrorHandler
    
    SetWaitCursor True

    DoEvents  'give the cursor a chance to display
    
    'verify ShipTo selection
    sContext = "Verifying Shipping Method"
    If cboShipVia.ListIndex = 0 Then
        msg "You must select a shipping method first.", vbExclamation, "Commit"
        SetWaitCursor False
        Exit Sub
    End If

    On Error GoTo RollbackHandler
    
   
    sContext = "Beginning Transaction"
    g_DB.Connection.BeginTrans
    
    sContext = "Creating temp files"
    CreateTempFiles
    
    sContext = "Getting PO API options"
    GetPOAPIOptions lSpid
    
    sContext = "Getting PO defaults"
    GetPurchOrdDflts lSpid
             
    sContext = "Looping through lines"
    With m_rstLines
        If Not .BOF Then .MoveFirst
        Do While Not .EOF
            sContext = "Accessing line"
            If .Fields("QtyToOrder") > 0 Then
                sContext = "Getting PO line"
                GetPOItem m_rstLines, lSpid
                
                sContext = "Getting PO line distributions"
                GetPOLineDist m_rstLines, lSpid
                
                sContext = "Getting PO line amounts"
                POLineAmts lSpid
            End If
            sContext = "Looping to next line"
           
            .MoveNext
        Loop
    End With
        
    'VL 11/19/2015 fix provided by Jimmy Thomas from BlytheCo
    UpdateShipToAddress
    
    sContext = "Creating Purchase Order"
    CreatePurchOrder lSpid
            
    sContext = "Dropping temp tables"
    DropTempFiles
    
    sContext = "Committing transaction"
    g_DB.Connection.CommitTrans

'added 11/17/04 LR
    On Error GoTo ErrorHandler

    UpdateRequestDate m_lPOKey
    
    LogOAEvent "Create PO", GetUserID, m_lPOKey, , , "Create PO " & m_lPOKey & ". The transaction ID is " & m_sTranID
    
    SetWaitCursor False
    'Msg m_sTranID, , "PO Created in Sage"
    
    cmdCommit.Visible = False
    cmdCancel.Visible = False
    lblPONum.Visible = True
    cmdOK.Visible = True
    cmdOK.SetFocus
    lblPONum.Caption = m_sTranID
    
    Exit Sub
    
RollbackHandler:
    On Error Resume Next
    m_lPOKey = 0
    m_sTranID = ""
    g_DB.Connection.RollbackTrans
    'drop through to continue error handling
    
ErrorHandler:
    Dim sErrMsg As String
    Dim sTitle As String
    
    ClearWaitCursor
    sErrMsg = SageError.ExtractSageErrorInfo(lSpid)
    If Len(sErrMsg) > 0 Then
        sTitle = "Unexpected Sage Error In PO Commit while " & sContext
        ErrorUI.DisplayMsgBox sErrMsg, vbOKOnly, elError, sTitle
    Else
        sTitle = "Unexpected Error In PO Commit while " & sContext
'TODO: study this function        ErrorUI.DisplayError sTitle
'added 11/17/04 LR
        msg sTitle & vbCrLf & Err.Number & " " & Err.Description
    End If
    
    m_bCancel = True
    Me.Hide

End Sub


Private Sub UpdateRequestDate(ByVal i_lPOKey As Long)
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdText
        .CommandText = "Update tpoPurchOrder SET DfltRequestDate = '" & dtRequest.Value & "' where POKey = " & i_lPOKey
        .Execute
    End With
    
    Set cmd = Nothing
End Sub


Private Sub CreateTempFiles()
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP(ReadFile(g_CreatePOTempTables), adCmdText)
    cmd.Execute
    Set cmd = Nothing
End Sub


Private Sub DropTempFiles()
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP(ReadFile(g_DropPOTempTables), adCmdText)
    cmd.Execute
    Set cmd = Nothing
End Sub


Private Sub GetPOAPIOptions(ByRef o_lSpid As Long)
    Dim iRetVal As Integer
    Dim cmd As ADODB.Command
    
    Set cmd = CreateCommandSP("sppoGetPOAPIOptions")
    With cmd
        .Parameters("@_iCompanyID").Value = "CPC"
        .Parameters("@_iLogSuccessful").Value = Null
        .Execute
        iRetVal = .Parameters("@_oRetVal").Value
        o_lSpid = .Parameters("@_oSpid").Value
    End With
    Set cmd = Nothing
    If iRetVal > 2 Then
        DisplayPOAPIError "GetPOAPIOptions", o_lSpid
        Err.Raise -1, "GetPurchOAPIOptions", "Unexpected error in GetPurchOAPIOptions"
    End If
End Sub


Private Sub GetPurchOrdDflts(i_lSpid As Long)
    Dim iRetVal As Integer
    Dim cmd As ADODB.Command
    Dim userid As String
    
    On Error GoTo EH
    
    'wrapper SP
    Set cmd = CreateCommandSP("spcpGetPurchOrdDflts")
    With cmd
        .Parameters("@_iVendKey").Value = m_lVendKey

        If g_DB.IsDevelopment Then
            userid = InputBox("Enter BuyerID:", "Create PO (Development Only)")
            DoEvents
        Else
            userid = m_sUserID
        End If
        
        .Parameters("@_iUserID").Value = userid
        '.Parameters("@_iDfltShipToAddrKey").value = User.GetUserWhseShipAddrKey(userid)

        .Parameters("@_iStatus").Value = 0      'unissued
        .Parameters("@_iTranCmnt").Value = Left(txtComment.text, 50)
        .Parameters("@_iMatchToleranceKey").Value = m_lMatchToleranceKey
        .Parameters("@_iDSTWhseID").Value = m_sWhseID
        .Parameters("@_iDfltShipMethKey").Value = cboShipVia.ItemData(cboShipVia.ListIndex)
        
        'These support the service that fixes Unassigned Buyer
        .Parameters("@_iUserFld1").Value = CStr(UserNameToBuyerKey(m_sUserID))
        .Parameters("@_iUserFld2").Value = CStr(cboShipVia.ItemData(cboShipVia.ListIndex))

        .Execute
        
        m_lPOKey = .Parameters("@POKey").Value
        m_sTranID = .Parameters("@TranID").Value
        iRetVal = .Parameters("@RetVal").Value
    End With
    
    If iRetVal > 2 Then
        GoTo EH
    End If
    
    Set cmd = Nothing
    Exit Sub
EH:
    If g_DB.Connection.Errors.Count > 0 Then
        Dim serr As String
        Dim i As Integer
        For i = 0 To g_DB.Connection.Errors.Count - 1
            serr = g_DB.Connection.Errors(i).Description
            MsgBox serr, , "ADO/SQL Error " & i + 1
        Next i
    Else
        DisplayPOAPIError "GetPurchOrdDflts", i_lSpid
    End If
    Err.Raise -1, "GetPurchOrdDflts", "Unexpected error in GetPurchOrdDflts"
    Set cmd = Nothing
End Sub


Private Sub GetPOItem(i_oRst As ADODB.Recordset, i_lSpid As Long)
    Dim iRetVal As Integer
    Dim ocmd As ADODB.Command
    Dim sComment As String
    Dim sMsg As String
    Dim orst As ADODB.Recordset
    Dim lOPKey As Long
    Dim lSOKey As Long

    'if this item is a SPO, build an ExtComment for the Acuity PO
    If i_oRst.Fields("isSPO") Then
        If i_oRst.Fields("SOLineKey") > 0 Then
            Set ocmd = CreateCommandSP("spcpcPOGetSPODetail")
            ocmd.Parameters("@_iSOLineKey").Value = i_oRst.Fields("SOLineKey").Value
            
            Set orst = New ADODB.Recordset
            Set orst = ocmd.Execute
            If Not orst.EOF Then
                'cache these for logging
                lOPKey = orst.Fields("OPKey").Value
                lSOKey = Trim(orst.Fields("TranKey").Value)
                
                If Len(orst.Fields("ModelNbr").Value) > 0 Then
                    sComment = sComment & "Model # " & Trim(orst.Fields("ModelNbr").Value)
                End If
                If Len(orst.Fields("SerialNbr").Value) > 0 Then
                    sComment = sComment & "; Serial # " & Trim(orst.Fields("SerialNbr").Value)
                End If
                sComment = sComment & "; OP " & lOPKey
                sComment = sComment & "; SO " & lSOKey
            End If
            Set ocmd = Nothing
            Set orst = Nothing
        End If
    End If
    
    'wrapper SP
    Set ocmd = CreateCommandSP("spcpGetPOitem")
    With ocmd
        .Parameters("@_iPOKey").Value = m_lPOKey
        .Parameters("@_iClosedForRcvg").Value = 0
        .Parameters("@_iClosedForInvc").Value = 0
        .Parameters("@_iDescription").Value = i_oRst.Fields("Descr").Value
        .Parameters("@_iStatus").Value = 1      'Open
        .Parameters("@_iItemKey").Value = i_oRst.Fields("ItemKey").Value
        .Parameters("@_iUnitCost").Value = i_oRst.Fields("UnitCost").Value
        If Len(sComment) > 0 Then
            .Parameters("@_iExtCmnt").Value = Left(sComment, 254)
        End If
        .Execute
        m_lLineKey = .Parameters("@POLineKey").Value
        iRetVal = .Parameters("@RetVal").Value
    End With
    
    If iRetVal > 2 Then
        DisplayPOAPIError "GetPOItem", i_lSpid
        Set ocmd = Nothing
        Err.Raise -1, "GetPOItem", "Unexpected error in GetPOItem"
    End If

    '*** Auto-Freeze logic ***
    'if this item is a SPO, freeze it
    If i_oRst.Fields("isSPO") Then
        Debug.Print "SOLKey = " & i_oRst.Fields("SOLineKey") & ", POLKey " & m_lLineKey
        Set ocmd = CreateCommandSP("spCPCInsertSPLPOFreeze")
        ocmd.Parameters("@_iSOLineKey").Value = i_oRst.Fields("SOLineKey").Value
        ocmd.Parameters("@_iPOLineKey").Value = m_lLineKey
        ocmd.Parameters("@_iPOKey").Value = m_lPOKey
        ocmd.Parameters("@_iPOTranNo").Value = Right$(m_sTranID, 10)
        ocmd.Execute
        sMsg = "Freeze item - " & Trim(i_oRst.Fields("Descr")) & " on PO line " & m_lLineKey & ". The vendor is " & Trim(i_oRst.Fields("VendName").Value)
        LogDB.LogOAEvent "Auto Freeze", GetUserID, m_lLineKey, i_oRst.Fields("SOLineKey").Value, , sMsg
        LogDB.LogActivity "SA", sMsg, lOPKey, lSOKey, , i_oRst.Fields("SOLineKey").Value, m_lPOKey, Right$(m_sTranID, 10), m_lLineKey
    End If
End Sub


'Uses module level variables:
'm_lPOKey
'm_lLineKey
'm_lLineDistKey
    
Private Sub GetPOLineDist(i_oRst As ADODB.Recordset, i_lSpid As Long)
    Dim iRetVal As Integer
    Dim ocmd As ADODB.Command
    Dim CtrlIdx As Integer
    
    Set ocmd = CreateCommandSP("spcpGetPOLineDist")
    
    With ocmd
        .Parameters("@_iPOKey").Value = m_lPOKey
        .Parameters("@_iPOLineKey").Value = m_lLineKey
        .Parameters("@_iStatus").Value = 1
        .Parameters("@_iQtyOrd") = i_oRst.Fields("QtyToOrder").Value
        .Parameters("@_iGLAcctKey").Value = Null
        .Execute
        m_lLineDistKey = .Parameters("@POLineDistKey").Value
        iRetVal = .Parameters("@RetVal").Value
    End With
    Set ocmd = Nothing
    If iRetVal > 2 Then
        DisplayPOAPIError "GetPOLineDist", i_lSpid
        Err.Raise -1, "GetPOLineDist", "Unexpected error in GetPOLineDist"
    End If
End Sub


'Uses module level variables:
'm_lPOKey
'm_lLineKey

Private Sub POLineAmts(i_lSpid As Long)
    Dim iRetVal As Integer
    Dim ocmd As ADODB.Command
    
    Set ocmd = CreateCommandSP("sppoLineAmts")
    
    With ocmd
        .Parameters("@_iPOKey").Value = m_lPOKey
        .Parameters("@_iPOLineKey").Value = m_lLineKey
        .Execute
        iRetVal = .Parameters("@_oRetVal").Value
    End With
    Set ocmd = Nothing
    If iRetVal > 2 Then
        DisplayPOAPIError "POLineAmts", i_lSpid
        Err.Raise -1, "POLineAmts", "Unexpected error in POLineAmts"
    End If
End Sub

'Uses module level variables:
'm_lPOKey

Private Sub CreatePurchOrder(ByRef o_lSpid As Long)
    Dim iRetVal As Integer
    Dim ocmd As ADODB.Command
    
    Set ocmd = CreateCommandSP("sppoCreatePurchOrder")
    With ocmd
        .Parameters("@_iPOKey").Value = m_lPOKey
        .Parameters("@_iPurchAmt").Value = Null
        .Parameters("@_iFreightAmt").Value = Null
        .Parameters("@_iSTaxAmt").Value = Null
        .Parameters("@_iTranAmt").Value = Null
        .Parameters("@_iOpenAmt").Value = Null
        .Parameters("@_iAmtInvcd").Value = Null
        .Parameters("@_iUseTemp").Value = 0
        .Execute
        iRetVal = .Parameters("@_oRetVal").Value
    End With
    Set ocmd = Nothing
    If iRetVal > 2 Then
        DisplayPOAPIError "CreatePurchOrder", o_lSpid
        Err.Raise -1, "CreatePurchOrder", "Unexpected error in CreatePurchOrder"
    End If
End Sub


Private Sub DisplayPOAPIError(sSPName As String, i_lSpid As Long)
    Dim sSQL As String
    Dim rst As ADODB.Recordset
'SQL:
    sSQL = "SELECT * FROM tciErrorLog WHERE SessionID = " & i_lSpid
    
    Set rst = New ADODB.Recordset
    Set rst = LoadDiscRst(sSQL)

    If Not rst.EOF Then
        'Handle this special case specifically
        If rst.Fields("StringNo") = "220325" Then
            msg Trim(rst.Fields("StringData1")) & " is marked as Discontinued. Remove it from the PO and try again."
        Else
            msg "SPROC - " & sSPName & vbCrLf & _
                "ErrorCmnt - " & rst.Fields("ErrorCmnt").Value & vbCrLf & _
                "StringNo - " & rst.Fields("StringNo") & vbCrLf & _
                "StringData1 - " & rst.Fields("StringData1") & vbCrLf & _
                "StringDate2 - " & rst.Fields("StringData2") & vbCrLf & _
                "StringData3 - " & rst.Fields("StringData3") & vbCrLf & _
                "StringDate5 - " & rst.Fields("StringData4"), , "Sage PO API Error"
        End If
    Else
        msg "There is no error in tciErrorLog", , "Sage PO API Error"
    End If
    
    Set rst = Nothing
End Sub

Private Sub UpdateShipToAddress()
    Dim sSQL As String

    sSQL = "UPDATE #tpoAPIValid SET DfltShipToAddrKey = (SELECT ShipAddrKey FROM timWarehouse (NOLOCK) WHERE WhseKey = (SELECT DfltShipToWhseKey FROM #tpoAPIValid)) " & _
            "UPDATE #tpoPOLineDist SET ShipToAddrKey = (SELECT DfltShipToAddrKey FROM #tpoAPIValid) " & _
            "UPDATE #tpoPOLineDist SET ShipToWhseKey = (SELECT DfltShipToWhseKey FROM #tpoAPIValid) "
      
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP(sSQL, adCmdText)
    cmd.Execute
    Set cmd = Nothing
    
End Sub

'Private Function GetWhseShipAddrKey(whseid As String) As Long
'    Dim orst As ADODB.Recordset
''SQL:
'
'    Set orst = New ADODB.Recordset
'    Set orst = LoadDiscRst("SELECT ShipAddrKey FROM timWarehouse WHERE WhseID='" & whseid & "'")
'
'    With orst
'        If Not .EOF Then
'            GetWhseShipAddrKey = .Fields("ShipAddrKey").value
'        Else
'            GetWhseShipAddrKey = -1  'error - not found
'        End If
'    End With
'    Set orst = Nothing
'End Function


'*******    End Commit the PO to Sage     *******


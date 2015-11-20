VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lWhseID As Long
Private m_sBatchDescr As String
Private m_collAdjments As Collection
Private m_lBatchKey As Long
Private m_sBatchID As String

Public Property Let whseid(value As Long)
    m_lWhseID = value
End Property

Public Property Let BatchDescr(value As String)
    m_sBatchDescr = value
End Property

Public Property Get Count() As Integer
    Count = m_collAdjments.Count
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Procedures
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Add(pItemID As String, pQuantity As Integer, pTranDescr As String) As Boolean
    Dim lItemKey As Long
    Dim cmd As ADODB.Command
    Dim lors As ADODB.Recordset
    Dim lsSql As String
    Dim loAdj As Adjustment

    Add = False
    
    'GetItemType function gets the part#'s item type (5 finished good or 7 kit)
    'Input: ItemID (part#) - Output: Item Key (this value is used for a kit)
    
    Select Case GetItemType(pItemID, lItemKey)
        Case 5 'Valid Finished Good
            Set loAdj = New Adjustment
            loAdj.ItemID = pItemID
            loAdj.Quantity = pQuantity
            loAdj.TranDescr = pTranDescr
            m_collAdjments.Add loAdj
            Add = True
            
        Case 7 'Valid Kit - get & process each finished good in a kit
            lsSql = "Select itemID, CompItemQty from timkitcomplist, timitem Where kititemkey = " & lItemKey & _
                " and timkitcomplist.compitemkey = timitem.itemkey"
            Set lors = New ADODB.Recordset
            lors.Source = lsSql
            Set lors.ActiveConnection = g_DB.Connection
            lors.Open
            Do Until lors.EOF = True
                Set loAdj = New Adjustment
                loAdj.ItemID = Trim(lors!ItemID)
                loAdj.Quantity = pQuantity * lors!CompItemQty '# kits * # item in each kit
                loAdj.TranDescr = pTranDescr
                m_collAdjments.Add loAdj
                Add = True
                lors.MoveNext
            Loop
    End Select
End Function


Public Sub ProcessBatch()
    
    Dim lSpid As Long
    Dim sContext As String
    Dim sSageErrMsg As String
    Dim sOAErrMsg As String
    Dim sGeneralErr As String
    Dim sErrMsg As String
    Dim sDesc As String
    Dim anAdj As Adjustment
    Dim lTrankey As Long
    
On Error GoTo ErrorHandler
        
    If m_collAdjments.Count > 0 Then
        On Error GoTo ErrorHandler
            
        sContext = "Creating temp files"
        Call CreateTempTables
        
        sContext = "Setup batch"
        Call SetUpBatch(lSpid)
        
        'process each adjustment
        For Each anAdj In m_collAdjments
            lTrankey = 0
            
            'GetTranDetails passes in Item and Quantity and returns TranKey
            sContext = "Get tran details"
            Call GetTranDetails(anAdj.ItemID, anAdj.Quantity, anAdj.TranDescr, lSpid, lTrankey)
            
            'GetDistDetails passes in TranKey and Quantity
            sContext = "Get dist details"
            Call GetDistDetails(lTrankey, anAdj.Quantity, lSpid)
            
            'Create Pending Tran passes in TranKey
            sContext = "Create pending tran"
            Call CreatePendingTran(lTrankey, lSpid)
        Next
    
        sContext = "Batch cleanup"
        Call BatchCleanUp(lSpid)
        
        sContext = "Drop temp tables"
        Call DropTempFiles
                
        MsgBox "Batch " & Mid$(m_sBatchID, InStr(1, m_sBatchID, "-") + 1, Len(m_sBatchID)) & " is ready to post.", vbInformation, ""
    End If
Exit Sub

ErrorHandler:
    '************ADO Query Time-Out********************
    If Err.Number = -2147217871 Then
        sDesc = Err.Description & " (ADO query timed-out)"
    Else
        sDesc = Err.Description
    End If
    
    sGeneralErr = "* " & sContext & " encountered error " & vbCrLf & _
                    vbTab & "Err Num: " & vbTab & Err.Number & vbCrLf & _
                    vbTab & "Err Source: " & vbTab & Err.Source & vbCrLf & _
                    vbTab & "Err Desc: " & vbTab & Trim$(sDesc) & vbCrLf
                    
    '*****
    sSageErrMsg = SageError.ExtractSageErrorInfo(lSpid)
    
    On Error Resume Next
    
    ' Log error to file
    If sSageErrMsg = "" Then sSageErrMsg = "* No Error Msg is available in Sage" & vbCrLf & vbCrLf
    sOAErrMsg = "* IM information in SageAssistant" & vbCrLf & _
            vbTab & "Batch:" & vbTab & m_lBatchKey & vbCrLf

    sErrMsg = sGeneralErr & _
                vbCrLf & _
                sSageErrMsg & _
                sOAErrMsg & _
                "* Source: " & SetErrSource("Batch", "Create", False) & vbCrLf & _
                vbCrLf & _
                "* Current database is " & g_DB.SERVER & "." & g_DB.DATABASE & vbCrLf & _
                vbCrLf & _
                "* Timestamp: " & Now & " " & Timer

    LogEventExt "InvAdjustments", "ProcessBatch", "", sErrMsg
    
    ' Log error to OA Event log - spCPCeventInsert
    LogOAEvent "InvAdj", GetUserID, m_lBatchKey, , sContext, "Batch was not created."
    MsgBox "Batch was not created." & chr(13) & chr(13) & "Error: " & Err.Number & chr(13) & Err.Description, vbInformation, "Inventory Adjustment"
    
    Err.Raise -1, "ProcessBatch", "Batch was not created"
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Procedures/Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    Set m_collAdjments = New Collection
End Sub


Private Sub SetUpBatch(ByRef o_lSpid As Long)

    On Error GoTo ErrorHandler

    Dim iRetVal As Integer
    Dim bSPError As Boolean
    Dim cmd As ADODB.Command
    'Dim sBatchID As String
    Dim lSpid As Long
    
    Set cmd = CreateCommandSP("spimAPISetupBatch")
    
    With cmd
        .Parameters("@_iBatchCmnt").value = m_sBatchDescr
        .Parameters("@_iBatchType").value = 701
        .Parameters("@_iUserID").value = GetUserName
        .Parameters("@_iPostDate").value = Now
        .Parameters("@_iSourceCompanyID").value = "CPC"
        .Parameters("@_iWhseKey").value = m_lWhseID
        .Execute
        iRetVal = .Parameters("@_oRetVal").value
        m_lBatchKey = .Parameters("@_oBatchKey").value
        m_sBatchID = .Parameters("@_oBatchID").value
        o_lSpid = .Parameters("@_oSpid").value
    End With
    
    If iRetVal <> SageErrType.aetNone Then
        'SageError.PopulateErrorComment o_lSpid
        If iRetVal <> SageErrType.aetWarning Then
             bSPError = True
             GoTo ErrorHandler
        End If
    End If
    
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    Set cmd = Nothing
    If bSPError Then
        Err.Raise -1, "SetUpBatch", "Unexpected error in spimAPISetupBatch"
    Else
        ThrowError "InvAdj", "SetUpBatch (Line " & Erl & ")", ""
    End If
End Sub


Private Sub GetTranDetails(pItemID As String, pQuantity As Integer, pTranDescr As String, ByRef o_lSpid As Long, pTranKey As Long)
    On Error GoTo ErrorHandler

    Dim iRetVal As Integer
    Dim bSPError As Boolean
    Dim cmd As ADODB.Command
    Dim lSpid As Long
    Dim lTranID As String
    
    Set cmd = CreateCommandSP("spimAPIGetTranDetails")
    
    With cmd
        .Parameters("@_iBatchKey").value = m_lBatchKey
        .Parameters("@_iItemID").value = pItemID
        .Parameters("@_iReasonCodeKey").value = 20 'Inventory Adj. Reason Code
        .Parameters("@_iTranCmnt").value = pTranDescr
        .Parameters("@_iTranQty").value = pQuantity
        .Parameters("@_iTranType").value = 710  'Adjustment
        'When DefaultIfNull is set to 1, certain fields can be omitted or passed as null
        .Parameters("@_iDefaultIfNull").value = 1
        .Execute
        iRetVal = .Parameters("@_oRetVal").value
        pTranKey = .Parameters("@_oInvtTranKey").value
        lTranID = .Parameters("@_oTranID").value
    End With
    
    'if return code is not equal to none
    If iRetVal <> SageErrType.aetNone Then
        'populates sparPopulateErrorComment
        ''SageError.PopulateErrorComment o_lSpid
        'test if the return value is not an Sage warning
        If iRetVal <> SageErrType.aetWarning Then
             bSPError = True
             GoTo ErrorHandler
        End If
    End If
    
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    Set cmd = Nothing
    If bSPError Then
        Err.Raise -1, "GetTranDetails", "Unexpected error in spimAPIGetTranDetails"
    Else
        ThrowError "InvAdj", "SetUpBatch (Line " & Erl & ")", ""
    End If
End Sub


Private Sub GetDistDetails(pTranKey As Long, pQuantity As Integer, ByRef o_lSpid As Long)
    On Error GoTo ErrorHandler

    Dim iRetVal As Integer
    Dim bSPError As Boolean
    Dim cmd As ADODB.Command
    Dim lSpid As Long
    
    Set cmd = CreateCommandSP("spimAPIGetDistDetails")
    
    With cmd
        .Parameters("@_iInvtTranKey").value = pTranKey
        .Parameters("@_iWhseBinID").value = Null
        .Parameters("@_iWhseBinKey").value = Null
        .Parameters("@_iLotNo").value = Null
        .Parameters("@_iInvtLotKey").value = Null
        .Parameters("@_iLotExpirationDate").value = Null
        .Parameters("@_iSerialNo").value = Null
        .Parameters("@_iInvtSerialKey").value = Null
        .Parameters("@_iDistQty").value = pQuantity
        'When DefaultIfNull is set to 1, certain fields can be omitted or passed as null
        .Parameters("@_iDefaultIfNull").value = 1
        .Execute
        iRetVal = .Parameters("@_oRetVal").value
    End With
    
    If iRetVal <> SageErrType.aetNone Then
        ''SageError.PopulateErrorComment o_lSpid
        If iRetVal <> SageErrType.aetWarning Then
             bSPError = True
             GoTo ErrorHandler
        End If
    End If
    
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    Set cmd = Nothing
    If bSPError Then
        Err.Raise -1, "GetDistDetails", "Unexpected error in spimAPIGetDistDetails"
    Else
        ThrowError "InvAdj", "SetUpBatch (Line " & Erl & ")", ""
    End If
End Sub


Private Sub CreatePendingTran(pTranKey As Long, ByRef o_lSpid As Long)
    On Error GoTo ErrorHandler

    Dim iRetVal As Integer
    Dim bSPError As Boolean
    Dim cmd As ADODB.Command
    Dim lSpid As Long
    
    Set cmd = CreateCommandSP("spimAPICreatePendingTran")
    
    With cmd
        .Parameters("@_iInvtTranKey").value = pTranKey
        .Execute
        iRetVal = .Parameters("@_oRetVal").value
    End With
    
    If iRetVal <> SageErrType.aetNone Then
       '' SageError.PopulateErrorComment o_lSpid
        If iRetVal <> SageErrType.aetWarning Then
             bSPError = True
             GoTo ErrorHandler
        End If
    End If
    
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    Set cmd = Nothing
    If bSPError Then
        Err.Raise -1, "CreatePendingTran", "Unexpected error in spimAPICreatePendingTran"
    Else
        ThrowError "InvAdj", "SetUpBatch (Line " & Erl & ")", ""
    End If
End Sub


Private Sub BatchCleanUp(ByRef o_lSpid As Long)
    On Error GoTo ErrorHandler

    Dim iRetVal As Integer
    Dim bSPError As Boolean
    Dim cmd As ADODB.Command
    Dim lSpid As Long
    
    Set cmd = CreateCommandSP("spimAPIBatchCleanup")
    
    With cmd
        .Parameters("@_iBatchKey").value = m_lBatchKey
        .Execute
        iRetVal = .Parameters("@_oRetVal").value
    End With
    
    If iRetVal <> SageErrType.aetNone Then
        ''SageError.PopulateErrorComment o_lSpid
        If iRetVal <> SageErrType.aetWarning Then
             bSPError = True
             GoTo ErrorHandler
        End If
    End If
    
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    Set cmd = Nothing
    If bSPError Then
        Err.Raise -1, "BatchCleanUp", "Unexpected error in spimAPIBatchCleanup"
    Else
        ThrowError "InvAdj", "SetUpBatch (Line " & Erl & ")", ""
    End If
End Sub


'convert this to command returning scalar

Private Function GetItemType(pItemID As String, pItemKey As Long) As Integer
    Dim lors As ADODB.Recordset
    Dim lsSql As String
    GetItemType = 0
    
    lsSql = "Select itemtype, itemkey from timitem where companyID = 'CPC' and itemid = '" & pItemID & "'"
    
    Set lors = New ADODB.Recordset
    lors.Source = lsSql
    Set lors.ActiveConnection = g_DB.Connection
    lors.Open
    
    If lors.EOF = False Then
        GetItemType = lors!ItemType
        pItemKey = lors!ItemKey
    End If
    lors.Close
End Function


Private Sub CreateTempTables()
    Dim cmd As ADODB.Command
    Set cmd = DATABASE.CreateCommandSP(ReadFile(g_CreateIMTempTables), adCmdText)
    cmd.Execute
    Set cmd = Nothing
End Sub


Private Sub DropTempFiles()
    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP(ReadFile(g_DropIMTempTables), adCmdText)
    cmd.Execute
    Set cmd = Nothing
End Sub



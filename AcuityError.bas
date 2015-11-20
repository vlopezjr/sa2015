Attribute VB_Name = "SageError"
Option Explicit

Public Enum SageErrType
    aetUnknown = 0
    aetNone = 1
    aetWarning = 2
    aetError = 3
    aetFatalErr = 4
End Enum

Public Function ExtractSageErrorInfo(i_lSpid As Long) As String
    Dim sSQL As String
    Dim rst As ADODB.Recordset
    Dim sTemp As String
    
    sSQL = "select errortype, severity, tciErrorLog.StringNo, ConstantName, " & _
        "isnull(StringData1, '') as StringData1, " & _
        "isnull(StringData2, '') as StringData2, " & _
        "isnull(StringData3, '') as StringData3, " & _
        "isnull(StringData4, '') as StringData4, " & _
        "isnull(StringData5, '') as StringData5, " & _
        "rtrim(ltrim(replace(replace(replace(replace(replace(ErrorCmnt, StringData1, ''), StringData2, ''), StringData3, ''), StringData4, ''), StringData5, ''))) as ErrorCmnt " & _
        "from tciErrorLog left join tsmString on tcierrorlog.stringno = tsmstring.stringno where SessionID = " & i_lSpid
    
    Set rst = LoadRst(sSQL)
    
    While Not rst.EOF

        Select Case Trim(rst.Fields("StringNo").value)
            ' Ignore the followings warnings
            ' 220373 = PO Created Successfully
            ' 220619 = Unit Cost Calculated to zero
            ' 250223 = Unit Price was passed in as zero
            ' 250544 = Qty available is less than Qty ordered.
            ' 250639 = Cust. Addr. does not require Acknowledgment  Status set to Open.
            Case "220373", "220619", "250223", "250544", "250639"
            Case Else
                sTemp = "Sage Error Information" & vbCrLf
                
                sTemp = sTemp & "  " & rst.Fields("ConstantName").value & " (StringNo: " & rst.Fields("StringNo").value & ") Severity=" & rst.Fields("severity").value & vbCrLf

                If Trim(rst.Fields("StringData1").value & "") <> "" Then
                    sTemp = sTemp & "  " & "StringData1: " & Trim(rst.Fields("StringData1").value) & vbCrLf
                End If
                If Trim(rst.Fields("StringData2").value & "") <> "" Then
                    sTemp = sTemp & "  " & "StringData2: " & Trim(rst.Fields("StringData2").value) & vbCrLf
                End If
                If Trim(rst.Fields("StringData3").value & "") <> "" Then
                    sTemp = sTemp & "  " & "StringData3: " & Trim(rst.Fields("StringData3").value) & vbCrLf
                End If
                If Trim(rst.Fields("StringData4").value & "") <> "" Then
                    sTemp = sTemp & "  " & "StringData4: " & Trim(rst.Fields("StringData4").value) & vbCrLf
                End If
                If Trim(rst.Fields("StringData5").value & "") <> "" Then
                    sTemp = sTemp & "  " & "StringData5: " & Trim(rst.Fields("StringData5").value) & vbCrLf
                End If
                If Trim(rst.Fields("ErrorCmnt").value & "") <> "" Then
                    sTemp = sTemp & "  " & "ErrorCmnt: " & Trim(rst.Fields("ErrorCmnt").value) & vbCrLf
                End If
                ExtractSageErrorInfo = ExtractSageErrorInfo & sTemp & vbCrLf
        End Select
        rst.MoveNext
    Wend
    
    rst.Close
    Set rst = Nothing
End Function


'!!! Not used

Private Sub PopulateErrorComment(i_lSpid As Long)
    Dim iRetVal As Integer
    Dim bSPError As Boolean
    Dim cmd As ADODB.Command
    
    On Error GoTo ErrorHandler
    
    Set cmd = CreateCommandSP("sparPopulateErrorComment")
    With cmd
        .Parameters("@_iBatchKey").value = Null
        .Parameters("@_iSpid").value = i_lSpid
        .Parameters("@_iLanguageID").value = 1033
        .Execute
        iRetVal = .Parameters("@_oRetVal").value
    End With
    
    If iRetVal > SageErrType.aetWarning Then
        bSPError = True
        GoTo ErrorHandler
    End If
    
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    Set cmd = Nothing
    
    If bSPError Then
        Err.Raise -1, "sparPopulateErrorComment", "Unexpected error in sparPopulateErrorComment"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub





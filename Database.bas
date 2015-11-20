Attribute VB_Name = "Database"
Option Explicit

' This is a collection a ADO helper functions

'Public Function LoadDiscRst()
'Public Function LoadRst()
'Public Sub CloseRst()
'Public Function CreateCommandSP()
'Public Function CallSP()
'Public Sub ExecuteSP()
'Public Function GetSurrogateKey()
'Public Function PrepSQLText()



Public Function LoadDiscRst( _
        ByVal i_sSQL As String, _
        Optional i_cn As ADODB.Connection = Nothing, _
        Optional i_lLockType As ADODB.LockTypeEnum = adLockReadOnly, _
        Optional i_lMaxRows As Long = 0 _
) As ADODB.Recordset

    On Error GoTo ErrorHandler
    
    Set LoadDiscRst = LoadRst(i_sSQL, i_cn, adUseClient, i_lLockType, adOpenStatic, i_lMaxRows)
    Set LoadDiscRst.ActiveConnection = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise -1, "Database.LoadDiscRst", "LoadDiscRst" & vbCrLf & Err.Description
    
End Function


'By default this uses the global connection g_DB.Connection.
'You can override this by passing in a reference to a different connection.

Public Function LoadRst( _
        ByVal i_sSQL As String, _
        Optional i_cn As ADODB.Connection = Nothing, _
        Optional i_lCursorLocation As ADODB.CursorLocationEnum = adUseClient, _
        Optional i_lLockType As ADODB.LockTypeEnum = adLockReadOnly, _
        Optional i_lCursorType As ADODB.CursorTypeEnum = adOpenStatic, _
        Optional i_lMaxRows As Long = 0 _
) As ADODB.Recordset

    Dim Cn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim lErrorCount As Long
    
    If i_lMaxRows > 0 Then
        i_sSQL = "SET ROWCOUNT " & i_lMaxRows & vbCrLf & i_sSQL & " SET ROWCOUNT 0 "
    End If

    If i_cn Is Nothing Then
        Set Cn = g_DB.Connection
    Else
        Set Cn = i_cn
    End If
    
    On Error GoTo ErrorHandler
    
    Cn.CursorLocation = i_lCursorLocation

    Set rst = New ADODB.Recordset
    
RetryPoint:
    
    rst.Open i_sSQL, Cn, i_lCursorType, i_lLockType

    'when would there be errors in the ADO connection when there is no VB runtime error?
    If Cn.Errors.Count > 0 Then GoTo ErrorHandler
    
    Set LoadRst = rst
    
    Exit Function

ErrorHandler:
    Dim sMsg As String
    Dim obj As Object

    'Retry up to 3 times
    If lErrorCount < 3 Then
        lErrorCount = lErrorCount + 1
        Resume RetryPoint
    End If

    If Err.Number <> 0 Then
        sMsg = "LoadRst " & "SQL=" & i_sSQL & vbCrLf & _
                Err.Number & " (" & Err.Source & ") " & Err.Description
    End If

    If Cn.Errors.Count > 0 Then
        Dim i As Long
        sMsg = sMsg & vbCrLf & "Additional Information:" & vbCrLf
        For Each obj In Cn.Errors
            sMsg = sMsg & "ErrMsg " & i & ": " & obj.Description & vbCrLf
        Next
    End If
    
    Err.Raise -1, "Database.LoadRst", sMsg
End Function


Public Sub CloseRst(ByRef rst As ADODB.Recordset)
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            Set rst.ActiveConnection = Nothing
            rst.Close
        End If
    End If
    Set rst = Nothing
End Sub


Public Function CreateCommandSP(i_sSPName As String, Optional i_lCommandType As ADODB.CommandTypeEnum = adCmdStoredProc) As ADODB.Command
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    With cmd
        g_DB.Connection.CursorLocation = adUseClient
        .ActiveConnection = g_DB.Connection
        .CommandType = i_lCommandType
        .CommandText = i_sSPName
    End With
    Set CreateCommandSP = cmd
End Function


Public Function CallSP( _
        ByVal i_sName As String, _
        Optional i_sParm1 As String = "", _
        Optional i_vParm1 As Variant, _
        Optional i_sParm2 As String = "", _
        Optional i_vParm2 As Variant, _
        Optional i_sParm3 As String = "", _
        Optional i_vParm3 As Variant, _
        Optional i_sParm4 As String = "", _
        Optional i_vParm4 As Variant, _
        Optional i_sParm5 As String = "", _
        Optional i_vParm5 As Variant, _
        Optional i_sParm6 As String = "", _
        Optional i_vParm6 As Variant _
) As ADODB.Recordset
    Dim cmd As ADODB.Command

    On Error GoTo ErrorHandler
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = i_sName
        .CommandTimeout = 120        'we've some expensive sprocs
    End With

    'Deal with parameters, if any
    If Len(i_sParm1) > 0 Then cmd.Parameters(i_sParm1).value = i_vParm1
    If Len(i_sParm2) > 0 Then cmd.Parameters(i_sParm2).value = i_vParm2
    If Len(i_sParm3) > 0 Then cmd.Parameters(i_sParm3).value = i_vParm3
    If Len(i_sParm4) > 0 Then cmd.Parameters(i_sParm4).value = i_vParm4
    If Len(i_sParm5) > 0 Then cmd.Parameters(i_sParm5).value = i_vParm5
    If Len(i_sParm6) > 0 Then cmd.Parameters(i_sParm6).value = i_vParm6

    Set CallSP = cmd.Execute
    
    If g_DB.Connection.Errors.Count > 0 Then GoTo ErrorHandler

    Exit Function

ErrorHandler:
    Dim sMsg As String
    Dim obj As Object
    
    If Err.Number <> 0 Then
        sMsg = "CallSP " & i_sName & vbCrLf & _
                Err.Number & " (" & Err.Source & ") " & Err.Description
    End If

    If g_DB.Connection.Errors.Count > 0 Then
        Dim i As Long
        sMsg = sMsg & vbCrLf & "Additional Information:" & vbCrLf
        For Each obj In g_DB.Connection.Errors
            sMsg = sMsg & "ErrMsg " & i & ": " & obj.Description & vbCrLf
        Next
    End If

    Err.Raise Number:=-1, Source:="Database.CallSP", Description:=sMsg

End Function


Public Sub ExecuteSP( _
        ByVal i_sName As String, _
        Optional i_sParm1 As String = "", _
        Optional i_vParm1 As Variant, _
        Optional i_sParm2 As String = "", _
        Optional i_vParm2 As Variant, _
        Optional i_sParm3 As String = "", _
        Optional i_vParm3 As Variant, _
        Optional i_sParm4 As String = "", _
        Optional i_vParm4 As Variant, _
        Optional i_sParm5 As String = "", _
        Optional i_vParm5 As Variant, _
        Optional i_sParm6 As String = "", _
        Optional i_vParm6 As Variant _
)
    Dim cmd As ADODB.Command

    On Error GoTo ErrorHandler
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = i_sName
    End With

    'Deal with parameters, if any
    If Len(i_sParm1) > 0 Then cmd.Parameters(i_sParm1).value = i_vParm1
    If Len(i_sParm2) > 0 Then cmd.Parameters(i_sParm2).value = i_vParm2
    If Len(i_sParm3) > 0 Then cmd.Parameters(i_sParm3).value = i_vParm3
    If Len(i_sParm4) > 0 Then cmd.Parameters(i_sParm4).value = i_vParm4
    If Len(i_sParm5) > 0 Then cmd.Parameters(i_sParm5).value = i_vParm5
    If Len(i_sParm6) > 0 Then cmd.Parameters(i_sParm6).value = i_vParm6

    cmd.Execute
    
    If g_DB.Connection.Errors.Count > 0 Then GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    Dim sMsg As String
    Dim obj As Object
    
    If Err.Number <> 0 Then
        sMsg = "ExecuteSP " & i_sName & vbCrLf & _
                Err.Number & " (" & Err.Source & ") " & Err.Description
    End If

    If g_DB.Connection.Errors.Count > 0 Then
        Dim i As Long
        sMsg = sMsg & vbCrLf & "Additional Information:" & vbCrLf
        For Each obj In g_DB.Connection.Errors
            sMsg = sMsg & "ErrMsg " & i & ": " & obj.Description & vbCrLf
        Next
    End If

    Err.Raise -1, "Database.ExecuteSP", sMsg

End Sub


Public Function GetSurrogateKey(i_sTableName As String) As Long
    Dim cmd As ADODB.Command
    Dim lErrorCount As Long
    
    On Error GoTo ErrorHandler

RetryPoint:
    Set cmd = CreateCommandSP("spGetNextSurrogateKey")
    
    With cmd
        .Parameters("@iTableName") = i_sTableName
        .Execute
        GetSurrogateKey = .Parameters("@oNewKey").value
    End With
        
    If g_DB.Connection.Errors.Count > 0 Then GoTo ErrorHandler
   
    'Clean-Up
    Set cmd = Nothing
    Exit Function
    
ErrorHandler:
    Dim sMsg As String
    Dim obj As Object
    
    If lErrorCount = 0 Then
        lErrorCount = 1
        g_DB.Disconnect
        g_DB.Connect
        Resume RetryPoint
    End If
    
    If Err.Number <> 0 Then
        sMsg = "GetSurrogateKey " & i_sTableName & vbCrLf & _
                Err.Number & " (" & Err.Source & ") " & Err.Description
    End If

    If g_DB.Connection.Errors.Count > 0 Then
        Dim i As Long
        sMsg = sMsg & vbCrLf & "Additional Information:" & vbCrLf
        For Each obj In g_DB.Connection.Errors
            sMsg = sMsg & "ErrMsg " & i & ": " & obj.Description & vbCrLf
        Next
    End If
    
    Err.Raise -1, "Database.GetSurrogateKey", sMsg
End Function


Public Function PrepSQLText(sInput As String) As String
    PrepSQLText = Replace(sInput, "'", "''")
End Function


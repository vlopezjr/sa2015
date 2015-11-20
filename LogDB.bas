Attribute VB_Name = "LogDB"
'******************************************************************
' Event logging functions
' Created 2/3/06 LR
' as part of build 463
'
' 2/16/06 LR
' LogDB uses CreateCommandSP from Database
' but Database requires DBConnect.cls
' These in turn probably drag in modRegistry, modError, modErrorUI, ErrorInfo.cls
' In the interest of keeping this simple (adding logging to the ARStatus DLL)
' I'm editing the shared file LogdDB and replacing the calls to CreateCommandSP
' with the ADO code.
'
' This file also requires
' GetUserName()
' g_DB
'******************************************************************

Option Explicit

Private Declare Function GetComputerName Lib "kernel32" _
    Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Sub LogEvent(ByRef module As String, ByRef procedure As String, ByRef message As String)

    Dim oCmd As ADODB.Command
    
    Set oCmd = CreateCommandSP("spcpcLogEvent")
    
    With oCmd
        .Parameters("@UserID") = GetUserName
        .Parameters("@ComputerName") = vbGetComputerName
        .Parameters("@DBServer") = g_DB.server
        .Parameters("@DBName") = g_DB.database
        .Parameters("@AppPath") = VB.App.path
        .Parameters("@ProgName") = App.EXEName
        .Parameters("@Version") = App.Major & "." & App.Minor & "." & App.Revision
        .Parameters("@Module") = module
        .Parameters("@Procedure") = procedure
        .Parameters("@Message") = message
        .Execute
    End With
    
    Set oCmd = Nothing

End Sub


Public Sub LogEventExt(ByRef module As String, ByRef procedure As String, ByRef message As String, ByRef ExtData As String)

    Dim oCmd As ADODB.Command
   
    Set oCmd = CreateCommandSP("spcpcLogEventExt")
    
    With oCmd
        .Parameters("@UserID") = GetUserName
        .Parameters("@ComputerName") = vbGetComputerName
        .Parameters("@DBServer") = g_DB.server
        .Parameters("@DBName") = g_DB.database
        .Parameters("@AppPath") = VB.App.path
        .Parameters("@ProgName") = App.EXEName
        .Parameters("@Version") = App.Major & "." & App.Minor & "." & App.Revision
        .Parameters("@Module") = module
        .Parameters("@Procedure") = procedure
        .Parameters("@Message") = message
        .Parameters("@ExtData") = ExtData
        .Execute
    End With
    
    Set oCmd = Nothing

End Sub


'(byref err as ErrObject) doesn't work

Public Sub LogError(ByRef module As String, ByRef procedure As String, ByRef message As String, Source As String, Number As Long, Descr As String)

    Dim oCmd As ADODB.Command
    
    Set oCmd = CreateCommandSP("spcpcLogError")
    
    With oCmd
        .Parameters("@UserID") = GetUserName
        .Parameters("@ComputerName") = vbGetComputerName
        .Parameters("@DBServer") = g_DB.server
        .Parameters("@DBName") = g_DB.database
        .Parameters("@AppPath") = VB.App.path
        .Parameters("@ProgName") = App.EXEName
        .Parameters("@Version") = App.Major & "." & App.Minor & "." & App.Revision
        .Parameters("@Module") = module
        .Parameters("@Procedure") = procedure
        .Parameters("@Message") = message
        '.Parameters("@ErrCategory") = typErrInfo.ErrType(0)
        .Parameters("@ErrSource") = Source
        '.Parameters("@ErrLineNo") = CInt(LTrim$(typErrInfo.ErrLine(0)))
        .Parameters("@ErrNo") = Number
        .Parameters("@ErrDescr") = Descr
        .Execute
    End With
    
    Set oCmd = Nothing

End Sub


'copped from Failsafe.bas
'gets the name of the machine

Private Function vbGetComputerName() As String
  Const MAXSIZE As Integer = 256
  Dim sTmp As String * MAXSIZE
  Dim lLen As Long

  On Error GoTo ERR_HANDLER

  lLen = MAXSIZE - 1
  If (GetComputerName(sTmp, lLen)) Then
    vbGetComputerName = Left$(sTmp, lLen)
  Else
    vbGetComputerName = ""
  End If

  Exit Function
  
ERR_HANDLER:
  vbGetComputerName = ""

End Function


Public Sub LogOAEvent( _
    sEventID As String, _
    sEventUser As String, _
    lEntityPK1 As Long, _
    Optional lEntityPK2 As Variant, _
    Optional lEventIntValue As Variant, _
    Optional sEventStrValue As Variant, _
    Optional oEventDateValue As Variant, _
    Optional lEventRefKey As Variant)

    'Don't let error in LogOAEvent affect Order commit or order saving
    On Error Resume Next
    
    Dim cmd As ADODB.Command

    'The StringBuilder used when logging order events will sometimes return an empty string
    'and pass it in as sEventStrValue. when this happens, there is nothing to log.
    If Trim(sEventStrValue) = "" Then Exit Sub
    
    SetWaitCursor True
    Set cmd = CreateCommandSP("spCPCeventInsert")
    
    With cmd
        .Parameters("@_iEventID").value = sEventID
        .Parameters("@_iEntityPK1").value = lEntityPK1
        .Parameters("@_iEventUser").value = sEventUser
        
        If Not IsMissing(lEntityPK2) Then .Parameters("@_iEntityPK2").value = lEntityPK2
        If Not IsMissing(lEventIntValue) Then .Parameters("@_iEventIntValue").value = lEventIntValue
        If Not IsMissing(sEventStrValue) Then .Parameters("@_iEventStrValue").value = sEventStrValue
        If Not IsMissing(oEventDateValue) Then .Parameters("@_iEventDateValue").value = oEventDateValue
        If Not IsMissing(lEventRefKey) Then .Parameters("@_iEventRefKey").value = lEventRefKey
        .Execute
    End With
    Set cmd = Nothing
    SetWaitCursor False
End Sub


Public Sub LogActivity( _
    sProcess As String, _
    sDescr As String, _
    Optional lOPKey As Variant, _
    Optional lSOKey As Variant, _
    Optional sSOTranNo As Variant, _
    Optional lSOLineKey As Variant, _
    Optional lPOKey As Variant, _
    Optional sPOTranNo As Variant, _
    Optional lPOLineKey As Variant, _
    Optional lWhseKey As Variant)

    Dim cmd As ADODB.Command
    Set cmd = CreateCommandSP("spcpcActivityLogInsert")
    With cmd
        .Parameters("@Process").value = sProcess
        .Parameters("@Description").value = sDescr
        If Not IsMissing(lOPKey) Then .Parameters("@OPKey").value = lOPKey
        If Not IsMissing(lSOKey) Then .Parameters("@SOKey").value = lSOKey
        If Not IsMissing(sSOTranNo) Then .Parameters("@SOTranNo").value = sSOTranNo
        If Not IsMissing(lSOLineKey) Then .Parameters("@SOLineKey").value = lSOLineKey
        If Not IsMissing(lPOKey) Then .Parameters("@POKey").value = lPOKey
        If Not IsMissing(sPOTranNo) Then .Parameters("@POTranNo").value = sPOTranNo
        If Not IsMissing(lPOLineKey) Then .Parameters("@POKey").value = lPOLineKey
        If Not IsMissing(lWhseKey) Then .Parameters("@WhseKey").value = lWhseKey
        .Execute
    End With
    Set cmd = Nothing
End Sub


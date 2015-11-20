Attribute VB_Name = "User"
Option Explicit

' the logged in user
Public LoggedInUserId As String

Public g_rstUsers As ADODB.Recordset
Public g_rstBuyers As ADODB.Recordset
Public g_rstCSRs As ADODB.Recordset
Public g_rstCollectors As ADODB.Recordset

Public g_sDefaultUserID As String
Public g_bWillCallUser As Boolean

Private m_lUserKey As Long

Private Declare Function GetNTUserName Lib "advapi32.dll" _
    Alias "GetUserNameA" _
    (ByVal lpBuffer As String, _
    nSize As Long) As Long


'Called by Main()

Public Sub Init()
    Dim sSQL As String


    sSQL = "SELECT tcpUser.UserID, tcpUser.UserKey, tcpUser.BranchID, tcpBranch.WhseKey " _
        & "FROM tcpUser INNER JOIN " _
        & "tcpBranch ON tcpUser.BranchID = tcpBranch.BranchID " _
        & "ORDER BY tcpUser.UserID"
                         
    Set g_rstUsers = LoadDiscRst(sSQL)
    
    LogOnUser
    
    g_bWillCallUser = IsWillCallGroupUser

    sSQL = "SELECT tcpUser.UserID,tcpUser.UserKey, tcpUser.BranchID, tcpUser.IsActive " & _
        "FROM  tcpUser INNER JOIN tcpGroupMember ON tcpUser.UserKey = tcpGroupMember.UserKey " & _
        "INNER JOIN tcpGroup ON tcpGroupMember.GroupKey = tcpGroup.GroupKey " & _
        "WHERE (tcpGroup.GroupID = 'CSR' or tcpGroup.GroupID = 'CSR Will Call' or tcpGroup.GroupID = 'WillCall') ORDER BY tcpUser.UserID"
    Set g_rstCSRs = LoadDiscRst(sSQL)

    sSQL = "SELECT dbo.timBuyer.BuyerKey, dbo.timBuyer.BuyerID, dbo.tcpUser.UserKey " _
        & "FROM dbo.timBuyer " _
        & "INNER JOIN dbo.tcpUser ON dbo.timBuyer.BuyerID = dbo.tcpUser.UserID " _
        & "WHERE dbo.timBuyer.CompanyID = 'cpc' " _
        & "ORDER BY BuyerID"
    Set g_rstBuyers = LoadDiscRst(sSQL)


    sSQL = "SELECT tcpUser.UserID, tcpUser.UserKey, tcpUser.BranchID " _
        & "FROM  tcpUser INNER JOIN " _
        & "tcpGroupMember ON tcpUser.UserKey = tcpGroupMember.UserKey INNER JOIN " _
        & "tcpGroup ON tcpGroupMember.GroupKey = tcpGroup.GroupKey " _
        & "WHERE tcpGroup.GroupID = 'Collectors' and tcpUser.IsActive = -1 " _
        & "ORDER BY tcpUser.UserID"
    Set g_rstCollectors = LoadDiscRst(sSQL)
End Sub


Public Sub SetUpUsers(cbo As ComboBox, rst As ADODB.Recordset)
    LoadCombo cbo, rst, "UserID", "UserKey"
End Sub


Public Sub LoadActiveCSRs(cbo As ComboBox)
    Dim rst As ADODB.Recordset
    Set rst = LoadDiscRst("select distinct tcpso.userid, tcpuser.userkey from tcpso " & _
        "inner join tcpuser on tcpso.userid = tcpuser.userid " & _
        "where createdate > DATEADD(year,-1,GETDATE()) order by tcpso.userid")
    LoadCombo cbo, rst, "UserID", "UserKey"
    Set rst = Nothing
End Sub


Public Sub LoadActiveUsers(cbo As ComboBox)
    Dim rst As ADODB.Recordset
    Set rst = LoadDiscRst("select userid, userkey from tcpuser where isactive = -1 order by branchid, userid")
    LoadCombo cbo, rst, "UserID", "UserKey"
    Set rst = Nothing
End Sub


Public Function GetUserName() As String
    Dim lResult As Long
    Dim lsize As Long
    Dim sBuffer As String
    
    If LoggedInUserId = "" Then
        sBuffer = String(256, "*")
        lResult = GetNTUserName(sBuffer, 256)
        If lResult > 0 Then
            lsize = InStr(1, sBuffer, chr(0)) - 1
            If lsize > 0 Then
                LoggedInUserId = LCase(Mid(sBuffer, 1, lsize))
            End If
        End If
    End If
    
    GetUserName = LoggedInUserId
End Function


Public Function GetUserID(Optional ByVal i_lUserKey As Long = -1) As String
    If i_lUserKey = -1 Then
        i_lUserKey = m_lUserKey
    End If
    
    g_rstUsers.Filter = "UserKey=" & i_lUserKey
    If Not g_rstUsers.EOF Then
        GetUserID = Trim(g_rstUsers.Fields("UserID").value)
    End If
    g_rstUsers.Filter = adFilterNone
End Function


Public Function GetUserKey(Optional ByVal i_sUserId As String = "") As Long
    If i_sUserId = "" Then
        If g_sDefaultUserID = "" Then
            GetUserKey = m_lUserKey
            Exit Function
        Else
            i_sUserId = g_sDefaultUserID
        End If
    End If
    
    g_rstUsers.Filter = "UserID='" & i_sUserId & "'"
    If Not g_rstUsers.EOF Then
        GetUserKey = g_rstUsers.Fields("UserKey").value
    End If
    g_rstUsers.Filter = adFilterNone
End Function


Public Function GetUserWhseKey(Optional ByVal i_lUserKey As Long = -1) As Long
    Dim prevFilter As String
    
    If i_lUserKey < 0 Then
        i_lUserKey = m_lUserKey
    End If
    
    With g_rstUsers
        .Filter = "UserKey=" & i_lUserKey
        If g_rstWhses.Filter <> adFilterNone Then prevFilter = g_rstWhses.Filter
        g_rstWhses.Filter = "WhseID = '" & .Fields("BranchID").value & "'"
        If Not g_rstWhses.EOF Then
            GetUserWhseKey = Trim(g_rstWhses.Fields("WhseKey").value)
        End If
        'g_rstWhses.Filter = adFilterNone
        If prevFilter <> vbNullString Then g_rstWhses.Filter = prevFilter
        .Filter = adFilterNone
    End With
End Function


Public Function GetUserWhseID(Optional ByVal i_lUserKey As Long = -1) As String
    If i_lUserKey < 0 Then
        i_lUserKey = m_lUserKey
    End If
    
    With g_rstUsers
        .Filter = "UserKey=" & i_lUserKey
        GetUserWhseID = .Fields("BranchID").value
        .Filter = adFilterNone
    End With
End Function


'This is not called from anywhere in OfficeAssit.
'Is it used by another project?

Public Function GetUserCommittedOrder(lOPKey As Long) As String
    Dim ocmd As ADODB.Command
    Set ocmd = CreateCommandSP("spcpcGetUserCommittedOrder")
    With ocmd
        .Parameters("@_iOPKey").value = lOPKey
        .Execute
        GetUserCommittedOrder = .Parameters("@_oUserName").value
    End With
    Set ocmd = Nothing
End Function


Public Function GetUserWhseShipAddrKey(i_sUserId As String) As Long
    g_rstUsers.Filter = "UserId='" & i_sUserId & "'"
    g_rstWhses.Filter = "whsekey=" & g_rstUsers.Fields("whsekey").value
    If Not g_rstWhses.EOF Then
        GetUserWhseShipAddrKey = g_rstWhses.Fields("ShipAddrKey").value
    End If
    g_rstWhses.Filter = adFilterNone
    g_rstUsers.Filter = adFilterNone
End Function


'Convert UserName to BuyerKey.
'If the user is not a buyer in Sage, then return 0.
Public Function UserNameToBuyerKey(ByVal UserName As String) As Long

    g_rstBuyers.Filter = "BuyerID='" & UserName & " '"
    If g_rstBuyers.EOF Then
        UserNameToBuyerKey = 0
    Else
        UserNameToBuyerKey = g_rstBuyers.Fields("BuyerKey")
    End If
    g_rstBuyers.Filter = adFilterNone
    
End Function


Public Function BuyerKeyToUserID(ByVal BuyerKey As Long) As String
    
    g_rstBuyers.Filter = "BuyerKey=" & BuyerKey
    If g_rstBuyers.EOF Then
        BuyerKeyToUserID = vbNullString
    Else
        BuyerKeyToUserID = g_rstBuyers.Fields("BuyerID")
    End If
    g_rstBuyers.Filter = adFilterNone

End Function


Public Function GetUserSalesAcctKey(Optional ByVal i_lUserKey As Long = -1) As Long
    If i_lUserKey < 0 Then
        i_lUserKey = m_lUserKey
    End If
    
    With g_rstUsers
        .Filter = "UserKey = " & i_lUserKey
        g_rstWhses.Filter = "WhseID='" & .Fields("BranchID").value & "'"
        If Not g_rstWhses.EOF Then
            GetUserSalesAcctKey = Trim(g_rstWhses.Fields("SalesAcctKey").value)
        End If
        g_rstWhses.Filter = adFilterNone
        .Filter = adFilterNone
    End With
End Function


Public Function GetUserShipMethKey(Optional ByVal i_lUserKey As Long = -1) As Long
    If i_lUserKey < 0 Then
        i_lUserKey = m_lUserKey
    End If
    
    With g_rstUsers
        .Filter = "UserKey=" & i_lUserKey
        g_rstWhses.Filter = "WhseID ='" & .Fields("BranchID").value & "'"
        If Not g_rstWhses.EOF Then
            GetUserShipMethKey = Trim(g_rstWhses.Fields("ShipMethKey").value)
        End If
        g_rstWhses.Filter = adFilterNone
        .Filter = adFilterNone
    End With
End Function


'*** 4/7/08 LR: Why isn't this determined by looking at cached global record sets?

Public Function IsWillCallGroupUser()
    Dim sSQL As String
    Dim rst As ADODB.Recordset
    
    sSQL = "select tcpUser.* from tcpUser inner join tcpGroupMember " _
        & "on tcpGroupMember.UserKey = tcpUser.UserKey " _
        & "inner join tcpGroup on tcpGroup.GroupKey = tcpGroupMember.GroupKey " _
        & "where tcpGroup.GroupID = 'WillCall' and tcpUser.UserKey = " & GetUserKey
    
    Set rst = LoadDiscRst(sSQL)
    
    If Not rst.EOF Then IsWillCallGroupUser = True
    Set rst = Nothing
End Function


Public Function IsAdmin(ByVal UserName As String) As Boolean
    Dim lc_username As String
    
    lc_username = LCase(UserName)
    
    If lc_username = "lennyr" Or lc_username = "victorl" Or lc_username = "dannyh" Then
        IsAdmin = True
    Else
        IsAdmin = False
    End If
End Function


'*** Ugly
'This is a very odd place to put this dependency
'It's initialized in this routine
'    m_lUserKey = g_rstUsers.Fields("UserKey").Value
'And used by
'GetUserID
'GetUserKey
'GetUserWhseKey
'GetUserWhseID
'GetUserSalesAcctKey
'GetUserShipMethKey

Public Sub LogOnUser()
    g_rstUsers.Filter = "UserID='" & GetUserName & "'"
    If Not g_rstUsers.EOF Then
        m_lUserKey = g_rstUsers.Fields("UserKey").value
        LogUser True
    End If
    g_rstUsers.Filter = adFilterNone
End Sub


Public Sub LogOffUser()
    If m_lUserKey <> 0 Then
        LogUser False
    End If
End Sub


Private Sub LogUser(IsLoggingOn As Boolean)
    Dim cmd As ADODB.Command

    Set cmd = CreateCommandSP("spCPCLogUser")
   
    With cmd
        .Parameters("@UserKey").value = m_lUserKey
        .Parameters("@IsLoggingOn").value = IsLoggingOn
        .Execute
    End With
    Set cmd = Nothing

End Sub


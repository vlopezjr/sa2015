VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccessRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_oRecordSet As ADODB.Recordset
Private m_sUserID As String


Public Function HasRight(ByVal i_sRightID As String) As Boolean
    LoadOnDemand
    
    m_oRecordSet.Filter = "RightID = '" & Trim(i_sRightID) & "'"
    
    HasRight = (m_oRecordSet.RecordCount <> 0)
End Function


Private Sub LoadOnDemand()
    If m_oRecordSet Is Nothing Then
        Set m_oRecordSet = CallSP("spCPCShowRights", "@UserID", UserId)
        Set m_oRecordSet.ActiveConnection = Nothing
    End If
End Sub


Public Property Get UserId() As String
    UserId = m_sUserID
End Property


Public Property Let UserId(ByVal sUserID As String)
    Set m_oRecordSet = Nothing
    m_sUserID = Trim(sUserID)
End Property


Private Sub Class_Initialize()
    m_sUserID = GetUserName
End Sub


Private Sub Class_Terminate()
    Set m_oRecordSet = Nothing
End Sub

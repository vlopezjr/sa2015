VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colNotifications As Collection
Private m_colNotifications_Backup As Collection

Public Property Get Count() As Integer
    Count = m_colNotifications.Count
End Property


Private Sub Class_Initialize()
    Set m_colNotifications = New Collection
    Set m_colNotifications_Backup = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colNotifications = Nothing
    Set m_colNotifications_Backup = Nothing
End Sub

Public Sub Load(CntctKey As Long)
    Dim oCmd As New ADODB.Command
    Dim oRst As New ADODB.Recordset
    
    With oCmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spcpcContactLoadNotify"
        Set oRst = .Execute(, CntctKey)
        Do Until oRst.EOF
            m_colNotifications.Add oRst.Fields("NotifyKey").value
            oRst.MoveNext
        Loop
    End With
    oRst.Close
    Set oCmd = Nothing
End Sub

Public Sub Backup()
    Dim i As Integer
    
    ClearCollection m_colNotifications_Backup
    For i = 1 To m_colNotifications.Count
        m_colNotifications_Backup.Add m_colNotifications.Item(i)
    Next
End Sub

Public Sub Update(CntctKey As Long)
    Dim i As Integer
    Dim oCmd As ADODB.Command
    'delete notifications from db for this contact.
    Delete CntctKey
    
    Set oCmd = CreateCommandSP("spcpcContactInsertNotify")
    
    'write the notifications to the db from the current collection.
    For i = 1 To m_colNotifications.Count
        oCmd.Parameters("@CntctKey").value = CntctKey
        oCmd.Parameters("@NotifyKey").value = m_colNotifications(i)
        oCmd.Execute
    Next
End Sub

' checkbox control arrays are 0-based
' tcpwebnotifications keys are 1-based
' therefore we must increment the index

Public Sub Add(Item As Long)
    m_colNotifications.Add Item + 1
End Sub

Public Sub Clear()
    'Need to test for empty
    ClearCollection m_colNotifications
End Sub

'delete all of a contact's notificatitons from the database

Public Sub Delete(CntctKey As Long)
    Dim oCmd As ADODB.Command

    Set oCmd = CreateCommandSP("DELETE tcpWebContactNotify WHERE cntctkey=" & CntctKey, adCmdText)
    oCmd.Execute
End Sub

'remove a notification from the collection

Public Sub Remove(NotifyKey As Integer)
    Dim i As Integer
    
    NotifyKey = NotifyKey + 1
    For i = 1 To m_colNotifications.Count
        If m_colNotifications(i) = NotifyKey Then
            Exit For
        End If
    Next
    m_colNotifications.Remove i
End Sub


Public Sub Restore()
    Dim i As Integer
    ClearCollection m_colNotifications
    For i = 1 To m_colNotifications_Backup.Count
        m_colNotifications.Add m_colNotifications_Backup.Item(i)
    Next
End Sub


Public Function Export(Optional bSplitOrder As Boolean) As JDMPDXML.XMLNode
    Dim oXMLchild As JDMPDXML.XMLNode
    Dim i As Integer
    
    Set oXMLchild = New JDMPDXML.XMLNode
    oXMLchild.Tag = "Notifications"
    For i = 1 To m_colNotifications.Count
        oXMLchild.Add "Notification", m_colNotifications(i)
    Next
    
    Set Export = oXMLchild
End Function


Public Sub Import(ByVal i_oXML As JDMPDXML.XMLNode)
    Dim oSubXMLNode As JDMPDXML.XMLNode
    Dim oSubSubXMLNode As JDMPDXML.XMLNode
    
    Set oSubXMLNode = New JDMPDXML.XMLNode
    Set oSubSubXMLNode = New JDMPDXML.XMLNode

    For Each oSubXMLNode In i_oXML
        If oSubXMLNode.Tag = "Notifications" Then
            For Each oSubSubXMLNode In oSubXMLNode
            'Need to have a Notifications node for structure purposes
                If oSubSubXMLNode.Tag = "Notification" Then
                    m_colNotifications.Add oSubXMLNode.Item("Notification").value
                End If
            Next
        End If
    Next
End Sub


'this is the default method (see Tools Menu, Procedure Attributes)
'Decrement the key value to match the 0 based checkbox control array.
Public Function Item(ByVal Index As Variant) As Integer
Attribute Item.VB_UserMemId = 0
    Item = m_colNotifications.Item(Index) - 1
End Function


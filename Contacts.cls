VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colContacts As Collection

Public WithEvents m_oSelContact As Contact
Attribute m_oSelContact.VB_VarHelpID = -1

Private m_iSelIndex As Integer

Private m_sOwnerName As String
Private m_sOwnerID As String
Private m_lOwnerKey As Long

'***DH 4/1/08
'Private m_sUserID As String

' Events

Public Event StateChange(NewState As BitMap)

' Properties

Public Property Get OwnerName() As String
    OwnerName = m_sOwnerName
End Property

Public Property Let OwnerName(ByVal value As String)
    m_sOwnerName = value
End Property

Public Property Get OwnerID() As String
    OwnerID = m_sOwnerID
End Property

Public Property Let OwnerID(ByVal value As String)
    m_sOwnerID = value
End Property

Public Property Get OwnerKey() As Long
    OwnerKey = m_lOwnerKey
End Property

Public Property Let OwnerKey(ByVal value As Long)
    m_lOwnerKey = value
End Property


Public Property Get selContact() As Contact
    Set selContact = m_oSelContact
End Property

Public Property Get SelIndex() As Integer
    SelIndex = m_iSelIndex
End Property

Public Property Get Contacts() As Collection
    Set Contacts = m_colContacts
End Property

Public Function Count() As Long
    Count = m_colContacts.Count
End Function

'***DH 4/1/08
'Public Sub Delete()
'    m_oSelContact.Delete
'    Remove
'End Sub

'Removes the contact from the collection.
Public Sub Remove()
    m_colContacts.Remove m_iSelIndex
End Sub

'This is the default method (see Tools Menu, Procedure Attributes)
Public Function Item(ByVal Index As Variant) As Contact
Attribute Item.VB_UserMemId = 0
    Set Item = m_colContacts.Item(Index)
End Function

'This is hidden in the typelib (see Tools Menu, Procedure Attributes)
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colContacts.[_NewEnum]
End Function


' Methods

'This will generate an error if the contact is not found.
Public Function GetContactByKey(Key As String) As Contact
    SelectContact Key
    Set GetContactByKey = m_oSelContact
End Function


Public Sub SelectContact(Key As String)
    On Error Resume Next
    Set m_oSelContact = m_colContacts(Key)

    If Err.Number > 0 Then
        Set m_oSelContact = Nothing
        m_iSelIndex = 0
    Else
        SetIndex
    End If
End Sub


Private Sub SetIndex()
    Dim i As Integer
    For i = 1 To m_colContacts.Count
        If m_oSelContact.Key = m_colContacts(i).Key Then
            m_iSelIndex = i
            Exit For
        End If
    Next
End Sub


Public Sub LoadContacts()
    Dim oContact As Contact
    Dim orst As ADODB.Recordset

    Dim oCmd As New ADODB.Command
    With oCmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spcpcContactsLoad"
        .Parameters.Append .CreateParameter("@CntctOwnerKey", adInteger, adParamInput, , m_lOwnerKey)
    End With
    Set orst = oCmd.Execute
   
    If Not orst.EOF Then
        'Leave the recordset connected
        Do Until orst.EOF
            Set oContact = New Contact
            oContact.Load orst
            m_colContacts.Add oContact, CStr(Trim(orst.Fields("CntctKey").value))
            orst.MoveNext
        Loop
    End If
    
    Set oContact = Nothing
    Set oCmd = Nothing
    orst.Close
    Set orst = Nothing
End Sub


'When adding a new contact put it into the collection in alphabetical order.

Public Sub InsertNewContact(ByRef oNewContact As Contact)
    Set m_oSelContact = oNewContact
    InsertContact oNewContact
End Sub

Public Sub InsertContact(ByRef oNewContact As Contact)
    Dim oContact As Contact
    Dim Count As Integer
    
    Count = 1
    For Each oContact In m_colContacts
        If StrComp(oContact.Name, oNewContact.Name, vbTextCompare) = 1 Then
            m_colContacts.Add Item:=oNewContact, Key:=CStr(oNewContact.Key), Before:=Count
            SetIndex
            Exit Sub
        End If
        Count = Count + 1
    Next
    m_colContacts.Add Item:=oNewContact, Key:=CStr(oNewContact.Key)
    SetIndex
End Sub


Public Sub Clear()
    Dim i As Integer
    For i = 1 To m_colContacts.Count
        m_colContacts.Remove 1
    Next i
End Sub

'***DH 4/1/08
''note: This contact is not in the collection
'Public Sub NewContact()
'    Set m_oSelContact = New Contact
'    m_iSelIndex = 0
'    m_oSelContact.UserID = m_sUserID
'End Sub

'***DH 4/1/08
''Add or update the Selected Contact
'Public Sub AddUpdate()
'
'    If m_oSelContact.IsNew Then
'        'Insert into the database
'        m_oSelContact.Insert
'        'Add into the collection in alphabetical order
'        InsertNewContact m_oSelContact
'    Else
'        'Update the database
'        m_oSelContact.Update
'    End If
'
'End Sub


Public Function Export(Optional bSplitOrder As Boolean) As JDMPDXML.XMLNode
    Dim oXML As New JDMPDXML.XMLNode
    Dim i As Integer
    Dim oContact As Contact

    If bSplitOrder Then
        With oXML
            .Tag = "Contacts"
            .Add "SelIndex", m_iSelIndex
            .Add "SelContactKey", m_oSelContact.Key
            
            For i = 1 To m_colContacts.Count
                Set oContact = m_colContacts(CStr(m_colContacts(i).Key))
                .AddChild oContact.Export(bSplitOrder:=True)
            Next
            
        End With
        Set Export = oXML
    Else
        Set Export = Nothing
    End If
    
End Function


Public Function Import(ByVal i_oXML As JDMPDXML.XMLNode)
    Dim lSelContactkey  As Long
    Dim oSubXMLNode As JDMPDXML.XMLNode
    
    With i_oXML
        If .Tag = "Contacts" Then
            m_iSelIndex = .Item("SelIndex").value
            lSelContactkey = .Item("SelContactKey").value
            
            'Each Contact
            For Each oSubXMLNode In i_oXML
                If oSubXMLNode.Tag = "Contact" Then
                    Dim oContact As Contact
                    Set oContact = New Contact
                                      
                    'oSubXMLNode.ExportFile "c:\XMLContactsTest" & Format(Now, "ddyy nnss") & ".txt"
                    oContact.Import oSubXMLNode
                    m_colContacts.Add oContact, oSubXMLNode.Item("cntctkey").value
                    
                    If oContact.Key = lSelContactkey Then
                        Set m_oSelContact = m_colContacts(CStr(oContact.Key))
                    End If
                End If
            Next
            
        End If
    End With
    
End Function

'***LR 11/19/07
Public Sub Edit(i_sUserID As String)
'    FContactsMgr.Init Me, bSelect:=False
'***DH 4/1/08
    'm_sUserID = i_sUserID
    
    If m_colContacts.Count > 0 Then
        'Select the first contact in the collection.
        Set m_oSelContact = m_colContacts(1)
        m_iSelIndex = 1
    End If
'***DH 4/1/08
    'FContactsMgr.Init Me
    FContactMgr.Init Me, i_sUserID
End Sub


'*************************************************************************
'Event Handlers
'*************************************************************************

Private Sub Class_Initialize()
    Set m_colContacts = New Collection
End Sub


Private Sub Class_Terminate()
    'do I need to destroy the contact objects?
    Set m_colContacts = Nothing
    Set m_oSelContact = Nothing
End Sub


Private Sub m_oSelContact_NameChange()
    'Remove contact from collection
    Remove
    InsertContact m_oSelContact
End Sub

'Bubble up
Private Sub m_oSelContact_StateChange(NewState As BitMap)
    RaiseEvent StateChange(NewState)
End Sub



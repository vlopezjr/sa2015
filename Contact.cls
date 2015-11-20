VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum OwnerType
    opNone = 0
    opOrder = 30801
    opCustomer = 501
End Enum

Private Type ContactProperties
    CntctKey As Long
    CntctOwnerKey As Long
    FirstName As String
    LastName As String
    CreateUserID As String
    CreateDate As Date
    UpdateUserID As String
    'This is a variant to support Empty values
    UpdateDate As Variant
    EntityType As OwnerType
    emailaddr As String
    EMailFormat As Integer
    Fax As String
    FaxExt As String
    Phone As String
    PhoneExt As String
    CellPhone As String
    Title As String
    DeclinedEmailAddr As Boolean
End Type

Private m_udtContact As ContactProperties
Private WithEvents m_oState As BitMap
Attribute m_oState.VB_VarHelpID = -1
Private m_oNotifications As Notifications
Private m_udtContact_Backup As ContactProperties
Private m_oState_Backup As BitMap
Private m_sUserID As String
Private sErrMsg As String


'************************************************************************
'PUBLIC EVENTS
'************************************************************************
Public Event StateChange(NewState As BitMap)
Public Event NameChange()


Private Sub Class_Initialize()
    Set m_oState = New BitMap
    Set m_oState_Backup = New BitMap
    Set m_oNotifications = New Notifications
    'The initial state of a New Contacts is NEW
    Clear
End Sub


Private Sub Class_Terminate()
    Set m_oState = Nothing
       Set m_oNotifications = Nothing
End Sub

Public Property Get State() As BitMap
    Set State = m_oState
End Property

Public Property Get Key() As Long
    Key = m_udtContact.CntctKey
End Property

Public Property Let Key(ByVal lNewValue As Long)
    m_udtContact.CntctKey = lNewValue
End Property

Public Property Get OwnerType() As OwnerType
    OwnerType = m_udtContact.EntityType
End Property

Public Property Let OwnerType(ByVal eNewValue As OwnerType)
    m_udtContact.EntityType = eNewValue
End Property

Public Property Get OwnerKey() As Long
    OwnerKey = m_udtContact.CntctOwnerKey
End Property

Public Property Let OwnerKey(ByVal lNewValue As Long)
    m_udtContact.CntctOwnerKey = lNewValue
End Property

Public Property Get Name() As String
    Name = Trim$(m_udtContact.FirstName + " " + m_udtContact.LastName)
End Property

Public Property Get FirstName() As String
    FirstName = m_udtContact.FirstName
End Property

Public Property Let FirstName(ByVal sNewValue As String)
    m_udtContact.FirstName = Trim$(sNewValue)
End Property

Public Property Get LastName() As String
    LastName = m_udtContact.LastName
End Property

Public Property Let LastName(ByVal sNewValue As String)
    m_udtContact.LastName = Trim$(sNewValue)
End Property

Public Property Get emailaddr() As String
    emailaddr = Trim(m_udtContact.emailaddr)
End Property

Public Property Let emailaddr(ByVal sNewValue As String)
    m_udtContact.emailaddr = sNewValue
End Property


Public Property Get EMailFormat() As Integer
    EMailFormat = m_udtContact.EMailFormat
End Property

Public Property Let EMailFormat(ByVal iNewValue As Integer)
    m_udtContact.EMailFormat = iNewValue
End Property


Public Property Get Phone() As String
    Phone = Trim(m_udtContact.Phone)
End Property

Public Property Let Phone(ByVal sNewValue As String)
    m_udtContact.Phone = sNewValue
End Property


Public Property Get PhoneExt() As String
    PhoneExt = Trim$(m_udtContact.PhoneExt)
End Property

Public Property Let PhoneExt(ByVal sNewValue As String)
    m_udtContact.PhoneExt = sNewValue
End Property


Public Property Get CellPhone() As String
    CellPhone = Trim(m_udtContact.CellPhone)
End Property

Public Property Let CellPhone(ByVal sNewValue As String)
    m_udtContact.CellPhone = sNewValue
End Property

Public Property Get Fax() As Variant
    Fax = m_udtContact.Fax
End Property

Public Property Let Fax(ByVal sNewValue As Variant)
    m_udtContact.Fax = sNewValue
End Property

Public Property Get FaxExt() As String
    FaxExt = Trim$(m_udtContact.FaxExt)
End Property

Public Property Let FaxExt(ByVal sNewValue As String)
    m_udtContact.FaxExt = sNewValue
End Property

Public Property Get DeclinedEmailAddr() As Boolean
    DeclinedEmailAddr = m_udtContact.DeclinedEmailAddr
End Property

Public Property Let DeclinedEmailAddr(ByVal bNewValue As Boolean)
    m_udtContact.DeclinedEmailAddr = bNewValue
End Property

Public Property Get CreateUserID() As String
    CreateUserID = m_udtContact.CreateUserID
End Property

Public Property Get CreateDate() As Date
    CreateDate = m_udtContact.CreateDate
End Property

Public Property Get UpdateUserID() As String
    UpdateUserID = m_udtContact.UpdateUserID
End Property

Public Property Get UpdateDate() As Date
    UpdateDate = m_udtContact.UpdateDate
End Property

Public Property Let UserID(ByVal sNewValue As String)
    m_sUserID = sNewValue
End Property

Public Property Get Notifications() As Notifications
    Set Notifications = m_oNotifications
End Property

Public Function IsDirty() As Boolean
    IsDirty = m_oState.TestBits(eMask.IsDirty)
    
End Function

Public Function IsNew() As Boolean
    IsNew = m_oState.TestBits(eMask.IsNew)
End Function

Public Function IsValid() As Boolean
    IsValid = m_oState.TestBits(eMask.IsValid)
End Function

'Public Methods

Public Sub LoadByKey(ByVal CntctKey As Long)
    Dim orst As New ADODB.Recordset
    Dim oCmd As New ADODB.Command
    
    Clear

    With oCmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spcpcContactLoad"
        .Parameters.Append .CreateParameter("@CntctKey", adInteger, adParamInput, , CntctKey)
    End With
    
    Set orst = oCmd.Execute
    If orst.EOF Then Exit Sub
    
    With m_udtContact
        .CntctKey = CntctKey
        .CntctOwnerKey = orst.Fields("CntctOwnerKey").value
        .EntityType = orst.Fields("EntityType").value
        .CreateUserID = Trim(orst.Fields("CreateUserID").value)
        .UpdateUserID = Trim(orst.Fields("UpdateUserID").value)
        .CreateDate = orst.Fields("CreateDate").value
        .CellPhone = IIf(IsNull(orst.Fields("CellPhone").value), "", Trim(orst.Fields("CellPhone").value))
        .emailaddr = IIf(IsNull(orst.Fields("EMailAddr").value), "", Trim(orst.Fields("EMailAddr").value))
        .EMailFormat = orst.Fields("EMailFormat").value
        .Fax = IIf(IsNull(orst.Fields("Fax").value), "", Trim(orst.Fields("Fax").value))
        .FaxExt = IIf(IsNull(orst.Fields("FaxExt").value), "", Trim(orst.Fields("FaxExt").value))
        .Phone = IIf(IsNull(orst.Fields("Phone").value), "", Trim(orst.Fields("Phone").value))
        .PhoneExt = IIf(IsNull(orst.Fields("PhoneExt").value), "", Trim(orst.Fields("PhoneExt").value))
        
        If IsNull(orst.Fields("UpdateDate").value) Then
            .UpdateDate = Empty
        Else
            .UpdateDate = orst.Fields("UpdateDate").value
        End If

        .DeclinedEmailAddr = IIf(IsNull(orst.Fields("DeclinedEmailAddr").value), 0, orst.Fields("DeclinedEmailAddr").value)
        .FirstName = IIf(IsNull(orst.Fields("FirstName").value), "", Trim(orst.Fields("FirstName").value))
        .LastName = IIf(IsNull(orst.Fields("LastName").value), "", Trim(orst.Fields("LastName").value))
    End With

    orst.Close
    
    m_oNotifications.Load CntctKey
    m_oState.SetBits eMask.IsValid
    m_oState.ClearBits eMask.IsNew + eMask.IsDirty

    Backup
End Sub


'Note: this routine uses the client's database connection within oirst.

Public Sub Load(ByRef oirst As ADODB.Recordset)

    With oirst
        m_udtContact.CntctKey = .Fields("CntctKey").value
        m_udtContact.CntctOwnerKey = .Fields("CntctOwnerKey").value
        m_udtContact.CreateUserID = Trim(.Fields("CreateUserID").value)
        m_udtContact.CreateDate = .Fields("CreateDate").value
        m_udtContact.UpdateUserID = Trim(.Fields("UpdateUserID").value)

        If IsNull(.Fields("UpdateDate").value) Then
            m_udtContact.UpdateDate = Empty
        Else
            m_udtContact.UpdateDate = .Fields("UpdateDate").value
        End If
        
        m_udtContact.EntityType = opCustomer
        m_udtContact.emailaddr = IIf(IsNull(.Fields("EMailAddr").value), "", Trim(.Fields("EMailAddr").value))
        m_udtContact.EMailFormat = .Fields("EMailFormat").value
        m_udtContact.Fax = Trim(.Fields("Fax").value)
        m_udtContact.FaxExt = Trim(.Fields("FaxExt").value)
        m_udtContact.Phone = IIf(IsNull(.Fields("Phone").value), "", Trim(.Fields("Phone").value))
        m_udtContact.PhoneExt = IIf(IsNull(.Fields("PhoneExt").value), "", Trim(.Fields("PhoneExt").value))
        m_udtContact.CellPhone = IIf(IsNull(.Fields("CellPhone").value), "", Trim(.Fields("CellPhone").value))
        m_udtContact.DeclinedEmailAddr = IIf(IsNull(.Fields("DeclinedEmailAddr").value), 0, .Fields("DeclinedEmailAddr").value)
        m_udtContact.FirstName = IIf(IsNull(.Fields("FirstName").value), "", Trim(.Fields("FirstName").value))
        m_udtContact.LastName = IIf(IsNull(.Fields("LastName").value), "", Trim(.Fields("LastName").value))
    End With
       
    m_oNotifications.Load m_udtContact.CntctKey

    With m_udtContact
        If .FirstName = "" Or .LastName = "" Or .Phone = "" Then
            m_oState.ClearBits eMask.IsValid
        Else
            m_oState.SetBits eMask.IsValid
        End If
    End With
    m_oState.ClearBits eMask.IsNew + eMask.IsDirty
    
    Backup
End Sub


Public Sub Backup()
    m_udtContact_Backup = m_udtContact
    m_oState_Backup.value = m_oState.value
    m_oNotifications.Backup
End Sub


Public Sub Restore()
    m_udtContact = m_udtContact_Backup
    m_oState.value = m_oState_Backup.value
    m_oNotifications.Restore
End Sub


Public Sub Insert()
    Dim oCmd As ADODB.Command
    Dim RetVal As Integer
    
    Set oCmd = CreateCommandSP("spcpcContactInsert")
    With oCmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)

        .Parameters.Append .CreateParameter("@CntctOwnerKey", adInteger, adParamInput, 0, m_udtContact.CntctOwnerKey)
        .Parameters.Append .CreateParameter("@Name", adVarChar, adParamInput, 40, m_udtContact.FirstName + " " + m_udtContact.LastName)
        .Parameters.Append .CreateParameter("@CntctKey", adInteger, adParamOutput)
        .Parameters.Append .CreateParameter("@CreateUserID", adChar, adParamInput, 30, m_sUserID)
        .Parameters.Append .CreateParameter("@EmailAddr", adVarChar, adParamInput, 40, m_udtContact.emailaddr)
        .Parameters.Append .CreateParameter("@EmailFormat", adSmallInt, adParamInput, , m_udtContact.EMailFormat)
        .Parameters.Append .CreateParameter("@EntityType", adSmallInt, adParamInput, , m_udtContact.EntityType)
        .Parameters.Append .CreateParameter("@Fax", adChar, adParamInput, 17, m_udtContact.Fax)
        .Parameters.Append .CreateParameter("@FaxExt", adChar, adParamInput, 4, m_udtContact.FaxExt)
        .Parameters.Append .CreateParameter("@Phone", adChar, adParamInput, 17, m_udtContact.Phone)
        .Parameters.Append .CreateParameter("@PhoneExt", adChar, adParamInput, 4, m_udtContact.PhoneExt)
        .Parameters.Append .CreateParameter("@CellPhone", adChar, adParamInput, 17, m_udtContact.CellPhone)
        .Parameters.Append .CreateParameter("@DeclinedEmailAddr", adSmallInt, adParamInput, , m_udtContact.DeclinedEmailAddr)
        .Parameters.Append .CreateParameter("@FirstName", adVarChar, adParamInput, 30, m_udtContact.FirstName)
        .Parameters.Append .CreateParameter("@LastName", adVarChar, adParamInput, 30, m_udtContact.LastName)
        .Execute
        m_udtContact.CntctKey = Trim(.Parameters("@CntctKey").value)
        RetVal = .Parameters("RETURN_VALUE").value
    End With

    If RetVal = 0 Then
        Err.Raise 1, "Create", "This contact name already exists."
        Exit Sub
    End If

    m_oNotifications.Update m_udtContact.CntctKey
    m_udtContact.CreateUserID = m_sUserID
    m_udtContact.CreateDate = Date
    m_oState.ClearBits eMask.IsNew + eMask.IsDirty
    
    Backup
End Sub


Public Sub Update()
    Dim oCmd As ADODB.Command
    Dim RetVal As Integer
    
    Set oCmd = CreateCommandSP("spcpcContactUpdate")
    With oCmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@CntctOwnerKey", adInteger, adParamInput, 0, m_udtContact.CntctOwnerKey)
        .Parameters.Append .CreateParameter("@EntityType", adSmallInt, adParamInput, , m_udtContact.EntityType)
        .Parameters.Append .CreateParameter("@CntctKey", adInteger, adParamInput, 0, m_udtContact.CntctKey)
        .Parameters.Append .CreateParameter("@EmailAddr", adVarChar, adParamInput, 40, m_udtContact.emailaddr)
        .Parameters.Append .CreateParameter("@EmailFormat", adSmallInt, adParamInput, , m_udtContact.EMailFormat)
        .Parameters.Append .CreateParameter("@Fax", adChar, adParamInput, 17, m_udtContact.Fax)
        .Parameters.Append .CreateParameter("@FaxExt", adChar, adParamInput, 4, m_udtContact.FaxExt)
        .Parameters.Append .CreateParameter("@Name", adVarChar, adParamInput, 40, m_udtContact.FirstName + " " + m_udtContact.LastName)
        .Parameters.Append .CreateParameter("@Phone", adChar, adParamInput, 17, m_udtContact.Phone)
        .Parameters.Append .CreateParameter("@PhoneExt", adChar, adParamInput, 4, m_udtContact.PhoneExt)
        .Parameters.Append .CreateParameter("@CellPhone", adChar, adParamInput, 17, m_udtContact.CellPhone)
        .Parameters.Append .CreateParameter("@UpdateUserID", adChar, adParamInput, 30, m_sUserID)
        .Parameters.Append .CreateParameter("@DeclinedEmailAddr", adSmallInt, adParamInput, , m_udtContact.DeclinedEmailAddr)
        .Parameters.Append .CreateParameter("@FirstName", adVarChar, adParamInput, 30, m_udtContact.FirstName)
        .Parameters.Append .CreateParameter("@LastName", adVarChar, adParamInput, 30, m_udtContact.LastName)
        .Execute
        RetVal = .Parameters("RETURN_VALUE").value
    End With

    If RetVal = 0 Then
        Err.Raise 1, "Update", "This contact name already exists."
        Exit Sub
    End If

    m_oNotifications.Update m_udtContact.CntctKey
    m_udtContact.UpdateDate = Date
    m_udtContact.UpdateUserID = m_sUserID
    
    If m_udtContact.FirstName <> m_udtContact_Backup.FirstName Then
        RaiseEvent NameChange
    End If
    
    m_oState.ClearBits eMask.IsDirty
    
    Backup
End Sub


Public Sub Delete(Optional bFullDelete As Boolean = False)
    Dim oCmd As ADODB.Command
    
    m_oNotifications.Delete m_udtContact.CntctKey

    If bFullDelete Then
        Set oCmd = CreateCommandSP("DELETE FROM tciContact WHERE cntctkey=" & m_udtContact.CntctKey, adCmdText)
    Else
        Set oCmd = CreateCommandSP("UPDATE tciContact SET Deleted=-1, UpdateDate=getdate(), UpdateUserID ='" _
        & m_sUserID & "' WHERE cntctkey=" & m_udtContact.CntctKey, adCmdText)
    End If
    
    oCmd.Execute
    Set oCmd = Nothing
End Sub


Public Sub Clear()
    With m_udtContact
        .CntctKey = 0
        .CntctOwnerKey = 0
        .EMailFormat = 3    'default
        .emailaddr = vbNullString
        .Fax = vbNullString
        .FaxExt = vbNullString
        .Phone = vbNullString
        .PhoneExt = vbNullString
        .CellPhone = vbNullString
        .Title = vbNullString
        .DeclinedEmailAddr = False
        .FirstName = vbNullString
        .LastName = vbNullString
        .UpdateDate = Empty
    End With
       
    m_oNotifications.Clear

    m_oState.ClearBits eMask.IsDirty
    m_oState.SetBits eMask.IsNew
End Sub


'Event Handlers

Private Sub m_oState_StateChange(NewState As BitMap)
    RaiseEvent StateChange(NewState)
End Sub


'***466 Modified
' Note: bSplitOrder is optional, but has no default value
' does it default to False?
' optional parameters, in order to work with IsMissing, need to be typed as variant

Public Function Export(Optional bSplitOrder As Boolean = False) As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Dim oXMLchild As JDMPDXML.XMLNode
    Dim i As Integer
    
    Set oXML = New JDMPDXML.XMLNode
    Set oXMLchild = New JDMPDXML.XMLNode
    
    With oXML
        .Tag = "Contact"
        .Add "FullName", m_udtContact.FirstName & " " & m_udtContact.LastName
        .Add "FirstName", m_udtContact.FirstName
        .Add "LastName", m_udtContact.LastName
        
        If Len(m_udtContact.Fax) > 0 Then
            .Add "FaxNbr", m_udtContact.Fax
            If Len(m_udtContact.FaxExt) > 0 Then
                .Add "FaxExt", m_udtContact.FaxExt
            End If
        End If
        
        If Len(m_udtContact.Phone) > 0 Then
            .Add "PhoneNbr", m_udtContact.Phone
            If Len(m_udtContact.PhoneExt) > 0 Then
                .Add "PhoneExt", m_udtContact.PhoneExt
            End If
        End If
        
        If Len(m_udtContact.CellPhone) > 0 Then
            .Add "CellPhone", m_udtContact.CellPhone
        End If
        
        If bSplitOrder Then
            .Add "CntctKey", m_udtContact.CntctKey
            .Add "CntctOwnerKey", m_udtContact.CntctOwnerKey
            .Add "CreateUserID", m_udtContact.CreateUserID
            .Add "CreateDate", m_udtContact.CreateDate
            .Add "UpdateUserID", m_udtContact.UpdateUserID
            .Add "UpdateDate", m_udtContact.UpdateDate
            .Add "EntityType", m_udtContact.EntityType
            .Add "EMailAddr", m_udtContact.emailaddr
            .Add "EMailFormat", m_udtContact.EMailFormat
            .Add "CellPhone", m_udtContact.CellPhone
            .Add "Title", m_udtContact.Title
            .Add "DeclinedEmailAddr", m_udtContact.DeclinedEmailAddr
            .Add "State", m_oState.value
            oXML.AddChild m_oNotifications.Export()
        End If
    End With
            
    Set Export = oXML
End Function




Public Function XmlCommit() As JDMPDXML.XMLNode
    Dim oXML As JDMPDXML.XMLNode
    Set oXML = New JDMPDXML.XMLNode
    
    With oXML
        .Tag = "contact"
        .Add "key", m_udtContact.CntctKey
        .Add "ownertype", m_udtContact.EntityType
    End With
            
    Set XmlCommit = oXML
End Function


Public Function Serialize() As IXMLDOMElement
    Dim dom As New DOMDocument
    Dim element As IXMLDOMElement
    Dim property As IXMLDOMElement

    Set element = dom.createElement("contact")
    
    Set property = dom.createElement("key")
    element.appendChild property
    property.text = m_udtContact.CntctKey
    
    Set property = dom.createElement("ownertype")
    element.appendChild property
    property.text = m_udtContact.EntityType
        
    Set Serialize = element
End Function


'***466 Added
' Note: 4/23/08 LR we need to review this

Public Sub Import(ByVal i_oXML As JDMPDXML.XMLNode)
    Dim oSubXMLNode As JDMPDXML.XMLNode
    Dim oSubSubXMLNode As JDMPDXML.XMLNode
    Dim i As Integer
    
    Set oSubXMLNode = New JDMPDXML.XMLNode
    Set oSubSubXMLNode = New JDMPDXML.XMLNode
    
    With i_oXML
        m_udtContact.CntctKey = .Item("CntctKey").value
        m_udtContact.CntctOwnerKey = .Item("CntctOwnerKey").value
        m_udtContact.FirstName = Trim$(.Item("FirstName").value)
        m_udtContact.LastName = Trim$(.Item("LastName").value)
        m_udtContact.CreateUserID = Trim$(.Item("CreateUserID").value)
        m_udtContact.CreateDate = .Item("CreateDate").value
        m_udtContact.UpdateUserID = Trim$(.Item("UpdateUserID").value)
        m_udtContact.UpdateDate = .Item("UpdateDate").value
        m_udtContact.EntityType = .Item("EntityType").value
        m_udtContact.emailaddr = Trim$(.Item("EMailAddr").value)
        m_udtContact.EMailFormat = .Item("EMailFormat").value
        m_udtContact.Fax = Trim$(.Item("PhoneExt").value)
        m_udtContact.FaxExt = Trim$(.Item("FaxExt").value)
        m_udtContact.Phone = Trim$(.Item("PhoneNbr").value)
        m_udtContact.PhoneExt = Trim$(.Item("PhoneExt").value)
        m_udtContact.CellPhone = Trim$(.Item("CellPhone").value)
        m_udtContact.Title = Trim$(.Item("Title").value)
        m_udtContact.DeclinedEmailAddr = .Item("DeclinedEmailAddr").value
        m_oState.value = .Item("State").value
        m_oNotifications.Import i_oXML
    End With
    
End Sub


Public Function Edit(i_sUserID As String, Optional sName As String = "", Optional oType As OwnerType = opNone, Optional ByVal lOwnerKey As Long = 0) As VbMsgBoxResult
    Dim frm As New FContactMgr
    Dim sOperation As String

    'This variable is set on entry and used later on updates and inserts.
    m_sUserID = i_sUserID
    
    'New Contact
    If m_udtContact.CntctKey = 0 Then
        If oType = opNone Then
            'Raise error
            Debug.Print "no ownertype provided"
            Exit Function
        End If
        If lOwnerKey = 0 Then
            'Raise error
            Debug.Print "no ownerkey provided"
            Exit Function
        End If
    End If

    frm.Edit Me, sName, oType, lOwnerKey
    frm.Show vbModal

    If frm.Cancel Then
        Edit = VbMsgBoxResult.vbCancel
    Else
        Edit = VbMsgBoxResult.vbOK
    End If
    
    Unload frm
End Function


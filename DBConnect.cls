VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sServer As String
Private m_sDatabase As String
Private m_sUserID As String
Private m_sPassword As String

Private m_bInitialized As Boolean

Private m_GlobalConfig As Configuration
Private m_UserConfig As Configuration

Private m_AppCn As ADODB.Connection


Public Property Get Connection() As ADODB.Connection
    Set Connection = m_AppCn
End Property


Public Property Get GlobalConfig() As Configuration
    Set GlobalConfig = m_GlobalConfig
End Property

Public Property Let GlobalConfig(ByRef config As Configuration)
    Set m_GlobalConfig = config
End Property


Public Property Get UserConfig() As Configuration
    Set UserConfig = m_UserConfig
End Property

Public Property Let UserConfig(ByRef config As Configuration)
    Set m_UserConfig = config
End Property


Public Property Get ConnectionString() As String
    ConnectionString = "Provider=SQLOLEDB.1;" _
                     & "Data Source=" & Server & "; " _
                     & "Initial Catalog=" & Database & "; " _
                     & "User ID=" & m_sUserID & "; " _
                     & "Password=" & m_sPassword
End Property


Public Function Connect() As Boolean
    On Error GoTo ErrorHandler
    
    Set m_AppCn = OpenDatabase
        
    Connect = True
    Exit Function

ErrorHandler:
    Connect = False
End Function


Public Sub Disconnect()
    On Error GoTo ErrorHandler

    m_AppCn.Close
    Set m_AppCn = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise -1, "Database.Disconnect", "Unable to close connection to database(s)."
End Sub


Public Function IsDevelopment() As Boolean
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xpathquery As String

    xpathquery = "/configuration/selectedserver/add"
    Set xNode = m_UserConfig.DOM.SelectSingleNode(xpathquery)

    If xNode Is Nothing Then
        IsDevelopment = False   'this is an end-user whose user confgi file is missing this node
    Else
        IsDevelopment = "development" = LCase(xNode.Attributes(0).NodeValue)
    End If
End Function


Public Function Database() As String
    Initialize
    Database = m_sDatabase
End Function


Public Function Server() As String
    Initialize
    Server = m_sServer
End Function


Public Function Password() As String
    Initialize
    Password = m_sPassword
End Function


Public Sub Reset()
    m_bInitialized = False
    Initialize
End Sub


Private Function OpenDatabase()
    Dim Cn As ADODB.Connection

    On Error GoTo ErrorHandler
    
    Set Cn = New ADODB.Connection
    With Cn
        .ConnectionString = ConnectionString
        .CursorLocation = adUseClient
        .Open
    End With
    Set OpenDatabase = Cn
    Exit Function

ErrorHandler:
    Dim sMsg As String
    Dim obj As Object
    
    sMsg = "Could not open database " & g_DB.Server & ":" & g_DB.Database _
         & vbCrLf & Err.Description
         
    Err.Raise -1, "Database.OpenDatabase", sMsg
   
    If Err.Number = 0 Then
        sMsg = "Unknown error in Connect"
    Else
        sMsg = Err.Description
    End If

    If Cn.Errors.Count > 0 Then
        Dim i As Long
        sMsg = sMsg & vbCrLf & "Additional Information:" & vbCrLf
        For Each obj In Cn.Errors
            sMsg = sMsg & "ErrMsg " & i & ": " & obj.Description & vbCrLf
        Next
    End If
    
    Err.Raise -1, "Database.OpenDatabase", sMsg
End Function


Private Sub Initialize()
    If m_bInitialized Then Exit Sub
    
    Dim nodelist As MSXML2.IXMLDOMNodeList
    Dim xAttrib As MSXML2.IXMLDOMAttribute
    Dim xpathquery As String
    
    xpathquery = "//databaseConnections/add[@name='sage']"
    Set nodelist = m_GlobalConfig.DOM.documentElement.SelectNodes(xpathquery)
    For Each xAttrib In nodelist.Item(0).Attributes
        Select Case xAttrib.Name
            Case "server":
                m_sServer = xAttrib.value
            Case "database"
                m_sDatabase = xAttrib.value
            Case "userid"
                m_sUserID = xAttrib.value
            Case "password"
                m_sPassword = xAttrib.value
        End Select
    Next xAttrib
    
    m_bInitialized = True
End Sub


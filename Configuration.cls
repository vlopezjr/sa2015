VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_filepath As String
Private m_targetserver As String
Private m_document As MSXML2.DOMDocument

Private m_localpathname As String
Private m_remotepathname As String

Private m_localconfigfound As Boolean


Public Property Get LocalConfigFound() As Boolean
    LocalConfigFound = m_localconfigfound
End Property


Public Property Get LocalPathname() As String
    LocalPathname = m_localpathname
End Property

Public Property Let LocalPathname(ByVal path As String)
    m_localpathname = path
End Property


Public Property Get RemotePathname() As String
    RemotePathname = m_remotepathname
End Property

Public Property Let RemotePathname(ByVal path As String)
    m_remotepathname = path
End Property


Public Property Get FilePath() As String
    FilePath = m_filepath
End Property

Public Property Let FilePath(ByVal path As String)
    m_filepath = path
End Property


Public Property Get TargetServer() As String
    TargetServer = m_targetserver
End Property

Public Property Let TargetServer(ByVal name As String)
    m_targetserver = name
End Property


Public Property Get DOM() As MSXML2.DOMDocument
    Set DOM = m_document
End Property


Public Sub Load(username As String)
    Set m_document = New MSXML2.DOMDocument
    m_document.async = False
    
    m_filepath = GetConfigFileLocation(username)
    If Not m_document.Load(m_filepath) Then
        Err.Raise vbObjectError + 513 + 1002, "LoadXmlConfigFile", "XML Parser Error. Call computer guys."
    End If
End Sub


Public Sub Save()
    m_document.Save m_filepath
End Sub


Public Function GetKeyValue(section As String, key As String) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xAttrib As MSXML2.IXMLDOMAttribute
    Dim xpathquery As String

    xpathquery = "/configuration/" & section & "/add[@key='" & key & "']"
    Set xNode = m_document.documentElement.SelectSingleNode(xpathquery)

    For Each xAttrib In xNode.Attributes
        If xAttrib.name = "value" Then
            GetKeyValue = xAttrib.value
        End If
    Next xAttrib
End Function


Public Function SetKeyValue(section As String, key As String, ByVal value As String) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xAttrib As MSXML2.IXMLDOMAttribute
    Dim xpathquery As String

    xpathquery = "/configuration/" & section & "/add[@key='" & key & "']"
    Set xNode = m_document.documentElement.SelectSingleNode(xpathquery)
    
    For Each xAttrib In xNode.Attributes
        If xAttrib.name = "value" Then
            xAttrib.value = value
        End If
    Next xAttrib
End Function


Public Function GetServerList() As Dictionary
    Dim xGroupNode As MSXML2.IXMLDOMNode
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xAttrib As MSXML2.IXMLDOMAttribute
    Dim caption As String
    Dim name As String
    
    Dim list As Dictionary
    
    Set list = New Dictionary
    
    Set xGroupNode = m_document.documentElement.SelectSingleNode("/configuration/servers")
    For Each xNode In xGroupNode.ChildNodes
        Set xAttrib = xNode.Attributes.Item(0)
        caption = xAttrib.value
        Set xAttrib = xNode.Attributes.Item(1)
        name = xAttrib.value
        list.Add caption, name
    Next
    Set GetServerList = list
End Function


'check in the app folder first to support development (typically c:\WS\<projectfolder>)
'if config file isn't here, go to the server (WEB15) - generalize this for developers

Private Function GetConfigFileLocation(username As String) As String
    Dim ofs As FileSystemObject
    Set ofs = New FileSystemObject
    
    Dim configFileName As String
    configFileName = username & ".xml"
    
    m_localpathname = App.path & "\" & configFileName
    m_remotepathname = "\\" & m_targetserver & "\officeassistant\configfiles\" & configFileName
    
    If ofs.FileExists(m_localpathname) Then
        GetConfigFileLocation = m_localpathname
        m_localconfigfound = True
    ElseIf ofs.FileExists(m_remotepathname) Then
        GetConfigFileLocation = m_remotepathname
        m_localconfigfound = False
    Else
        Err.Raise vbObjectError + 513 + 1001, "GetConfigFileLocation", "Missing config.ini file. Call computer guys."
    End If
    Set ofs = Nothing
End Function


'this uses <username>.xml

Public Function GetControlRev(section As String, name As String) As Integer

End Function

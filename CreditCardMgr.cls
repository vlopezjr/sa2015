VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreditCardMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colCreditCards As Collection
Private m_oPrevCC As CreditCard
Private WithEvents m_oSelCC As CreditCard
Attribute m_oSelCC.VB_VarHelpID = -1

Private m_lCustKey As Long

Private m_orstCardTypes As ADODB.Recordset


Public Event StateChange(newstate As BitMap)


'should this be read-only?
Public Property Get PrevCC() As CreditCard
    Set PrevCC = m_oPrevCC
End Property

Public Property Let PrevCC(value As CreditCard)
    Set m_oPrevCC = value
End Property


'should this be read-only?
Public Property Get SelCC() As CreditCard
    Set SelCC = m_oSelCC
End Property

Public Property Let SelCC(value As CreditCard)
    Set m_oSelCC = value
End Property


Public Property Get CreditCards() As Collection
    Set CreditCards = m_colCreditCards
End Property


Public Sub SelectCreditCard(Key As String)
    Set m_oSelCC = m_colCreditCards(CStr(Key))
End Sub


' Class Event Handlers

Private Sub Class_Initialize()
    Set m_colCreditCards = New Collection
    InitCardTypes
End Sub


Private Sub Class_Terminate()
    Set m_orstCardTypes = Nothing
    Set m_colCreditCards = Nothing
    Set m_oPrevCC = Nothing
    Set m_oSelCC = Nothing
End Sub


' Object Event Handlers

'bubble up the CreditCard state change event

Private Sub m_oSelCC_StateChange(newstate As BitMap)
    RaiseEvent StateChange(newstate)
End Sub


'Note: there are two ways to determine if this fails to load any credit cards
'1. collection.count = 0 or
'2. SelCC is nothing

'why do we pass in CustKey & CustID?
'If CustID is used simply to determine is the Cust is MISC,
'why not pass in a flag instead indicating MISC?

Public Sub LoadCreditCards(oCust As Customer, PrevCC As CreditCard, Optional oOrder As Variant)
    
    m_lCustKey = oCust.Key
    Set m_oPrevCC = PrevCC
    
    'If InStr(1, CustID, "-MISC") > 0 Or CustKey = 0 Then
    If oCust.HasAccount Then
        If IsMissing(oOrder) Then
            LoadAcctCC
        Else
            LoadAcctCC oOrder
        End If
    Else
        LoadNonAcctCC
    End If
    
End Sub


Private Sub LoadNonAcctCC()
    'if the order has a previously assigned CC, put it in our collection
    'and make it the selected CC
    If Not (m_oPrevCC Is Nothing) Then
        m_colCreditCards.Add m_oPrevCC, CStr(m_oPrevCC.Key)
        Set m_oSelCC = m_oPrevCC
    End If
    'otherwise, what's the state of the collection and m_oSelCC?
End Sub


'if the query is empty, the resulting collection.count = 0

Private Sub LoadAcctCC(Optional oOrder As Variant)
    Dim oCC As CreditCard
    Dim oRst As ADODB.Recordset
    Dim sSQL As String
    Dim lKey As Long

    Set oRst = New ADODB.Recordset
    sSQL = "SELECT CCKey FROM tcpCreditCard WHERE CustKey = " & m_lCustKey & " and status = " & 1
    oRst.Open sSQL, g_DB.Connection, adOpenStatic, adLockReadOnly

    If Not oRst.EOF Then
        While Not oRst.EOF
            Set oCC = New CreditCard
            If IsMissing(oOrder) Then
                oCC.Load oRst!CCKey
            Else
                oCC.Load oRst!CCKey, oOrder
            End If

            If m_oPrevCC Is Nothing Then
                m_colCreditCards.Add oCC, CStr(oCC.Key)
                If oCC.Preferred = True Then
                    Set m_oSelCC = m_colCreditCards(CStr(oCC.Key))
                End If
            ElseIf oCC.Key <> m_oPrevCC.Key Then
                m_colCreditCards.Add oCC, CStr(oCC.Key)
                If oCC.Preferred = True Then
                    Set m_oSelCC = m_colCreditCards(CStr(oCC.Key))
                End If
            ElseIf oCC.Key = m_oPrevCC.Key Then
                'add the Previous Credit Card
                m_colCreditCards.Add m_oPrevCC, CStr(m_oPrevCC.Key)
            End If

            Set oCC = Nothing

            oRst.MoveNext
        Wend

        'If a Prev Card exist, set the Sel CC to the Prev CC
        'If Not (m_oPrevCC Is Nothing) Then Set m_oSelCC = m_oPrevCC
        If Not (m_oPrevCC Is Nothing) Then
            Set m_oSelCC = m_colCreditCards(CStr(m_oPrevCC.Key))
            'couldn't you just say
            'Set m_oSelCC = m_oPrevCC ?
        End If

        'catch all
        If m_oSelCC Is Nothing Then
            Set m_oSelCC = m_colCreditCards(1)
        End If
    End If

    oRst.Close
    Set oRst = Nothing
    
End Sub


'Note: we're not placing this in the collection
'When is this invoked?

Public Sub NewCreditCard()
    Set m_oSelCC = New CreditCard
    m_oSelCC.Init
End Sub


Public Sub CancelChanges()
    m_oSelCC.Restore
End Sub


Public Sub DeleteCreditCard()
    m_oSelCC.Delete m_lCustKey
    m_colCreditCards.Remove CStr(m_oSelCC.Key)
End Sub


Public Sub AddUpdateCreditCard()
    If m_oSelCC.IsNew Then
        m_oSelCC.Add (m_lCustKey)
        m_colCreditCards.Add m_oSelCC, CStr(m_oSelCC.Key)
    Else
        m_oSelCC.Update
    End If
End Sub


Public Sub DupCreditCardNo(DupCCInfo As String)
    Dim oRst As ADODB.Recordset
    Dim sSQL As String
    
    Set oRst = New ADODB.Recordset
    sSQL = "SELECT * from tcpCreditCard where CrCardNo = '" & m_oSelCC.CardNo & "' and "
    sSQL = sSQL & "CCKey <> " & m_oSelCC.Key
    oRst.Open sSQL, g_DB.Connection
    
    If oRst.EOF = True Then
        DupCCInfo = ""
    Else
        Do Until oRst.EOF = True
            DupCCInfo = DupCCInfo & vbTab & "- " & Trim$(oRst!CardHolderName)
            If oRst!Status = 1 Then
                DupCCInfo = DupCCInfo & " (active)"
            Else
                DupCCInfo = DupCCInfo & " (deleted)"
            End If
            oRst.MoveNext
            If oRst.EOF = False Then DupCCInfo = DupCCInfo & chr(10) & chr(13)
        Loop
    End If
    
    oRst.Close
    Set oRst = Nothing
End Sub


'***********************************************************************************
'
' Card Type 'object'
' These methods operate on the disconnected recordset m_orstCardTypes
'
'***********************************************************************************

Private Sub InitCardTypes()
    Dim sSQL As String
    
    Set m_orstCardTypes = New ADODB.Recordset
    sSQL = "SELECT * from tcpCreditCardType"
    m_orstCardTypes.Open sSQL, g_DB.Connection, adOpenStatic, adLockReadOnly
    Set m_orstCardTypes.ActiveConnection = Nothing
End Sub


Public Function GetCardTypeMask(ByVal CardTypeKey As Long) As String
    m_orstCardTypes.MoveFirst
    Do Until m_orstCardTypes.EOF
        If m_orstCardTypes!CrCardTypeKey = CardTypeKey Then
            GetCardTypeMask = m_orstCardTypes!CrCardTypeMask
            Exit Function
        End If
        m_orstCardTypes.MoveNext
    Loop
End Function


Public Sub GetCardTypeDescr(ByVal CardTypeKey As Long, ByRef CardTypeDescr As String)
    m_orstCardTypes.MoveFirst
    If m_orstCardTypes.EOF = False Then
        Do Until m_orstCardTypes.EOF = True
            If m_orstCardTypes!CrCardTypeKey = CardTypeKey Then
                CardTypeDescr = m_orstCardTypes!CrCardTypeName
                Exit Sub
            End If
            m_orstCardTypes.MoveNext
        Loop
    End If
End Sub


Public Sub LoadCardTypeList(cbo As ComboBox)
    Dim llCounter As Long

    'load cbo using the private disconnected RST
    cbo.Clear
    m_orstCardTypes.MoveFirst
    If m_orstCardTypes.EOF = False Then
        Do Until m_orstCardTypes.EOF = True
            cbo.AddItem m_orstCardTypes!CrCardTypeName
            cbo.ItemData(cbo.NewIndex) = m_orstCardTypes!CrCardTypeKey
            m_orstCardTypes.MoveNext
        Loop
    End If
    
    'set SelCC CardTypeKey
    For llCounter = 0 To cbo.ListCount - 1
        If cbo.ItemData(llCounter) = m_oSelCC.TypeKey Then
            cbo.ListIndex = llCounter
            Exit Sub
        End If
    Next
End Sub


'***466 SMR 05/23/2006

Public Function Export() As JDMPDXML.XMLNode
    Dim i_oXML As JDMPDXML.XMLNode
    Dim i As Integer
    Dim oCCard As CreditCard
    
    Set i_oXML = New JDMPDXML.XMLNode
    
    With i_oXML
        .Tag = "CreditCards"
        .Add "SelCCardKey", m_oSelCC.Key
        
        If Not m_oPrevCC Is Nothing Then
            .Add "PrevCCardKey", m_oPrevCC.Key
        Else
            .Add "PrevCCardKey", 0
        End If
        
        .Add "CustKey", m_lCustKey
        
        For i = 1 To m_colCreditCards.Count
            Set oCCard = m_colCreditCards.Item(i)
            .AddChild oCCard.Export
        Next
    End With

    Set Export = i_oXML

End Function


'***466 SMR 05/23/2006

Public Function Import(ByVal i_oXML As JDMPDXML.XMLNode)

    Dim lSelCCardkey  As Long
    Dim lPrevCCardkey As Long
    Dim oSubXMLNode As JDMPDXML.XMLNode
    Dim oCCard As CreditCard
    
    With i_oXML
        If .Tag = "CreditCards" Then
            lSelCCardkey = .Item("SelCCardKey").value
            lPrevCCardkey = .Item("PrevCCardKey").value
            m_lCustKey = .Item("CustKey").value
        
            'Each Credit Card
            For Each oSubXMLNode In i_oXML
                If oSubXMLNode.Tag = "CreditCard" Then
                    
                    Set oCCard = New CreditCard
                    
                    oCCard.Import oSubXMLNode
                               
                    m_colCreditCards.Add oCCard, CStr(oCCard.Key)
                    
                    If oCCard.Key = lSelCCardkey Then
                        Set m_oSelCC = m_colCreditCards(CStr(oCCard.Key))
                    End If

                    If oCCard.Key = lPrevCCardkey Then
                        Set m_oPrevCC = m_colCreditCards(CStr(oCCard.Key))
                    End If
                End If
            Next
        End If
        
    End With

End Function


Public Sub PrintCCardXML()
    Dim oXML As JDMPDXML.XMLNode
    
    Set oXML = Export
    oXML.ExportFile "c:\XMLCCardsTest" & Format(Now, "ddyy nnss") & ".txt"
    
End Sub



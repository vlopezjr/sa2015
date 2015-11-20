VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sInput As String
Private m_sDelimiter As String
Private m_lCharPos As Long


' used by modError

Public Property Get ParseString() As Variant
    ParseString = m_sInput
End Property

Public Property Let ParseString(ByVal i_sInput As Variant)
    m_sInput = i_sInput
    m_lCharPos = 1
End Property


Public Property Get Delimiter() As Variant
    Delimiter = m_sDelimiter
End Property

Public Property Let Delimiter(ByVal i_sDelimiter As Variant)
    m_sDelimiter = i_sDelimiter
End Property



Public Sub Reset()
    m_lCharPos = 1
End Sub


Public Function Done() As Boolean
    If m_lCharPos > Len(m_sInput) Then
        Done = True
    End If
End Function


' called by functions in modError

Public Function GetNextToken() As String
    Dim lCharPos As Long
    Dim lTokenLen As Long
    
    lCharPos = InStr(m_lCharPos, m_sInput, m_sDelimiter, vbBinaryCompare)
    If lCharPos > 0 Then
        lTokenLen = lCharPos - m_lCharPos
        If lTokenLen > 0 Then
            GetNextToken = Mid$(m_sInput, m_lCharPos, lTokenLen)
            m_lCharPos = lCharPos + Len(m_sDelimiter)
        End If
    Else
        GetNextToken = Mid$(m_sInput, m_lCharPos)
        m_lCharPos = Len(m_sInput) + 1
    End If
End Function

' not called

Public Function GetToken(ByVal i_lIndex As Long) As String
    Dim lIndex As Long
    Dim sToken As String

    Reset
    sToken = GetNextToken
    lIndex = 1
    While i_lIndex > lIndex And Len(sToken) > 0
        sToken = GetNextToken
        lIndex = lIndex + 1
    Wend

    If i_lIndex = lIndex Then
        GetToken = sToken
    End If
End Function

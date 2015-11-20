VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lNumber As Long
Private m_sSource As String
Private m_sDescription As String


Public Property Get Number() As Long
    Number = m_lNumber
End Property

Public Property Let Number(ByVal i_lNumber As Long)
    m_lNumber = i_lNumber
End Property


Public Property Get Source() As String
    If Len(Trim(m_sSource)) = 0 Then
        Source = "<No error source specified>"
    Else
        Source = m_sSource
    End If
End Property

Public Property Let Source(ByVal i_sSource As String)
    m_sSource = i_sSource
End Property


Public Property Get Description() As String
    If Len(Trim(m_sDescription)) = 0 Then
        Description = "No error description available"
    Else
        Description = m_sDescription
    End If
End Property

Public Property Let Description(ByVal i_sDescription As String)
    m_sDescription = i_sDescription
End Property



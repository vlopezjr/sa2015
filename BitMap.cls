VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event StateChange(NewState As BitMap)

Private m_prevbitmap As Integer
Private m_bitmap As Integer

'***DH 2/20/09 This is meant to replace the constants in modGlobal
Public Enum eMask
    IsNew = &H1
    IsValid = &H2
    IsDirty = &H4
End Enum


Public Function SetBits(mask As Integer) As Integer
    m_prevbitmap = m_bitmap
    m_bitmap = m_bitmap Or mask
    SetBits = m_bitmap
    If m_bitmap <> m_prevbitmap Then RaiseEvent StateChange(Me)
End Function


Public Function ClearBits(mask As Integer) As Integer
    m_prevbitmap = m_bitmap
    m_bitmap = m_bitmap And (Not mask)
    ClearBits = m_bitmap
    If m_bitmap <> m_prevbitmap Then RaiseEvent StateChange(Me)
End Function


Public Function TestBits(mask As Integer) As Boolean
Attribute TestBits.VB_UserMemId = 0
    TestBits = ((m_bitmap And mask) = mask)
End Function


Public Function ToString() As String
    ToString = CStr(m_bitmap)
End Function


Public Property Get value() As Integer
    value = m_bitmap
End Property

Public Property Let value(ByVal iNewValue As Integer)
    m_prevbitmap = m_bitmap
    m_bitmap = iNewValue
    If m_bitmap <> m_prevbitmap Then RaiseEvent StateChange(Me)
End Property

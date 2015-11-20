VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "State"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oState As BitMap
Public Event StateChange()


'****
'Init
'****
Private Sub Class_Initialize()
    Set m_oState = New BitMap
End Sub


'**********
'Properties
'**********
Public Property Get IsNew() As Boolean
    IsNew = m_oState.TestBits(eMask.IsNew)
End Property

Public Property Let IsNew(bValue As Boolean)
    If bValue Then
        m_oState.SetBits (eMask.IsNew)
    Else
        m_oState.ClearBits (eMask.IsNew)
    End If
End Property


Public Property Get IsValid() As Boolean
    IsValid = m_oState.TestBits(eMask.IsValid)
End Property

Public Property Let IsValid(bValue As Boolean)
    If bValue Then
        m_oState.SetBits (eMask.IsValid)
    Else
        m_oState.ClearBits (eMask.IsValid)
    End If
End Property


Public Property Get IsDirty() As Boolean
    IsDirty = m_oState.TestBits(eMask.IsDirty)
End Property

Public Property Let IsDirty(bValue As Boolean)
    If bValue Then
        m_oState.SetBits (eMask.IsDirty)
    Else
        m_oState.ClearBits (eMask.IsDirty)
    End If
End Property


'*************
'Event Handler
'*************
Private Sub m_oState_StateChange(newstate As BitMap)
    RaiseEvent StateChange
End Sub

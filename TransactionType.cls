VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransactionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PRIVATE VARIABLES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_iTrxTypeKey As Long
Private m_sDescr As String
Private m_bIsDebit As Boolean



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PROPERITES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get TrxTypeKey() As Long
    TrxTypeKey = m_iTrxTypeKey
End Property

Public Property Let TrxTypeKey(value As Long)
    m_iTrxTypeKey = value
End Property


Public Property Get Descr() As String
    Descr = m_sDescr
End Property

Public Property Let Descr(value As String)
    m_sDescr = value
End Property


Public Property Get IsDebit() As Boolean
    IsDebit = m_bIsDebit
End Property

Public Property Let IsDebit(value As Boolean)
    m_bIsDebit = value
End Property



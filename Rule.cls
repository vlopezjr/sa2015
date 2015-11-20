VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is part of Petty Cashier

Dim m_iRuleKey As Long
Dim m_sRuleDescr As String


Public Property Get Key() As Long
    Key = m_iRuleKey
End Property

Public Property Let Key(ByVal value As Long)
    m_iRuleKey = value
End Property


Public Property Get Descr() As String
    Descr = m_sRuleDescr
End Property

Public Property Let Descr(ByVal value As String)
    m_sRuleDescr = value
End Property

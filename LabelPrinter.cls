VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabelPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'NOTES:
' Label printers must have the word "Zebra" in their device name.
' This class works in conjunction with FNumLabels.frm.

Private Const DEVICE_NAME = "Zebra"

Private Const TWIPS_PER_INCH = 1440

Private m_LabelHeight As Double
Private m_LabelWidth As Double
Private m_LabelTop As Double
Private m_LabelLeft As Double
Private m_FontSize As Integer
Private m_FontBold As Boolean

Private m_sText() As String
Private m_iNumLabels As Integer


Public Property Get NumLabels() As Integer
    NumLabels = m_iNumLabels
End Property

Public Property Let NumLabels(ByVal iNewValue As Integer)
    m_iNumLabels = iNewValue
End Property


Public Property Let LabelHeight(ByVal dNewValue As Double)
    m_LabelHeight = dNewValue
End Property


Public Property Let LabelWidth(ByVal dNewValue As Double)
    m_LabelWidth = dNewValue
End Property


Public Property Let LabelTop(ByVal dNewValue As Double)
    m_LabelTop = dNewValue
End Property


Public Property Let LabelLeft(ByVal dNewValue As Double)
    m_LabelLeft = dNewValue
End Property

Public Property Let FontSize(ByVal iNewValue As Integer)
    m_FontSize = iNewValue
End Property

Public Property Let FontBold(ByVal bNewValue As Boolean)
    m_FontBold = bNewValue
End Property


'for testing
Public Property Get Label() As String
    Dim i As Integer
    For i = 1 To UBound(m_sText)
        Label = Label & m_sText(i) & vbCrLf
    Next i
End Property

Public Sub AddLine(s As String)
    Dim Index As Integer
    Index = UBound(m_sText) + 1
    ReDim Preserve m_sText(Index)
    m_sText(Index) = s
End Sub

Private Sub Class_Initialize()
    ReDim m_sText(0)

    m_LabelHeight = 1#
    m_LabelWidth = 4#
    m_LabelTop = 0.1
    m_LabelLeft = 0.75
    m_FontSize = 10
    m_FontBold = True
End Sub

Public Sub Clear()
    ReDim m_sText(0)
End Sub


Public Sub PrintLabel()
    Dim p As Printer
    Dim i As Integer
    Dim j As Integer
    Dim oFrm As FNumLabels
    Dim NumLabels As Integer

    Set oFrm = New FNumLabels
    NumLabels = oFrm.NumLabels(m_iNumLabels)
    Unload oFrm
    Set oFrm = Nothing
    
    If NumLabels > 0 Then
    
        For Each p In Printers
            If InStr(1, p.DeviceName, DEVICE_NAME) Then
                Set Printer = p
                Exit For
            End If
        Next
        
        If InStr(1, Printer.DeviceName, DEVICE_NAME) Then
            Debug.Print "Current: " & Printer.DeviceName
            Printer.FontSize = m_FontSize
            Printer.FontBold = m_FontBold

            For i = 1 To NumLabels
                Printer.width = m_LabelWidth * TWIPS_PER_INCH
                Printer.Height = m_LabelHeight * TWIPS_PER_INCH
                Printer.CurrentY = m_LabelTop * TWIPS_PER_INCH
                For j = 1 To UBound(m_sText)
                    Debug.Print m_sText(j)
                    Printer.CurrentX = m_LabelLeft * TWIPS_PER_INCH
                    Printer.Print m_sText(j)
                Next j
                Printer.EndDoc
            Next i
        Else
            Err.Raise -1, "LabelPrinter", "No Zebra Printer found"
        End If
        
    End If
    
    ReDim m_sText(0)
    
End Sub



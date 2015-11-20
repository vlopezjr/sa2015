VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FViewOrder 
   Caption         =   "View Order"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   10710
   Begin VB.CheckBox chkWarehouse 
      Caption         =   "Warehouse"
      Height          =   252
      Left            =   4920
      TabIndex        =   11
      Top             =   450
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkShowRemarks 
      Caption         =   "Manufacturing"
      Height          =   252
      Index           =   3
      Left            =   3390
      TabIndex        =   9
      Top             =   450
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton optCustomerReceipt 
      Caption         =   "Customer Receipt"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optCPCReceipt 
      Caption         =   "CPC Receipt"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton optReceipt 
      Caption         =   "2-Copy Receipts"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optCPC 
      Caption         =   "Order Snapshot"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkShowRemarks 
      Caption         =   "Purchasing"
      Height          =   252
      Index           =   2
      Left            =   2100
      TabIndex        =   8
      Top             =   450
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkShowRemarks 
      Caption         =   "Private"
      Height          =   252
      Index           =   1
      Left            =   1050
      TabIndex        =   7
      Top             =   450
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkShowRemarks 
      Caption         =   "Public"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   450
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   9480
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   3975
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   10665
      ExtentX         =   18812
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CheckBox chkNoPricing 
      Caption         =   "No Pricing"
      Height          =   252
      Left            =   2100
      TabIndex        =   5
      Top             =   450
      Width           =   1095
   End
End
Attribute VB_Name = "FViewOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_lWindowID As Long

Dim m_bLoading As Boolean

Dim m_bPrintOnly As Boolean
Dim m_oOrder As Order


'*****************************************************************************
' Form event handlers
'*****************************************************************************

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If m_lWindowID <> 0 Then
        MDIMain.UnloadTool m_lWindowID
    End If
    MDIMain.DoRefresh
End Sub


Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub
    
    If width < 9000 Then width = 9000
    If Height < 4755 Then Height = 4755
    
    WB.Move WB.Left, WB.Top, ScaleWidth, ScaleHeight - 680
    cmdPrint.Left = width - 1300
End Sub


'*****************************************************************************
' Public Properties and Methods
'*****************************************************************************

Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property


Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property


Public Sub SetCaption(ByRef i_sTitle As String)
    Me.caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub


Public Sub ShowOrder(ByVal i_oOrder As Order, Optional ByVal i_bPrintOnly As Boolean = True)
    Set m_oOrder = i_oOrder
    m_bPrintOnly = i_bPrintOnly
    
    If Not m_bPrintOnly Then
        MDIMain.AddNewWindow Me
        SetCaption "Print Preview OP " & m_oOrder.OPKey
    
        m_bLoading = True
    
        optCPC.Visible = True
        optReceipt.Visible = True
        optCPCReceipt.Visible = True
        optCustomerReceipt.Visible = True
        
        'set the default based on user type
        If g_bWillCallUser Then
            optReceipt.Value = True
        Else
            optCPC.Value = True
        End If
        
        chkShowRemarks(0).Visible = optCPC.Value
        chkShowRemarks(1).Visible = optCPC.Value
        chkShowRemarks(2).Visible = optCPC.Value
        chkShowRemarks(3).Visible = optCPC.Value
        chkNoPricing.Visible = Not optCPC.Value
    
        m_bLoading = False
        
    End If
    
    LoadPage bFirstPass:=True
    
End Sub


'*****************************************************************************
' control event handlers
'*****************************************************************************

Private Sub optCPC_Click()
    If m_bLoading = True Then Exit Sub

    If optCPC.Value = True Then
        chkShowRemarks(0).Visible = True
        chkShowRemarks(1).Visible = True
        chkShowRemarks(2).Visible = True
        chkShowRemarks(3).Visible = True
        chkNoPricing.Visible = False
        LoadPage False
    End If
End Sub


Private Sub optCPCReceipt_Click()
    If m_bLoading = True Then Exit Sub
    
    If optCPCReceipt.Value = True Then
        chkShowRemarks(0).Visible = False
        chkShowRemarks(1).Visible = False
        chkShowRemarks(2).Visible = False
        chkShowRemarks(3).Visible = False
        chkNoPricing.Visible = True
        LoadPage False
    End If
End Sub


Private Sub optCustomerReceipt_Click()
    If m_bLoading = True Then Exit Sub
    
    If optCustomerReceipt.Value = True Then
        chkShowRemarks(0).Visible = False
        chkShowRemarks(1).Visible = False
        chkShowRemarks(2).Visible = False
        chkShowRemarks(3).Visible = False
        chkNoPricing.Visible = True
        LoadPage False
    End If
End Sub


Private Sub optReceipt_Click()
    If m_bLoading = True Then Exit Sub
    
    If optReceipt.Value = True Then
        chkShowRemarks(0).Visible = False
        chkShowRemarks(1).Visible = False
        chkShowRemarks(2).Visible = False
        chkShowRemarks(3).Visible = False
        chkNoPricing.Visible = True
        LoadPage False
    End If
End Sub


Private Sub chkWarehouse_Click()
        LoadPage False
End Sub


Private Sub chkNoPricing_Click()
        LoadPage False
End Sub


Private Sub chkShowRemarks_Click(Index As Integer)
        LoadPage False
End Sub


'9/20/09 LR added checks

Private Sub cmdPrint_Click()
    'if the WebBrowser is still busy, exit
    If WB.Busy = True Then Exit Sub

    'if we're printing a 2-copy receipt
    If optReceipt.Value = True Then
        
        If OrderHasShipment(m_oOrder.OPKey) Then
            MsgBox "A 2-copy receipt cannot be printed" & vbCrLf & "because the order has already been picked up or shipped.", vbExclamation, "Print 2-Copy Receipt for OP-" & m_oOrder.OPKey
        ElseIf OrderHasReceipt(m_oOrder.OPKey) Then
            MsgBox "A 2-copy receipt cannot be printed" & vbCrLf & "because one has already been printed.", vbExclamation, "Print 2-Copy Receipt for OP-" & m_oOrder.OPKey
        Else
            WB.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
            
            'Can we determine if the user clicked OK or Cancel on the Print Dialog Box?

            LogOAEvent "Order", GetUserID, m_oOrder.OPKey, , , "Printed 2-Copy Receipt"
            SetReceiptPrinted m_oOrder.OPKey
        End If
    
    Else
        WB.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
            
        If optCustomerReceipt.Value = True Then
            LogOAEvent "Order", GetUserID, m_oOrder.OPKey, , , "Printed Customer Receipt"
        End If
    End If

End Sub


'this fires only in response to the Navigate2 method, not the Refresh method

Private Sub WB_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

    If LCase(CStr(URL)) = LCase(XMLURL) Then     '3/16/05 LR I don't think this is necessary
        
        WB.Refresh      '3/16/05 LR force a refresh
        Screen.MousePointer = vbDefault
        
    Else
        MsgBox "Unexpected URL " & CStr(URL), vbExclamation, "FViewOrder.WB_DocumentComplete"
    End If
End Sub


'*****************************************************************************
' Private functions
'*****************************************************************************

Private Function GetStrXML() As String
    Dim oXMLNode As JDMPDXML.XMLNode

    Set oXMLNode = m_oOrder.Export(chkShowRemarks(0).Value, chkShowRemarks(1).Value, chkShowRemarks(2).Value, chkNoPricing.Value, chkWarehouse.Value, (chkShowRemarks(3).Visible And chkShowRemarks(3).Value = vbChecked))
    
    oXMLNode.IndentWidth = 2    'need to have a width >= 0 to make xml well-formed
    GetStrXML = oXMLNode.ExportString
End Function


Private Sub LoadPage(ByVal bFirstPass As Boolean)
    Dim sErrMsg As String
    Dim sXML As String
    sXML = GetStrXML
    
    On Error GoTo EH
    
    If optCPC.Value = True Then
        sErrMsg = "Order.xsl"
        SaveToFile XMLURL, XslHeader(g_XsltPath + "Order.xsl") + sXML
        
    ElseIf optReceipt.Value = True Then
        sErrMsg = "Receipt.xsl"
        SaveToFile XMLURL, XslHeader(g_XsltPath + "Receipt.xsl") + sXML
        
    ElseIf optCPCReceipt.Value = True Then
        sErrMsg = "CPCOnly.xsl"
        SaveToFile XMLURL, XslHeader(g_XsltPath + "CPCOnly.xsl") + sXML
        
    ElseIf optCustomerReceipt.Value = True Then
        sErrMsg = "CustomerOnly.xsl"
        SaveToFile XMLURL, XslHeader(g_XsltPath + "CustomerOnly.xsl") + sXML
        
    End If
    
    If bFirstPass Then
        Screen.MousePointer = vbHourglass
        Debug.Print "Navigate to: " & XMLURL
        sErrMsg = sErrMsg & vbCrLf & "FirstPass Navigate " & XMLURL
        WB.Navigate2 XMLURL
    Else
        Debug.Print "Refresh: " & WB.LocationURL
        sErrMsg = sErrMsg & vbCrLf & "Not FirstPass Refresh"
        WB.Refresh
    End If
    
    Exit Sub
EH:
    msg sErrMsg & vbCrLf & Err.Number & " - " & Err.Description & vbCrLf & _
    "Show this to the Computer Guys", vbExclamation, "FViewOrder.LoadPage"

End Sub






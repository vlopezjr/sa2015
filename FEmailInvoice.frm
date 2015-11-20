VERSION 5.00
Begin VB.Form FEmailInvoice 
   Caption         =   "Email Invoice"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInvoice 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Text            =   "FEmailInvoice.frx":0000
      Top             =   2640
      Width           =   10215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   10215
      Begin VB.TextBox txtNotes 
         Height          =   555
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   9615
      End
      Begin VB.CheckBox chkCC 
         Caption         =   "CC: me a copy of this email"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   3315
      End
      Begin VB.CheckBox chkLogThis 
         Caption         =   "Log this in the customer's Collection History"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   3435
      End
      Begin VB.Label Label2 
         Caption         =   "Add a note to the customer (Not saved in our system):"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   6915
      End
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   4635
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   8400
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   9480
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "To..."
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "FEmailInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_invckey As Long
Private m_sCustID As String
Private m_sDocDescr As String

Public Property Get InvcKey() As Long
    InvcKey = m_invckey
End Property

Public Property Let InvcKey(ByVal lNewValue As Long)
    m_invckey = lNewValue
End Property

Private Sub Form_Activate()
    txtEmail.SetFocus
End Sub

Public Sub Init()
    SetWaitCursor True
    txtInvoice.text = GenerateInvoice
    SetWaitCursor False
    
    Me.Show vbModal
End Sub

Private Sub txtEmail_Change()
    If Len(txtEmail.text) > 0 Then
        cmdSend.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSend_Click()
    Dim emailaddr As String
    Dim subject As String
    Dim body As String
    Dim cc As String
    Dim oRemarkContext As MemoMeister.RemarkContext
    Dim remark As String

    SetWaitCursor True
    
    If Len(txtNotes.text) > 0 Then body = txtNotes.text & vbCrLf & vbCrLf
    body = body & txtInvoice.text
    
    If chkCC.value = vbChecked Then cc = GetUserName & "@caseparts.com"
    
    emailaddr = txtEmail.text
    subject = "Case Parts " & m_sDocDescr
    
    If chkCC.value = vbChecked Then cc = GetUserName & "@caseparts.com"
        
    EMail.Send GetUserName & "@caseparts.com", emailaddr, subject, body, False, cc
    
    If chkLogThis.value = vbChecked Then
        remark = GetUserName & " emailed " & m_sDocDescr & " to " & emailaddr
        Set oRemarkContext = New RemarkContext
        oRemarkContext.Load "ARCustLoad", m_sCustID
        oRemarkContext.AddRemark "Cust.AR.Coll", remark
        oRemarkContext.Save True
        Set oRemarkContext = Nothing
    End If
    
    SetWaitCursor False
    
    Me.Hide
End Sub


Private Function GenerateInvoice() As String
    Dim ocon As ADODB.Connection
    Dim ocmd As ADODB.Command
    Dim orstHeader As ADODB.Recordset
    Dim orstDetail As ADODB.Recordset
    Dim addrFormat As Integer
    Dim pubOrderRemark As String
    Dim salesAmt As String
    Dim staxAmt As String
    Dim shipAmt As String
    Dim total As String
    
    Set ocmd = New ADODB.Command
    With ocmd
        .ActiveConnection = g_DB.Connection
        .CommandText = "spcpcGetInvoice"
        .CommandType = adCmdStoredProc
        .Parameters("@invckey").value = m_invckey
        Set orstHeader = .Execute
    End With
    With orstHeader
        m_sDocDescr = .Fields("DocType") & ": " & .Fields("TranNo")
        StringBuilder.AddLine m_sDocDescr & " (eMail Copy)"
        StringBuilder.AddLine .Fields("TranDate")
        
        StringBuilder.AddLine
        StringBuilder.AddLine "Case Parts Company"
        StringBuilder.AddLine "Headquarters/Accounting"
        StringBuilder.AddLine "877 Monterey Pass Road, Monterey Park, CA 91754"
        StringBuilder.AddLine "Tel: (323)729-6000 Tel (800)421-0271 Fax: (800)972-2441"
        
        StringBuilder.AddLine
        StringBuilder.Add LJustify("Bill To:", 40)
        StringBuilder.AddLine "Ship To:"
        StringBuilder.Add LJustify(RTrim$(.Fields("BillAddrName")), 40)
        StringBuilder.AddLine RTrim$(.Fields("ShipAddrName"))
        StringBuilder.Add LJustify(RTrim$(.Fields("BillAddrLine1")), 40)
        StringBuilder.AddLine RTrim$(.Fields("ShipAddrLine1"))
        
        If Len(RTrim$(.Fields("BillAddrLine2"))) > 0 Then addrFormat = 1
        If Len(RTrim$(.Fields("ShipAddrLine2"))) > 0 Then addrFormat = addrFormat + 2
        Select Case addrFormat
            Case 0:
                StringBuilder.Add LJustify(RTrim$(.Fields("BillCity")) & ", " & RTrim$(.Fields("BillState")) & " " & RTrim$(.Fields("BillZip")), 40)
                StringBuilder.AddLine RTrim$(.Fields("ShipCity")) & ", " & RTrim$(.Fields("ShipState")) & " " & RTrim$(.Fields("ShipZip"))
            Case 1:
                StringBuilder.Add LJustify(RTrim$(.Fields("BillAddrLine2")), 40)
                StringBuilder.AddLine RTrim$(.Fields("ShipCity")) & ", " & RTrim$(.Fields("ShipState")) & " " & RTrim$(.Fields("ShipZip"))
                StringBuilder.AddLine RTrim$(.Fields("BillCity")) & ", " & RTrim$(.Fields("BillState")) & " " & RTrim$(.Fields("BillZip"))
            Case 2:
                StringBuilder.Add LJustify(RTrim$(.Fields("BillCity")) & ", " & RTrim$(.Fields("BillState")) & " " & RTrim$(.Fields("BillZip")), 40)
                StringBuilder.AddLine RTrim$(.Fields("ShipAddrLine2"))
                StringBuilder.Add LJustify("", 40)
                StringBuilder.AddLine RTrim$(.Fields("ShipCity")) & ", " & RTrim$(.Fields("ShipState")) & " " & RTrim$(.Fields("ShipZip"))
            Case 3:
                StringBuilder.Add LJustify(RTrim$(.Fields("BillAddrLine2")), 40)
                StringBuilder.AddLine RTrim$(.Fields("ShipAddrLine2"))
                StringBuilder.Add LJustify(RTrim$(.Fields("BillCity")) & ", " & RTrim$(.Fields("BillState")) & " " & RTrim$(.Fields("BillZip")), 40)
                StringBuilder.AddLine RTrim$(.Fields("ShipCity")) & ", " & RTrim$(.Fields("ShipState")) & " " & RTrim$(.Fields("ShipZip"))
        End Select
        
        StringBuilder.AddLine
        StringBuilder.AddLine "Account # " & " : " & .Fields("CustID")
        StringBuilder.AddLine "Terms     " & " : " & .Fields("Terms")
        StringBuilder.AddLine "Order #   " & " : OP " & .Fields("OPKey")
        StringBuilder.AddLine "Order Date" & " : " & Format$(.Fields("OrderDate"), "mm/dd/yyyy")
        StringBuilder.AddLine "CSR       " & " : " & .Fields("CSR")
        StringBuilder.AddLine "PO        " & " : " & .Fields("CustPONo")
        StringBuilder.AddLine "Ship Via  " & " : " & .Fields("ShipMethDesc")
    
        m_sCustID = .Fields("CustID") 'cache for MM remark injection
        pubOrderRemark = RTrim$(.Fields("PubOrdRemark"))
        salesAmt = Format$(.Fields("SalesAmt"), "####0.00")
        staxAmt = Format$(.Fields("STaxAmt"), "####0.00")
        shipAmt = Format$(.Fields("ShipAmt"), "####0.00")
        total = Format$(.Fields("TranAmt"), "####0.00")
        
    End With
    
    StringBuilder.AddLine
    '           1234567890123456789012345678901234567890123456789012345678901234567890123456
    StringBuilder.AddLine "Qty Shipped B/O PartNo       Description                    Unit    Ext"
    '           1   1       0   xxxxxxxxxxxx xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx xxxx.xx xxxxx.xx
    '                           12           30
    
    ' 1/25/12 LR modified to handle NULLs in Credit Memos
    
    Set orstDetail = orstHeader.NextRecordset
    With orstDetail
        Do While Not .EOF
            If IsNull(.Fields("QtyOrd")) Then
                StringBuilder.Add LJustify("    ", 4)
            Else
                StringBuilder.Add LJustify(.Fields("QtyOrd"), 4)
            End If
            
            If .Fields("QtyShipped") >= 0 Then
                StringBuilder.Add LJustify(.Fields("QtyShipped"), 8)
            Else
                StringBuilder.Add LJustify("        ", 8)
            End If
            
            If IsNull(.Fields("QtyOnBO")) Then
                StringBuilder.Add LJustify("    ", 4)
            Else
                StringBuilder.Add LJustify(.Fields("QtyOnBO"), 4)
            End If
            
            StringBuilder.Add LJustify(Left(.Fields("ItemID"), 12), 13)
            StringBuilder.Add LJustify(Left(.Fields("Description"), 30), 31)
            StringBuilder.Add RJustify(Format$(.Fields("UnitPrice"), "###0.00"), 7)
            StringBuilder.AddLine RJustify(Format$(.Fields("ExtPrice"), "####0.00"), 9)
            .MoveNext
        Loop
    End With
    
    If Len(pubOrderRemark) > 0 Then
        StringBuilder.AddLine
        StringBuilder.AddLine
        StringBuilder.AddLine "Remarks:"
        StringBuilder.AddLine pubOrderRemark
    End If
    
    StringBuilder.AddLine
    StringBuilder.AddLine
    StringBuilder.AddLine RJustify("Sales Amount " & " : " & RJustify(salesAmt, 8), 76)
    StringBuilder.AddLine RJustify("Sales Tax    " & " : " & RJustify(staxAmt, 8), 76)
    StringBuilder.AddLine RJustify("Freight      " & " : " & RJustify(shipAmt, 8), 76)
    StringBuilder.AddLine RJustify("Invoice Total" & " : " & RJustify(total, 8), 76)

    GenerateInvoice = StringBuilder.ToString
    StringBuilder.Clear
    
End Function


Private Function LJustify(value As String, width As Integer)
    LJustify = Format(value, "!" & String(width, "@"))
End Function


Private Function RJustify(value As String, width As Integer)
    RJustify = Format(value, String(width, "@"))
End Function


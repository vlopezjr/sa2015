VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FEmailQuote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email Quote"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAddNotes 
      Caption         =   "Add my notes to the Order remark"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   720
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox chkLogThis 
      Caption         =   "Log this in Private Order remarks"
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      Top             =   420
      Value           =   1  'Checked
      Width           =   2715
   End
   Begin VB.TextBox txtBody 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2580
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   9660
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   435
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkCC 
      Caption         =   "CC: me a copy"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   60
      Width           =   1395
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   540
      TabIndex        =   0
      Top             =   120
      Width           =   4635
   End
   Begin VB.TextBox txtNotes 
      Height          =   915
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   7455
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5535
      Left            =   180
      TabIndex        =   6
      Top             =   2580
      Visible         =   0   'False
      Width           =   10455
      ExtentX         =   18441
      ExtentY         =   9763
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
      Location        =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Notes to customer (This is optional. It is sent to your customer but is not saved in your order.):"
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   6675
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   315
   End
End
Attribute VB_Name = "FEmailQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWindowID  As Long
Private m_emailaddr As String
Private m_isHtml As Boolean
Private m_quotenumber As String
Private m_ponumber As String
Private m_subject As String
Private m_html As String
Private m_logEvent As Boolean
Private m_CustomerNotes As String

' MDI child interface stuff

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
    If Len(txtEmail.text) > 0 Then
        txtNotes.SetFocus
    Else
        txtEmail.SetFocus
    End If
End Sub

Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property

Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property

Public Sub SetCaption(ByRef i_sTitle As String)
    Me.Caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub

' Form properties

Public Property Get emailaddr() As String
    emailaddr = m_emailaddr
End Property

Public Property Let emailaddr(ByVal sNewValue As String)
    m_emailaddr = sNewValue
End Property

Public Property Get IsHtml() As Boolean
    IsHtml = m_isHtml
End Property

Public Property Let IsHtml(ByVal bNewValue As Boolean)
    m_isHtml = bNewValue
End Property

Public Property Get QuoteNumber() As String
    QuoteNumber = m_quotenumber
End Property

Public Property Let QuoteNumber(ByVal sNewValue As String)
    m_quotenumber = sNewValue
End Property

Public Property Get PONumber() As String
    PONumber = m_ponumber
End Property

Public Property Let PONumber(ByVal sNewValue As String)
    m_ponumber = sNewValue
End Property

Public Property Get LogEvent() As Boolean
    LogEvent = m_logEvent
End Property

Public Property Let LogEvent(ByVal bNewValue As Boolean)
    m_logEvent = bNewValue
End Property

'Private Sub chkLogThis_Click()
'    m_logEvent = (chkLogThis.value = vbChecked)
'End Sub


Private Sub txtEmail_Change()
    If Len(txtEmail.text) > 0 Then
        cmdSend.Enabled = True
    End If
End Sub



Public Sub Init()
    Dim Source As New MSXML2.DOMDocument30
    Dim stylesheet As New MSXML2.DOMDocument30
    Dim myErr
    
    MDIMain.AddNewWindow Me
    SetCaption "Email Quote# " & m_quotenumber
    
    m_subject = "CaseParts Quote# " & m_quotenumber
    If Len(m_ponumber) > 0 Then
        m_subject = m_subject & " (PO# " & m_ponumber & ")"
    End If
    
    txtEmail.text = m_emailaddr
    
    'm_logEvent = chkLogThis.value
    m_logEvent = False
    
    If Len(txtEmail.text) > 0 Then
        cmdSend.Enabled = True
    End If
    
    If m_isHtml Then
        webBrowser1.Visible = True
    Else
        txtBody.Visible = True
    End If
    
    Source.async = False
    Source.Load g_QuoteEmailPath & GetUserName & ".xml"
    
    If (Source.parseError.errorCode <> 0) Then
        Set myErr = Source.parseError
        MsgBox ("XML error " & myErr.Reason)
    Else
        stylesheet.async = False
        stylesheet.Load GetXsl()
        
        If (stylesheet.parseError.errorCode <> 0) Then
            Set myErr = stylesheet.parseError
            MsgBox ("XSLT error " & myErr.Reason)
        Else
            m_html = Source.transformNode(stylesheet)
            If m_isHtml Then
                'save to an intermediate html file
                SaveAsTextFile GetQuoteHtmlName(), m_html
                'then view it
                webBrowser1.Navigate GetQuoteUri(), 4  'flag means don't cache
            Else
                txtBody.text = m_html
            End If
       End If
    End If
        
    Me.Show vbModal
    
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSend_Click()
    Dim body As String
    Dim cc As String
    
    If Len(txtNotes.text) > 0 Then body = txtNotes.text & IIf(m_isHtml, "<br/><br/>", vbCrLf & vbCrLf)
    body = body & "This quote does not include applicable freight or sales tax." & IIf(m_isHtml, "<br/>", vbCrLf) & m_html
    
    If chkCC.value = vbChecked Then cc = GetUserName & "@caseparts.com"
    
    m_emailaddr = txtEmail.text
    
    EMail.Send GetUserName & "@caseparts.com", txtEmail.text, m_subject, body, m_isHtml, cc
       
    m_logEvent = (chkLogThis.value = vbChecked)
    
    Me.Hide
End Sub

Private Function GetXsl() As String
    If m_isHtml Then
        GetXsl = g_QuoteEmailPath & "HtmlQuote.xsl"
    Else
        GetXsl = g_QuoteEmailPath & "TextQuote.xsl"
    End If
End Function

Private Function GetQuoteUri() As String
    GetQuoteUri = g_QuoteEmailUri & GetUserName & ".htm"
End Function


Private Function GetQuoteHtmlName() As String
    GetQuoteHtmlName = g_QuoteEmailPath & GetUserName & ".htm"
End Function


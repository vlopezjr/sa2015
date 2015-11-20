VERSION 5.00
Begin VB.Form FPickConflictNotifications 
   Caption         =   "Pick Conflict Notifications"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3638.839
   ScaleMode       =   0  'User
   ScaleWidth      =   8769.229
   Begin VB.Timer tmrScheduler 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   120
      Top             =   2640
   End
   Begin VB.CommandButton btnOpenConflicts 
      Caption         =   "Resolve Conflicts"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ListBox lstEntries 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8295
   End
   Begin VB.Label lblRefreshing 
      Caption         =   "Refreshing Conflicts..."
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "FPickConflictNotifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FormWidth = 8670
Private Const FormHeight = 3540


Private m_lWindowID As Long


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


Private Sub btnOpenConflicts_Click()
    Dim oFrm As FPickConflicts
    Set oFrm = New FPickConflicts
    oFrm.Show
End Sub

Private Sub Form_Load()
    LoadPickConflictOrdersLog
    tmrScheduler.Enabled = True
End Sub


Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Me.width = FormWidth
    Me.Height = FormHeight
End Sub

Private Sub LoadPickConflictOrdersLog()
    Dim orst As ADODB.Recordset
    Set orst = CallSP("spcpcGetConflictOrdersLog", "@WhseKey", User.GetUserWhseKey)
    
    
    Dim iOpKey As Long
    Dim iSoKey As Long
    Dim iItemKey As Long
    Dim sTranNo As String
    Dim sItemId As String
    Dim iWhseKey As Integer
    Dim sAction As String
    Dim dCreatedDate As Date
  
    lstEntries.Clear
    
    With orst
        If Not .EOF Then
            .MoveFirst

            Do While Not .EOF
                
                iOpKey = .Fields("OpKey").value
                iSoKey = .Fields("SoKey").value
                iItemKey = .Fields("SoKey").value
                sTranNo = .Fields("TranNo").value
                sItemId = .Fields("ItemId").value
                iWhseKey = .Fields("WhseKey").value
                sAction = .Fields("Description").value
                dCreatedDate = .Fields("CreatedDate").value
                
                Dim entry As String
                
                entry = TimeValue(dCreatedDate) & "  " & PadRight(sAction, 9) & " OP-" & iOpKey & ", SO-" & Trim(sTranNo) & ", Item-" & sItemId
                
                lstEntries.AddItem Trim(entry)

                .MoveNext
            Loop
        End If
    End With
    
End Sub

Private Sub tmrScheduler_Timer()
    lblRefreshing.Visible = True
    LoadPickConflictOrdersLog
    lblRefreshing.Visible = False
End Sub

Function PadRight(text As Variant, totalLength As Integer) As String
    Dim temp As String
    Dim tempLength As Integer
    Dim length As Integer
    
    temp = Trim(CStr(text))
    tempLength = Len(temp)
    length = totalLength - tempLength
    
    If (length > 0) Then
        PadRight = temp & String(length, " ")
    Else
        PadRight = temp
    End If
    
End Function

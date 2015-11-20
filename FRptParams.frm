VERSION 5.00
Begin VB.Form FRptParams 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Parameters"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDisplayRpt 
      Caption         =   "Display"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FRptParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbShowRpt As Boolean
Private moParent As FViewer
Private msDefaultcboTxt As String

Private WithEvents cbo As ComboBox
Attribute cbo.VB_VarHelpID = -1

Public Property Get UserName() As String
    UserName = GetUserName
End Property

Public Property Let DefaultTest(ByVal lNewValue As String)
    msDefaultcboTxt = lNewValue
End Property

Private Sub cbo_Validate(Cancel As Boolean)
    Dim lscboText As String
    Dim liCounter As Integer
    
    lscboText = cbo.Text
    
    For liCounter = 0 To cbo.ListCount - 1
        cbo.ListIndex = liCounter
        If cbo.Text = lscboText Then Exit Sub
        Cancel = False
    Next
    
    cbo.ListIndex = 1
    Cancel = True
End Sub

Public Function ShowParam(lorst As Recordset, oParent As FViewer) As Boolean
    Dim lbl As Label
    Dim opt As OptionButton
    'Dim cbo As ComboBox
    Dim cal As SOTACalendar
    Dim liCounter As Integer
    Set moParent = oParent
        
    For liCounter = 1 To lorst.RecordCount
        Select Case lorst.Fields("VBCtrlType")
            Case "opt"
                Set opt = Controls.Add("VB.OptionButton", "opt" & liCounter)
                Call SetControlProps(opt, liCounter * 500, 2000, 100)
                opt.Caption = Trim(lorst.Fields("OADisplayField"))
                If lorst.Fields("OADefaultValue") = 1 Then opt.value = True
            Case "cbo"
                Set lbl = Controls.Add("VB.label", "lbl" & liCounter)
                Call SetControlProps(lbl, liCounter * 500, 2000, 100)  'ctrl, top, width, left
                lbl.Caption = lorst.Fields("OADisplayField")
                
                Set cbo = Controls.Add("VB.ComboBox", "cbo" & liCounter)
                Call SetControlProps(cbo, liCounter * 500, 2000, 800)
                cbo.Tag = Trim(lorst.Fields("RptParamName"))
                                           
                Dim arst As ADODB.Recordset
                
                If Not IsNull(lorst.Fields("VBCtrlSqlQuery")) Then
                    Set arst = LoadDiscRst(lorst.Fields("VBCtrlSqlQuery"))
                    Helpers.LoadCombo cbo, arst, Trim(Replace(lorst.Fields("OADisplayField"), " ", "")), Trim(lorst.Fields("OAKeyField"))
                
                    'SMR - New - Get cbo default value
                    If Not IsNull(lorst.Fields("OADefaultValue")) Then
                        Dim sc As ScriptControl
                        Dim CodeString As String
                        
                        Set sc = New ScriptControl
                        sc.Language = "VBScript"
                        sc.AddObject "FRptParams", Me  'Me.Name, Me
                        
                        'CodeString = "sub Main() " & Me.Name & ".Controls(""cbo" & liCounter & """).Text=" & Me.Name & "." & lorst.Fields("OADefaultValue") & " end sub"
                        CodeString = "sub Main() FRptParams.DefaultTest=" & Me.Name & "." & lorst.Fields("OADefaultValue") & " end sub"
                        sc.AddCode CodeString
                        sc.Run "Main"
                        Call SetComboByText(cbo, msDefaultcboTxt)
                    End If
                Else
                    Set arst = LoadDiscRst("select * from tcpRptParamcboValues where ParamName = '" & lorst.Fields("OADisplayField") & "'")
                    Helpers.LoadCombo cbo, arst, "cboItemDisplayString", "cboItemDataInt"
                End If
                Set arst = Nothing
                
            Case "cal"
                Set lbl = Controls.Add("VB.label", "lbl" & liCounter)
                Call SetControlProps(lbl, liCounter * 500, 2000, 100)
                lbl.Caption = lorst.Fields("OADisplayField")
                
                Set cal = Controls.Add("SOTACalendarControl.SOTACalendar", "cal" & liCounter)
                Call SetControlProps(cal, liCounter * 500, 1500, 1000)
                cal.Tag = Trim(lorst.Fields("RptParamName"))
                cal.value = DateAdd("d", lorst.Fields("OADefaultValue"), Date)
        End Select
        lorst.MoveNext
    Next
    
    Me.Height = 500 * liCounter + 1200
    Me.Width = 3000
    cmdDisplayRpt.Top = Me.Height - 1000
    cmdDisplayRpt.Left = (Me.Width / 2) - (cmdDisplayRpt.Width / 2)
    
    mbShowRpt = False
    Me.Show vbModal
    
    'If cmdDisplay was not clicked, mbShowRpt will be false
    ShowParam = mbShowRpt
End Function


Private Sub cmdDisplayRpt_Click()
    mbShowRpt = True
    Dim oControl As Control
    For Each oControl In Controls
        If InStr(1, oControl.Name, "cbo") > 0 Then
            If oControl.ItemData(oControl.ListIndex) = 0 Then
                msg "Please make a selection.", vbInformation, oControl.Tag
                oControl.SetFocus
                Exit Sub
            End If
        End If
    Next

    'The control's tag property contains the rptParamName
    For Each oControl In Controls
        If InStr(1, oControl.Name, "cbo") > 0 Then
            moParent.ParamAdd Mid$(oControl.Name, 4), oControl.Tag, oControl.ItemData(oControl.ListIndex)
        ElseIf InStr(1, oControl.Name, "opt") > 0 Then
            If oControl.value = True Then
                moParent.ParamAdd Mid$(oControl.Name, 4), oControl.Caption, 1
            End If
        ElseIf InStr(1, oControl.Name, "cal") > 0 Then
            If oControl.Visible = True Then
                moParent.ParamAdd Mid$(oControl.Name, 4), oControl.Tag, oControl.value
            End If
        End If
    Next
    Unload Me
End Sub


Private Sub SetControlProps(psControl As Control, psTop As Long, psWidth As Long, _
        psLeft As Long)
    psControl.Top = psTop
    psControl.Width = psWidth
    psControl.Left = psLeft
    psControl.Visible = True
End Sub


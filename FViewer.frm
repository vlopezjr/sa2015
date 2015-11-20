VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Office Assistant Viewer"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   12120
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer crxViewer 
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12135
      _cx             =   21405
      _cy             =   11668
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
      EnableInteractiveParameterPrompting=   0   'False
   End
   Begin VB.PictureBox Viewer 
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   12075
      TabIndex        =   0
      Top             =   0
      Width           =   12135
   End
   Begin SHDocVwCtl.WebBrowser wbViewer 
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12105
      ExtentX         =   21352
      ExtentY         =   11668
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
End
Attribute VB_Name = "FViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim crxApp As CRAXDRT.Application
Dim crxRpt As CRAXDRT.Report
Dim crxTables As CRAXDRT.DatabaseTables
Dim crxTable As CRAXDRT.DatabaseTable
Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim crxSubReport As CRAXDRT.Report
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.section

Private mcolParams As Collection
Private msRptName As String
Private msRptPath As String
Private msRptExt As String

Private msRptDbName As String
Private msRptDSN As String
Private m_bPrintRpt As Boolean

Private m_lWindowID  As Long

Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property

Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property



Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_Initialize()
    Set mcolParams = New Collection
End Sub

Private Sub Form_Load()
    MDIMain.AddNewWindow Me
    SetCaption "Report: " & msRptName
End Sub


Private Sub Form_Resize()
    'WindowState: 0 (Orginal Size); 1 (Min); 2 (Max)
    If Me.WindowState = 1 Then Exit Sub
        
    Viewer.width = Me.width - 100
    Viewer.Height = Me.Height - 390

    wbViewer.width = Me.width - 100
    wbViewer.Height = Me.Height - 390
    
    crxViewer.width = Me.width
    crxViewer.Height = Me.Height
        
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.UnloadTool m_lWindowID
    Set mcolParams = Nothing
End Sub


Private Sub SetCaption(ByRef i_sTitle As String)
    Me.caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub


'-----Public Procedures/Functions----------------------

Public Sub ViewReportFromMenu(ReportName As String)
    'disconnected recordset
    Dim lsSql As String
    Dim lorst As ADODB.Recordset
    lsSql = "SELECT * FROM dbo.tcpRpt LEFT OUTER JOIN tcpRptParams ON " & _
        "dbo.tcpRpt.RptKey = dbo.tcpRptParams.RptKey WHERE dbo.tcpRpt.RptName = '" & ReportName & "'"
    Set lorst = LoadDiscRst(lsSql)

    'check if no records exist
    If lorst.EOF And lorst.BOF Then msg "This report is not set up for display.": Exit Sub

    'cache report values
    msRptName = Replace(lorst.Fields("RptName"), " ", "")
    msRptPath = g_ReportPath
    msRptExt = lorst.Fields("RptExt")

    'SMR - August 17, 2005
    msRptDbName = lorst.Fields("RptDbName").value
    msRptDSN = lorst.Fields("RptDSN").value

    'check if no parameters exist
    If IsNull(lorst.Fields("RptParamNbr")) Then
        Call ViewReportByType(ReportName)
        Exit Sub
    End If

    'We have parameters for this report.....Show Parameter Dialog
    Dim oParamFrm As FRptParams
    Set oParamFrm = New FRptParams
    If oParamFrm.ShowParam(lorst, Me) Then Call ViewReportByType(msRptName)
    Set oParamFrm = Nothing
End Sub


Public Sub ViewReportByType(ReportName As String)
    If Len(msRptName) = 0 Then
        If GetPathExt(ReportName) = False Then msg "This report is not set up for display.": Exit Sub
    End If

    Select Case UCase(msRptExt)
        Case "RPT": Call ViewRPT
        Case "ASP", "ASPX": Call ViewASP
        Case Else: MsgBox "Please add this report type to SageAssistant.": Exit Sub
    End Select
End Sub


Private Sub ViewASP()
    Dim liCounter As Integer
    Dim loViewerParams As ViewerParams
    Dim sURL As String
    Dim bOrderbyOldest As Boolean
    Dim lsCSR As String
    Dim ldStartDate As Date
    Dim ldEndDate As Date
    Dim lsFilter As String

    For liCounter = 1 To mcolParams.Count
        Set loViewerParams = mcolParams.Item(liCounter)
        With loViewerParams
            Select Case .ParamName
                Case "CSR": lsCSR = GetUserID(.ParamValue)
                Case "StartDate": ldStartDate = .ParamValue
                Case "EndDate": ldEndDate = .ParamValue
                Case "Oldest": If .ParamValue = 1 Then bOrderbyOldest = True
                Case "Newest": If .ParamValue = 1 Then bOrderbyOldest = False
                Case "Filter": lsFilter = "&qfilter=" & .ParamValue
            End Select
        End With
    Next

    sURL = msRptPath & "." & msRptExt & "?qdatabase=" & g_DB.DATABASE & _
    "&qopenquote=1&qcsr=" & lsCSR & "&qcreatedate=" & ldStartDate & "&qenddate=" & ldEndDate & "&qorderby=" & True
    If msRptName = "OpenOrder" Then sURL = sURL & lsFilter

    wbViewer.Navigate2 CStr(sURL)
    Viewer.Visible = False
    wbViewer.Visible = True
End Sub


Public Sub ViewRPT()
    Dim strServerOrDSNName As String
    Dim strDBNameOrPath As String
    Dim strUserID As String
    Dim strPassword As String
    Dim i As Long
    Dim j As Long

    Call SetWaitCursor(True)

    strServerOrDSNName = msRptDSN
    strDBNameOrPath = msRptDbName

    strUserID = "sa"
    strPassword = "C3l5ius"

    Set crxApp = New CRAXDRT.Application
    Set crxRpt = crxApp.OpenReport(g_ReportPath & msRptName & "." & msRptExt)

    crxRpt.DATABASE.Tables(1).SetLogOnInfo strServerOrDSNName, _
        strDBNameOrPath, strUserID, strPassword

    'This removes the schema from the Database Table's Location property.
    Set crxTables = crxRpt.DATABASE.Tables
    For Each crxTable In crxTables
        With crxTable
             .Location = .Name
        End With
    Next

    'Loop through the Report's Sections to find any subreports, and change them as well
    Set crxSections = crxRpt.Sections

    For i = 1 To crxSections.Count
        Set crxSection = crxSections(i)

        For j = 1 To crxSection.ReportObjects.Count

            If crxSection.ReportObjects(j).Kind = crSubreportObject Then
                Set crxSubreportObject = crxSection.ReportObjects(j)

                'Open the subreport, and treat like any other report
                Set crxSubReport = crxSubreportObject.OpenSubreport
                Set crxTables = crxSubReport.DATABASE.Tables

                For Each crxTable In crxTables
                    With crxTable
                        .SetLogOnInfo strServerOrDSNName, _
                            strDBNameOrPath, strUserID, strPassword
                        .Location = .Name
                    End With
                Next
            End If
        Next j
    Next i

    Dim liCounter As Integer
    Dim loViewerParams As ViewerParams

    For liCounter = 1 To mcolParams.Count
        Set loViewerParams = mcolParams.Item(liCounter)
        crxRpt.ParameterFields(liCounter).AddCurrentValue (loViewerParams.ParamValue)
    Next

    wbViewer.Visible = False
    Viewer.Visible = True

    'View the report
    crxViewer.ReportSource = crxRpt
    
    
    crxViewer.ViewReport
    Me.Show
    
    
    Call SetWaitCursor(False)
End Sub


Public Sub ParamAdd(ParamNo As Long, ParamName As String, ParamValue As Variant)
    Dim loViewerParams As ViewerParams

    Set loViewerParams = New ViewerParams
    With loViewerParams
        .ParamName = ParamName
        .ParamValue = ParamValue
        mcolParams.Add loViewerParams, CStr(ParamNo)
        Debug.Print ParamNo, ParamName, ParamValue
    End With
End Sub


'-----Public Procedures/Functions----------------------

Private Function GetPathExt(ReportName As String) As Boolean
    GetPathExt = False

    'disconnected recordset
    Dim lsSql As String
    Dim lorst As ADODB.Recordset

    lsSql = "select * from tcpRpt where RptName = '" & ReportName & "'"
    Set lorst = LoadDiscRst(lsSql)

    'check if no records exist
    If lorst.EOF And lorst.BOF Then: Exit Function

    'cache report values
    msRptName = Replace(ReportName, " ", "")
    msRptPath = g_ReportPath
    msRptExt = lorst.Fields("RptExt")

    'SMR - August 17, 2005
    msRptDbName = lorst.Fields("RptDbName").value
    msRptDSN = lorst.Fields("RptDSN").value

    GetPathExt = True
End Function



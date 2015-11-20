Attribute VB_Name = "PmtTerms"
Option Explicit

Public g_orstTerms As ADODB.Recordset


Public Sub Init()
    Dim sSQL As String
    
    sSQL = "SELECT PmtTermsKey, PmtTermsID, DueDayOrMonth " _
         & "FROM tciPaymentTerms " _
         & "WHERE CompanyID = 'CPC' AND DiscDateOption = 1 " _
         & "ORDER BY PmtTermsID"
    Set g_orstTerms = LoadDiscRst(sSQL)
    
End Sub

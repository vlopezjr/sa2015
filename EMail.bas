Attribute VB_Name = "EMail"
Option Explicit


Public Sub Send(ByVal mailFrom As String, ByVal mailTo As String, _
                    ByVal subject As String, ByVal body As String, _
                    Optional ByVal isBodyHtml As Boolean = True, _
                    Optional ByVal cc As String = "", _
                    Optional ByVal bcc As String = "")
                    
    If g_MailProxy Is Nothing Then Set g_MailProxy = CreateMailProxy
    
    g_MailProxy.EMail mailFrom, mailTo, subject, body, isBodyHtml, cc, bcc
    
End Sub
                    
                    
Public Sub SendToList(ByVal listCode As String, ByVal mailFrom As String, _
                        ByVal subject As String, ByVal body As String, _
                        Optional ByVal isBodyHtml As Boolean = True)

    If g_MailProxy Is Nothing Then Set g_MailProxy = CreateMailProxy

    g_MailProxy.EMailList listCode, mailFrom, subject, body, isBodyHtml

End Sub


Private Function CreateMailProxy() As MSSOAPLib30.SoapClient30
    Set CreateMailProxy = New MSSOAPLib30.SoapClient30
    CreateMailProxy.MSSoapInit g_MailServiceUrl & "?WSDL"
End Function

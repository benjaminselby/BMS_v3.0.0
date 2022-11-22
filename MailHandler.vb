Imports System.Net.Mail


Namespace BmsShared

    Public Class MailHandler

        Private Shared smtpClient As SmtpClient = New SmtpClient(ConfigHandler.mailServerName)

        Public Overloads Shared Sub SendMail(
                            senderAddress As String, recipients As List(Of String),
                            subject As String, messageBody As String)

            Dim Mail As MailMessage = New MailMessage()
            For Each recipientAddress As String In recipients

                '===========================================================================
                ' DEBUG MODE - CHANGE FOR PRODUCTION!!!
                'Mail.To.Add("selby_b@woodcroft.sa.edu.au");
                Mail.To.Add(recipientAddress)
                '===========================================================================
            Next

            Mail.From = New MailAddress(senderAddress)
            Mail.Subject = subject
            Mail.Body = messageBody
            Mail.IsBodyHtml = True

            smtpClient.Send(Mail)

        End Sub



        Public Overloads Shared Sub SendMail(senderAddress As String,
                                   emailRecipient As String, subject As String, messageBody As String)

            Dim emailRecipients As List(Of String) = New List(Of String) From {emailRecipient}
            SendMail(senderAddress, emailRecipients, subject, messageBody)

        End Sub


    End Class

End Namespace
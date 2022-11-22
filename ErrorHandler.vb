

Namespace BmsShared

    Public Class ErrorHandler

        Public Shared Sub HandleError(exception As Exception,
                                      rethrow As Boolean,
                                      sendEmail As Boolean,
                                      Optional addToLogTable As Boolean = True,
                                      Optional showMsgBox As Boolean = False,
                                      Optional source As String = "",
                                      Optional message As String = "")

            Dim errorMessage As String = "ERROR FOR USER " _
                & ConfigHandler.currentUser.id & vbCrLf _
                & If(source = "", "", " - " & source & " - ") _
                & If(message = "", "", " - " & message & " - ") _
                & exception.Message & vbCrLf & exception.StackTrace

            If addToLogTable Then
                DataHandler.logMessage(message:=errorMessage, status:="ERROR", source:=source)
            End If


            If showMsgBox Then

                Dim msgBoxErrorMessage As String _
                    = "Please take a screen-shot of this error message and contact Data Management for assistance. " _
                    & vbCrLf & vbCrLf & errorMessage
                MessageBox.Show(msgBoxErrorMessage, "ERROR in BMS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End If


            If sendEmail Then

                Dim subjectLine As String =
                    "ERROR - BMS" & If(source = "", "", ": " + source)

                MailHandler.SendMail(
                    senderAddress:=ConfigHandler.systemEmailSender,
                    recipients:=ConfigHandler.dataManagementEmails,
                    subject:=subjectLine,
                    messageBody:=errorMessage)

            End If

            If rethrow Then Throw exception

        End Sub

    End Class

End Namespace
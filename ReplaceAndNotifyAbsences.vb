Imports System.Configuration
Imports BMS.BmsShared



Public Class ReplaceAndNotifyAbsences

    Private Const minimumFormHeight = 300

    Private Sub FormLoad(sender As Object, e As System.EventArgs) Handles Me.Load

        PopulateAbsenteesDataGrid()

    End Sub



    ' ======================================================================================
    ' FUNCTIONS
    ' ======================================================================================
#Region "Functions"


    Private Sub PopulateAbsenteesDataGrid()
        Try

            absenteeListDgv.DataSource = DataHandler.getDataTable(
                procedureName:=ConfigurationManager.AppSettings("GetSanctionAbsencesProc"))

            For Each column As DataGridViewColumn In absenteeListDgv.Columns
                column.ReadOnly = True
                column.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            With absenteeListDgv
                .Columns(0).Visible = False     ' DateOfBooking
                .Columns(1).Visible = True      ' Sanction Date
                .Columns(2).Visible = False     ' StaffReason
                .Columns(3).Visible = False     ' Completed
                .Columns(4).Visible = False     ' Removed
                .Columns(5).Visible = False     ' Present
                .Columns(6).Visible = False     ' StaffId
                .Columns(7).Visible = True      ' StaffName
                .Columns(8).Visible = False     ' StaffEmail
                .Columns(9).Visible = False     ' StudentId
                .Columns(10).Visible = True     ' StudentName
                .Columns(11).Visible = True     ' StudentTutorGroup
                .Columns(12).Visible = True     ' StudentYearLevel
                .Columns(13).Visible = True     ' Sanction Code
                .Columns(14).Visible = True     ' HoyName
                .Columns(15).Visible = False    ' HoyEmail
                .Columns(16).Visible = False    ' Sanction record sequence number.
            End With

            absenteeListDgv.Width = absenteeListDgv.Columns.GetColumnsWidth(DataGridViewElementStates.Visible) + 1
            absenteeListDgv.Height = absenteeListDgv.Rows.GetRowsHeight(DataGridViewElementStates.Visible) _
                    + absenteeListDgv.ColumnHeadersHeight

            Me.Width = absenteeListDgv.Width + 64
            Dim myHeight As Integer = absenteeListDgv.Height + Me.ReplaceAndNotifyBtn.Height _
                + Me.Height - Me.ClientSize.Height + 84
            Me.Height = If(myHeight < minimumFormHeight, minimumFormHeight, myHeight)

        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=False, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True, source:="BMS Admin - ReplaceAndNotifyAbsences")
        End Try


    End Sub

#End Region



    ' ======================================================================================
    ' EVENT HANDLERS
    ' ======================================================================================
#Region "EventHandlers"



    Private Sub ReplaceAndNotifyBtn_Click(
            sender As System.Object, e As System.EventArgs) _
            Handles ReplaceAndNotifyBtn.Click

        If absenteeListDgv.Rows.Count <= 0 Then
            MessageBox.Show("No sanction absences to process.", "Nothing to do")
            Return
        End If

        Me.Cursor = Cursors.WaitCursor

        Dim creationFailedFlag As Boolean = False

        Try

            Dim nSanctionsToProcess = CType(absenteeListDgv.DataSource, DataTable).Rows.Count
            Dim nSanctionsProcessed As Integer = 0
            Dim newSanctionSeq As String
            Dim newSanctionDate As Date

            For Each sanctionRow As DataRow In CType(absenteeListDgv.DataSource, DataTable).Rows

                ' Remove the absence that was not attended by the student, and 
                ' create a new one in the future for them to attend instead. 

                Dim commandParameters = New Dictionary(Of String, String) _
                    From {{"AbsentSanctionSeq", sanctionRow("Seq")}}

                DataHandler.executeNonQuery(
                    procedureName:=ConfigurationManager.AppSettings("RecreateAbsentSanctionProc"),
                    commandParameters:=commandParameters,
                    returnValue:=newSanctionSeq)

                ' If creation of new sanction fails we don't stop loop but flag event so 
                ' we can alert user after processing. 
                If CInt(newSanctionSeq) = 0 _
                        Or Integer.TryParse(newSanctionSeq, New Integer) = False Then
                    creationFailedFlag = True
                    Continue For
                End If


                ' Get details of the new sanction which has been created

                Dim queryParameters = New Dictionary(Of String, String) _
                    From {{"SanctionSeq", newSanctionSeq}}

                Dim newSanctionDateStr = DataHandler.getScalarValue(
                    command:=ConfigurationManager.AppSettings("GetSanctionDateSql"),
                    queryParameters:=queryParameters,
                    isStoredProcedure:=False,
                    columnName:="SanctionDate")

                newSanctionDate = CDate(newSanctionDateStr)

                ' ====================================================================================
                ' EMAILS SENT HERE!
                ' ====================================================================================
                SendEmailsNewSanctionCreated(
                    sanction:=sanctionRow,
                    newSanctionDate:=newSanctionDate)
                ' ====================================================================================

                nSanctionsProcessed += 1
                ProgressBar1.Value = (nSanctionsProcessed / nSanctionsToProcess) * 100

            Next

        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=True, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True, source:="BMS Admin - ReplaceAndNotifyAbsences")
        End Try

        Me.Cursor = Cursors.Default
        ReplaceAndNotifyBtn.Enabled = False

        If creationFailedFlag = True Then

            MessageBox.Show(text:="Unable to create new sanctions for some students. Please update the remaining sanctions manually." _
                   & "You may need to delete and re-create them. " & vbCrLf _
                   & "If you continue to have problems, contact Data Management for assistance.",
                   caption:="Could not Re-create all Sanctions",
                   icon:=MessageBoxIcon.Warning,
                   buttons:=MessageBoxButtons.OK)

            PopulateAbsenteesDataGrid()

        Else
            MessageBox.Show(text:="New sanctions were created for all students, And the relevant staff have all been notified.",
                   caption:="All Sanctions Processed",
                   buttons:=MessageBoxButtons.OK)
            Me.Close()
        End If

    End Sub



    Private Function GetStaffEmailBody(studentName As String, staffName As String,
                                       sanctionType As String, reason As String,
                                       oldDate As Date, newDate As Date
                                      ) As String

        Dim header = ConfigHandler.staffEmailElements.Select("Name='HeaderText'")(0)("Value").ToString()
        Dim introduction = ConfigHandler.staffEmailElements.Select("Name='Introduction'")(0)("Value").ToString()
        Dim detailsTable = String.Format(
            ConfigHandler.staffEmailElements.Select("Name='DetailsTable'")(0)("Value").ToString(),
            studentName,
            staffName,
            sanctionType,
            reason,
            oldDate.ToString("yyyy-MM-dd"))
        Dim body = String.Format(
            ConfigHandler.staffEmailElements.Select("Name='Body'")(0)("Value").ToString(),
            newDate.ToString("yyyy-MM-dd"))
        Dim conclusion = ConfigHandler.staffEmailElements.Select("Name='Conclusion'")(0)("Value").ToString()

        Return header & introduction & detailsTable & body & conclusion

    End Function


    'Private Function GetStudentEmailBody(studentName As String, staffName As String,
    '                                   sanctionType As String, reason As String,
    '                                   oldDate As Date, newDate As Date
    '                                  ) As String
    'Dim header = ConfigHandler.studentEmailElements.Select("Name='HeaderText'")(0)("Value").ToString()
    'Dim introduction = ConfigHandler.studentEmailElements.Select("Name='Introduction'")(0)("Value").ToString()
    'Dim detailsTable = String.Format(
    '        ConfigHandler.studentEmailElements.Select("Name='DetailsTable'")(0)("Value").ToString(),
    '        studentName,
    '        staffName,
    '        sanctionType,
    '        reason,
    '        oldDate.ToString("yyyy-MM-dd"))
    'Dim body = String.Format(
    '        ConfigHandler.studentEmailElements.Select("Name='Body'")(0)("Value").ToString(),
    '        newDate.ToString("yyyy-MM-dd"))
    'Dim conclusion = ConfigHandler.studentEmailElements.Select("Name='Conclusion'")(0)("Value").ToString()

    'Return header & introduction & detailsTable & body & conclusion
    'End Function




    Private Sub SendEmailsNewSanctionCreated(sanction As DataRow, newSanctionDate As Date)

        Try
            ' Determine which staff member should be notified. Serious sanctions go to HoYs. 
            Dim staffEmailAddress As String = If(
                 Strings.Left(sanction("SanctionCode"), 9) = "Detention" _
                     Or sanction("SanctionCode") = "Internal Suspension",
             sanction("HoyEmail"),
             sanction("StaffEmail"))

            Dim originalSanctionDate As String = CDate(sanction("SanctionDate")).ToString("yyyy-MM-dd")
            Dim reason As String = sanction("StaffReason")
            Dim staffName As String = sanction("StaffName")
            Dim studentName As String = sanction("StudentName")
            Dim studentEmailAddress As String = sanction("StudentEmail")
            Dim sanctionType As String = sanction("SanctionCode")

            Dim mailClient As New Net.Mail.SmtpClient("ns1.woodcroft.sa.edu.au")
            mailClient.Port = 25

            Dim mailMessage As New Net.Mail.MailMessage()
            mailMessage.IsBodyHtml = True
            mailMessage.From = New Net.Mail.MailAddress(ConfigHandler.systemEmailSender)
            mailMessage.Subject = "Notice of non-attendance for " & studentName

            For Each recipient In ConfigHandler.dataManagementEmails
                mailMessage.Bcc.Add(recipient)
            Next


            ' ====================================================================================
            ' SEND TO STAFF
            mailMessage.Body = GetStaffEmailBody(studentName:=studentName,
                        staffName:=staffName, sanctionType:=sanctionType,
                        reason:=reason, oldDate:=originalSanctionDate,
                        newDate:=newSanctionDate)
            'mailMessage.To.Add("selby_b@woodcroft.sa.edu.au")
            mailMessage.To.Add(staffEmailAddress)
            mailClient.Send(mailMessage)
            ' ====================================================================================

            ' ====================================================================================
            ' SEND TO STUDENT
            ' [2022.11.11 selby_b] Discussed with Adrian and he would prefer staff to notify
            ' students personally, to prevent notifications from being ignored by students. 
            ' Disabling for now, may put back in later or remove. 
            '
            'mailMessage.Body = GetStudentEmailBody(studentName:=studentName,
            '            staffName:=staffName, sanctionType:=sanctionType,
            '            reason:=reason, oldDate:=originalSanctionDate,
            '            newDate:=newSanctionDate)
            'mailMessage.To.Add(studentEmailAddress)
            'mailClient.Send(mailMessage)
            ' ====================================================================================

        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=True, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True, source:="BMS Admin - ReplaceAndNotifyAbsences")
        End Try


    End Sub


#End Region

End Class

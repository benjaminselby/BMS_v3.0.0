Imports System.Configuration
Imports BMS.BmsShared


Public Class StudentSanctions

    Private currentStudent As Student


    Private Sub StudentSanctions_Load(sender As System.Object, e As System.EventArgs) _
                Handles MyBase.Load

        Try
            Me.WindowState = FormWindowState.Maximized

            studentListCbx.Items.Clear()

            Dim allStudents = DataHandler.getDataTable(
                procedureName:=ConfigurationManager.AppSettings("GetAllStudentsProc"))

            For Each studentRow As DataRow In allStudents.Rows

                studentListCbx.Items.Add(New Student(
                    FirstName:=studentRow("Preferred"),
                    Surname:=studentRow("Surname"),
                    YearLevel:=studentRow("YearLevel"),
                    ID:=studentRow("ID"),
                    emailAddress:=studentRow("EmailAddress"),
                    networkLogin:=studentRow("NetworkLogin")))

            Next

        Catch ex As Exception
            ErrorHandler.HandleError(
                exception:=ex, rethrow:=False, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True,
                source:="BMS Admin - StudentSanctions",
                message:="Could not obtain student list data.")
        End Try

    End Sub


    ' ======================================================================================
    ' FUNCTIONS
    ' ======================================================================================
#Region "Functions"


    Private Sub ConfirmAndSaveChanges()

        If currentStudent IsNot Nothing _
                AndAlso sanctionListDgv.DataSource IsNot Nothing _
                AndAlso CType(sanctionListDgv.DataSource, DataTable).GetChanges IsNot Nothing Then

            If MessageBox.Show(text:=String.Format(
                        "Do you wish to save your changes to the sanctions for {0}?",
                        currentStudent.firstName),
                    buttons:=MessageBoxButtons.YesNo,
                    caption:="Save changes to Synergy",
                    icon:=MessageBoxIcon.Question) = MsgBoxResult.Yes Then

                SaveChanges()

            End If
        End If

    End Sub



    Private Sub PopulateSanctionsDgv()

        If studentListCbx.SelectedItem IsNot Nothing Then

            sanctionListDgv.DataSource = Nothing

            Dim queryParameters = New Dictionary(Of String, String) _
                From {{"StudentId", CType(studentListCbx.SelectedItem, Student).id}}

            sanctionListDgv.DataSource = DataHandler.getDataTable(
                procedureName:=ConfigurationManager.AppSettings("GetSanctionsForStudentProc"),
                queryParameters:=queryParameters)

            sanctionListDgv.Columns(0).Visible = False ' Sanction Sequence Number
            sanctionListDgv.Columns(1).Visible = False ' Staff ID
            sanctionListDgv.Columns(4).Visible = False ' Student ID

            ' Limit the width of the StaffReason comment column because it can 
            ' contain some long text.
            sanctionListDgv.Columns(8).Width = 400

            For Each column As DataGridViewColumn In sanctionListDgv.Columns
                With column
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                    .Resizable = DataGridViewTriState.True
                    .ReadOnly = If(column.Index <= 3, True, False)
                End With
            Next

        End If

    End Sub



    Private Sub SaveChanges()

        Try

            Dim sanctionChanges = CType(sanctionListDgv.DataSource, DataTable).GetChanges()

            If sanctionChanges IsNot Nothing Then

                For Each sanctionRow As DataRow In sanctionChanges.Rows

                    Dim updateParameters = New Dictionary(Of String, String) _
                        From {
                            {"Seq", sanctionRow("SanctionSeq").ToString()},
                            {"StaffID", sanctionRow("StaffId").ToString()},
                            {"StudentID", sanctionRow("StudentId").ToString()},
                            {"SanctionType", sanctionRow("ReasonType").ToString()},
                            {"SanctionDate", CType(sanctionRow("SanctionDate"), DateTime).ToString("yyyy-MM-dd")},
                            {"ClassCode", sanctionRow("ClassCode").ToString()},
                            {"CurrentGrade", sanctionRow("CurrentGrade").ToString()},
                            {"SummativeTaskName", sanctionRow("SummativeTaskName").ToString()},
                            {"SummativeDueDate", sanctionRow("SummativeDueDate").ToString()},
                            {"Comment", sanctionRow("StaffReason").ToString()},
                            {"Completed", sanctionRow("Completed").ToString()},
                            {"Present", sanctionRow("Present").ToString()},
                            {"Removed", sanctionRow("Removed").ToString()},
                            {"HOYNotes", sanctionRow("HOYNotes").ToString()},
                            {"HOSNotes", sanctionRow("HOSNotes").ToString()},
                            {"CounsNotes", sanctionRow("CounsNotes").ToString()},
                            {"OtherNotes", sanctionRow("OtherNotes").ToString()},
                            {"ModifiedBy", ConfigHandler.currentUser.id}}

                    DataHandler.update(
                        procedureName:=ConfigurationManager.AppSettings("UpdateSanctionProc"),
                        commandParameters:=updateParameters)

                Next

            End If

        Catch ex As Exception

            ErrorHandler.HandleError(
                exception:=ex, rethrow:=False, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True,
                source:="BMS Admin - Student Sanctions",
                message:=ConfigHandler.dataSaveErrorMsg)

        End Try

    End Sub



#End Region
    ' ======================================================================================
    ' EVENT HANDLERS
    ' ======================================================================================
#Region "Event Handlers"



    Private Sub sanctionListDgv_DataError(
            sender As Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) _
            Handles sanctionListDgv.DataError

        MsgBox("The value you have just entered is not in the correct format. " _
               & "Please try again, or press ESC to cancel." & vbCrLf _
               & "If this problem persists, please contact Data Management for assistance.",
                MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Data Format Error")

    End Sub


    Private Sub StudentSanctions_ConfirmAndSaveChanges(
                sender As Object, e As System.Windows.Forms.FormClosingEventArgs) _
                Handles Me.FormClosing

        sanctionListDgv.CurrentCell = Nothing
        ConfirmAndSaveChanges()

    End Sub


    Private Sub studentListCbx_SelectedIndexChanged(
                sender As System.Object, e As System.EventArgs) _
                Handles studentListCbx.SelectedIndexChanged


        ConfirmAndSaveChanges()

        currentStudent = studentListCbx.SelectedItem

        PopulateSanctionsDgv()


    End Sub


#End Region



End Class
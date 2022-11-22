Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Reporting.WinForms
Imports System.Configuration
Imports BMS.BmsShared


Public Class BmsMain

    Private sanctionsForUser As SanctionsForUserForm


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        Me.Text += " v" + ConfigHandler.versionNumber


        For Each reason As String In ConfigHandler.sanctionReasons
            ReasonCbx.Items.Add(reason)
        Next


        ' The StudentsMbx is a custom control which enables better text-matching than an ordinary ComboBox. 
        ' It can't be successfully added to this form in the Design window because its code keeps disappearing 
        ' from the InitialiseComponent() function, so I have added it here. 

        Me.StudentsMbx = New BMS.MatchingComboBox()
        Me.StudentsMbx.FilterRule = Nothing
        Me.StudentsMbx.Font = New System.Drawing.Font("Gill Sans MT", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StudentsMbx.FormattingEnabled = True
        Me.StudentsMbx.Name = "StudentsMbx"
        Me.StudentsMbx.PropertySelector = Nothing
        Me.StudentsMbx.Location = New System.Drawing.Point(84, 232)
        Me.StudentsMbx.Size = New System.Drawing.Size(493, 29)
        Me.StudentsMbx.SuggestBoxHeight = 96
        Me.StudentsMbx.SuggestListOrderRule = Nothing
        Me.StudentsMbx.Anchor = AnchorStyles.Bottom + AnchorStyles.Left
        Me.StudentsMbx.TabIndex = 1
        Me.Controls.Add(Me.StudentsMbx)
        AddHandler StudentsMbx.SelectedIndexChanged, AddressOf StudentsMbx_SelectedIndexChanged
        Me.StudentsMbx.Select()


        PopulateStudentsMbx()
        PopulateStaffCbx()
        PopulateClassesCbx()


    End Sub



    ' ==================================================================================
    ' FUNCTIONS
    ' ==================================================================================
#Region "Functions"


    Private Function Reset()

        ' Clears the form in preparation for user to commence new sanction booking. 

        HideAllSanctionControls()

        StudentsMbx.SelectedItem = Nothing
        StudentsMbx.Text = Nothing

        ' Don't reset the Staff Member combobox selection to enable 
        ' multiple sanctions for the same staff member. 
        StaffMemberCbx.Hide()
        StaffMemberLbl.Hide()

        ClassLbl.Hide()
        ClassCbx.SelectedItem = Nothing
        ClassCbx.Text = Nothing
        ClassCbx.Hide()

        SanctionTypeLbl.Hide()
        SanctionTypeCbx.SelectedItem = Nothing
        SanctionTypeCbx.Text = Nothing
        SanctionTypeCbx.Hide()

        StudentPhotoPbx.Hide()
        ReportViewer1.Hide()

        SubmitBtn.Hide()

        Me.Refresh()

    End Function




    Private Sub PopulateStaffCbx()

        ' The Staff Member Combo Box defaults to the current user, but a user can change it if they 
        ' choose to set sanctions on behalf of another staff member (e.g. sometimes Alison Williams
        ' will do this).

        StaffMemberCbx.Items.Clear()

        Try

            Dim staffList As DataTable = ConfigHandler.staffList

            For Each staffMember As DataRow In staffList.Rows
                StaffMemberCbx.Items.Add(
                    New User(
                        networkLogin:=staffMember("NetworkLogin").ToString,
                        firstName:=staffMember("Preferred").ToString,
                        surname:=staffMember("Surname").ToString,
                        id:=staffMember("Id").ToString,
                        emailAddress:=staffMember("Email").ToString))
            Next


            ' Select the current user's entry in the Staff List by default. 
            StaffMemberCbx.SelectedIndex = -1
            For i = 0 To StaffMemberCbx.Items.Count - 1
                Dim IndexUser As User = TryCast(StaffMemberCbx.Items(i), User)
                If IndexUser IsNot Nothing AndAlso IndexUser.id = ConfigHandler.currentUser.id Then
                    StaffMemberCbx.SelectedIndex = i
                    Exit For
                End If
            Next

        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=True, sendEmail:=True,
                                     addToLogTable:=True, showMsgBox:=True)
        End Try

    End Sub



    Private Sub PopulateClassesCbx()

        ClassCbx.Items.Clear()
        ClassCbx.Items.Add("NOT APPLICABLE")

        Try

            ' =========================================================================================
            ' 1. Add classes for the current user to the class selection ComboBox.
            ' =========================================================================================

            For Each staffClass As DataRow In ConfigHandler.currentStaffClasses.Rows

                ClassCbx.Items.Add(New SchoolClass(
                        staffClass("ClassCode").ToString,
                        staffClass("Description").ToString))

            Next

            ClassCbx.Items.Add(ConfigHandler.comboBoxSpacerText)



            ' =========================================================================================
            ' 2. Append all classes below the current user's class list.  
            ' =========================================================================================

            ' Originally the system only included classes for the current user, but 
            ' there were so many requests from people to view other classes for various 
            ' reasons that I decided to include all classes as well. 

            For Each schoolClass As DataRow In ConfigHandler.allClasses.Rows
                ClassCbx.Items.Add(New SchoolClass(
                    schoolClass("ClassCode").ToString,
                    schoolClass("Description").ToString))
            Next

        Catch Ex As Exception
            ErrorHandler.HandleError(exception:=Ex, rethrow:=True, sendEmail:=True,
                            addToLogTable:=True, showMsgBox:=True)
        End Try

        ClassCbx.SelectedItem = Nothing

    End Sub



    Private Sub PopulateStudentsMbx()

        ' The student selector MatchingComboBox is populated with ALL students at the school. 

        StudentsMbx.Items.Clear()
        StudentsMbx.Text = ""
        StudentsMbx.SelectedItem = Nothing

        Try
            For Each studentDataRow As DataRow In ConfigHandler.allStudents.Rows
                StudentsMbx.Items.Add(New Student(
                    FirstName:=studentDataRow("Preferred").ToString,
                    Surname:=studentDataRow("Surname").ToString,
                    ID:=studentDataRow("ID").ToString,
                    emailAddress:=studentDataRow("EmailAddress").ToString,
                    networkLogin:=studentDataRow("NetworkLogin").ToString,
                    YearLevel:=CInt(studentDataRow("YearLevel"))))
            Next
        Catch Ex As Exception
            ErrorHandler.HandleError(exception:=Ex, rethrow:=True, sendEmail:=True,
                        addToLogTable:=True, showMsgBox:=True)
        End Try

    End Sub



    Private Sub PopulateSanctionsCbx()

        SanctionTypeCbx.Items.Clear()
        SanctionTypeCbx.Text = ""
        SanctionTypeCbx.SelectedItem = Nothing

        If StudentsMbx.SelectedItem Is Nothing Then
            ' No student selected - Clear sanctions combo box.
            SanctionTypeCbx.Items.Clear()
        End If

        Try

            ' Sanction types can vary based on the year level of the selected student. 

            Dim sanctionTypes As List(Of String) =
                ConfigHandler.availableSanctionTypes(CType(StudentsMbx.SelectedItem, Student).yearLevel)

            If sanctionTypes.Count < 1 Then
                Throw New Exception(String.Format(
                    "Could not get sanction types from config table for student {0} {1} [{2}]",
                    CType(StudentsMbx.SelectedItem, Student).firstName,
                    CType(StudentsMbx.SelectedItem, Student).surname,
                    CType(StudentsMbx.SelectedItem, Student).id))
            End If

            For Each sanctionType As String In sanctionTypes
                SanctionTypeCbx.Items.Add(sanctionType)
            Next

        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=True, sendEmail:=True,
                            addToLogTable:=True, showMsgBox:=True, source:="PopulateSanctionsCbx")
        End Try

    End Sub



    Private Sub PopulateDateSelectorCbx()

        DateSelectorCbx.Items.Clear()

        If SanctionTypeCbx.SelectedItem = Nothing Then
            Return
        End If

        Try

            Dim alreadyBooked As Boolean
            Dim bookedByUser As User

            Dim availableSanctionDates As DataTable = ConfigHandler.availableSanctionDatesForStudent(
                studentId:=CType(StudentsMbx.SelectedItem, Student).id,
                sanctionType:=SanctionTypeCbx.Text)

            For Each sanctionDate As DataRow In availableSanctionDates.Rows

                alreadyBooked = If(IsDBNull(sanctionDate("StaffId")), False, True)
                bookedByUser = If(IsDBNull(sanctionDate("StaffId")),
                        Nothing,
                        New User(
                            id:=sanctionDate("StaffId"),
                            firstName:=sanctionDate("StaffPreferred"),
                            surname:=sanctionDate("StaffSurname"),
                            emailAddress:=sanctionDate("StaffEmail")))

                DateSelectorCbx.Items.Add(New SanctionDate(
                    day:=Convert.ToDateTime(sanctionDate("Date").ToString).Date,
                    alreadyBooked:=alreadyBooked,
                    bookedBy:=bookedByUser))

            Next

        Catch Ex As Exception
            ErrorHandler.HandleError(exception:=Ex, rethrow:=True, sendEmail:=True,
                            addToLogTable:=True, showMsgBox:=True)
        End Try

    End Sub




    Private Sub HideAllSanctionControls()

        ' This hides the controls which are specialised relating to particular sanction types. 
        CurrentGradeLbl.Hide()
        CurrentGradeCbx.Text = ''
        CurrentGradeCbx.SelectedItem = Nothing
        CurrentGradeCbx.Hide()

        DueDateCbx.Items.Clear()
        DueDateLbl.Hide()
        DueDateCbx.SelectedItem = Nothing
        DueDateCbx.Text = Nothing
        DueDateCbx.Hide()

        TaskNameLbl.Hide()
        TaskNameTbx.Text = Nothing

        TaskNameTbx.Hide()

        IncidentDateLbl.Hide()
        SanctionDateLbl.Hide()
        DateSelectorCbx.SelectedItem = Nothing
        DateSelectorCbx.Text = Nothing
        DateSelectorCbx.Hide()

        ReasonLbl.Hide()
        ReasonCbx.SelectedItem = Nothing
        ReasonCbx.Text = Nothing
        ReasonCbx.Hide()

        InformationOnlyChk.Checked = False
        InformationOnlyLbl.Hide()
        InformationOnlyChk.Hide()

        CommentLbl.Hide()
        CommentTbx.Text = Nothing
        CommentTbx.Hide()

        Me.Refresh()

    End Sub

#End Region




    ' ==================================================================================
    ' EVENT HANDLERS
    ' ==================================================================================
#Region "Event Handlers"


    Private Sub SanctionTypeCbx_SelectedIndexChanged(
            sender As Object, e As EventArgs) Handles SanctionTypeCbx.SelectedIndexChanged

        If SanctionTypeCbx.SelectedIndex = -1 Then
            Return
        End If

        ' Hide the specialised controls by default, reveal later based on chosen sanction type. 
        HideAllSanctionControls()

        Try

            Select Case (SanctionTypeCbx.Text)

                Case "Report of Concern"

                    CurrentGradeLbl.Show()
                    CurrentGradeCbx.Show()

                Case "Non Submission: Summative"

                    DueDateCbx.Items.Clear()
                    DueDateCbx.Items.Add("NOT APPLICABLE")
                    ' Add previous dates to the Due Date combobox, excluding weekends (DayOfWeek = 6 or 0)
                    ' (This should be moved to a DB function but I doubt it's used much so 
                    ' that might be a waste of time.) 
                    For dayCounter = -21 To 0
                        If (Now().AddDays(dayCounter)).DayOfWeek <> 6 And (Now().AddDays(dayCounter)).DayOfWeek <> 0 Then
                            DueDateCbx.Items.Add(String.Format("{0:ddd, d MMM}", Now().AddDays(dayCounter)))
                        End If
                    Next

                    TaskNameLbl.Show()
                    TaskNameTbx.Show()
                    DueDateLbl.Show()
                    DueDateCbx.Show()

                Case Else

                    If CType(StudentsMbx.SelectedItem, Student).yearLevel <= ConfigHandler.juniorSchoolFinalYear Then
                        IncidentDateLbl.Show()
                        SanctionDateLbl.Hide()
                    Else
                        IncidentDateLbl.Hide()
                        SanctionDateLbl.Show()
                    End If

                    DateSelectorCbx.Show()
                    ReasonLbl.Show()
                    ReasonCbx.Show()
                    InformationOnlyLbl.Show()
                    InformationOnlyLbl.Show()
                    InformationOnlyChk.Show()

                    PopulateDateSelectorCbx()
            End Select

        Catch Ex As Exception
            ErrorHandler.HandleError(exception:=Ex, rethrow:=True, sendEmail:=True,
                            addToLogTable:=True, showMsgBox:=True)
        End Try

        CommentLbl.Show()
        CommentTbx.Show()
        SubmitBtn.Show()

    End Sub



    Private Sub StudentsMbx_SelectedIndexChanged(sender As Object, e As EventArgs)

        ' Added as event handler in control declaration code. 

        If StudentsMbx.SelectedItem Is Nothing Then
            Return
        End If

        HideAllSanctionControls()

        Try

            ' Display student photo.

            Dim studentPhoto As Bitmap = ConfigHandler.studentPhoto(
                studentId:=CType(StudentsMbx.SelectedItem, Student).id)

            If studentPhoto Is Nothing Then
                ' No photo found: empty and hide the student photo picture box. 
                StudentPhotoPbx.Image = Nothing
                StudentPhotoPbx.Hide()
            Else
                StudentPhotoPbx.Image = studentPhoto
                StudentPhotoPbx.Show()
            End If


            ' Display summary report of sanctions for the selected student. 

            Dim Params As New List(Of ReportParameter)
            Params.Add(New ReportParameter("StudentID", CType(StudentsMbx.SelectedItem, Student).id.ToString))
            Me.ReportViewer1.ServerReport.SetParameters(Params)
            Me.ReportViewer1.RefreshReport()
            Me.ReportViewer1.Visible = True

            ' Different sanction types are available for different student year levels. 
            ' So we re-populate the available sanctions list every time a new student is selected. 
            PopulateSanctionsCbx()

            StaffMemberLbl.Show()
            StaffMemberCbx.Show()
            ClassLbl.Show()
            ClassCbx.Show()
            SanctionTypeLbl.Show()
            SanctionTypeCbx.Show()

        Catch Ex As Exception
            ErrorHandler.HandleError(exception:=Ex, rethrow:=True, sendEmail:=True,
                            addToLogTable:=True, showMsgBox:=True)
        End Try

    End Sub



    Private Sub SubmitBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                    Handles SubmitBtn.Click

        Try

            ' =======================================================================================
            ' FORM VALIDATION
            ' =======================================================================================
            If StaffMemberCbx.SelectedItem Is Nothing Then
                ' This should never happen (only if current user is not active in [dbo.Staff] Synergy table). 
                MessageBox.Show("A sanction cannot be created without a staff name. Please select a staff member.",
                "No Staff Member Selected!",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            ElseIf SanctionTypeCbx.Text = "" Then
                MessageBox.Show("Please select the category of sanction.", "No Category Selected!",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            ElseIf StudentsMbx.Text = "" Then
                MessageBox.Show("Please select a student.", "No Student Selected!",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            ElseIf ClassCbx.Text = "" Or ClassCbx.Text = ConfigHandler.comboBoxSpacerText Then
                MessageBox.Show("Please select an appropriate class or NOT APPLICABLE.", "No Class Selected!",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Select Case (SanctionTypeCbx.Text)

                Case "Report of Concern"

                    If CurrentGradeCbx.Text = "" Then
                        MessageBox.Show("Please select the student's current grade or NOT APPLICABLE.", "No Current Grade Selected!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    ElseIf SanctionTypeCbx.Text = "" Then
                        MessageBox.Show("Please select the category of sanction.", "No Category Selected!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    End If

                Case "Non Submission: Summative"

                    If TaskNameTbx.Text = "" Then
                        MessageBox.Show("Please enter the title of the summative task that has not been submitted.", "No Task Entered!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    ElseIf DueDateCbx.Text = "" Then
                        MessageBox.Show("Please select the date the task was due for submission.", "No Due Date Selected!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    End If

                Case Else

                    ' Catchups and both detention types. 

                    If DateSelectorCbx.Text = "" Then
                        MessageBox.Show("Please select the date of the sanction.", "No Date Selected!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    ElseIf CType(DateSelectorCbx.SelectedItem, SanctionDate).alreadyBooked Then
                        MessageBox.Show("Please select a date that is not already booked by another staff member.", "Booking Clash!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    ElseIf (SanctionTypeCbx.Text = "Catch up class" Or SanctionTypeCbx.Text = "Detention (Level 1: Lunchtime)") _
                            And (CDate(DateSelectorCbx.Text).DayOfWeek = 2 Or CDate(DateSelectorCbx.Text).DayOfWeek = 4) _
                            And InformationOnlyChk.Checked = False Then
                        MessageBox.Show("Catch up classes and lunchtime detentions can only be booked on a Tuesday or Thursday if administered in full by the setting teacher. Please signify that this is the case by ticking the 'Information Only' box, or change the date of the booking to a Monday, Wednesday or Friday.", "Day of Week Issue!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    ElseIf Trim(ReasonCbx.Text) = "" Then
                        MessageBox.Show("Please select the reason for the sanction.", "No Reason Selected!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    End If

            End Select

            ' If we reach this point then the form is valid. 


            ' =========================================================================================
            ' CREATE NEW SANCTION RECORD IN DATABASE
            ' =========================================================================================

            ' Organise the query parameters which will be used to update the database. 
            ' Most fields are generic, some are set only for specific sanction types. 
            Dim summativeDueDate As String = ""
            Dim summativeTaskName As String = ""
            Dim sanctionDate As String = ""
            Dim sanctionCode As String = ""
            Dim reasonType As String = ""
            Dim currentGrade As String = ""

            Select Case (SanctionTypeCbx.Text)

                Case "Report of Concern"

                    sanctionCode = "Report of Concern"
                    reasonType = ReasonCbx.Text
                    currentGrade = CurrentGradeCbx.Text

                Case "Non Submission: Summative"

                    summativeTaskName = TaskNameTbx.Text
                    sanctionCode = "Non Submission"
                    reasonType = "Summative Work"

                    If DueDateCbx.Text = "NOT APPLICABLE" Or DueDateCbx.Text = "" Then
                        summativeDueDate = CDate("1900-1-1")
                    Else
                        summativeDueDate = CType(DueDateCbx.SelectedItem, Date).ToString("yyyy-MM-dd")
                    End If

                Case Else

                    ' Detentions and catch up classes.

                    sanctionCode = SanctionTypeCbx.Text
                    reasonType = ReasonCbx.Text
                    sanctionDate = CType(DateSelectorCbx.Text, Date).ToString("yyyy-MM-dd")

            End Select


            Dim insertParams As Dictionary(Of String, String) =
                New Dictionary(Of String, String) From {
                    {"StaffId", CType(StaffMemberCbx.SelectedItem, User).id},
                    {"StudentId", CType(StudentsMbx.SelectedItem, Student).id},
                    {"SanctionCode", sanctionCode},
                    {"SanctionReason", reasonType},
                    {"SanctionDate", sanctionDate},
                    {"ClassCode", ClassCbx.Text},
                    {"CurrentGrade", currentGrade},
                    {"SummativeTaskName", summativeTaskName},
                    {"SummativeDueDate", summativeDueDate},
                    {"Comment", CommentTbx.Text},
                    {"InformationOnly", CInt(InformationOnlyChk.Checked)},
                    {"ModifiedBy", ConfigHandler.currentUser.id}}

            Dim rowsInserted As Integer = DataHandler.insert(
                procedureName:=ConfigurationManager.AppSettings("CreateSanctionProc"),
                commandParameters:=insertParams)


            If rowsInserted <= 0 Then
                MessageBox.Show("There was a problem saving this sanction to the Synergy database. " _
                    & "Make sure that you are currently connected to the Woodcroft network." _
                    & "If this problem persists, please notify the Data Manager.",
                    "Sanction not saved!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If


            If MessageBox.Show("Booking saved, would you like to make another?",
                "Booking Saved!",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                Reset()
            Else
                Me.Close()
            End If

        Catch Ex As Exception
            ErrorHandler.HandleError(exception:=Ex, rethrow:=True, sendEmail:=True,
                            addToLogTable:=True, showMsgBox:=True)
        End Try

    End Sub



    Private Sub SanctionDateCbx_SelectedIndexChanged(
            sender As Object, e As EventArgs) Handles DateSelectorCbx.SelectedIndexChanged

        If DateSelectorCbx.SelectedItem Is Nothing Then
            Return
        End If

        Dim selectedDate As SanctionDate = CType(DateSelectorCbx.SelectedItem, SanctionDate)
        If selectedDate.alreadyBooked Then
            MsgBox(
                "Sorry, that date has already been booked by another staff member. " +
                "Please select a different date.",
                MsgBoxStyle.Exclamation)

            DateSelectorCbx.SelectedIndex = -1
            DateSelectorCbx.SelectedItem = Nothing
        End If

    End Sub



    Private Sub ViewCurrentBookingsBtn_Click(sender As Object, e As EventArgs) _
            Handles ViewCurrentBookingsBtn.Click

        sanctionsForUser = New SanctionsForUserForm()
        sanctionsForUser.ShowDialog()

    End Sub

#End Region


End Class

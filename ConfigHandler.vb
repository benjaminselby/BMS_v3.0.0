Imports System.Data.SqlClient
Imports System.Configuration
Imports BMS.BmsShared



Class ConfigHandler


    Public Const comboBoxSpacerText As String = "==============================================="
    Public Const dateFormat As String = "dd/MM/yyyy"

    Public Const dataSaveErrorMsg As String =
        "Database Error: Your changes may not have been saved correctly." & vbCrLf _
        & "Please make sure you are connected to the Woodcroft network." & vbCrLf _
        & "If this problem persists, contact Data Management for assistance." & vbCrLf _
        & "========" & vbCrLf _
        & "EXCEPTION DETAILS:" & vbCrLf




    Private Shared _systemEmailSender As String = Nothing
    Public Shared ReadOnly Property systemEmailSender As String
        Get
            If _systemEmailSender Is Nothing Then
                Dim queryParams As Dictionary(Of String, String) = New Dictionary(Of String, String) _
                From {
                    {"Key1", "config"},
                    {"Key2", "systemEmailSender"}}

                _systemEmailSender = DataHandler.getScalarValue(
                        command:=ConfigurationManager.AppSettings("GetConfigValuesProc"),
                        queryParameters:=queryParams,
                        columnName:="value")
            End If

            Return _systemEmailSender
        End Get
    End Property




    Private Shared _mailServerName As String = Nothing
    Public Shared ReadOnly Property mailServerName As String
        Get
            If _mailServerName Is Nothing Then

                Dim queryParams As Dictionary(Of String, String) = New Dictionary(Of String, String) _
                From {
                    {"Key1", "config"},
                    {"Key2", "mailServerName"}}

                _mailServerName = DataHandler.getScalarValue(
                        command:=ConfigurationManager.AppSettings("GetConfigValuesProc"),
                        queryParameters:=queryParams,
                        columnName:="value")

            End If

            Return _mailServerName
        End Get
    End Property





    Private Shared _dataManagementEmails As List(Of String) = Nothing
    Public Shared ReadOnly Property dataManagementEmails As List(Of String)
        Get
            If _dataManagementEmails Is Nothing Then

                Dim queryParams As Dictionary(Of String, String) = New Dictionary(Of String, String) _
                From {
                    {"Key1", "config"},
                    {"Key2", "dataManagementEmail"}}

                _dataManagementEmails = DataHandler.getList(
                        procedureName:=ConfigurationManager.AppSettings("GetConfigValuesProc"),
                        queryParameters:=queryParams,
                        columnName:="value")

            End If

            Return _dataManagementEmails
        End Get
    End Property


    Private Shared _dbConnectionString As String = Nothing
    Public Shared ReadOnly Property dbConnectionString As String
        Get
            If _dbConnectionString Is Nothing Then
                _dbConnectionString = ConfigurationManager.ConnectionStrings("SynergyOne").ToString()
            End If
            Return _dbConnectionString
        End Get
    End Property



    Private Shared _versionNumber As String = Nothing
    Public Shared ReadOnly Property versionNumber As String
        Get
            If _versionNumber Is Nothing Then
                _versionNumber = ConfigurationManager.AppSettings("VersionNumber")
            End If
            Return _versionNumber
        End Get
    End Property



    Private Shared _staffEmailElements As DataTable = Nothing
    Public Shared ReadOnly Property staffEmailElements As DataTable
        Get
            If _staffEmailElements Is Nothing Then

                Dim queryParameters = New Dictionary(Of String, String) _
                From {{"Type", "Staff"}}

                _staffEmailElements = DataHandler.getDataTable(
                        procedureName:=ConfigurationManager.AppSettings("GetNonAttendanceEmailElementsProc"),
                        queryParameters:=queryParameters)

            End If

            Return _staffEmailElements
        End Get
    End Property




    Private Shared _studentEmailElements As DataTable = Nothing
    Public Shared ReadOnly Property studentEmailElements As DataTable
        Get
            If _studentEmailElements Is Nothing Then

                Dim queryParameters = New Dictionary(Of String, String) _
                From {{"Type", "Student"}}

                _studentEmailElements = DataHandler.getDataTable(
                        procedureName:=ConfigurationManager.AppSettings("GetNonAttendanceEmailElementsProc"),
                        queryParameters:=queryParameters)

            End If

            Return _studentEmailElements
        End Get
    End Property





    Private Shared _currentUser As User = Nothing
    Public Shared ReadOnly Property currentUser As User
        Get
            If _currentUser Is Nothing Then

                ' Extracts information about the current user from the DB. Uses this information 
                ' to create a User object. If no user information can be obtained from DB, 
                ' returns nothing. 

                Dim usernamePrefix As String

                Dim queryParams As Dictionary(Of String, String) = New Dictionary(Of String, String) _
                    From {
                        {"Key1", "config"},
                        {"Key2", "usernamePrefix"}}

                usernamePrefix = DataHandler.getScalarValue(
                    command:=ConfigurationManager.AppSettings("GetConfigValuesProc"),
                    queryParameters:=queryParams,
                    columnName:="value")

                Dim currentUserWinIdentity As Security.Principal.WindowsIdentity _
                    = Security.Principal.WindowsIdentity.GetCurrent()


                Dim currentUserNetworkLogin As String _
                    = Strings.Right(currentUserWinIdentity.Name,
                                    Len(currentUserWinIdentity.Name) - Len(usernamePrefix))

                ' ------------------------------------------------------------------------------
                ' If testing, set a debug user Network Login here. 
                'MessageBox.Show("RUNNING IN DEBUG MODE!")
                'currentUserNetworkLogin = "elliston_a"
                ' ------------------------------------------------------------------------------

                queryParams = New Dictionary(Of String, String) _
                    From {{"NetworkLogin", currentUserNetworkLogin}}

                Dim userInfo As Dictionary(Of String, String)
                userInfo = DataHandler.getDictionary(
                            procedureName:=ConfigurationManager.AppSettings("GetUserDetailsFromNetworkLoginProc"),
                            queryParameters:=queryParams)

                If userInfo Is Nothing Then
                    Throw New Exception(
                        "Could not find information for current user with network login " _
                        + currentUserNetworkLogin.ToUpper() + " in database.")
                End If

                Dim newUser = New User(
                        networkLogin:=currentUserNetworkLogin,
                        firstName:=userInfo("Preferred").ToString(),
                        surname:=userInfo("Surname").ToString(),
                        id:=userInfo("ID").ToString(),
                        emailAddress:=userInfo("OccupEmail").ToString())

                _currentUser = newUser

            End If

            Return _currentUser

        End Get
    End Property



    Private Shared _juniorSchoolFinalYear As Integer? = Nothing
    Public Shared ReadOnly Property juniorSchoolFinalYear As Integer
        Get
            If _juniorSchoolFinalYear Is Nothing Then

                _juniorSchoolFinalYear = Integer.Parse(DataHandler.getScalarValue(
                    command:=ConfigurationManager.AppSettings("GetJuniorSchoolFinalYearProc")))
            End If

            Return _juniorSchoolFinalYear
        End Get
    End Property




    Private Shared _sanctionReasons As List(Of String) = Nothing
    Public Shared ReadOnly Property sanctionReasons As List(Of String)
        Get
            If _sanctionReasons Is Nothing Then

                Dim queryParams As Dictionary(Of String, String) = New Dictionary(Of String, String) _
                    From {{"Key1", "Reason"}}

                ' The stored procedure returns these values in correct display order. 
                ' I'm hoping that this will be reliable in the future... 
                _sanctionReasons = DataHandler.getList(
                    procedureName:=ConfigurationManager.AppSettings("GetConfigValuesProc"),
                    queryParameters:=queryParams,
                    columnName:="value")

            End If
            Return _sanctionReasons
        End Get
    End Property



    Private Shared _staffList As DataTable = Nothing
    Public Shared ReadOnly Property staffList As DataTable
        Get
            If _staffList Is Nothing Then

                _staffList = DataHandler.getDataTable(
                    procedureName:=ConfigurationManager.AppSettings("GetStaffListProc"))

            End If
            Return _staffList
        End Get
    End Property




    Private Shared _currentStaffClasses As DataTable = Nothing
    Public Shared ReadOnly Property currentStaffClasses As DataTable
        Get
            If _currentStaffClasses Is Nothing Then

                Dim queryParams As Dictionary(Of String, String) = New Dictionary(Of String, String) _
                    From {{"StaffId", ConfigHandler.currentUser.id}}

                _currentStaffClasses = DataHandler.getDataTable(
                    procedureName:=ConfigurationManager.AppSettings("GetClassesForStaffProc"),
                    queryParameters:=queryParams)

            End If
            Return _currentStaffClasses
        End Get
    End Property




    Private Shared _allClasses As DataTable = Nothing
    Public Shared ReadOnly Property allClasses As DataTable
        Get
            If _allClasses Is Nothing Then

                _allClasses = DataHandler.getDataTable(
                    procedureName:=ConfigurationManager.AppSettings("GetAllClassesProc"))

            End If
            Return _allClasses
        End Get
    End Property



    Private Shared _allStudents As DataTable = Nothing
    Public Shared ReadOnly Property allStudents As DataTable
        Get
            If _allStudents Is Nothing Then

                _allStudents = DataHandler.getDataTable(
                    procedureName:=ConfigurationManager.AppSettings("GetAllStudentsProc"))

            End If
            Return _allStudents
        End Get
    End Property



    Public Shared ReadOnly Property availableSanctionTypes(yearLevel As Integer) As List(Of String)
        Get
            Dim queryParams As Dictionary(Of String, String) =
                New Dictionary(Of String, String) From {
                    {"Key1", "SanctionType"},
                    {"Key2", yearLevel.ToString}}

            Return DataHandler.getList(
                    procedureName:=ConfigurationManager.AppSettings("GetConfigValuesProc"),
                    queryParameters:=queryParams)
        End Get
    End Property


    Public Shared ReadOnly Property availableSanctionDatesForStudent(
                                studentId As String, sanctionType As String) _
                                As DataTable
        ' Sanction dates can vary based on sanction type (ie. some kinds of sanction 
        ' bookings may only be made for certain days in the future) and also based on 
        ' student year level (e.g. Junior School requested that sanction dates be in THE PAST
        ' because they use them only as a record of when a punishment was applied, not 
        ' to book a sanction in the future. 
        Get

            Dim queryParams As Dictionary(Of String, String) =
                New Dictionary(Of String, String) From {
                    {"StudentId", studentId},
                    {"SanctionType", sanctionType}}

            Return DataHandler.getDataTable(
                    procedureName:=ConfigurationManager.AppSettings("GetSanctionDatesForStudentProc"),
                    queryParameters:=queryParams)

        End Get
    End Property



    Public Shared ReadOnly Property studentPhoto(studentId As String) As Bitmap
        Get
            Dim queryParams As Dictionary(Of String, String) =
                New Dictionary(Of String, String) From {{"ID", studentId}}

            Return DataHandler.getImage(
                procedureName:=ConfigurationManager.AppSettings("GetPhotoDataForUserProc"),
                queryParameters:=queryParams,
                columnName:="ImageData")

        End Get
    End Property


    Public Shared ReadOnly Property currentSanctionsForUser() As DataTable
        Get
            Dim queryParams As Dictionary(Of String, String) = New Dictionary(Of String, String) _
                From {{"StaffId", ConfigHandler.currentUser.id}}

            Return DataHandler.getDataTable(
                procedureName:=ConfigurationManager.AppSettings("GetCurrentSanctionsForStaff"),
                queryParameters:=queryParams)
        End Get
    End Property



    Public Shared ReadOnly Property sanctionBookingsForDate(
                        studentId As String, sanctionType As String, sanctionDate As Date
                        ) As DataTable
        Get

            Dim queryParams As Dictionary(Of String, String) =
                New Dictionary(Of String, String) From {
                    {"StudentId", studentId},
                    {"SanctionType", sanctionType},
                    {"SanctionDate", sanctionDate.ToString("yyyy-MM-dd")}}

            Return DataHandler.getDataTable(
                procedureName:=ConfigurationManager.AppSettings("GetSanctionBookingsForDateProc"),
                queryParameters:=queryParams)

        End Get
    End Property


End Class

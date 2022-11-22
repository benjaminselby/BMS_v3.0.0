Imports System.Configuration
Imports BMS.BmsShared


Public Class BmsAdminMain

    Private intScreenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
    Private intScreenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
    Private decWidthRatio As Decimal
    Private decHeightRatio As Decimal


    Private Sub Form_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Me.Text += " v" + ConfigHandler.versionNumber

        Try

            ' Scale form components based on display resolution. 

            decWidthRatio = FormatNumber(intScreenWidth / Me.Width, 3)
            decHeightRatio = FormatNumber(intScreenHeight / Me.Height, 3)

            Me.WindowState = FormWindowState.Maximized

            For Each ctrl As Control In Me.Controls

                ResizeRelocateControl(ctrl)

                If TypeOf ctrl Is GroupBox Then
                    For Each childCtrl As Control In ctrl.Controls
                        ResizeRelocateControl(childCtrl)
                    Next
                End If

            Next

            ShowNamesBtn.Left = SanctionCountsChartRpt.Left + SanctionCountsChartRpt.Width - ShowNamesBtn.Width
            ShowNamesBtn.Top = SanctionCountsChartRpt.Top + SanctionCountsChartRpt.Height + 14

            ShowUnhandledReportsOfConcernCount()

            Me.SanctionCountsChartRpt.RefreshReport()

        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=True, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True, source:="BMS Admin - Main")
        End Try

    End Sub



    ' =======================================================================================
    ' FUNCTIONS
    ' =======================================================================================


#Region "Functions"



    Private Sub ResizeRelocateControl(ByRef ctl As Control)

        Dim intOriginalWidth As Integer = ctl.Width
        Dim intOriginalHeight As Integer = ctl.Height

        ctl.Font = New Font(ctl.Font.Name, CInt(ctl.Font.Size * decHeightRatio), ctl.Font.Style)
        ctl.Width = CInt(ctl.Width * decWidthRatio)
        ctl.Height = CInt(ctl.Height * decHeightRatio)
        ctl.Location = New Point(CInt(ctl.Location.X * decWidthRatio), CInt(ctl.Location.Y * decHeightRatio))

    End Sub


    Public Sub ShowUnhandledReportsOfConcernCount()

        ' Add number of unhandled Reports of Concern to label under RoC button.
        Dim nReportsUnhandled As Integer = DataHandler.getScalarValue(
                command:=ConfigurationManager.AppSettings("GetUnhandledReportsOfConcernCountProc"))

        UnhandledReportsOfConcernLbl.Text = nReportsUnhandled & " currently unhandled"
        UnhandledReportsOfConcernLbl.Left = ReportsOfConcernBtn.Left _
                    + ReportsOfConcernBtn.Width - UnhandledReportsOfConcernLbl.Width

        Me.Refresh()

    End Sub

#End Region



    ' =======================================================================================
    ' EVENT HANDLERS
    ' =======================================================================================


#Region "EventHandlers"


    Private Sub PrepareEmailRunBtn_Click(sender As System.Object, e As System.EventArgs) _
                Handles PrepareEmailRunBtn.Click
        ReplaceAndNotifyAbsences.Show()
    End Sub


    Private Sub ReportsOfConcernBtn_Click(sender As System.Object, e As System.EventArgs) _
                Handles ReportsOfConcernBtn.Click

        Dim reportsOfConcern As New ReportsOfConcern()
        reportsOfConcern.ShowDialog(Me)
        reportsOfConcern.Dispose()

    End Sub


    Private Sub MakeBookingBtn_Click(sender As System.Object, e As System.EventArgs) Handles MakeBookingBtn.Click

        Dim bms As New BmsMain()
        bms.ShowDialog(Me)
        bms.Dispose()

    End Sub


    Private Sub ElectronicRollBtn_Click(sender As Button, e As System.EventArgs) _
       Handles InternalSuspensionBtn.Click,
                CatchUpClassBtn.Click,
                DetentionLunchBtn.Click,
                DetentionAfterSchoolBtn.Click

        Dim electronicRoll As New ElectronicRoll()
        electronicRoll.sanctionType = sender.Tag.ToString()
        electronicRoll.ShowDialog(Me)
        electronicRoll.Dispose()

    End Sub


    Private Sub StudentProfileBtn_Click(sender As System.Object, e As System.EventArgs) Handles StudentProfileBtn.Click

        Dim studentProfile As New StudentProfile()
        studentProfile.ShowDialog(Me)
        studentProfile.Dispose()

    End Sub


    Private Sub ShowNamesBtn_Click(sender As System.Object, e As System.EventArgs) Handles ShowNamesBtn.Click
        If ShowNamesBtn.Text = "Show Names" Then
            SanctionCountsChartRpt.ServerReport.ReportPath = "/Behaviour Monitoring System/70DaysNamed"
            SanctionCountsChartRpt.RefreshReport()
            ShowNamesBtn.Text = "Hide Names"
        Else
            SanctionCountsChartRpt.ServerReport.ReportPath = "/Behaviour Monitoring System/70DaysNoNames"
            SanctionCountsChartRpt.RefreshReport()
            ShowNamesBtn.Text = "Show Names"
        End If
    End Sub



    Private Sub TutorGroupsBtn_Click(sender As System.Object, e As System.EventArgs) _
            Handles TutorGroupsBtn.Click

        Dim tutorGroup As New TutorGroup()
        tutorGroup.ShowDialog(Me)
        tutorGroup.Dispose()

    End Sub


    Private Sub RoleManagementBtn_Click(sender As System.Object, e As System.EventArgs) _
            Handles RoleManagementBtn.Click

        Dim roleManagement = New RoleManagement()
        roleManagement.ShowDialog(Me)
        roleManagement.Dispose()

    End Sub


    Private Sub StudentSanctionsBtn_Click(sender As System.Object, e As System.EventArgs) _
            Handles StudentSanctionsBtn.Click

        Dim studentSanctions = New StudentSanctions()
        studentSanctions.ShowDialog(Me)
        studentSanctions.Dispose()

    End Sub


    Private Sub PointsBtn_Click(sender As Object, e As EventArgs) _
            Handles PointsBtn.Click

        Dim points As New Points()
        points.ShowDialog(Me)
        points.Dispose()

    End Sub


    Private Sub ViewRollsBtn_Click(sender As Object, e As EventArgs) _
                Handles ViewRollsBtn.Click

        Dim viewRolls As New ViewRolls()
        viewRolls.ShowDialog(Me)
        viewRolls.Dispose()

    End Sub

#End Region

End Class
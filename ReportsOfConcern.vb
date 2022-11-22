Imports Microsoft.Reporting.WinForms
Imports System.Configuration
Imports BMS.BmsShared



Public Class ReportsOfConcern

    Private currentYearLevel As Integer



    ' =======================================================================================
    ' FUNCTIONS
    ' =======================================================================================


#Region "Functions"


    Private Sub ShowReportForYear(YearLevel As Integer)

        Try
            Me.currentYearLevel = YearLevel
            Me.ReportViewer1.Visible = True
            Dim Params As New List(Of ReportParameter) From {
                New ReportParameter("YearLevel", Me.currentYearLevel)
            }
            Me.ReportViewer1.ServerReport.SetParameters(Params)
            Me.ReportViewer1.RefreshReport()
            MarkAsCompleteBtn.Enabled = True
        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=False, sendEmail:=True,
                            addToLogTable:=True, showMsgBox:=True, source:="BmsAdmin.ReportsOfConcern Form")
        End Try

    End Sub


#End Region




    ' =======================================================================================
    ' EVENT HANDLERS
    ' =======================================================================================


#Region "EventHandlers"



    Private Sub RoC_Deactivate(sender As Object, e As System.EventArgs) Handles Me.Deactivate

        CType(Owner, BmsAdminMain).ShowUnhandledReportsOfConcernCount()

    End Sub



    Private Sub MarkAsCompleteBtn_Click(sender As System.Object, e As System.EventArgs) _
            Handles MarkAsCompleteBtn.Click


        If MsgBox(String.Format(
                  "This will set all Reports of Concern for Year {0} to 'Completed'. " _
                    & vbCrLf _
                    & "Are you sure you want to proceed?", currentYearLevel),
                MsgBoxStyle.Question + MsgBoxStyle.OkCancel) <> MsgBoxResult.Ok Then
            Return
        End If


        Try
            'Update ustudent sanctions so that all Reports of Concern for a single Year level
            ' are marked as complete (ie handled).

            Dim updateParameters = New Dictionary(Of String, String) _
                From {{"YearLevel", Me.currentYearLevel.ToString()}}

            DataHandler.update(
                    procedureName:=ConfigurationManager.AppSettings("SetReportsOfConcernCompleteProc"),
                    commandParameters:=updateParameters)

            ReportViewer1.RefreshReport()

        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=False, sendEmail:=True,
                    addToLogTable:=True, showMsgBox:=True, source:="BmsAdmin.ReportsOfConcern Form")
        End Try
    End Sub



#End Region




    Private Sub RadioButton_CheckedChanged(sender As RadioButton, e As System.EventArgs) _
        Handles _
            Year7Rbtn.CheckedChanged,
            Year8Rbtn.CheckedChanged,
            Year9Rbtn.CheckedChanged,
            Year10Rbtn.CheckedChanged,
            Year11Rbtn.CheckedChanged,
            Year12Rbtn.CheckedChanged

        If sender.Checked Then
            ShowReportForYear(sender.Tag)
        End If

    End Sub

End Class
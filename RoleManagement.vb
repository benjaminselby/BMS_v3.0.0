Imports System.Configuration
Imports BMS.BmsShared


Public Class RoleManagement

    Private headsOfYear As New DataTable

    Private Sub RoleManagement_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try

            HeadsOfYearDgv.DataSource = DataHandler.getDataTable(
                procedureName:=ConfigurationManager.AppSettings("GetHeadsOfYearProc"))

            With HeadsOfYearDgv.Columns(1)
                .ReadOnly = True
                .DefaultCellStyle.BackColor = Color.LightGray
                .DefaultCellStyle.Font = New Font(HeadsOfYearDgv.DefaultCellStyle.Font, FontStyle.Bold)
                .DefaultCellStyle.SelectionBackColor = Color.LightGray
                .DefaultCellStyle.SelectionForeColor = HeadsOfYearDgv.DefaultCellStyle.ForeColor
            End With

        Catch ex As Exception
            ErrorHandler.HandleError(exception:=ex, rethrow:=False, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True, source:="BMS Admin - Role Management")
        End Try

    End Sub


    Private Sub HeadsOfYearDgv_DataError(sender As Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) _
            Handles HeadsOfYearDgv.DataError

        MsgBox("The value you have just entered is not in the correct format. Please review your edit and try again." &
               vbCrLf & "If this problem persists, please contact Data Management for assistance.",
                MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Data Format Error")

    End Sub



    Private Sub RoleManagement_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) _
            Handles Me.FormClosing

        Try
            HeadsOfYearDgv.CurrentCell = Nothing
            Dim headsOfYearChanges = CType(HeadsOfYearDgv.DataSource, DataTable).GetChanges
            If headsOfYearChanges IsNot Nothing Then

                If MsgBox("Do you want to save your changes?",
                          MsgBoxStyle.YesNo + MsgBoxStyle.Question,
                          "Save Changes?") = MsgBoxResult.Yes Then

                    For Each headsOfYearRow As DataRow In headsOfYearChanges.Rows

                        Dim updateParameters = New Dictionary(Of String, String) _
                            From {
                                {"YearLevel", headsOfYearRow("Year Level").ToString()},
                                {"StaffId", headsOfYearRow("HOY_ID").ToString()}}

                        DataHandler.update(
                            procedureName:=ConfigurationManager.AppSettings("UpdateHeadOfYearProc"),
                            commandParameters:=updateParameters)

                    Next

                End If
            End If

        Catch ex As Exception
            ErrorHandler.HandleError(
                exception:=ex, rethrow:=False, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True,
                source:="BMS Admin - Role Management",
                message:=ConfigHandler.dataSaveErrorMsg)
        End Try

    End Sub

End Class
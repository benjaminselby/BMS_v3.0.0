Imports System.Data.SqlClient
Imports System.Configuration
Imports BMS.BmsShared


Public Class ElectronicRoll

    Public sanctionType As String
    'Private sanctionRollTable As DataTable
    Private Const minimumFormHeight As Integer = 300

    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load

        Try

            headingLbl.Text = "Electronic Roll for " & FormatDateTime(Today, DateFormat.LongDate) & ": " & sanctionType

            Dim queryParameters = New Dictionary(Of String, String) _
                From {{"SanctionCode", Me.sanctionType}}

            rollDgv.DataSource = DataHandler.getDataTable(
                procedureName:=ConfigurationManager.AppSettings("GetElectronicRollProc"),
                queryParameters:=queryParameters)

            For Each col As DataGridViewColumn In rollDgv.Columns
                With col
                    .SortMode = DataGridViewColumnSortMode.NotSortable
                    .Resizable = DataGridViewTriState.False
                    .Frozen = True
                    .ReadOnly = True
                End With
            Next

            With rollDgv

                .Columns("SanctionSeq").Visible = False
                .Columns("StudentId").Visible = False

                .Columns("Preferred").Resizable = True
                .Columns("Preferred").DisplayIndex = 0

                .Columns("Surname").Resizable = True
                .Columns("Surname").DisplayIndex = 1

                .Columns("Present").ReadOnly = False
                .Columns("Present").DisplayIndex = 2
                .Columns("Present").DefaultCellStyle.BackColor = Color.LightGray
                .Columns("Present").DefaultCellStyle.SelectionBackColor = Color.LightGray

                .Columns("Tutor").HeaderText = "Tutor Group"
                .Columns("SanctionDate").Visible = False
                .Columns("ReasonType").HeaderText = "Type"
                .Columns("SanctionCode").Visible = False

                .Columns("StaffReason").HeaderText = "Notes"
                .Columns("StaffReason").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("StaffReason").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                .Columns("StaffReason").Width = 300

                .Columns("HOYNotes").HeaderText = "YLM Notes (optional)"
                .Columns("HOYNotes").ReadOnly = False
                .Columns("HOYNotes").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("HOYNotes").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                .Columns("HOYNotes").Width = 300

            End With

            rollDgv.Width = rollDgv.Columns.GetColumnsWidth(DataGridViewElementStates.Visible) + 1
            rollDgv.Height = rollDgv.Rows.GetRowsHeight(DataGridViewElementStates.None) + rollDgv.ColumnHeadersHeight

            Me.Width = Math.Max(rollDgv.Width, headingLbl.Width) + 44
            Dim myHeight As Integer = rollDgv.Height + Me.headingLbl.Height + Me.Height - Me.ClientSize.Height + 44
            Me.Height = Math.Max(minimumFormHeight, myHeight)

        Catch ex As Exception
            ErrorHandler.HandleError(
                exception:=ex, rethrow:=False, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True,
                source:="BMS Admin - ElectronicRoll",
                message:="Could not obtain electronic roll data.")
        End Try

    End Sub


    Private Sub ElectronicRoll_FormClosing(
            ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) _
            Handles Me.FormClosing

        Try
            rollDgv.CurrentCell = Nothing

            Dim changesDset As DataTable = CType(rollDgv.DataSource, DataTable).GetChanges
            If changesDset IsNot Nothing Then

                For Each row As DataRow In changesDset.Rows

                    Dim updateParameters As New Dictionary(Of String, String) _
                        From {
                            {"Present", row("Present").ToString()},
                            {"HoyNotes", row("HoyNotes").ToString()},
                            {"SanctionSeq", row("SanctionSeq").ToString()}}
                    DataHandler.update(
                        procedureName:=ConfigurationManager.AppSettings("UpdateSanctionAttendanceProc"),
                        commandParameters:=updateParameters)

                Next

            End If

        Catch ex As Exception
            ErrorHandler.HandleError(
                exception:=ex, rethrow:=False, sendEmail:=True,
                addToLogTable:=True, showMsgBox:=True,
                source:="BMS Admin - ElectronicRoll",
                message:=ConfigHandler.dataSaveErrorMsg)
        End Try

    End Sub

End Class


Imports System.Configuration
Imports BMS.BmsShared


Public Class SanctionsForUserForm

    Private sanctionListDtbl As New DataTable


    '======================================================================================
    ' WINDOW INITIALISATION
    '======================================================================================


    Private Sub SanctionsForUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "Current sanctions for " _
            + ConfigHandler.currentUser.firstName + " " + ConfigHandler.currentUser.surname

        PopulateSanctionListDgv()

    End Sub



    '======================================================================================
    ' FUNCTIONS
    '======================================================================================
#Region "Functions"

    Private Sub PopulateSanctionListDgv()

        sanctionListDgv.Columns.Clear()
        sanctionListDgv.DataSource = Nothing
        sanctionListDtbl.Clear()

        sanctionListDtbl = ConfigHandler.currentSanctionsForUser
        ' Add a checkbox column to the data table so users can select particular sanctions. 
        If Not sanctionListDtbl.Columns.Contains("Selected") Then
            sanctionListDtbl.Columns.Add("Selected", GetType(Boolean))
        End If
        sanctionListDgv.DataSource = sanctionListDtbl



        ' Default Settings for all columns. 
        For columnIndex As Integer = 0 To sanctionListDgv.Columns.Count - 1
            With sanctionListDgv.Columns(columnIndex)
                .SortMode = DataGridViewColumnSortMode.NotSortable
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .Visible = False
                ' column 0 is the checkbox column
                .ReadOnly = If(columnIndex <> 0, True, False)
            End With
        Next


        With sanctionListDgv

            .Columns("Selected").DisplayIndex = 0
            .Columns("StudentName").DisplayIndex = 1
            .Columns("Reason").DisplayIndex = 2
            .Columns("SanctionType").DisplayIndex = 3
            .Columns("SanctionDate").DisplayIndex = 4
            .Columns("Comment").DisplayIndex = 5

            .Columns("Selected").Visible = True
            .Columns("StudentName").Visible = True
            .Columns("Reason").Visible = True
            .Columns("SanctionType").Visible = True
            .Columns("SanctionDate").Visible = True
            .Columns("Comment").Visible = True

            .Columns("Selected").ReadOnly = False

            .Columns("Selected").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .Columns("Selected").Width = 50
            .Columns("Selected").HeaderText = "Delete?"

            .Columns("StudentName").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .Columns("StudentName").Width = 100

            .Columns("Comment").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            .Columns("StudentName").DefaultCellStyle.BackColor = Color.LightGray
            .Columns("SanctionType").DefaultCellStyle.BackColor = Color.White
            .Columns("Reason").DefaultCellStyle.BackColor = Color.White

        End With

        sanctionListDgv.Refresh()

    End Sub

#End Region



    '======================================================================================
    ' EVENT HANDLERS
    '======================================================================================
#Region "Event Handlers"


    Private Sub DeleteBtn_Click(sender As Object, e As EventArgs) Handles DeleteBtn.Click

        If sanctionListDtbl IsNot Nothing AndAlso sanctionListDtbl.Rows.Count > 0 Then

            Dim sanctionsToDelete() As DataRow = sanctionListDtbl.Select("Selected=1")
            If sanctionsToDelete.Count > 0 Then

                If MsgBox(String.Format(
                        "{0} sanction{1} will be deleted. Do you wish to proceed?",
                        sanctionsToDelete.Count,
                        If(sanctionsToDelete.Count > 1, "s", "")),
                    MsgBoxStyle.YesNo,
                    "Delete Selected Sanctions?") = MsgBoxResult.Yes Then


                    For Each sanctionRow As DataRow In sanctionsToDelete

                        Dim rowsAffected As Integer = DataHandler.delete(
                            procedureName:=ConfigurationManager.AppSettings("DeleteSanctionProc"),
                            commandParameters:=New Dictionary(Of String, String) _
                                From {{"Seq", sanctionRow("Seq")}})

                        ' TODO: confirm rows affected here? Or count total number? 
                    Next


                End If
            End If
        End If

        PopulateSanctionListDgv()

    End Sub




    Private Sub SaveAndExitBtn_Click(sender As Object, e As EventArgs) Handles SaveAndExitBtn.Click

        Try

            Dim modifiedSanctions As DataTable = sanctionListDtbl.GetChanges()

            If modifiedSanctions IsNot Nothing AndAlso modifiedSanctions.Rows.Count > 0 Then
                For Each modifiedRow As DataRow In modifiedSanctions.Rows

                    Dim sanctionDate As DateTime = CType(modifiedRow("SanctionDate"), DateTime)

                    Dim updateParameters As New Dictionary(Of String, String) _
                        From {
                            {"Seq", modifiedRow("Seq")},
                            {"SanctionDate", sanctionDate.ToString("yyyy-MM-dd")},
                            {"SanctionType", modifiedRow("SanctionType")},
                            {"SanctionReason", modifiedRow("Reason")},
                            {"Comment", modifiedRow("Comment")}}

                    Dim rowsUpdated As Integer? = DataHandler.update(
                                    procedureName:=ConfigurationManager.AppSettings("UpdateSanctionProc"),
                                    commandParameters:=updateParameters)

                Next
            End If

        Catch ex As Exception

            MsgBox(
                "WARNING: Your changes may not have been saved correctly!" & vbCrLf &
                "Please ensure that you are connected to the Woodcroft network and try again." & vbCrLf &
                "If this problem persists, contact Data Management for assistance." & vbCrLf & vbCrLf &
                "========" & vbCrLf &
                "EXCEPTION DETAILS:" & vbCrLf &
                ex.ToString, MsgBoxStyle.Critical, "Database error!")

        End Try

        MyBase.Close()

    End Sub

#End Region

End Class
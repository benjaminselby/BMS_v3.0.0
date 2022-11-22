Public Class SanctionBooking

    Public DateTime As DateTime
    Public StaffName As String

    Public Sub New(DateTime As DateTime, StaffName As String)
        Me.DateTime = DateTime
        Me.StaffName = StaffName
    End Sub

End Class

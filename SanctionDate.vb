Public Class SanctionDate

    Public day As Date
    Public alreadyBooked As Boolean
    Public bookedBy As User

    Public Sub New(day As Date, alreadyBooked As Boolean, bookedBy As User)
        Me.day = day
        Me.alreadyBooked = alreadyBooked
        Me.bookedBy = bookedBy
    End Sub

    Public Overrides Function ToString() As String
        Dim returnString As String = Me.day.ToString("ddd dd MMM")
        If Me.alreadyBooked Then
            returnString += ": Booked by " + bookedBy.ToString()
        End If
        Return returnString
    End Function

End Class


Public Class Student
    Inherits User

    Public yearLevel As Integer

    Public Sub New(FirstName As String, Surname As String,
                   ID As String, YearLevel As Integer,
                   Optional emailAddress As String = Nothing,
                   Optional networkLogin As String = Nothing)

        MyBase.New(firstName:=FirstName, surname:=Surname, id:=ID,
                    emailAddress:=emailAddress, networkLogin:=networkLogin)
        Me.yearLevel = YearLevel

    End Sub

    Public Overrides Function ToString() As String
        Return String.Format("{0} {1} - Year {2} - ID: {3}",
            Me.firstName, Me.surname, Me.yearLevel, Me.id)
    End Function

End Class

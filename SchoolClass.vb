Imports System.Diagnostics.Eventing.Reader
Imports System.Security.Cryptography

Public Class SchoolClass

    Public ClassCode As String
    Public Description As String

    Public Sub New(ClassCode As String, Description As String)
        Me.ClassCode = ClassCode
        Me.Description = Description
    End Sub

    Public Overrides Function ToString() As String
        If Me.ClassCode = "" Then
            Return Me.Description
        ElseIf Me.Description = "" Then
            Return Me.ClassCode
        Else
            Return Me.ClassCode & " - " & Me.Description
        End If
    End Function

End Class

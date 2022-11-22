Imports BMS.BmsShared
Imports System.Configuration

Public Class User

    Public ReadOnly networkLogin As String
    Public ReadOnly firstName As String
    Public ReadOnly surname As String
    ' ID is a string in case future requires non-integer values. 
    Public ReadOnly id As String
    Public ReadOnly emailAddress As String

    Public Sub New(firstName As String, surname As String,
                   id As String, emailAddress As String,
                   Optional networkLogin As String = Nothing)

        Me.firstName = firstName
        Me.surname = surname
        Me.id = id
        Me.emailAddress = emailAddress
        Me.networkLogin = networkLogin

    End Sub


    Public Overrides Function ToString() As String
        Return String.Format("{0} {1} - ID: {2}", Me.firstName, Me.surname, Me.id)
    End Function


    Private _isBmsAdmin As Boolean? = Nothing
    Public ReadOnly Property isBmsAdmin
        Get
            If Me._isBmsAdmin Is Nothing Then
                Dim queryParameters = New Dictionary(Of String, String) _
                    From {{"UserId", Me.id}}
                Me._isBmsAdmin = DataHandler.getScalarValue(
                    command:=ConfigurationManager.AppSettings("CheckBmsAdminUserProc"),
                    queryParameters:=queryParameters)
            End If
            Return _isBmsAdmin
        End Get
    End Property


End Class

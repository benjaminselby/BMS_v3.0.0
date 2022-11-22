Imports System.Data.SqlClient
Imports System.Configuration


Namespace BmsShared

    Public Class DataHandler

        ' This class forms the principal connection to any permanent data store
        ' (ie. database). 


        ' Size in characters of default buffer generally used for initialising
        ' output parameters. 
        Private Const varcharBufferSize As Integer = 1000


        Public Shared Sub logMessage(message As String,
                                      Optional status As String = "",
                                      Optional source As String = "")

            Dim insertParameters = New Dictionary(Of String, String) _
                From {
                    {"Application", "BMS - " & ConfigHandler.versionNumber},
                    {"Component", source},
                    {"Status", status},
                    {"UserId", ConfigHandler.currentUser.id},
                    {"Message", message}}


            DataHandler.insert(procedureName:=ConfigurationManager.AppSettings("AddLogMessageProc"),
                               commandParameters:=insertParameters)

        End Sub


        Public Shared Function getScalarValue(command As String,
                                Optional isStoredProcedure As Boolean = True,
                                Optional queryParameters As Dictionary(Of String, String) = Nothing,
                                Optional columnName As String = Nothing
                                ) As String

            Dim returnValue As String = Nothing

            Using dbConnection = New SqlConnection(ConfigHandler.dbConnectionString)
                Using dbCommand = New SqlCommand(command, dbConnection)

                    If isStoredProcedure Then
                        dbCommand.CommandType = CommandType.StoredProcedure
                    Else
                        dbCommand.CommandType = CommandType.Text
                    End If

                    If queryParameters IsNot Nothing Then
                        For Each parameter As KeyValuePair(Of String, String) In queryParameters
                            dbCommand.Parameters.AddWithValue(parameter.Key, parameter.Value)
                        Next
                    End If

                    Try
                        dbConnection.Open()

                        Dim dbDataReader As SqlDataReader = dbCommand.ExecuteReader()
                        If (dbDataReader.HasRows) Then

                            dbDataReader.Read()
                            ' If column name is not specified we assume the first column. 
                            If columnName Is Nothing Then
                                returnValue = dbDataReader.GetValue(0)
                            Else
                                returnValue = dbDataReader.Item(columnName)
                            End If

                        End If

                    Finally
                        dbConnection.Close()
                    End Try

                End Using
            End Using

            Return returnValue

        End Function



        Public Shared Function getList(procedureName As String,
                                    Optional queryParameters As Dictionary(Of String, String) = Nothing,
                                    Optional columnName As String = Nothing
                                    ) As List(Of String)

            ' Returns a list of string values representing all values in a single column 
            ' returned from a database query. 

            Dim returnList As List(Of String) = Nothing

            Using dbConnection = New SqlConnection(ConfigHandler.dbConnectionString)
                Using dbCommand = New SqlCommand(procedureName, dbConnection)

                    dbCommand.CommandType = CommandType.StoredProcedure

                    For Each parameter As KeyValuePair(Of String, String) In queryParameters
                        dbCommand.Parameters.AddWithValue(parameter.Key, parameter.Value)
                    Next

                    Try
                        dbConnection.Open()

                        Dim dbDataReader As SqlDataReader = dbCommand.ExecuteReader()
                        If (dbDataReader.HasRows) Then

                            returnList = New List(Of String)

                            Do While (dbDataReader.Read())

                                If columnName Is Nothing Then
                                    returnList.Add(dbDataReader.GetValue(0))
                                Else
                                    returnList.Add(dbDataReader.Item(columnName))
                                End If

                            Loop

                        End If

                    Finally
                        dbConnection.Close()
                    End Try

                End Using
            End Using

            Return returnList

        End Function



        Public Shared Function getDictionary(procedureName As String,
                                    Optional queryParameters As Dictionary(Of String, String) = Nothing
                                    ) As Dictionary(Of String, String)

            ' Returns a dictionary of {ColumnName, Value} based on the first row 
            ' returned from the database. Ignores further rows. 

            Dim returnDictionary As Dictionary(Of String, String) = Nothing

            Using dbConnection = New SqlConnection(ConfigHandler.dbConnectionString)
                Using dbCommand = New SqlCommand(procedureName, dbConnection)

                    dbCommand.CommandType = CommandType.StoredProcedure

                    For Each parameter As KeyValuePair(Of String, String) In queryParameters
                        dbCommand.Parameters.AddWithValue(parameter.Key, parameter.Value)
                    Next

                    Try
                        dbConnection.Open()

                        Dim dbDataReader As SqlDataReader = dbCommand.ExecuteReader()
                        If (dbDataReader.HasRows) Then

                            returnDictionary = New Dictionary(Of String, String)

                            dbDataReader.Read()

                            For i As Integer = 0 To dbDataReader.FieldCount - 1

                                returnDictionary.Add(
                                    dbDataReader.GetName(i),
                                    dbDataReader.GetValue(i).ToString())

                            Next

                        End If

                    Finally
                        dbConnection.Close()
                    End Try

                End Using
            End Using

            Return returnDictionary

        End Function



        Public Shared Function getDataTable(procedureName As String,
                                Optional queryParameters As Dictionary(Of String, String) = Nothing
                                ) As DataTable

            Dim returnTable = New DataTable()

            Using dbConnection = New SqlConnection(ConfigHandler.dbConnectionString)
                Using dbCommand = New SqlCommand(procedureName, dbConnection)

                    dbCommand.CommandType = CommandType.StoredProcedure

                    If queryParameters IsNot Nothing Then
                        For Each parameter As KeyValuePair(Of String, String) In queryParameters
                            dbCommand.Parameters.AddWithValue(parameter.Key, parameter.Value)
                        Next
                    End If

                    Try
                        dbConnection.Open()

                        Using dataAdapter As New SqlDataAdapter(dbCommand)
                            dataAdapter.Fill(returnTable)
                        End Using

                    Finally
                        dbConnection.Close()
                    End Try

                End Using
            End Using

            Return returnTable

        End Function



        Public Shared Function getImage(procedureName As String,
                                Optional queryParameters As Dictionary(Of String, String) = Nothing,
                                Optional columnName As String = Nothing
                                ) As Bitmap

            Dim returnImage As Bitmap = Nothing

            Using dbConnection = New SqlConnection(ConfigHandler.dbConnectionString)
                Using dbCommand = New SqlCommand(procedureName, dbConnection)

                    dbCommand.CommandType = CommandType.StoredProcedure

                    For Each parameter As KeyValuePair(Of String, String) In queryParameters
                        dbCommand.Parameters.AddWithValue(parameter.Key, parameter.Value)
                    Next

                    Try
                        dbConnection.Open()

                        Using dataReader As SqlDataReader = dbCommand.ExecuteReader()

                            If dataReader.HasRows() Then

                                Dim imageData As Byte()
                                dataReader.Read()
                                If columnName Is Nothing Then
                                    imageData = CType(dataReader.GetValue(0), Byte())
                                Else
                                    imageData = CType(dataReader(columnName), Byte())
                                End If

                                returnImage = New Bitmap(New System.IO.MemoryStream(imageData))

                            End If
                        End Using

                    Finally
                        dbConnection.Close()
                    End Try

                End Using
            End Using

            Return returnImage

        End Function



        Public Shared Function insert(procedureName As String,
                                Optional commandParameters As Dictionary(Of String, String) = Nothing
                                ) As Integer?

            Dim rowsInserted As Integer = executeNonQuery(
                        procedureName:=procedureName,
                        commandParameters:=commandParameters)
            Return rowsInserted

        End Function




        Public Shared Function update(procedureName As String,
                                Optional commandParameters As Dictionary(Of String, String) = Nothing
                                ) As Integer?

            Dim rowsUpdated As Integer = executeNonQuery(
                        procedureName:=procedureName,
                        commandParameters:=commandParameters)
            Return rowsUpdated

        End Function


        Public Shared Function delete(procedureName As String,
                        Optional commandParameters As Dictionary(Of String, String) = Nothing
                        ) As Integer?

            Dim rowsDeleted As Integer = executeNonQuery(
                        procedureName:=procedureName,
                        commandParameters:=commandParameters)
            Return rowsDeleted

        End Function



        Public Shared Function executeNonQuery(procedureName As String,
                        Optional commandParameters As Dictionary(Of String, String) = Nothing,
                        Optional ByRef returnValue As String = Nothing,
                        Optional ByRef outputParameters As Dictionary(Of String, String) = Nothing
                        ) As Integer?

            ' Run a procedure which either inserts, deletes, or updates values.
            ' Function returns number of rows affected. 
            '
            ' If a return value from the procedure is required, its value will be placed in the 
            ' returnValue string parameter and returned by reference. 
            '
            ' If output parameters are specified, they should be in a dictionary where the 
            ' key is the parameter name and the value is an empty string which will be 
            ' populated with any values returned from the procedure. 
            ' 
            ' For simplicity we treat all output and return parameters as VARCHAR and send
            ' them to string buffers. This seems to work OK even when the SQL PROCEDURE defines
            ' the output parameters as something else e.g. INT. The caller can cast them to 
            ' whatever data type is required. 


            Dim returnRowCount As Integer? = Nothing
            Dim outputParameterObjects = New Dictionary(Of String, SqlParameter)
            Dim returnParameterObject As SqlParameter


            Using dbConnection As New SqlConnection(ConfigHandler.dbConnectionString)
                Using dbCommand As New SqlCommand(procedureName, dbConnection)

                    dbCommand.CommandType = CommandType.StoredProcedure

                    For Each parameter As KeyValuePair(Of String, String) In commandParameters
                        dbCommand.Parameters.AddWithValue(parameter.Key, parameter.Value)
                    Next

                    ' Add output parameters to a collection so we can reference
                    ' their return values later. 
                    If outputParameters IsNot Nothing Then
                        For Each parameterName As String In outputParameters.Keys
                            Dim newOutputParameter As SqlParameter =
                            dbCommand.Parameters.Add(parameterName, SqlDbType.VarChar, DataHandler.varcharBufferSize)
                            newOutputParameter.Direction = ParameterDirection.Output
                            outputParameterObjects.Add(parameterName, newOutputParameter)
                        Next
                    End If

                    returnParameterObject = dbCommand.Parameters.Add("ReturnValue", SqlDbType.VarChar, DataHandler.varcharBufferSize)
                    returnParameterObject.Direction = ParameterDirection.ReturnValue


                    dbConnection.Open()
                    returnRowCount = dbCommand.ExecuteNonQuery()
                    dbConnection.Close()

                    ' Populate the return values dictionary using output parameter objects
                    ' to send back to caller. 
                    If outputParameters IsNot Nothing Then
                        For Each parameterName As String In outputParameters.Keys
                            outputParameters(parameterName) = outputParameterObjects(parameterName).Value.ToString()
                        Next
                    End If

                    returnValue = returnParameterObject.Value.ToString()

                End Using
            End Using

            Return returnRowCount

        End Function

    End Class

End Namespace

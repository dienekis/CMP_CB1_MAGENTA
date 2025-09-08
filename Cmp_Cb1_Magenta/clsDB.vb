
Public Class clsDB
    'Class created for handling all Database related interactions

    Public RecordCount As Long = 0
    Dim oOleDbConnection As New System.Data.OleDb.OleDbConnection()
    Dim oOleDbCommand As OleDb.OleDbCommand
    Dim oOleDbDataReader As OleDb.OleDbDataReader

    Public Function initConnection(ByVal Database As String, ByVal Username As String, ByVal Password As String) As Boolean
        Dim Provider As String
        Try

            If modMain.g_DebugMode Then
                Console.WriteLine("[D] Trying to initialize connection to " & Database & " [tns entry] with Username " & Username & " and Password " & Password)
                modLogger.WriteToLog("[D] Trying to initialize connection to " & Database & " [tns entry] with Username " & Username & " and Password " & Password)
            End If

            If Environment.Is64BitOperatingSystem Then
                Provider = "ORAOLEDB.ORACLE" '64bit
            Else
                Provider = "MSDAORA.1"  '32bit
            End If

            oOleDbConnection.ConnectionString = "Provider=" & Provider & ";Password=" & Password & ";Persist Security Info=False;User ID= " & Username & ";Data Source= " & Database & ""
            initConnection = OpenConnection()


        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in initConnection of clsdb. Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in initConnection of modDB. Specific exception is ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            Return False
        End Try
    End Function





    Public Function CloseConnection() As Boolean
        Try

            If Not oOleDbConnection.State = ConnectionState.Closed Then
                oOleDbConnection.Close()
                oOleDbConnection.Dispose()
                Debug.Print("Connection to the Database was closed successfully")
                If modMain.g_DebugMode Then
                    modLogger.WriteToLog("Connection to the Database was closed successfully")
                End If
            End If
            Return True
        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in CloseConnection of modDB. Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in CloseConnection of modDB. Specific exception is ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            Return False
        End Try
    End Function


    Public Function OpenConnection() As Boolean
        Try

            If Not oOleDbConnection.State = ConnectionState.Open Then
                oOleDbConnection.Open()
                Debug.Print("Connection to the Database Opened !")
            End If
            If modMain.g_DebugMode Then
                modLogger.WriteToLog("Connection to the Database was Opened successfully")
            End If
            Return True
        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in OpenConnection of modDB. Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in OpenConnection of modDB. Specific exception is ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            Return False
        End Try
    End Function


    Public Sub BeginTransaction()
        Try

            oOleDbConnection.BeginTransaction().Begin()
            Debug.Print("Transaction Began!")
        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in BeginTransaction of modDB. Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in BeginTransaction of modDB. Specific exception is ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E] Inner Exception is ")
            modLogger.WriteToLog(ex.InnerException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E] Stack Trace is ")
            modLogger.WriteToLog(ex.StackTrace)
            modLogger.WriteToLog("")
        End Try
    End Sub


    Public Sub CommitTransaction()
        Try

            oOleDbConnection.BeginTransaction().Commit()
            Debug.Print("Transaction Commited!")
        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in CommitTransaction of modDB. Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in CommitTransaction of modDB. Specific exception is ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E] Inner Exception is ")
            modLogger.WriteToLog(ex.InnerException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E] Stack Trace is ")
            modLogger.WriteToLog(ex.StackTrace)
            modLogger.WriteToLog("")
        End Try
    End Sub


    Public Sub RollbackTransaction()
        Try

            oOleDbConnection.BeginTransaction().Rollback()
            Debug.Print("Transaction Rolledbacked!")
        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in RollbackTransaction of modDB. Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in RollbackTransaction of modDB. Specific exception is ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
        End Try
    End Sub

    Public Function ExecuteDelete(SqlString As String) As Boolean
        Try

            If oOleDbConnection.State = ConnectionState.Closed Then
                OpenConnection()
                Debug.Print("Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                    modLogger.WriteToLog("[D] Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                End If
            End If

            If modMain.g_DebugMode Then

                modLogger.WriteToLog("[D] Trying to execute SQL Query: " & SqlString)
            End If

            oOleDbCommand = New System.Data.OleDb.OleDbCommand(SqlString, oOleDbConnection)
            If oOleDbCommand.ExecuteNonQuery() = 0 Then 'Have to check how much logging will affect performance
                Debug.Print("ZERO rows were deleted by the execution of the query: " & SqlString)
                Console.WriteLine("[D] ZERO rows were deleted by the execution of the query: ")
                Console.WriteLine(SqlString)
                Console.WriteLine("")
                modLogger.WriteToLog("[D] ZERO rows were deleted by the execution of the query: ")
                modLogger.WriteToLog(SqlString)
                modLogger.WriteToLog("")
                Return True
            Else
                Debug.Print("ONE OR MORE rows were deleted by the execution of the query: " & SqlString)
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] ExecuteInsertUpdate - ONE OR MORE rows were deleted by the execution of the query [Success]")
                    modLogger.WriteToLog("[D] ExecuteInsertUpdate - ONE OR MORE rows were deleted by the execution of the query [Success]")
                End If
                Return True

            End If


        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in ExecuteInsertUpdate of modDB while trying to execute the query: ")
            Console.WriteLine(SqlString)
            Console.WriteLine("[E]Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in ExecuteInsertUpdate of modDB while trying to execute the query: ")
            modLogger.WriteToLog(SqlString)
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            Return False
        End Try
    End Function

    Public Function ExecuteInsertUpdate(SqlString As String) As Boolean
        Try

            If oOleDbConnection.State = ConnectionState.Closed Then
                OpenConnection()
                Debug.Print("Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                    modLogger.WriteToLog("[D] Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                End If
            End If

            If modMain.g_DebugMode Then

                modLogger.WriteToLog("[D] Trying to execute SQL Query: " & SqlString)
            End If

            oOleDbCommand = New System.Data.OleDb.OleDbCommand(SqlString, oOleDbConnection)
            If oOleDbCommand.ExecuteNonQuery() = 0 Then 'Have to check how much logging will affect performance
                Debug.Print("ZERO rows were affected by the execution of the query: " & SqlString)
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] ZERO rows were affected by the execution of the query: ")
                    Console.WriteLine(SqlString)
                    Console.WriteLine("")
                    modLogger.WriteToLog("[D] ZERO rows were affected by the execution of the query: ")
                    modLogger.WriteToLog(SqlString)
                    modLogger.WriteToLog("")
                End If
                Return False
            Else
                Debug.Print("ONE OR MORE rows were affected by the execution of the query: " & SqlString)
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] ExecuteInsertUpdate - ONE OR MORE rows were affected by the execution of the query [Success]")
                    modLogger.WriteToLog("[D] ExecuteInsertUpdate - ONE OR MORE rows were affected by the execution of the query [Success]")
                End If
                Return True
            End If


        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in ExecuteInsertUpdate of modDB while trying to execute the query: ")
            Console.WriteLine(SqlString)
            Console.WriteLine("[E]Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in ExecuteInsertUpdate of modDB while trying to execute the query: ")
            modLogger.WriteToLog(SqlString)
            modLogger.WriteToLog(ex.Message)
            Return False
        End Try
    End Function


    Public Function ExecuteSelect(SelectSql As String) As Boolean
        Try

            If oOleDbConnection.State = ConnectionState.Closed Then
                OpenConnection()
                Debug.Print("Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                    modLogger.WriteToLog("[D] Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                End If
            End If
            Debug.Print("Trying to execute SQL: " & SelectSql)

            If modMain.g_DebugMode Then
                modLogger.WriteToLog("[D] Trying to execute SQL Query: " & SelectSql)
            End If

            oOleDbCommand = New System.Data.OleDb.OleDbCommand(SelectSql, oOleDbConnection)
            oOleDbDataReader = oOleDbCommand.ExecuteReader()
            If oOleDbDataReader.HasRows Then
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] The ExecuteSelect Function populated the DataReader with result rows [Success]")
                End If
                Return True
            Else

                If modMain.g_DebugMode Then
                    Console.WriteLine("The ExecuteSelect Function Returned NO Results - Returning False")
                End If
                Return False
            End If

        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in ExecuteSelect of modDB while trying to execute the query: ")
            Console.WriteLine(SelectSql)
            Console.WriteLine("[E]Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in ExecuteSelect of modDB while trying to execute the query: ")
            modLogger.WriteToLog(SelectSql)
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            Return False
        End Try
    End Function



    Public Function ExecuteCount(CountSql As String) As Long

        Try

            If oOleDbConnection.State = ConnectionState.Closed Then
                OpenConnection()
                Debug.Print("Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                    modLogger.WriteToLog("[D] Connection was on a closed state when a query was passed for execution. Connection was re-opened (Not an error)")
                End If
            End If
            Debug.Print("Trying to execute SQL: " & CountSql)
            If modMain.g_DebugMode Then
                modLogger.WriteToLog("[D] Trying to execute SQL Query: " & CountSql)
            End If

            oOleDbCommand = New System.Data.OleDb.OleDbCommand(CountSql, oOleDbConnection)
            oOleDbDataReader = oOleDbCommand.ExecuteReader()
            oOleDbDataReader.Read()
            Return oOleDbDataReader.GetValue(0) ' <--- Is the Number that CountSql returns

        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in ExecuteCount of modDB while trying to execute the query: ")
            Console.WriteLine(CountSql)
            Console.WriteLine("[E]Specific exception is ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E]!!Exception occured in ExecuteCount of modDB while trying to execute the query: ")
            modLogger.WriteToLog(CountSql)
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            Return -1
        End Try

    End Function



    Private Sub SetRecordCount(newRecordCount As Long)
        RecordCount = newRecordCount
    End Sub


    Public Function GetRecordCount()
        GetRecordCount = RecordCount
    End Function


    Public Function GetResultSet() As OleDb.OleDbDataReader
        Debug.Print("Returning Data Rows From Query to the Import module...")
        GetResultSet = oOleDbDataReader
    End Function


    Public Sub DisposeObjects() 'Call when Database work is finished
        If modMain.g_DebugMode Then
            Console.WriteLine("[D] Disposing ClsDB Objects...")
        End If
        oOleDbCommand.Dispose()
        oOleDbDataReader.Close()
    End Sub


End Class

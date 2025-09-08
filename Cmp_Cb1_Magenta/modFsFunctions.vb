Imports System.IO
Module modFsFunctions
    Public Function ValidateCreateDirectory(FullDirectoryPath As String) As Boolean

        Try

            If IO.Directory.Exists(FullDirectoryPath) Then
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] The Directory " & FullDirectoryPath & " already exists(Not an error)")
                End If
                Return True
            Else
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Target Directory was Not found.The application will create it")
                End If
                IO.Directory.CreateDirectory(FullDirectoryPath)
                Return True
            End If

        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in ValidateCreateDirectory of modFsFunctions! Specific exception message:")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            Return False
        End Try



    End Function


    Public Function DeleteDirectory(FullDirectoryPath As String) As Boolean

        Try
            If IO.Directory.Exists(FullDirectoryPath) Then
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] The Directory " & FullDirectoryPath & " was found. The application will delete it")
                    modLogger.WriteToLog("[D] The Directory " & FullDirectoryPath & " was found. The application will delete it")
                End If

                IO.Directory.Delete(FullDirectoryPath)
                Return True
            Else
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Target Directory was Not found.The application cannot delete a non-existent dir")
                    modLogger.WriteToLog("[D] Target Directory was Not found.The application cannot delete a non-existent dir")
                End If

                Return True
            End If

        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in DeleteDirectory of modFsFunctions! Specific exception message:")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E] Exception occured in DeleteDirectory of modFsFunctions! Specific exception message:")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")

            ' AG Fix: Removed to avoid application execution error
            'modLogger.WriteToLog("[E] Inner Exception is ")
            'modLogger.WriteToLog(ex.InnerException.ToString)
            'modLogger.WriteToLog("")

            modLogger.WriteToLog("[E] Stack Trace is ")
            modLogger.WriteToLog(ex.StackTrace)
            modLogger.WriteToLog("")
            Return False
        End Try
    End Function


    Public Function DeleteFile(FileFullPath As String) As Boolean
        Try
            If modMain.g_DebugMode Then
                Console.WriteLine("[D] Trying to delete " & FileFullPath)
                modLogger.WriteToLog("[D] Trying to delete " & FileFullPath)
            End If
            IO.File.Delete(FileFullPath)
            Return True

        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in DeleteFile of modFsFunctions! Specific exception message:")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E] Exception occured in DeleteFile of modFsFunctions! Specific exception message:")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")

            ' AG Fix: Removed to avoid application execution error
            'modLogger.WriteToLog("[E] Inner Exception is ")
            'modLogger.WriteToLog(ex.InnerException.ToString)
            'modLogger.WriteToLog("")

            modLogger.WriteToLog("[E] Stack Trace is ")
            modLogger.WriteToLog(ex.StackTrace)
            modLogger.WriteToLog("")
            Return False
        End Try

    End Function


    Public Function FileExists(ByVal FileToCheck As String) As Boolean
        If IO.File.Exists(FileToCheck) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function FileHasReadAccess(ByVal FileToCheck As String) As Boolean
        'Check if file can be accessed for reading
        Try
            Using f1 As System.IO.FileStream = New System.IO.FileStream(FileToCheck, FileMode.Open, FileAccess.Read)
                Return True 'Return that the file can be read
            End Using

        Catch ex As Exception
            Return False 'Return false
        End Try

    End Function


    Public Sub CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String)

        Try
            IO.File.Copy(SourcePath, DestinationPath)

            If modMain.g_DebugMode Then
                Console.WriteLine("Copied file from " & SourcePath & " to " & DestinationPath)
            End If

        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in CopyFile of modFsFunctions! Specific exception message:")
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("[E] Exception occured in CopyFile of modFsFunctions! Specific exception message:")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")

            ' AG Fix: Removed to avoid application execution error
            'modLogger.WriteToLog("[E] Inner Exception is ")
            'modLogger.WriteToLog(ex.InnerException.ToString)
            'modLogger.WriteToLog("")

            modLogger.WriteToLog("[E] Stack Trace is ")
            modLogger.WriteToLog(ex.StackTrace)
            modLogger.WriteToLog("")
        End Try
    End Sub

    Public Function SubstringAfterSpecialChar(Path As String, SpecialCharacter As String) As String

        Try
            ' Get index of argument and return substring after its position.
            Dim posA As Integer = Path.LastIndexOf(SpecialCharacter)
            If posA = -1 Then
                Return ""
            End If
            Dim adjustedPosA As Integer = posA + SpecialCharacter.Length
            If adjustedPosA >= Path.Length Then
                Return ""
            End If

            Return Path.Substring(adjustedPosA)

        Catch ex As Exception
            Return "" 'Return false
        End Try
    End Function


End Module

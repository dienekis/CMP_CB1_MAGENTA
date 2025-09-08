Imports System.IO

Module modLogger

    Dim validateCounter As Long = 0
    Public Function getDate() As String
        'Retrieve the date with a specific format so that it can be used for the filename
        Return CStr(Now.Year) & "_" & CStr(Now.Month) & "_" & CStr(Now.Day)
    End Function


    Public Sub generateProcessLog()
        'Sub for creating the execution process log
        Dim fullPath As String = getFullLogFilePath()
        Dim initWriter As System.IO.StreamWriter
        Dim initFile As IO.FileStream

        Try
            Console.WriteLine("Generating Process Log for Application Execution")
            If Not IO.File.Exists(fullPath) Then
                initFile = IO.File.Create(fullPath)
                initFile.Close()
                initFile.Dispose()
                Console.WriteLine("Log file created successfully!")
            End If

            initWriter = My.Computer.FileSystem.OpenTextFileWriter(fullPath, True)
            initWriter.WriteLine("B2C_RENEWALS Documents Import Flow Application")
            initWriter.WriteLine("Start of execution at: " & System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
            initWriter.WriteLine("-------------------------------------------------------------------")

            initWriter.Close()
            initWriter.Dispose()

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in generateProcessLog of modLogger. Specific exception is ")
            Console.WriteLine(ex.GetBaseException)
            Console.WriteLine("")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
        End Try

    End Sub


    Public Sub WriteToLog(LineMessage As String)
        'Sub for writting a line to an existing log
        Dim logPath As String = getFullLogFilePath()
        Dim logWriter As System.IO.StreamWriter

        Try
            logWriter = My.Computer.FileSystem.OpenTextFileWriter(logPath, True)
            logWriter.WriteLine(LineMessage)

            logWriter.Close()
            logWriter.Dispose()

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in WriteToLog of modLogger. Specific exception is ")
            Console.WriteLine(ex.GetBaseException)
            Console.WriteLine("")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
            logWriter.Close()
            logWriter.Dispose()
        End Try

    End Sub

    Public Function getFullLogFilePath() As String
        'Returns the process log full path (log filename included)
        If validateCounter = 0 Then
            modFsFunctions.ValidateCreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory & "Logs\")
        End If
        validateCounter += 1
        getFullLogFilePath = System.AppDomain.CurrentDomain.BaseDirectory & "Logs\" & generateLogFileName()
    End Function

    Public Function generateLogFileName() As String
        Return "B2C_RENEWALS_DOCUMENTS_IMPORT_FLOW__" & modAppConfigSettings.ThreadId & "___" & getDate() & ".log"
    End Function


    Public Sub FinalizeLog(successCount As Long, repeatCount As Long, errorCount As Long, time As TimeSpan, zero As Long, dontexist As Long, insuf As Long, totalnum As Long, insuferror As Long, meta As Long, unerr As Long)
        'Sub for finalizing the process log
        Dim finalPath As String = getFullLogFilePath()
        Dim finalWriter As System.IO.StreamWriter
        'Dim totalDocs As Long = successCount + repeatCount + errorCount
        'The the current totalDocs variable has the total number of docs processed along with the insuficient at first occur.
        'this particular variable was successfully used but in this case it must be calculate the first occurance of documents with insuficient data.
        'at this point i will not delete the old method of calculation the total docs, instead i will keep it comment it in a possible future use of it

        Dim totalDocs As Long = totalnum
        Try
            finalWriter = My.Computer.FileSystem.OpenTextFileWriter(finalPath, True)
            finalWriter.WriteLine("----------------------------------------------------------------------------------------------------")
            finalWriter.WriteLine("END of Application Execution at: " & System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
            finalWriter.WriteLine("Statistics:")
            finalWriter.WriteLine("Total Documents Processed: " & totalDocs)
            finalWriter.WriteLine("")
            finalWriter.WriteLine("Main Statistics:")
            finalWriter.WriteLine("Remaining Documents For Import Processing: " & totalDocs - insuf)
            finalWriter.WriteLine("Documents Commited to FileNet: " & successCount)
            finalWriter.WriteLine("Documents set to REPEAT state: " & repeatCount)
            finalWriter.WriteLine("Documents set to ERROR state:  " & errorCount & "  ->  (" & errorCount - insuferror & "  from Main Import Processing)")
            finalWriter.WriteLine("                                      (" & insuferror & "  from Insufficient Processing)")
            finalWriter.WriteLine("Documents UNABLE to set to ERROR state (supposed to): " & unerr)
            finalWriter.WriteLine("")
            finalWriter.WriteLine("Extra Statistics:")
            finalWriter.WriteLine("Documents With Insufficient Data: " & insuf)
            finalWriter.WriteLine("Documents With Insufficient Metadata: " & meta)
            finalWriter.WriteLine("Documents Zero Size: " & zero)
            finalWriter.WriteLine("Documents Dont Exist: " & dontexist)
            finalWriter.WriteLine("Application Total Runtime: " & time.ToString)
            finalWriter.WriteLine("----------------------------------------------------------------------------------------------------")
            finalWriter.WriteLine("END OF FILE")
            finalWriter.WriteLine("")
            finalWriter.Close()
            finalWriter.Dispose()

        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in FinalizeLog of modLogger!")
            Console.WriteLine("[E] Specific BaseException:")
            Console.WriteLine(ex.GetBaseException)
            Console.WriteLine("[E] StackTrace:")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
        End Try

    End Sub


    'SUBS AND FUNCTIONS FOR NOT EXISTS LOG


    Public Sub generateNOTEXISTSLog()
        'Sub for generating a NOTEXISTS log file. This file is used for documenting files that have valid metadata in the ePOS Database but do not actually exist in the predefined storage path
        Dim fullPath As String = getNotExistsLogFilePath()
        Dim NotExistsWriter As System.IO.StreamWriter
        Dim NotExistsFile As IO.FileStream

        Try
            Console.WriteLine("Generating NOT EXISTS Log for Application Execution")
            If Not IO.File.Exists(fullPath) Then
                NotExistsFile = IO.File.Create(fullPath)
                NotExistsFile.Close()
                NotExistsFile.Dispose()
                Console.WriteLine("NOT EXISTS Log file created successfully!")
            End If

            NotExistsWriter = My.Computer.FileSystem.OpenTextFileWriter(fullPath, True)
            NotExistsWriter.WriteLine("B2C_RENEWALS Documents Import Flow NOT EXISTS Log file")
            NotExistsWriter.WriteLine("Created at: " & System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
            NotExistsWriter.WriteLine("--------------------------------------------------------------")
            NotExistsWriter.WriteLine("Entries in this log are of files that dont exist at the ")
            NotExistsWriter.WriteLine("specified location or of files that exist but have Zero size ")
            NotExistsWriter.WriteLine("--------------------------------------------------------------")

            NotExistsWriter.Close()
            NotExistsWriter.Dispose()

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in generateNOTEXISTSLog of modLogger. Specific exception is ")
            Console.WriteLine(ex.GetBaseException)
            Console.WriteLine("")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
        End Try

    End Sub


    Public Sub WriteToNOTEXISTSLog(LineMessage As String)
        'Sub for writting a line to an existing NOTEXISTS log
        Dim logPath As String = getNotExistsLogFilePath()
        Dim logWriter As System.IO.StreamWriter

        Try
            logWriter = My.Computer.FileSystem.OpenTextFileWriter(logPath, True)
            logWriter.WriteLine(LineMessage)

            logWriter.Close()
            logWriter.Dispose()

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in WriteToNOTEXISTSLog of modLogger. Specific exception is ")
            Console.WriteLine(ex.GetBaseException)
            Console.WriteLine("")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
        End Try

    End Sub


    Public Function getNotExistsLogFilePath() As String
        'Returns the NOTEXISTS log full path (log filename included)
        modFsFunctions.ValidateCreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory & "Logs\")
        getNotExistsLogFilePath = System.AppDomain.CurrentDomain.BaseDirectory & "Logs\" & generateNotExistsLogName()
    End Function

    Public Function generateNotExistsLogName() As String
        Return "B2C_RENEWALS_DOCUMENTS_IMPORT_FLOW__" & modAppConfigSettings.ThreadId & "___NOT_EXISTS_" & getDate() & ".log"
    End Function


    'SUBS AND FUNCTIONS FOR INSUFFICIENT DATA

    Public Function getInsufficientDataLogFilePath() As String
        'Returns the INSUFFICIENTDATA log full path (log filename included)
        modFsFunctions.ValidateCreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory & "Logs\")
        getInsufficientDataLogFilePath = System.AppDomain.CurrentDomain.BaseDirectory & "Logs\" & generateInsufficientDataLogName()
    End Function

    Public Function generateInsufficientDataLogName() As String
        Return "B2C_RENEWALS_DOCUMENTS_IMPORT_FLOW__" & modAppConfigSettings.ThreadId & "___INSUFFICIENT_DATA_" & getDate() & ".log"
    End Function


    Public Sub generateINSUFFICIENTDATALog()
        'Sub for generating a INSUFFICIENTDATA log file. This file is used for documenting files that have valid metadata in the ePOS Database but do not actually exist in the predefined storage path
        Dim fullPath As String = getInsufficientDataLogFilePath()
        Dim InsufficientDataWriter As System.IO.StreamWriter
        Dim InsufficientDataFile As IO.FileStream

        Try
            'Console.WriteLine("Generating INSUFFICIENT DATA Log for Application Execution")
            If Not IO.File.Exists(fullPath) Then
                InsufficientDataFile = IO.File.Create(fullPath)
                InsufficientDataFile.Close()
                InsufficientDataFile.Dispose()
                'Console.WriteLine("INSUFFICIENT DATA Log file created successfully!")
            End If

            InsufficientDataWriter = My.Computer.FileSystem.OpenTextFileWriter(fullPath, True)
            InsufficientDataWriter.WriteLine("B2C_RENEWALS Documents Import Flow INSUFFICIENT DATA Log file")
            InsufficientDataWriter.WriteLine("Created at: " & System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
            InsufficientDataWriter.WriteLine("-----------------------------------------------------------------------------")
            InsufficientDataWriter.WriteLine("The entries in this log are missing mandatory data for archival in Filenet")
            InsufficientDataWriter.WriteLine("-----------------------------------------------------------------------------")

            InsufficientDataWriter.Close()
            InsufficientDataWriter.Dispose()

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in generateINSUFFICIENTDATALog of modLogger. Specific exception is ")
            Console.WriteLine(ex.GetBaseException)
            Console.WriteLine("")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
        End Try

    End Sub


    Public Sub WriteToINSUFFICIENTDATALog(LineMessage As String)
        'Sub for writting a line to an existing NOTEXISTS log
        Dim logPath As String = getInsufficientDataLogFilePath()
        Dim logWriter As System.IO.StreamWriter

        Try
            logWriter = My.Computer.FileSystem.OpenTextFileWriter(logPath, True)
            logWriter.WriteLine(LineMessage)

            logWriter.Close()
            logWriter.Dispose()

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in WriteToINSUFFICIENTDATALog of modLogger. Specific exception is ")
            Console.WriteLine(ex.GetBaseException)
            Console.WriteLine("")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
        End Try

    End Sub

End Module

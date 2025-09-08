Imports System.Runtime.InteropServices

Module modFilenetHandler
    ' Class used for handling all Filenet related interactions (logon, document commital, logoff....)
    Public FnLib As New IDMObjects.Library
    Public fuser As New IDMObjects.User
    Private Username As String
    Private Password As String
    Public FirstLogonFilenet As Boolean = True

    Public Function initFilenet() As Boolean
        Try
            If modMain.g_DebugMode Then
                modLogger.WriteToLog("[D] Initializing connection to Filenet with Username " & modAppConfigSettings.FnUser & " and Password " & modAppConfigSettings.FnPwd)
            End If
            Username = modAppConfigSettings.FnUser
            Password = modAppConfigSettings.FnPwd
            initFilenet = FnLogon()

        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in initFilenet of modFilenetHandler! Could not connect to FileNet! Specific exception message: ")
            Console.WriteLine(ex.GetBaseException.ToString)
            modLogger.WriteToLog("[E] Exception occured in initFilenet of modFilenetHandler! Could not connect to FileNet! Specific exception message: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ex.StackTrace)
            Return False
        End Try
    End Function


    Public Function FnLogon() As Boolean
        Try

            If g_DebugMode Then Console.WriteLine("Trying to connect to Filenet Repository...")
            'Debug.Print(FnLib.Label)
            If Not FnLib.GetState(IDMObjects.idmLibraryState.idmLibraryLoggedOn) Then
                If Username = "" Then

                    FnLib.Logon(Username, Password)
                Else
                    FnLib.Logon(Username, Password)
                End If
                If FirstLogonFilenet Then
                    Console.WriteLine("Connected to: " & FnLib.Label & " with User: " & Username)
                    modLogger.WriteToLog("Connected to: " & FnLib.Label & " with User: " & Username)
                    modLogger.WriteToLog("")
                    FirstLogonFilenet = False
                End If

            Else
                If FirstLogonFilenet Then
                    Console.WriteLine("Connected to: " & FnLib.Label & " with User: " & Username)
                    modLogger.WriteToLog("Connected to: " & FnLib.Label & " with User: " & Username)
                    modLogger.WriteToLog("")
                    FirstLogonFilenet = False
                End If

            End If



            If g_DebugMode Then Console.WriteLine("Connected to Filenet Repository")

            Return True
        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in FnLogon of modFilenetHandler! Could not connect to FileNet! Specific exception message:")
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in FnLogon of modFilenetHandler! Could not connect to FileNet! Specific exception message: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ex.StackTrace)
            Return False
        End Try

    End Function


    Public Function FnLogoff() As Boolean
        Dim checkFNconnection As Boolean
        checkFNconnection = False
        Try
            If g_DebugMode Then Console.WriteLine("Trying to log off from FileNet...")
            checkFNconnection = FnLib.GetState(IDMObjects.idmLibraryState.idmLibraryLoggedOn)
            If checkFNconnection Then
                FnLib.Logoff()
                checkFNconnection = FnLib.GetState(IDMObjects.idmLibraryState.idmLibraryLoggedOn)
                If checkFNconnection Then
                    Debug.Print("error")
                End If

                Console.WriteLine("Successfully logged off from FileNet")
                modLogger.WriteToLog("Successfully logged off from FileNet")
            End If

            Return True
        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in FnLogoff of modFilenetHandler! Could not logoff from FileNet! Specific exception message:")
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in FnLogoff of modFilenetHandler! Could not connect to FileNet! Specific exception message: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ex.StackTrace)
            Return False
        End Try

    End Function


    Public Function DeleteDocumentFromFN(DocumentNumber As String) As Boolean
        Try
            Dim FnDocument As New IDMObjects.Document
            Console.WriteLine("Deleting Document " & DocumentNumber & "...")
            FnDocument = FnLib.GetObject(IDMObjects.idmObjectType.idmObjTypeDocument, DocumentNumber)
            FnDocument.Delete()
            Marshal.ReleaseComObject(FnDocument)
            Return True
        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in DeleteDocumentFromFN of modFilenetHandler! Could not delete specified document from FileNet! Specific exception message:")
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in DeleteDocumentFromFN of modFilenetHandler! Could not delete specified document from FileNet! Specific exception message: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            Return False
        End Try

    End Function


    Public Function GetFnFolderNumberFromName(ByVal FolderName As String, ByRef FolderNumber As Long) As Boolean
        Dim FnFolder As IDMObjects.Folder
        Try
            If modMain.g_DebugMode Then
                modLogger.WriteToLog("[D] Retrieving Filenet folder number for specified folder name...")
                Console.WriteLine("[D] Retrieving Filenet folder number for specified folder name...")
            End If

            FnFolder = FnLib.GetObject(IDMObjects.idmObjectType.idmObjTypeFolder, FolderName)
            FolderNumber = FnFolder.ID
            Marshal.ReleaseComObject(FnFolder)
            Return True
        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in GetFnFolderNumberFromName of modFilenetHandler! Could not retrieve a folder number for the specific name! Specific exception message:")
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in GetFnFolderNumberFromName of modFilenetHandler! Could not  retrieve a folder number for the specific name! Specific exception message: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ex.StackTrace)
            Marshal.ReleaseComObject(FnFolder)
            Return False
        End Try

    End Function


    Public Function FolderizeDocument(ByVal DocumentId As String, ByVal Folder As String, Optional EposFileName As String = "") As Boolean
        Dim objFnDoc As IDMObjects.Document
        Dim objFnFolder As IDMObjects.Folder
        Dim RetryCount As Integer = 0
        Dim IsSuccessful As Boolean = False
        Dim ExceptionPlaceholder As Exception

        If modMain.g_DebugMode Then
            Console.WriteLine("[D] Trying to folderize document " & DocumentId & " in folder " & Folder & "...")
            modLogger.WriteToLog("[D] Trying to folderize document " & DocumentId & " in folder " & Folder & "...")
        End If

        Do

            Try
                objFnDoc = FnLib.GetObject(IDMObjects.idmObjectType.idmObjTypeDocument, DocumentId)
                objFnFolder = FnLib.GetObject(IDMObjects.idmObjectType.idmObjTypeFolder, Folder)
                objFnFolder.File(objFnDoc)
                'Marshal.ReleaseComObject(objFnDoc)
                'Marshal.ReleaseComObject(objFnFolder)
                IsSuccessful = True
            Catch ex As Exception
                ExceptionPlaceholder = ex
                Console.WriteLine("[E] Failed to Folderize Document. The application will retry.")
                modLogger.WriteToLog("[E] Failed to Folderize Document. Sleeping for a few ms and attempting to retry...")
                Threading.Thread.Sleep(2500)
                RetryCount += 1
            End Try
        Loop Until IsSuccessful = True OrElse RetryCount >= 2

        If IsSuccessful Then
            If modMain.g_DebugMode Then
                Console.WriteLine("[D] FolderizeDocument: Successfully folderized document " & DocumentId & " in folder " & Folder & " !")
                modLogger.WriteToLog("[D] FolderizeDocument: Successfully folderized document " & DocumentId & " in folder " & Folder & " !")
            End If
            Marshal.ReleaseComObject(objFnFolder)
            Marshal.ReleaseComObject(objFnDoc)
            Return True
        Else
            Console.WriteLine("[E] Exception occured in FolderizeDocument of modFilenetHandler! Could not Folderize the document: " & DocumentId & " from folder:" & Folder & "! Specific exception message:")
            Console.WriteLine(ExceptionPlaceholder.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in FolderizeDocument of modFilenetHandler! Could not Folderize Document" & DocumentId & " from folder:" & Folder & " Specific exception message: ")
            modLogger.WriteToLog(ExceptionPlaceholder.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.StackTrace)
            Marshal.ReleaseComObject(objFnFolder)
            Marshal.ReleaseComObject(objFnDoc)
            Return False
        End If


    End Function


    Public Function UnFolderizeDocument(ByVal DocumentId As String, Folder As String) As Boolean
        Dim objFnDoc As IDMObjects.Document
        Dim objFnFolder As IDMObjects.Folder
        Dim RetryCount As Integer = 0
        Dim IsSuccessful As Boolean = False
        Dim ExceptionPlaceholder As Exception

        Do
            Try
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Trying to UN-Folderize document " & DocumentId & " from folder " & Folder & "...")
                    modLogger.WriteToLog("[D] Trying to UN-Folderize document " & DocumentId & " from folder " & Folder & "...")
                End If

                objFnDoc = FnLib.GetObject(IDMObjects.idmObjectType.idmObjTypeDocument, DocumentId)
                objFnFolder = FnLib.GetObject(IDMObjects.idmObjectType.idmObjTypeFolder, Folder)
                objFnFolder.Unfile(objFnDoc)
                Marshal.ReleaseComObject(objFnFolder)
                Marshal.ReleaseComObject(objFnDoc)
                IsSuccessful = True
            Catch ex As Exception
                Console.WriteLine("[E] Failed to UnFolderize Document. The application will retry.")
                modLogger.WriteToLog("[E] Failed to UnFolderize Document. The application will retry.")
                Threading.Thread.Sleep(2000)
                ExceptionPlaceholder = ex
                RetryCount += 1
            End Try
        Loop Until IsSuccessful = True OrElse RetryCount >= 1

        If IsSuccessful Then
            If modMain.g_DebugMode Then
                Console.WriteLine("[D] FolderizeDocument: Successfully Unfolderized document " & DocumentId & " from folder " & Folder & " !")
                modLogger.WriteToLog("[D] FolderizeDocument: Successfully Unfolderized document " & DocumentId & " from folder " & Folder & " !")
            End If
            Return True
        Else
            Console.WriteLine("[E] Exception occured in UnFolderizeDocument of modFilenetHandler! Could not Unfolderize the document: " & DocumentId & " from folder:" & Folder & "! Specific exception message:")
            Console.WriteLine(ExceptionPlaceholder.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in UnFolderizeDocument of modFilenetHandler! Could not Unfolderize Document" & DocumentId & " from folder:" & Folder & " Specific exception message: ")
            modLogger.WriteToLog(ExceptionPlaceholder.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.StackTrace)
            Marshal.ReleaseComObject(objFnFolder)
            Marshal.ReleaseComObject(objFnDoc)
            Return False
        End If


    End Function



    Public Function FnFolderExists(FnFolder As String, ByRef Response As Boolean) As Boolean

        Dim objFnFolder As New IDMObjects.Folder
        Dim RetryCount As Integer = 0
        Dim IsSuccessful As Boolean = False
        Dim ExceptionPlaceholder As Exception

        Do
            Try
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Checking (function) if Folder " & FnFolder & " exists in Filenet")
                    modLogger.WriteToLog("[D] Checking (function) if Folder " & FnFolder & " exists in Filenet")
                End If
                objFnFolder = FnLib.GetObject(IDMObjects.idmObjectType.idmObjTypeFolder, FnFolder)

                If objFnFolder Is Nothing Then
                    Response = False
                    RetryCount += 1
                Else
                    Response = True
                    If modMain.g_DebugMode Then Console.WriteLine("Filenet folder allready exists!")
                    ' Return True
                    IsSuccessful = True
                    Marshal.ReleaseComObject(objFnFolder)
                End If

            Catch ex As Exception
                If modMain.g_DebugMode Then Console.WriteLine("Failed to Retrieve Folder Information. The application will retry.")
                Threading.Thread.Sleep(1000)
                ExceptionPlaceholder = ex
                RetryCount += 1
            End Try

            If RetryCount > 1 Then
                If ExceptionPlaceholder.HResult = -2147215867 Then
                    'fnLogger.WriteToLog("HResult = -2147215867")
                    ' NEED TO HANDLE -2147215866
                    Return False
                End If
            Else
                'Console.WriteLine("Entering retry loop")
            End If
        Loop Until IsSuccessful = True OrElse RetryCount >= 3

        If IsSuccessful Then
            If modMain.g_DebugMode Then
                Console.WriteLine("[D] FnFolderExists: Successfully Retrieved Folder information")
                modLogger.WriteToLog("[D] FnFolderExists: Successfully Retrieved Folder information")
            End If
            Return True
        Else
            Console.WriteLine("[E] Exception occured in FnFolderExists of modFilenetHandler! Could not Retrieve folder information! Specific exception message:")
            Console.WriteLine(ExceptionPlaceholder.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in FnFolderExists of modFilenetHandler! Could not Retrieve folder information! Specific exception message:")
            modLogger.WriteToLog(ExceptionPlaceholder.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.StackTrace)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.HResult)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.Source)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.Message)
            'ExceptionPlaceholder.HResult
            'Marshal.ReleaseComObject(objFnFolder)
            Return False
        End If


    End Function


    Public Function CreateFnFolder(FolderName As String, FolderBasePath As String) As Boolean
        Dim objFnFolder As New IDMObjects.Folder
        Dim objFnParentFolder As New IDMObjects.Folder

        Try
            If modMain.g_DebugMode Then
                Console.WriteLine("[D] Attempting to create " & FolderName & " in " & FolderBasePath)
                modLogger.WriteToLog("[D] Attempting to create " & FolderName & " in " & FolderBasePath)
                Debug.Print("[D] Attempting to create " & FolderName & " in " & FolderBasePath)
            End If


            objFnParentFolder = FnLib.GetObject(IDMObjects.idmObjectType.idmObjTypeFolder, FolderBasePath)
            objFnFolder = objFnParentFolder.CreateSubFolder(FolderName)
            objFnFolder.SaveNew()

            If modMain.g_DebugMode Then
                Console.WriteLine("[D] Folder Successfully created! - Disposing objects and exiting function...")
                modLogger.WriteToLog("[D] Folder Successfully created! - Disposing objects and exiting function...")
            End If
            Marshal.ReleaseComObject(objFnFolder)
            Marshal.ReleaseComObject(objFnParentFolder)

            Return True

        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in CreateFnFolder of modFilenetHandler! Could not create a folder for the specific name! Specific exception message:")
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in CreateFnFolder of modFilenetHandler! Could not create a folder for the specific name! Specific exception message: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ex.StackTrace)
            Marshal.ReleaseComObject(objFnFolder)
            Marshal.ReleaseComObject(objFnParentFolder)
            Return False
        End Try

    End Function



    Public Function CommitFnDocument(ByVal FileName As String, ByRef DocumentNumber As Long, Optional EposName As String = "") As Boolean
        Dim objFnDoc As IDMObjects.Document
        Dim RetryCount As Integer = 0
        Dim IsSuccessful As Boolean = False
        Dim ExceptionPlaceholder As Exception
        Dim NewFileLocation As String
        Dim ImportFileInfo As IO.FileInfo

        Try

            If modMain.g_DebugMode Then
                Console.WriteLine("[D] Trying to commit document " & FileName & " to Filenet.")
                modLogger.WriteToLog("[D] Trying to commit document " & FileName & " to Filenet.")
            End If

            If modFsFunctions.FileExists(FileName) Then  'If file exists
                'before we even try to proccess the document we are going to see if it has the right permission
                If modFsFunctions.FileHasReadAccess(FileName) Then
                    ImportFileInfo = New IO.FileInfo(FileName)
                    If Not ImportFileInfo.Length = 0 Then 'check for zero size first
                        NewFileLocation = System.AppDomain.CurrentDomain.BaseDirectory & "LocalStage\" & EposName
                        'If Directory with LocalStage folder does NOT exists create it.
                        modFsFunctions.ValidateCreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory & "LocalStage")
                        If modMain.g_DebugMode Then Console.WriteLine("[D] Trying to copy target file to [local stage]: " & NewFileLocation)
                        modFsFunctions.CopyFile(FileName, NewFileLocation) 'copytostage
                        If modMain.g_DebugMode Then Console.WriteLine("[D] Copied to local stage.")
                    Else
                        modMain.global_ZERO_SIZE_COUNTER += 1
                        'Increase the number of zero size  Documents for the Finalize Log
                        modStatistics.increasezerosize()

                        If modMain.g_DebugMode Then
                            Console.WriteLine("[D] The application detected that " & FileName & " is a ZERO SIZE file and cannot be imported to Filenet...")
                            Console.WriteLine("[D] Appending the Filename to the NOT EXISTS log file.. (File will be skipped) ")
                            modLogger.WriteToLog("[D] The application detected that " & FileName & " is a ZERO SIZE file and cannot be imported to Filenet...")
                            modLogger.WriteToLog("[D] Appending the Filename to the NOT EXISTS log file.. (File will be skipped) ")
                        End If
                        modLogger.WriteToLog("For further information check the Not_Exists log file...")
                        'on the first zero size create log file. on the rest append to the file--
                        If modMain.global_ZERO_SIZE_COUNTER = 1 And modMain.global_NOT_EXISTS_COUNTER = 0 Then
                            modLogger.generateNOTEXISTSLog()
                            modLogger.WriteToNOTEXISTSLog(FileName & " [ZERO SIZE]")
                            Console.WriteLine(FileName & " is ZERO SIZE !!!!")

                            Return False
                        Else
                            modLogger.WriteToNOTEXISTSLog(FileName & " [ZERO SIZE]")
                            Console.WriteLine(FileName & " is ZERO SIZE !!!!")
                            Return False
                        End If
                    End If
                Else
                    modLogger.WriteToLog("Not the correct permissions to access the file: " & FileName)
                    Console.WriteLine("Not the correct permissions to access the file: " & FileName)
                    Return False
                    Exit Function
                End If

            Else
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] The application could not find the file: " & FileName & " at the specified location")
                    Console.WriteLine("[D] Appending the Filename to the NOT EXISTS log file.. (File will be skipped) ")
                    modLogger.WriteToLog("[D] The application could not find the file: " & FileName & " at the specified location")
                    modLogger.WriteToLog("[D] Appending the Filename to the NOT EXISTS log file.. (File will be skipped) ")
                End If
                modMain.global_NOT_EXISTS_COUNTER += 1

                'on the first not exists create log file. on the rest append to the file--

                'increase the number of Dont exist Documents for the Finalize Log
                modStatistics.IncreaseDontexistCounter()

                modLogger.WriteToLog("For further information check the Not_Exists log file...")
                If modMain.global_NOT_EXISTS_COUNTER = 1 And modMain.global_ZERO_SIZE_COUNTER = 0 Then
                    modLogger.generateNOTEXISTSLog()
                    modLogger.WriteToNOTEXISTSLog(FileName & " [DOES NOT EXIST]")
                    Console.WriteLine(FileName & " DOES NOT EXIST !!!!")
                Else
                    modLogger.WriteToNOTEXISTSLog(FileName & " [DOES NOT EXIST]")
                    Console.WriteLine(FileName & " DOES NOT EXIST !!!!")
                End If
                Return False
                Exit Function
            End If

        Catch ex As Exception
            Console.WriteLine("[E] Exception occured in CommitFnDocument of modFilenetHandler! Exception in pre-commital FS functions (file will be skipped)! Specific exception message:")
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("[E] Exception occured in CommitFnDocument of modFilenetHandler! Exception in pre-commital FS functions (file will be skipped)! Specific exception message: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ex.StackTrace)
            Return False
            Exit Function
        End Try

        'Perform Filenet commital 


        Do
            Try
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Trying to commit document " & FileName)
                    modLogger.WriteToLog("[D] Trying to commit document " & FileName)
                End If
                objFnDoc = FnLib.CreateObject(IDMObjects.idmObjectType.idmObjTypeDocument, modAppConfigSettings.ClassName)
                objFnDoc.SaveNew(NewFileLocation, IDMObjects.idmSaveNewOptions.idmDocSaveNewKeep)
                DocumentNumber = objFnDoc.ID

                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Document" & FileName & " was commited successfully, and it was assigned the document id= " & DocumentNumber)
                    modLogger.WriteToLog("[D] Document" & FileName & " was commited successfully, and it was assigned the document id= " & DocumentNumber)
                End If
                Marshal.ReleaseComObject(objFnDoc)
                IsSuccessful = True
                'DELETE From stage
                modFsFunctions.DeleteFile(NewFileLocation) 'Maybe handle deletion safely?
                If modMain.g_DebugMode Then
                    Console.WriteLine("[D] Document DELETED From Stage")
                    modLogger.WriteToLog("[D] Document DELETED From Stage")
                End If
                'Return True
            Catch e As Exception
                Console.WriteLine("[E] Exception occured in CommitFnDocument of modFilenetHandler! Exception in the file commital stage! Specific exception message:")
                Console.WriteLine(e.GetBaseException)
                ExceptionPlaceholder = e
                RetryCount += 1
                Threading.Thread.Sleep(2000)
            End Try
        Loop Until IsSuccessful = True OrElse RetryCount > 2


        If IsSuccessful Then
            Return True
        Else
            Console.WriteLine("[E] Exception occured in CommitFnDocument of modFilenetHandler! Could not commit the document: " & FileName & " to Filenet! Specific exception message:")
            Console.WriteLine(ExceptionPlaceholder.GetBaseException.ToString)
            Console.WriteLine(ExceptionPlaceholder.Source)
            Console.WriteLine(ExceptionPlaceholder.Message)
            modLogger.WriteToLog("[E] Exception occured in CommitFnDocument of modFilenetHandler! Could not commit the document: " & FileName & " to Filenet! Specific exception message:")
            modLogger.WriteToLog(ExceptionPlaceholder.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.Source)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.Message)
            modLogger.WriteToLog("")
            modLogger.WriteToLog(ExceptionPlaceholder.StackTrace)
            Return False
        End If
    End Function

End Module

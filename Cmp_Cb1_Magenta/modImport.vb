Imports System.Runtime.InteropServices

Module modImport
    Dim objEposDoc As clsEposDocument
    Dim OteDB As New clsDB
    Dim Map_SubCategory_Result As Integer
    Dim insuferrorcount As Long = 0
    Dim NullBilling As Boolean = False
    Dim DocumenstBool As Boolean = False
    Dim FoldersBool As Boolean = False
    Dim Folder_contentsBool As Boolean = False


    Public Function ImportBatch() As Boolean

        Dim Dbnullcounter As Integer = 0
        Dim i As Long
        Dim y As Long
        Dim x As Long
        Dim TotalDocumentsRC As Long
        Dim FileNamesArr(,) As String
        Dim FileNamesRS As OleDb.OleDbDataReader
        OteDB.initConnection(modAppConfigSettings.OterevDBname, modAppConfigSettings.OterevDBuser, modAppConfigSettings.OterevDBpwd)
        Dim firstofloop As Boolean = True
        Dim ImportSql As String
        Dim FilesToImportQueryTimeStart As New DateTime
        Dim FilesToImportQueryTimeStop As New DateTime
        Dim FilesToImportQueryTimeRunTime As New TimeSpan


        i = 0
        y = 0
        x = 0
        TotalDocumentsRC = 0
        modStatistics.ResetCounters()


        Try
            modLogger.WriteToLog("Initializing Main Import procedure...")
            modLogger.WriteToLog("Querying Database for the total number of documents to be imported...")
            Console.WriteLine("Initializing Main Import procedure...")
            Console.WriteLine("Querying Database for the total number of documents to be imported...")
            'Get the total number of documents to import using a count query (total number used for main loop)
            'TotalDocumentsRC = eposDB.ExecuteCount(modAppConfigSettings.TotalRecordCountSql)

            If modAppConfigSettings.RowNumEnable = 1 Then
                modLogger.WriteToLog("The RowNum Option is Enable")
                Console.WriteLine("The RowNum Option is Enable")
                ImportSql = modAppConfigSettings.EposFilesSql + " " + modAppConfigSettings.RowNum
            Else
                ImportSql = modAppConfigSettings.EposFilesSql
            End If

            'Get Start Time before query
            FilesToImportQueryTimeStart = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")

            OteDB.ExecuteSelect(ImportSql)

            FileNamesRS = OteDB.GetResultSet
            'Populate the array by looping through the result set

            'get stop time after query
            FilesToImportQueryTimeStop = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
            'get the total time of the query run
            FilesToImportQueryTimeRunTime = FilesToImportQueryTimeStop - FilesToImportQueryTimeStart

            modLogger.WriteToLog("The Query for Files to import run for: " & FilesToImportQueryTimeRunTime.ToString)

            If FileNamesRS.HasRows Then

                Do While FileNamesRS.Read() 'Convert input to String as a precaution ----CHECK for empty values?


                    If firstofloop Then 'flag the first loop in order to read the first row of the fetched data and read the count in the last col. 
                        'after that we read all the data normaly in loop 
                        TotalDocumentsRC = FileNamesRS.GetValue(13)
                        'Redim array according to expected input
                        If ((TotalDocumentsRC < 0) Or (TotalDocumentsRC = 0)) Then
                            'TotalDocumentsRC IF statement
                            'If no documents are found in epos db exit
                            modLogger.WriteToLog("Found " & TotalDocumentsRC & " for import. Exiting Function And terminating Import proccess...")
                            Console.WriteLine("Found " & TotalDocumentsRC & " for import. Exiting Function And terminating Import proccess...")
                            modFilenetHandler.FnLogoff()
                            Return True
                            Exit Function
                        Else
                            modLogger.WriteToLog("")
                            modLogger.WriteToLog("Found " & TotalDocumentsRC & " for import.")
                            Console.WriteLine("")
                            Console.WriteLine("Found " & TotalDocumentsRC & " for import. ")

                            ReDim FileNamesArr(TotalDocumentsRC - 1, 12)

                            'set the total number of Docs that the application will process, Regardless data accuracy
                            modStatistics.setTotalDocs(TotalDocumentsRC)

                            If IsDBNull(FileNamesRS.GetValue(0)) Or IsDBNull(FileNamesRS.GetValue(1)) Or IsDBNull(FileNamesRS.GetValue(2)) Or IsDBNull(FileNamesRS.GetValue(3)) Or IsDBNull(FileNamesRS.GetValue(4)) Or IsDBNull(FileNamesRS.GetValue(5)) Or IsDBNull(FileNamesRS.GetValue(6)) Or IsDBNull(FileNamesRS.GetValue(7)) Or IsDBNull(FileNamesRS.GetValue(8)) Or IsDBNull(FileNamesRS.GetValue(9)) Or IsDBNull(FileNamesRS.GetValue(10)) Then
                                Dbnullcounter += 1
                                modStatistics.IncreaseInsufficient()

                                'DIGITAL_ORDER_ID,DOCUMENT_FILE_INFO_ID,TOPIC,CATEGORY,SUBCATEGORY,NETAPP,FILE_PATH,FILE_TYPE,CUSTOMER_CODE,BILLING_ACCOUNT_ID,SHOP_CODE,DOC_DATE,IMPORT_STATUS

                                If Not IO.File.Exists(System.AppDomain.CurrentDomain.BaseDirectory & "Logs\" & modLogger.generateInsufficientDataLogName()) Then
                                    modLogger.generateINSUFFICIENTDATALog()
                                End If
                                modLogger.WriteToINSUFFICIENTDATALog("#" & Dbnullcounter)
                                modLogger.WriteToINSUFFICIENTDATALog("The application encountered null or empty values while trying to import document. (results from SQLFilesToImport query)")
                                modLogger.WriteToINSUFFICIENTDATALog("Specific File information is: ")
                                modLogger.WriteToINSUFFICIENTDATALog("!Rows with no data means DataBaseNULL value!")
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved DIGITAL_ORDER_ID is: " & FileNamesRS.GetValue(0))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved DOCUMENT_FILE_INFO_ID is: " & FileNamesRS.GetValue(1))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved TOPIC is: " & FileNamesRS.GetValue(2))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved CATEGORY is: " & FileNamesRS.GetValue(3))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved SUBCATEGORY is: " & FileNamesRS.GetValue(4))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved NETAPP is: " & FileNamesRS.GetValue(5))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved FILE_PATH is: " & FileNamesRS.GetValue(6))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved FILE_TYPE is: " & FileNamesRS.GetValue(7))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved CUSTOMER_CODE is: " & FileNamesRS.GetValue(8))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved BILLING_ACCOUNT_ID is: " & FileNamesRS.GetValue(9))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved SHOP_CODE is: " & FileNamesRS.GetValue(10))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved DOC_DATE is: " & FileNamesRS.GetValue(11))
                                modLogger.WriteToINSUFFICIENTDATALog("Retrieved IMPORT_STATUS is: " & FileNamesRS.GetValue(12))
                                modLogger.WriteToINSUFFICIENTDATALog("")

                                If (IsDBNull(FileNamesRS.GetValue(0)) Or IsDBNull(FileNamesRS.GetValue(1))) Then

                                    modLogger.WriteToINSUFFICIENTDATALog("The current document cannot set to repeat or error state due to essential null values!!!")
                                    modLogger.WriteToINSUFFICIENTDATALog("")
                                    modStatistics.increaseUnError()
                                Else

                                    SetDocumentToErrorStateDueNonData(CStr(FileNamesRS.GetValue(0)), CInt(FileNamesRS.GetValue(1)))
                                    insuferrorcount += 1
                                End If

                                firstofloop = False
                                Console.WriteLine("First Loop Complete")
                                Continue Do
                            End If
                            'DIGITAL_ORDER_ID,DOCUMENT_FILE_INFO_ID,TOPIC,CATEGORY,SUBCATEGORY,NETAPP,FILE_PATH,FILE_TYPE,CUSTOMER_CODE,BILLING_ACCOUNT_ID,SHOP_CODE,DOC_DATE,IMPORT_STATUS


                            FileNamesArr(i, 0) = CStr(FileNamesRS.GetValue(0)) 'DIGITAL_ORDER_ID
                            FileNamesArr(i, 1) = CStr(FileNamesRS.GetValue(1)) 'DOCUMENT_FILE_INFO_ID
                            FileNamesArr(i, 2) = CStr(FileNamesRS.GetValue(2)) 'TOPIC
                            FileNamesArr(i, 3) = CStr(FileNamesRS.GetValue(3)) 'CATEGORY
                            FileNamesArr(i, 4) = CStr(FileNamesRS.GetValue(4)) 'SUBCATEGORY
                            FileNamesArr(i, 5) = CStr(FileNamesRS.GetValue(5)) 'NETAPP
                            FileNamesArr(i, 6) = CStr(FileNamesRS.GetValue(6)) 'FILE_PATH
                            FileNamesArr(i, 7) = CStr(FileNamesRS.GetValue(7)) 'FILE_TYPE
                            FileNamesArr(i, 8) = CStr(FileNamesRS.GetValue(8)) 'CUSTOMER_CODE
                            FileNamesArr(i, 9) = CStr(FileNamesRS.GetValue(9)) 'BILLING_ACCOUNT_ID

                            If Not IsDBNull(FileNamesRS.GetValue(9)) Then
                                FileNamesArr(i, 10) = CStr(FileNamesRS.GetValue(10)) 'SHOP_CODE
                            Else
                                FileNamesArr(i, 10) = vbNullString 'SHOP_CODE
                            End If
                            FileNamesArr(i, 11) = CStr(FileNamesRS.GetValue(11)) 'DOC_DATE
                            FileNamesArr(i, 12) = CStr(FileNamesRS.GetValue(12)) 'IMPORT_STATUS



                            i += 1
                            firstofloop = False
                            Console.WriteLine("First Loop Complete")
                        End If

                        'break the loop and move to the next fetched row in order to continue the reading
                        Continue Do
                    End If


                    'after the read of the first loop the execution moves here.

                    If IsDBNull(FileNamesRS.GetValue(0)) Or IsDBNull(FileNamesRS.GetValue(1)) Or IsDBNull(FileNamesRS.GetValue(2)) Or IsDBNull(FileNamesRS.GetValue(3)) Or IsDBNull(FileNamesRS.GetValue(4)) Or IsDBNull(FileNamesRS.GetValue(5)) Or IsDBNull(FileNamesRS.GetValue(6)) Or IsDBNull(FileNamesRS.GetValue(7)) Or IsDBNull(FileNamesRS.GetValue(8)) Or IsDBNull(FileNamesRS.GetValue(9)) Or IsDBNull(FileNamesRS.GetValue(10)) Then
                        Dbnullcounter += 1
                        modStatistics.IncreaseInsufficient()

                        'DIGITAL_ORDER_ID,DOCUMENT_FILE_INFO_ID,TOPIC,CATEGORY,SUBCATEGORY,NETAPP,FILE_PATH,FILE_TYPE,CUSTOMER_CODE,BILLING_ACCOUNT_ID,SHOP_CODE,DOC_DATE,IMPORT_STATUS

                        If Not IO.File.Exists(System.AppDomain.CurrentDomain.BaseDirectory & "Logs\" & modLogger.generateInsufficientDataLogName()) Then
                            modLogger.generateINSUFFICIENTDATALog()
                        End If
                        modLogger.WriteToINSUFFICIENTDATALog("#" & Dbnullcounter)
                        modLogger.WriteToINSUFFICIENTDATALog("The application encountered null or empty values while trying to import document. (results from SQLFilesToImport query)")
                        modLogger.WriteToINSUFFICIENTDATALog("Specific File information is: ")
                        modLogger.WriteToINSUFFICIENTDATALog("!Rows with no data means DataBaseNULL value!")
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved DIGITAL_ORDER_ID is: " & FileNamesRS.GetValue(0))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved DOCUMENT_FILE_INFO_ID is: " & FileNamesRS.GetValue(1))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved TOPIC is: " & FileNamesRS.GetValue(2))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved CATEGORY is: " & FileNamesRS.GetValue(3))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved SUBCATEGORY is: " & FileNamesRS.GetValue(4))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved NETAPP is: " & FileNamesRS.GetValue(5))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved FILE_PATH is: " & FileNamesRS.GetValue(6))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved FILE_TYPE is: " & FileNamesRS.GetValue(7))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved CUSTOMER_CODE is: " & FileNamesRS.GetValue(8))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved BILLING_ACCOUNT_ID is: " & FileNamesRS.GetValue(9))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved SHOP_CODE is: " & FileNamesRS.GetValue(10))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved DOC_DATE is: " & FileNamesRS.GetValue(11))
                        modLogger.WriteToINSUFFICIENTDATALog("Retrieved IMPORT_STATUS is: " & FileNamesRS.GetValue(12))
                        modLogger.WriteToINSUFFICIENTDATALog("")

                        If (IsDBNull(FileNamesRS.GetValue(0)) Or IsDBNull(FileNamesRS.GetValue(1))) Then

                            modLogger.WriteToINSUFFICIENTDATALog("The current document cannot set to repeat or error state due to essential null values!!!")
                            modLogger.WriteToINSUFFICIENTDATALog("")
                            modStatistics.increaseUnError()
                        Else

                            SetDocumentToErrorStateDueNonData(CStr(FileNamesRS.GetValue(0)), CInt(FileNamesRS.GetValue(1)))
                            insuferrorcount += 1
                        End If
                    Else

                        FileNamesArr(i, 0) = CStr(FileNamesRS.GetValue(0)) 'DIGITAL_ORDER_ID
                        FileNamesArr(i, 1) = CStr(FileNamesRS.GetValue(1)) 'DOCUMENT_FILE_INFO_ID
                        FileNamesArr(i, 2) = CStr(FileNamesRS.GetValue(2)) 'TOPIC
                        FileNamesArr(i, 3) = CStr(FileNamesRS.GetValue(3)) 'CATEGORY
                        FileNamesArr(i, 4) = CStr(FileNamesRS.GetValue(4)) 'SUBCATEGORY
                        FileNamesArr(i, 5) = CStr(FileNamesRS.GetValue(5)) 'NETAPP
                        FileNamesArr(i, 6) = CStr(FileNamesRS.GetValue(6)) 'FILE_PATH
                        FileNamesArr(i, 7) = CStr(FileNamesRS.GetValue(7)) 'FILE_TYPE
                        FileNamesArr(i, 8) = CStr(FileNamesRS.GetValue(8)) 'CUSTOMER_CODE
                        FileNamesArr(i, 9) = CStr(FileNamesRS.GetValue(9)) 'BILLING_ACCOUNT_ID

                        If Not IsDBNull(FileNamesRS.GetValue(9)) Then
                            FileNamesArr(i, 10) = CStr(FileNamesRS.GetValue(10)) 'SHOP_CODE
                        Else
                            FileNamesArr(i, 10) = vbNullString 'SHOP_CODE
                        End If
                        FileNamesArr(i, 11) = CStr(FileNamesRS.GetValue(11)) 'DOC_DATE
                        FileNamesArr(i, 12) = CStr(FileNamesRS.GetValue(12)) 'IMPORT_STATUS

                        i += 1
                        Console.Write("{0}Rows Read: {1}/{2}", vbCr, i, TotalDocumentsRC - Dbnullcounter)
                    End If
                Loop
            Else
                Console.WriteLine("Found 0 for import. Exiting Function And terminating Import proccess...")
                modLogger.WriteToLog("Found 0 for import. Exiting Function And terminating Import proccess...")
                OteDB.CloseConnection()
                Return False
                Exit Function
            End If

            'Close OladbDataReader when finish reading to release all resources
            FileNamesRS.Close()

            'Connect to Filenet here after fetching document 
            If Not modFilenetHandler.initFilenet Then
                modLogger.WriteToLog("[E] Failed to Initialize connection to Filenet.")
                modLogger.WriteToLog("[E] The application will now terminate...")
                Console.WriteLine("[E] Failed to Initialize connection to Filenet.")
                Console.WriteLine("[E] The application will now terminate...")
                Return False
            End If

            Console.WriteLine("")
            modLogger.WriteToLog("Retrieved the Required File Name information.")
            Console.WriteLine("Retrieved the Required File Name information.")

            'Console.WriteLine("")
            modLogger.WriteToLog("Initializing Import procedure for all Documents...")
            Console.WriteLine("Initializing Import procedure for all Documents...")
            Console.WriteLine("")
            If Dbnullcounter > 0 Then
                modLogger.WriteToLog("Found: " & Dbnullcounter & " Documents with insufficient data! Not included in then main procedure!")
                modLogger.WriteToLog("For Further details check the INSUFFICIENT DATA Log file")
                modLogger.WriteToLog("Remaining Documentrs for processing:  " & TotalDocumentsRC - Dbnullcounter)
                Console.WriteLine("Found: " & Dbnullcounter & " Documents with insufficient data! Not included in then main procedure!")
                Console.WriteLine("INSUFFICIENT DATA Log file created successfully!")
                Console.WriteLine("Remaining Documents for processing:  " & TotalDocumentsRC - Dbnullcounter)
            End If


            'Loop for each Document and try to commit it.
            For y = 0 To (FileNamesArr.GetLength(0) - (Dbnullcounter + 1))
                modLogger.WriteToLog("")
                modLogger.WriteToLog("Proccessing Document " & (y + 1) & " of " & TotalDocumentsRC - Dbnullcounter)
                Console.WriteLine("")
                Console.Write("Proccessing Document " & (y + 1) & " of " & TotalDocumentsRC - Dbnullcounter)

                objEposDoc = New clsEposDocument
                NullBilling = False
                'NullBilling = True

                'Filter possible Null or empty Fields in Epos DB rows - If data is missing create and write to INSUFFICIENT DATA LOG
                If Not (String.IsNullOrEmpty(FileNamesArr(y, 0)) Or String.IsNullOrEmpty(FileNamesArr(y, 1))) Then

                    objEposDoc.DIGITAL_ORDER_ID = CStr(FileNamesArr(y, 0)) 'DIGITAL_ORDER_ID
                    objEposDoc.DOCUMENT_FILE_INFO_ID = CStr(FileNamesArr(y, 1)) 'DOCUMENT_FILE_INFO_ID
                    objEposDoc.Topic = CStr(FileNamesArr(y, 2)) 'TOPIC
                    objEposDoc.Category = CStr(FileNamesArr(y, 3)) 'CATEGORY
                    objEposDoc.SubCategory = CStr(FileNamesArr(y, 4)) 'SUBCATEGORY
                    objEposDoc.NetApp = CInt(FileNamesArr(y, 5)) 'NETAPP
                    objEposDoc.FilePath = Replace(CStr(FileNamesArr(y, 6)), "/", "\") 'FILE_PATH
                    objEposDoc.Extension = CStr(FileNamesArr(y, 7)) 'FILE_TYPE
                    objEposDoc.Customer_Code = CStr(FileNamesArr(y, 8)) 'CUSTOMER_CODE

                    'If Not String.IsNullOrEmpty(FileNamesArr(y, 9)) Then
                    objEposDoc.BILLING_ACCOUNT_ID = CStr(FileNamesArr(y, 9)) 'BILLING_ACCOUNT_ID
                    'Else
                    '    'GDPR
                    '    'Registry11_fixed
                    '    'Registry11_mobile
                    '    If objEposDoc.Topic.Equals("GDPR") Or objEposDoc.Topic.Equals("Registry11_fixed") Or objEposDoc.Topic.Equals("Registry11_mobile") Then 'Check if is acceptable to have null in billing acc. 
                    '        NullBilling = True
                    '        objEposDoc.BILLING_ACCOUNT_ID = "-1"

                    '    Else 'set to error state
                    '        Console.WriteLine(" Set Document to error state because of null Billing _account_id in order where expected.")
                    '        modLogger.WriteToLog("Set Document to error state because of null Billing _account_id in order where expected.")
                    '        modLogger.WriteToLog("DIGITAL_ORDER_ID = " & objEposDoc.DIGITAL_ORDER_ID)
                    '        modLogger.WriteToLog("DOCUMENT_FILE_INFO_ID = " & objEposDoc.DOCUMENT_FILE_INFO_ID)
                    '        modLogger.WriteToLog("")

                    '        SetDocumentToErrorState(objEposDoc)
                    '        Continue For
                    '    End If

                    'End If

                    If Not (String.IsNullOrEmpty(FileNamesArr(y, 10))) Then
                        objEposDoc.Shop_Code = CStr(FileNamesArr(y, 10)) 'SHOP_CODE
                    Else
                        objEposDoc.Shop_Code = ""
                    End If

                    objEposDoc.DocDate = CStr(FileNamesArr(y, 11)) 'DOC_DATE
                    objEposDoc.FileState = CInt(FileNamesArr(y, 12)) 'IMPORT_STATUS

                    'Un Comment below to add an "/" at the start of the file path if there is not
                    'If Not objEposDoc.FilePath.StartsWith("/", StringComparison.Ordinal) Then
                    '    Debug.Print("Not Found")
                    '    objEposDoc.FilePath = objEposDoc.FilePath.Insert(0, "/")
                    '    Debug.Print(objEposDoc.FilePath)
                    'End If

                    'Check For File Path validity. If not Set Document to Error State.
                    Dim nameNoExtension As String = modFsFunctions.SubstringAfterSpecialChar(objEposDoc.FilePath, "/")
                    If nameNoExtension.Equals("") Then
                        nameNoExtension = modFsFunctions.SubstringAfterSpecialChar(objEposDoc.FilePath, "\")
                        If nameNoExtension.Equals("") Then
                            Console.WriteLine("File_path Found: " & objEposDoc.FilePath)
                            Console.WriteLine("Set Document in Error State Due To Non Valid Format Of File Path!")
                            Console.WriteLine("DIGITAL_ORDER_ID = " & objEposDoc.DIGITAL_ORDER_ID)
                            Console.WriteLine("DOCUMENT_FILE_INFO_ID = " & objEposDoc.DOCUMENT_FILE_INFO_ID)
                            modLogger.WriteToLog("File_path Found: " & objEposDoc.FilePath)
                            modLogger.WriteToLog("Set Document in Error State Due To Non Valid Format Of File Path!")
                            modLogger.WriteToLog("DIGITAL_ORDER_ID = " & objEposDoc.DIGITAL_ORDER_ID)
                            modLogger.WriteToLog("DOCUMENT_FILE_INFO_ID = " & objEposDoc.DOCUMENT_FILE_INFO_ID)
                            modLogger.WriteToLog("")
                            SetDocumentToErrorState(objEposDoc)
                            Continue For
                        End If
                    Else

                    End If

                    objEposDoc.AbsoluteFilePath = modAppConfigSettings.Path & objEposDoc.FilePath & objEposDoc.Extension
                    objEposDoc.FileName = nameNoExtension & objEposDoc.Extension

                    Debug.Print("Doc Filename is: " & objEposDoc.FileName)
                    Console.Write(" Doc Filename is: " & objEposDoc.FileName)
                    Console.WriteLine()

                    If NullBilling Then 'Only customer Folder
                        Dim Folders(0) As String
                        ReDim objEposDoc.FolderNumber(0)

                        'Define Folders (1 in our case) with the matching anagnoristic
                        Folders(0) = objEposDoc.Customer_Code

                        If Folders(0).Contains("-") Then
                            Folders(0) = Replace(Folders(0), "-", "") & "_"
                        End If
                        If Folders(0).Contains(".") Then
                            Folders(0) = Replace(Folders(0), ".", "") & "_"
                        End If

                        objEposDoc.FillObjectFolders(Folders)

                    Else ' customer + Billing Folders
                        Dim Folders(1) As String
                        ReDim objEposDoc.FolderNumber(1)

                        'Define Folders (2 in our case) with the matching anagnoristic
                        Folders(0) = objEposDoc.Customer_Code

                        If Folders(0).Contains("-") Then
                            Folders(0) = Replace(Folders(0), "-", "") & "_"
                        End If
                        If Folders(0).Contains(".") Then
                            Folders(0) = Replace(Folders(0), ".", "") & "_"
                        End If

                        Folders(1) = Folders(0) & "/" & objEposDoc.BILLING_ACCOUNT_ID
                        If Folders(1).Contains("-") Then
                            Folders(1) = Replace(Folders(1), "-", "") & "_"
                        End If
                        If Folders(1).Contains(".") Then
                            Folders(1) = Replace(Folders(1), ".", "") & "_"
                        End If

                        objEposDoc.FillObjectFolders(Folders)

                    End If


                    If modMain.g_DebugMode Then
                        modLogger.WriteToLog("[DEBUG MODE]: Printing all data retrieved for document " & objEposDoc.FileName)
                        modLogger.WriteToLog("[D]=Retrieved DIGITAL_ORDER_ID is: " & objEposDoc.DIGITAL_ORDER_ID)
                        modLogger.WriteToLog("[D]=Retrieved DOCUMENT_FILE_INFO_ID is: " & objEposDoc.DOCUMENT_FILE_INFO_ID)
                        modLogger.WriteToLog("[D]=Retrieved TOPIC is: " & objEposDoc.Topic)
                        modLogger.WriteToLog("[D]=Retrieved CATEGORY is: " & objEposDoc.Category)
                        modLogger.WriteToLog("[D]=Retrieved SUBCATEGORY is: " & objEposDoc.SubCategory)
                        modLogger.WriteToLog("[D]=Retrieved NETAPP is: " & objEposDoc.NetApp)
                        modLogger.WriteToLog("[D]=Retrieved FILE_PATH is: " & objEposDoc.FilePath)
                        modLogger.WriteToLog("[D]=Retrieved FILE_TYPE is: " & objEposDoc.Extension)
                        modLogger.WriteToLog("[D]=Retrieved CUSTOMER_CODE is: " & objEposDoc.Customer_Code)
                        modLogger.WriteToLog("[D]=Retrieved BILLING_ACCOUNT_ID is: " & objEposDoc.BILLING_ACCOUNT_ID)
                        modLogger.WriteToLog("[D]=Retrieved SHOP_CODE is: " & objEposDoc.Shop_Code)
                        modLogger.WriteToLog("[D]=Retrieved DOC_DATE is: " & objEposDoc.DocDate)
                        modLogger.WriteToLog("[D]=Retrieved IMPORT_STATUS is: " & objEposDoc.FileState)
                        modLogger.WriteToLog("")
                    End If

                    If RetrieveMetadataFromEpos(objEposDoc) Then
                        modFilenetHandler.FnLogon() 'Try to Log On in case connection to Filenet was dropped during execution

                        If CreateFnFoldersForDocument(objEposDoc) Then 'If Filenet folders cannot be created for some reason, then skip the whole procces (If createFnFoldersForDocument was at a later step in the loop and the file was already commited it would require a rollback procedure)
                            If CommitDocument(objEposDoc) Then 'Include 3 commital stages.

                                Dim tempEndBoll As Boolean = NullBilling
                                For i = 0 To UBound(objEposDoc.FolderNumber) 'Write every folder that was created
                                    If Not tempEndBoll Then
                                        tempEndBoll = True
                                        Continue For
                                    End If
                                    Console.WriteLine("File: " & objEposDoc.FileName & " of order: " & objEposDoc.DIGITAL_ORDER_ID & " was commited Successfully into FileNet with Folder Number: " & objEposDoc.FolderNumber(i) & " and  Document ID: " & objEposDoc.F_Docnumber)
                                    modLogger.WriteToLog("File: " & objEposDoc.FileName & " of order: " & objEposDoc.DIGITAL_ORDER_ID & "  was commited Successfully into FileNet with Folder Number: " & objEposDoc.FolderNumber(i) & " and  Document ID: " & objEposDoc.F_Docnumber)
                                Next

                                modStatistics.IncreaseSuccessfulCounter()
                                OteDB.DisposeObjects()

                                'Delete document after it has been successfully commited!!!!!!!!
                                Dim DocumentDeleted As Boolean
                                Dim DeleteAttemptsCount As Integer

                                DocumentDeleted = False
                                DeleteAttemptsCount = 0

                                Do
                                    DocumentDeleted = modFsFunctions.DeleteFile(objEposDoc.AbsoluteFilePath) ' If return false the loop will try again to delete the file
                                    If Not DocumentDeleted Then 'if DocumentDeleted is false then it will increase the DeleteAttemptsCount +1 and put thread to sleep 
                                        DeleteAttemptsCount += 1
                                        Threading.Thread.Sleep(modAppConfigSettings.SleepInterval)
                                    End If

                                Loop Until DocumentDeleted OrElse DeleteAttemptsCount > 2

                                If DeleteAttemptsCount > 0 Then
                                    Console.WriteLine("Delete Attempts Count: " & DeleteAttemptsCount)
                                    modLogger.WriteToLog("Delete Attempts Count: " & DeleteAttemptsCount)
                                End If

                                ' If (modFsFunctions.DeleteFile(objEposDoc.AbsoluteFilePath)) Then 'Consider doing something with undeletable documents
                                If DocumentDeleted Then
                                    Debug.Print("File: " & objEposDoc.FileName & " was successfully deleted from the network storage")
                                    If g_DebugMode Then
                                        Console.WriteLine("[D] File: " & objEposDoc.FileName & " was successfully deleted from the network storage")
                                        modLogger.WriteToLog("[D] File: " & objEposDoc.FileName & " was successfully deleted from the network storage")
                                    End If

                                Else
                                    Debug.Print("File: " & objEposDoc.FileName & " could NOT be deleted from the network storage")
                                    Console.WriteLine("[D] File: " & objEposDoc.FileName & " could NOT be deleted from the network storage")
                                    modLogger.WriteToLog("[D] File: " & objEposDoc.FileName & " could NOT be deleted from the network storage")
                                    If g_DebugMode Then
                                        Console.WriteLine("[D] File: " & objEposDoc.FileName & " could NOT be deleted from the network storage")
                                        modLogger.WriteToLog("[D] File: " & objEposDoc.FileName & " could NOT be deleted from the network storage")
                                    End If
                                End If
                            Else
                                Console.WriteLine("File: " & objEposDoc.FileName & " of order: " & objEposDoc.DIGITAL_ORDER_ID & " could not be commited... Setting it to Repeat State.")
                                modLogger.WriteToLog("File: " & objEposDoc.FileName & " of order: " & objEposDoc.DIGITAL_ORDER_ID & " could not be commited... Setting it to Repeat State.")
                                SetDocumentToRepeatState(objEposDoc)

                            End If

                        Else
                            Console.WriteLine("Could not create Filenet folders for document: " & objEposDoc.FileName & " of order: " & objEposDoc.DIGITAL_ORDER_ID & " ... Setting it to Repeat State.")
                            modLogger.WriteToLog("Could not create Filenet folders for document: " & objEposDoc.FileName & " of order: " & objEposDoc.DIGITAL_ORDER_ID & " ... Setting it to Repeat State.")
                            SetDocumentToRepeatState(objEposDoc)

                        End If 'CreateFnFoldersForDocument If statement

                    Else 'LOGGED
                        Console.WriteLine("Could not retrieve MetaData for file: " & objEposDoc.FileName & " With DOCUMENT_FILE_INFO_ID = " & objEposDoc.DOCUMENT_FILE_INFO_ID & " ... Setting it to Repeat State.")
                        modLogger.WriteToLog("Could not retrieve MetaData for file: " & objEposDoc.FileName & " With DOCUMENT_FILE_INFO_ID = " & objEposDoc.DOCUMENT_FILE_INFO_ID & " ... Setting it to Repeat State.")
                        modStatistics.increaseMetadata()
                        SetDocumentToRepeatState(objEposDoc)

                    End If 'RetrieveMetadataFromEpos If statement

                Else
                    ' WRITE TO INSUFICIENT DATA LOG ALL FAILS
                    Console.WriteLine(" DIGITAL_ORDER_ID Or DOCUMENT_FILE_INFO_ID is null/empty! Set to error state. See The Insufficient Log for more Info.")
                    modLogger.WriteToLog(" DIGITAL_ORDER_ID Or DOCUMENT_FILE_INFO_ID is null/empty! Set to error state. See The Insufficient Log for more Info.")
                    modStatistics.IncreaseInsufficient()

                    If Not IO.File.Exists(System.AppDomain.CurrentDomain.BaseDirectory & "Logs\" & modLogger.generateInsufficientDataLogName()) Then
                        modLogger.generateINSUFFICIENTDATALog()
                    End If
                    modLogger.WriteToINSUFFICIENTDATALog("#" & Dbnullcounter)
                    modLogger.WriteToINSUFFICIENTDATALog("The application encountered null or empty values while trying to import document " & FileNamesArr(i, 0) & FileNamesArr(i, 3) & " (results from SQLFilesToImport query)")
                    modLogger.WriteToINSUFFICIENTDATALog("Specific File information is: ")
                    modLogger.WriteToINSUFFICIENTDATALog("Retrieved DIGITAL_ORDER_ID is: " & FileNamesArr(y, 0))
                    modLogger.WriteToINSUFFICIENTDATALog("Retrieved DOCUMENT_FILE_INFO_ID is: " & FileNamesArr(y, 1))
                    modLogger.WriteToINSUFFICIENTDATALog("The file will be skipped.")
                    modLogger.WriteToINSUFFICIENTDATALog("")

                    If Not SetDocumentToErrorState(objEposDoc) Then
                        Console.WriteLine("Not valid document data to set the document to error state!!!")
                        modLogger.WriteToLog("Not valid document data to set the document to ERROR state.")

                    Else
                        Console.WriteLine("Setting document to ERROR State and moving to the next file...")
                        insuferrorcount += 1
                    End If
                End If 'FILENAME Check IF statement

                'End of Current Document Proccessing.
                'Check if the time constraint is met
                If TimeIsUp(modAppConfigSettings.StopTime) Then
                    Console.WriteLine("-----------------------------------------")
                    Console.WriteLine("Permitted execution time is up! Must break execution due to time limitations set in the configuration file of this executable")
                    Console.WriteLine("-----------------------------------------")
                    modLogger.WriteToLog("-----------------------------------------")
                    modLogger.WriteToLog("Permitted execution time is up! Must break execution due to time limitations set in the configuration file of this executable")
                    modLogger.WriteToLog("-----------------------------------------")
                    modFilenetHandler.FnLogoff()
                    modMain.g_AbnormalLoopExit = True
                    Exit For
                End If


                If (Console.KeyAvailable) Then
                    If Console.ReadKey(True).KeyChar = "q"c Then
                        Console.WriteLine("USER ABORT COMMAND DETECTED BY Q KEYPRESS - CLEANING UP AND EXITING LOOP!")
                        Console.WriteLine("Aborted at : " & y + 1 & " documents. Cleaning up...")
                        modLogger.WriteToLog("USER ABORT COMMAND DETECTED BY Q KEYPRESS - CLEANING UP AND EXITING LOOP!")
                        modLogger.WriteToLog("Aborted at : " & y + 1 & " documents. Cleaning up...")
                        modFilenetHandler.FnLogoff()
                        modMain.g_AbnormalLoopExit = True
                        Exit For
                    End If

                End If

            Next y 'Move to Next document
            'When finished looping

            OteDB.CloseConnection()
            Return True ' exit import procedure.


            'Consider disposing objects and cleaning up here

        Catch ex As Exception
            Console.WriteLine("Exception occured in ImportBatch of modImport! MAIN LOOP EXECUTION WAS TERMINATED! Specific exception message: ")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            modLogger.WriteToLog("!THE MAIN IMPORT PROCEDURE WAS ABRUPTLY TERMINATED - EXITING IMPORT PROCCESS!")
            modLogger.WriteToLog("!Specific exception message: ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("!Base Exception: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("!Target Site is: ")
            modLogger.WriteToLog(ex.TargetSite.ToString)
            modLogger.WriteToLog("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            modLogger.WriteToLog("IMPORT TERMINATED - END OF FILE")
            modFilenetHandler.FnLogoff()
            Return False
        End Try


    End Function

    Public Function getInfuferror()
        Return insuferrorcount
    End Function


    Public Function SetProperCategoryFromSubCategory(ByRef SubCategory As String) As Integer
        Dim CategorySql As String
        Dim SubCatId As OleDb.OleDbDataReader
        Try
            CategorySql = "SELECT ELEMENT_ID FROM CATEGORIES_STRUCTURE WHERE ELEMENT_DESC= '" & CStr(SubCategory) & "' AND PARENT_ELEMENT_ID <> 0 and ACTIVE = 1 "

            If OteDB.ExecuteSelect(CategorySql) Then
                SubCatId = OteDB.GetResultSet()
                SubCatId.Read()
                Return CInt(SubCatId.GetValue(0))
            Else
                Console.WriteLine("Could Not query CATEGORIES_STRUCTURE!!")
                Return 0
            End If
        Catch ex As Exception
            Console.WriteLine("Exception occured in SetProperCategoryFromSubCategory of modImport! Specific exception message")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("Exception occured in SetProperCategoryFromSubCategory of modImport! Specific exception message")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return 0
        End Try

    End Function


    Private Function RetrieveMetadataFromEpos(ByRef EposDoc As clsEposDocument) As Boolean
        Try


            EposDoc.Department = modAppConfigSettings.Department
            EposDoc.Receiver = modAppConfigSettings.Receiver
            EposDoc.DocNaturalPosistion = modAppConfigSettings.DocNaturalPosition
            EposDoc.Pages = modAppConfigSettings.Pages
            EposDoc.vType = modAppConfigSettings.vType
            EposDoc.Author = modAppConfigSettings.Author
            EposDoc.View_mode = modAppConfigSettings.View_mode

            Map_SubCategory_Result = SetProperCategoryFromSubCategory(EposDoc.SubCategory)
            'Assign static config-settings values to document object

            If Map_SubCategory_Result = 0 Then
                modLogger.WriteToLog("Could NOT bind an ELEMNT_ID with the Given SubCategory (" & EposDoc.SubCategory & ")")
                Return False
            Else
                EposDoc.SubCategoryInteger = Map_SubCategory_Result
            End If

            Return True


        Catch ex As Exception

            Console.WriteLine("Exception occured in RetrieveMetadataFromEpos of modImport! Specific exception message")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("!Specific exception message ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("!Base Exception ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("!Target Site Is ")
            modLogger.WriteToLog(ex.TargetSite.ToString)
            Return False
        End Try

    End Function



    Private Function CreateFnFoldersForDocument(EposDoc As clsEposDocument) As Boolean

        Dim Folders() As String
        Dim i As Integer
        Dim FolderExistsResponse As Boolean
        Dim folderCountExist As Integer
        Dim folderCountCreate As Integer
        Dim ParentFolderPath As String = ""

        Try
            Folders = EposDoc.GetFoldersArray
            For i = 0 To UBound(Folders)
                FolderExistsResponse = False
                Debug.Print("Folders Is " & Folders(i))
                If Not String.IsNullOrEmpty(Folders(i)) Then
                    If g_DebugMode Then Console.WriteLine("Checking if " & modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i) & " exists")

                    Debug.Print("Checking if " & modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i) & " exists")
                    If Not modFilenetHandler.FnFolderExists(modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i), FolderExistsResponse) Then

                        If Not FolderExistsResponse Then

                            If i = 0 Then
                                ParentFolderPath = modAppConfigSettings.SchemaBaseFolder
                                If g_DebugMode Then Console.WriteLine("Folder " & modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i) & " does Not exist. The application will create it")
                                If Not modFilenetHandler.CreateFnFolder(Folders(i), ParentFolderPath) Then
                                    Console.WriteLine("!!! CreateFnFoldersForDocument Could Not create Filenet folder " & modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i) & " Current document will be skipped.")
                                    Return False
                                    'Exit Function
                                Else
                                    If g_DebugMode Then Console.WriteLine("Folder Successfully created!")
                                    'Return True
                                    folderCountCreate += 1
                                End If

                            ElseIf i > 0 Then
                                ParentFolderPath = modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i - 1)
                                Dim FolderBilling As String = ""
                                FolderBilling = CStr(EposDoc.BILLING_ACCOUNT_ID)

                                If FolderBilling.Contains("-") Then
                                    FolderBilling = Replace(FolderBilling, "-", "") & "_"
                                End If
                                If FolderBilling.Contains(".") Then
                                    FolderBilling = Replace(FolderBilling, ".", "") & "_"
                                End If

                                If g_DebugMode Then Console.WriteLine("Folder " & modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i) & " does Not exist. The application will create it")
                                If Not modFilenetHandler.CreateFnFolder(FolderBilling, CStr(ParentFolderPath)) Then
                                    Console.WriteLine("!!! CreateFnFoldersForDocument Could Not create Filenet folder " & modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i) & " Current document will be skipped.")
                                    Return False
                                    'Exit Function
                                Else
                                    If g_DebugMode Then Console.WriteLine("Folder Successfully created!")
                                    Debug.Print("Folder is: " & Replace(CStr(EposDoc.BILLING_ACCOUNT_ID), "-", "") & "_")
                                    Debug.Print("Parent Dir is: " & ParentFolderPath)
                                    'Return True
                                    folderCountCreate += 1
                                End If
                            End If

                            'If g_DebugMode Then Console.WriteLine("Folder " & modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i) & " does Not exist. The application will create it")
                            'If Not modFilenetHandler.CreateFnFolder(Folders(i), ParentFolderPath) Then
                            '    Console.WriteLine("!!! CreateFnFoldersForDocument Could Not create Filenet folder " & modAppConfigSettings.SchemaBaseFolder & "/" & Folders(i) & " Current document will be skipped.")
                            '    Return False
                            '    'Exit Function
                            'Else
                            '    If g_DebugMode Then Console.WriteLine("Folder Successfully created!")
                            '    'Return True
                            '    folderCountCreate += 1
                            'End If
                        End If
                    Else

                        folderCountExist += 1
                    End If

                End If

            Next

            If NullBilling Then
                ' chech if it creates 2 files or already 2 files exists or create one of them and the other exists
                If (folderCountCreate = 1) Or (folderCountExist = 1) Then
                    Return True
                Else
                    Return False
                End If
            Else

                ' chech if it creates 2 files or already 2 files exists or create one of them and the other exists
                If (folderCountCreate = 2) Or (folderCountExist = 2) Or (folderCountCreate = 1 And folderCountExist = 1) Then
                    'If (folderCountCreate = 1) Or (folderCountExist = 1) Then

                    Return True
                Else
                    Return False
                End If
            End If


        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in CreateFnFoldersForDocument of modImport! Specific exception message: ")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("!!!Exception occured in CreateFnFoldersForDocument of modImport! Specific exception message")
            modLogger.WriteToLog("!Specific exception message ")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("!Base Exception ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog("")
            modLogger.WriteToLog("!Target Site Is ")
            modLogger.WriteToLog(ex.TargetSite.ToString)
            Return False
        End Try

    End Function


    Private Function CommitDocument(EposDoc As clsEposDocument) As Boolean
        'General method - calls to all other commit functions

        Try
            If CommitDocumentToFilenet(EposDoc) Then

                If CommitDocumentToCosmoterev(EposDoc) Then
                    'IF documents is in table Epos_repeat for some reason the application after the import will find it and delete it.
                    If CheckImportAndEposRepeat(EposDoc) Then
                        'We will update the new table FN_EPOS_REL after the successful commit
                        'UpdateFN_EPOS_REL(objEposDoc)
                        Console.WriteLine("Document was successfully inserted in Filenet Databases!")
                        If NotifyEposDB(EposDoc) Then
                            Return True
                        Else
                            modFilenetHandler.DeleteDocumentFromFN(EposDoc.F_Docnumber)
                            Return False
                        End If
                    Else
                        modFilenetHandler.DeleteDocumentFromFN(EposDoc.F_Docnumber)
                        Return False
                    End If



                Else
                    Console.WriteLine("Failed to insert document metadata to Oterev Database.")
                    'OteDB.RollbackTransaction()
                    If modFilenetHandler.DeleteDocumentFromFN(EposDoc.F_Docnumber) Then
                        modLogger.WriteToLog("Document:" & EposDoc.F_Docnumber & " Deleted From Filenet!")
                        If Not DeleteDocumentFromOterev(EposDoc) Then
                            modLogger.WriteToLog("Fail to Delete From Oterev.")
                        Else
                            modLogger.WriteToLog("Document:" & EposDoc.F_Docnumber & " Deleted From Oterev Database!")
                        End If
                    Else
                        modLogger.WriteToLog("Fail to Delete From Filenet.")
                    End If
                    Return False
                End If
            Else
                Return False
            End If

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in CommitDocument of modImport! One of the commital stages failed! Specific exception message")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("!!!Exception occured in CommitDocument of modImport! One of the commital stages failed! Specific exception message")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try

    End Function


    Private Function CommitDocumentToFilenet(ByRef EposDoc As clsEposDocument) As Boolean

        Dim DocumentId As Long
        Dim FoldersArray() As String
        Dim i As Integer

        Try
            FoldersArray = EposDoc.GetFoldersArray()
            If modFilenetHandler.CommitFnDocument(EposDoc.AbsoluteFilePath, DocumentId, EposDoc.FileName) Then
                EposDoc.F_Docnumber = DocumentId
                Console.WriteLine("Document " & EposDoc.FileName & " was commited successfully to Filenet!")

                For i = 0 To UBound(FoldersArray)
                    If Not String.IsNullOrEmpty(FoldersArray(i)) Then

                        If Not modFilenetHandler.FolderizeDocument(EposDoc.F_Docnumber, modAppConfigSettings.SchemaBaseFolder & "/" & FoldersArray(i), EposDoc.FileName) Then
                            Console.WriteLine("Failed to folderize document " & EposDoc.FileName & " with doc id " & EposDoc.F_Docnumber & "in the filenet folder: " & modAppConfigSettings.SchemaBaseFolder & "/" & FoldersArray(i))
                            Return False
                            Exit Function
                        End If
                    End If
                Next
                Return True
            Else

                Console.WriteLine("Failed to commit document " & EposDoc.FileName & " to Filenet")
                Return False
            End If


        Catch ex As Exception
            Console.WriteLine("!!!Exception occured in CommitDocumentToFilenet of modImport! Failed to commit document! Specific exception message")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("!!!Exception occured in CommitDocumentToFilenet of modImport! Failed to commit document! Specific exception message")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try


    End Function

    Private Function DeleteDocumentFromOterev(EposDoc As clsEposDocument) As Boolean

        Try
            Dim DocSql As String
            Dim FolSql As String

            DocSql = modAppConfigSettings.DeleteDocuments
            DocSql = Replace(DocSql, "?F_DOC?", EposDoc.F_Docnumber)

            FolSql = modAppConfigSettings.DeleteFolderContents
            FolSql = Replace(FolSql, "?F_DOC?", EposDoc.F_Docnumber)

            If DocumenstBool Then
                OteDB.ExecuteDelete(DocSql)
            End If

            If Folder_contentsBool Then
                OteDB.ExecuteDelete(FolSql)
            End If

            Return True

        Catch ex As Exception
            modLogger.WriteToLog("!!!Exception occured in DeleteDocumentFromOterev of modImport! Failed to Delete document From DB! Specific exception message")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try
    End Function

    Private Function CommitDocumentToCosmoterev(EposDoc As clsEposDocument) As Boolean
        'Handle DB commital - Insert data into every necessary table
        Dim DocumentsSql As String
        Dim FolderContentSql As String
        Dim SelectFolderSql As String
        Dim InsertFoldersSql As String
        Dim FoldersArray() As String
        Dim FolderNumber As Long
        Dim RecordCount As Long

        DocumenstBool = False
        Folder_contentsBool = False
        FoldersBool = False

        Try

            If modMain.g_DebugMode Then
                Console.WriteLine("TESTING VALUE-PASSING FOR EPOS DOCUMENT OBJECT NAME = " & objEposDoc.FileName & " F_Docnumber = " & objEposDoc.F_Docnumber)
            End If


            'Documents
            DocumentsSql = "" 'Reset string
            DocumentsSql = "INSERT INTO DOCUMENTS (F_DOCNUMBER,F_ENTRYDATE,DOC_DATE,AUTHOR,RECEIVER,TOPIC,TYPE,CATEGORY_ELEMENT_ID,DEPT_ELEMENT_ID,UPDATE_BY_USER,DATE_OF_UPDATE,INVOICE_NO,DOC_NATURAL_POS,PAGES,ORDER_NO,SR_ID)"
            DocumentsSql = DocumentsSql & "VALUES(" & objEposDoc.F_Docnumber & ", sysdate, TO_DATE('" & objEposDoc.DocDate & "','DD/MM/YYYY'),'" & objEposDoc.Author & "',"
            DocumentsSql = DocumentsSql & "'" & objEposDoc.Receiver & "','" & objEposDoc.Topic & "'," & objEposDoc.vType & "," & objEposDoc.SubCategoryInteger & ","
            DocumentsSql = DocumentsSql & objEposDoc.Department & "," & modAppConfigSettings.UpdateUserCode & ", sysdate,"
            DocumentsSql = DocumentsSql & "null,'" & objEposDoc.DocNaturalPosistion & "'," & objEposDoc.Pages & ",'" & objEposDoc.DIGITAL_ORDER_ID & "','N/A')"

            If modMain.g_DebugMode Then
                Console.WriteLine("DocumentsSql query is: " & DocumentsSql)
            End If

            If OteDB.ExecuteInsertUpdate(DocumentsSql) = True Then
                DocumenstBool = True
                FoldersArray = objEposDoc.GetFoldersArray

                Dim tempDBBool As Boolean = NullBilling

                For i = 0 To UBound(FoldersArray)

                    If Not tempDBBool Then
                        tempDBBool = True
                        Continue For
                    End If
                    'FOLDER_CONTENTS
                    modFilenetHandler.GetFnFolderNumberFromName(modAppConfigSettings.SchemaBaseFolder & "/" & FoldersArray(i), FolderNumber)
                    FolderContentSql = ""
                    FolderContentSql = "INSERT INTO FOLDER_CONTENTS (F_FOLDERNUMBER,F_DOCNUMBER,INSERTED_BY_USER_ID,DATE_OF_INSERTION) VALUES ("
                    FolderContentSql = FolderContentSql & FolderNumber & "," & objEposDoc.F_Docnumber & "," & modAppConfigSettings.UpdateUserCode & ",sysdate)"

                    EposDoc.FolderNumber(i) = FolderNumber

                    If OteDB.ExecuteInsertUpdate(FolderContentSql) = False Then
                        Console.WriteLine("Failed when attempting to insert data into FOLDER_CONTENTS - EXITING FUNCTION...")
                        modLogger.WriteToLog("Failed when attempting to insert data into FOLDER_CONTENTS - EXITING FUNCTION...")
                        Return False
                        Exit Function
                    Else
                        Folder_contentsBool = True

                    End If


                    'FOLDERS
                    SelectFolderSql = ""
                    SelectFolderSql = "SELECT COUNT(*) FROM FOLDERS WHERE F_FOLDERNUMBER=" & FolderNumber & " AND CUSTOMER_ID='" & objEposDoc.Customer_Code & "'"
                    SelectFolderSql = SelectFolderSql & " AND BILLING_ACC_ID='" & objEposDoc.BILLING_ACCOUNT_ID & "'"
                    RecordCount = 0
                    RecordCount = OteDB.ExecuteCount(SelectFolderSql)

                    If RecordCount = 0 Then
                        'FOLDERS
                        InsertFoldersSql = ""
                        InsertFoldersSql = "INSERT INTO FOLDERS (F_FOLDERNUMBER,CUSTOMER_ID,BILLING_ACC_ID,PRODUCT_ID,CREATION_DATE,CREATED_BY ) VALUES ("
                        InsertFoldersSql = InsertFoldersSql & FolderNumber & ",'" & objEposDoc.Customer_Code & "','" & objEposDoc.BILLING_ACCOUNT_ID & "','-1',sysdate," & modAppConfigSettings.UpdateUserCode & ")"
                        'InsertFoldersSql = InsertFoldersSql & FolderNumber & ",'" & objEposDoc.Customer_Code & "','-1','-1',sysdate," & modAppConfigSettings.UpdateUserCode & ")"

                        If OteDB.ExecuteInsertUpdate(InsertFoldersSql) = False Then
                            Console.WriteLine("Failed when attempting to insert data into FOLDERS- EXITING FUNCTION...")
                            modLogger.WriteToLog("Failed when attempting to insert data into FOLDERS- EXITING FUNCTION...")
                            Return False
                            Exit Function
                        Else
                            FoldersBool = True
                        End If


                    ElseIf Not RecordCount = 1 And Not RecordCount = -1 Then
                        Console.WriteLine("Wrong number of records in table Folders for Folder Number: " & FolderNumber & " ,Subscriber Id: " & objEposDoc.Customer_Code)
                        modLogger.WriteToLog("Wrong number of records in table Folders for Folder Number: " & FolderNumber & " ,Subscriber Id: " & objEposDoc.Customer_Code)
                        Return False
                        Exit Function

                    End If



                    Debug.Print("At loop execution " & i & " of " & UBound(FoldersArray))

                Next

                Return True
            Else

                Console.WriteLine("Failed while trying to insert data into DOCUMENTS table")
                modLogger.WriteToLog("Failed while trying to insert data into DOCUMENTS table")
                Return False
            End If



        Catch ex As Exception
            Console.WriteLine("!!!Exception occured In CommitDocumentToCosmoterev Of modImport! CANNOT procced With document import! Specific exception message:")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            Console.WriteLine(ex.TargetSite)
            modLogger.WriteToLog("!!!Exception occured In CommitDocumentToCosmoterev Of modImport! CANNOT procced With document import! Specific exception message:")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try

    End Function


    Public Function CheckImportAndEposRepeat(EposDoc As clsEposDocument) As Boolean
        Dim checkSql As String
        Dim FixSql As String
        Try
            'DIGITAL_ORDER_ID, DOCUMENT_FILE_INFO_ID, FILE_NAME, INSERT_DATE, UPDATE_DATE, REPEAT_NUMBER, VIEW_MODE

            checkSql = ""
            checkSql = "SELECT * FROM CMP_DROP_REPEAT WHERE FILE_NAME='" & EposDoc.FileName & "' AND DIGITAL_ORDER_ID = '" & EposDoc.DIGITAL_ORDER_ID & "' AND DOCUMENT_FILE_INFO_ID= '" & EposDoc.DOCUMENT_FILE_INFO_ID & "' AND VIEW_MODE= '" & EposDoc.View_mode & "'"

            If OteDB.ExecuteSelect(checkSql) Then
                FixSql = ""
                FixSql = "DELETE FROM CMP_DROP_REPEAT WHERE FILE_NAME='" & EposDoc.FileName & "' AND DIGITAL_ORDER_ID = '" & EposDoc.DIGITAL_ORDER_ID & "' AND DOCUMENT_FILE_INFO_ID= '" & EposDoc.DOCUMENT_FILE_INFO_ID & "' AND VIEW_MODE= '" & EposDoc.View_mode & "'"
                If OteDB.ExecuteInsertUpdate(FixSql) Then
                    If modMain.g_DebugMode Then
                        Console.WriteLine("Document " & EposDoc.FileName & " deleted from CMP_DROP_REPEAT")
                        modLogger.WriteToLog("Document " & EposDoc.FileName & " deleted from CMP_DROP_REPEAT")
                    End If
                    Return True
                Else
                    Console.WriteLine("!!!Failed to update (delete) the document: " & EposDoc.FileName & " from CMP_DROP_REPEAT table")
                    Return False
                End If
            Else
                If modMain.g_DebugMode Then
                    Console.WriteLine("Document " & EposDoc.FileName & " Did not found in CMP_DROP_REPEAT")
                End If
                Return True
            End If

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured In CheckImportAndEposRepeat Of modImport!")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("!!!Exception occured In CheckImportAndEposRepeat Of modImport!")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try

    End Function


    Public Function SetDocumentToRepeatState(EposDoc As clsEposDocument) As Boolean
        ' Set a document to Repeat state  in the ePOS DB and in our own DB at the same time.
        Dim RepeatSql As String
        Dim RepeatTries As Integer
        Dim NotifySql As String
        Dim UpdateSql As String
        Dim ResultSet As OleDb.OleDbDataReader

        Try

            RepeatSql = ""
            RepeatSql = "SELECT * FROM CMP_DROP_REPEAT WHERE FILE_NAME='" & EposDoc.FileName & "' AND DIGITAL_ORDER_ID = '" & EposDoc.DIGITAL_ORDER_ID & "' AND DOCUMENT_FILE_INFO_ID= '" & EposDoc.DOCUMENT_FILE_INFO_ID & "' AND VIEW_MODE= '" & EposDoc.View_mode & "'"

            If Not OteDB.ExecuteSelect(RepeatSql) Then

                RepeatSql = ""
                RepeatSql = "INSERT INTO CMP_DROP_REPEAT (DIGITAL_ORDER_ID, DOCUMENT_FILE_INFO_ID, FILE_NAME, INSERT_DATE, UPDATE_DATE, REPEAT_NUMBER, VIEW_MODE) VALUES"
                RepeatSql = RepeatSql & "('" & EposDoc.DIGITAL_ORDER_ID & "','" & EposDoc.DOCUMENT_FILE_INFO_ID & "','" & EposDoc.FileName & "',SYSDATE,SYSDATE,0,'" & EposDoc.View_mode & "')"
                If OteDB.ExecuteInsertUpdate(RepeatSql) Then
                    If g_DebugMode Then
                        Console.WriteLine("Successfully inserted: " & EposDoc.FileName & " to CMP_DROP_REPEAT! - Notifying Filenet Database now...")
                    End If
                    'update CMP_IMPORT_REQUESTS_OTEVIEW set import_status=?STATE? where digital_order_id='?ID?' and document_file_info_id='?CMP_ID?'
                    NotifySql = ""
                    NotifySql = Replace(modAppConfigSettings.EposUpdateSql, "?ID?", EposDoc.DIGITAL_ORDER_ID)
                    NotifySql = Replace(NotifySql, "?CMP_ID?", EposDoc.DOCUMENT_FILE_INFO_ID)
                    NotifySql = Replace(NotifySql, "?F_DOCNUMBER?", "")
                    NotifySql = Replace(NotifySql, "?STATE?", 3) '  state 3 is translated to repeat state

                    If OteDB.ExecuteInsertUpdate(NotifySql) Then
                        Console.WriteLine("Successfully notified Filenet Database !")
                    Else
                        modLogger.WriteToLog("Fail to notify Filenet Database while trying to set " & EposDoc.FileName & " to repeat state !")
                    End If

                    modStatistics.IncreaseRepeatCounter()
                    Return True

                Else
                    Console.WriteLine("!!!Failed to update Document state (Repeat) in both Databases")
                    modLogger.WriteToLog("!!!Failed to update Document state (Repeat) in both Databases")
                    Return False
                End If

            Else
                ResultSet = OteDB.GetResultSet
                ResultSet.Read()
                If g_DebugMode Then
                    Console.WriteLine("The File " & EposDoc.FileName & " already exists in CMP_DROP_REPEAT TABLE")
                End If
                Console.WriteLine("The File " & EposDoc.FileName & " already in Repeat State!")
                RepeatTries = ResultSet.GetValue(5)


                If RepeatTries > modAppConfigSettings.RepeatNum Then
                    UpdateSql = ""
                    UpdateSql = "DELETE FROM CMP_DROP_REPEAT WHERE FILE_NAME='" & EposDoc.FileName & "' AND DIGITAL_ORDER_ID = '" & EposDoc.DIGITAL_ORDER_ID & "' AND DOCUMENT_FILE_INFO_ID= '" & EposDoc.DOCUMENT_FILE_INFO_ID & "' AND VIEW_MODE= '" & EposDoc.View_mode & "'"
                    If OteDB.ExecuteInsertUpdate(UpdateSql) Then
                        Console.WriteLine("Setting document " & EposDoc.FileName & " to ERROR state due to too many unsuccessful tries to commit it")
                        modLogger.WriteToLog("Setting document " & EposDoc.FileName & " to ERROR state due to too many unsuccessful tries to commit it")
                        SetDocumentToErrorState(EposDoc)
                    Else
                        Console.WriteLine("!!!Failed to update (delete) the document: " & EposDoc.FileName & " in the CMP_DROP_REPEAT table")
                        Return False
                    End If
                Else

                    'the following code is to ensure that if the EposDoc.FileName already exists in the Epos_repeat
                    'it also has updated  Filenet = 3 in the CMP_IMPORT_REQUESTS_OTEVIEW files table
                    If objEposDoc.FileState = 0 Then
                        NotifySql = ""
                        NotifySql = Replace(modAppConfigSettings.EposUpdateSql, "?ID?", EposDoc.DIGITAL_ORDER_ID)
                        NotifySql = Replace(NotifySql, "?CMP_ID?", EposDoc.DOCUMENT_FILE_INFO_ID)
                        NotifySql = Replace(NotifySql, "?F_DOCNUMBER?", "")
                        NotifySql = Replace(NotifySql, "?STATE?", 3) '  state 3 is translated to repeat state

                        If OteDB.ExecuteInsertUpdate(NotifySql) Then
                            Debug.Print("The " & EposDoc.FileName & " is in CMP_DROP_REPEAT but CMP_IMPORT_REQUESTS_OTEVIEW table has wrong Filenet Value! Its Updated to 3")
                            Console.WriteLine("The " & EposDoc.FileName & " is in CMP_DROP_REPEAT but CMP_IMPORT_REQUESTS_OTEVIEW table has wrong Filenet Value! Its Updated to 3")
                        Else
                            modLogger.WriteToLog("Fail to notify Filenet Database while trying to set " & EposDoc.FileName & " to repeat state !")
                        End If
                    End If

                    modStatistics.IncreaseRepeatCounter()
                    UpdateSql = ""
                    UpdateSql = "UPDATE CMP_DROP_REPEAT SET UPDATE_DATE=SYSDATE,REPEAT_NUMBER =" & RepeatTries + 1 & " WHERE FILE_NAME='" & EposDoc.FileName & "' AND DIGITAL_ORDER_ID = '" & EposDoc.DIGITAL_ORDER_ID & "' AND DOCUMENT_FILE_INFO_ID= '" & EposDoc.DOCUMENT_FILE_INFO_ID & "' AND VIEW_MODE= '" & EposDoc.View_mode & "'"
                    If OteDB.ExecuteInsertUpdate(UpdateSql) Then
                        If g_DebugMode Then
                            Console.WriteLine("REPEAT_NUMBER for document " & EposDoc.FileName & " has been successfully incremented to: " & RepeatTries + 1)
                        End If
                        Return True
                    Else
                        Console.WriteLine("!!!Failed to update CMP_DROP_REPEAT while trying to increment the repeat counter")
                        Return False
                    End If

                End If

            End If

        Catch ex As Exception
            Console.WriteLine("!!!Exception occured In SetDocumentToRepeatState Of modImport! Could not set document to repeat state! Specific exception message:")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException)
            modLogger.WriteToLog("!!!Exception occured In SetDocumentToRepeatState Of modImport! Could not set document to repeat state! Specific exception message:")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try

    End Function



    Private Function SetDocumentToErrorState(EposDoc As clsEposDocument) As Boolean
        Dim NotifySql As String
        Try
            NotifySql = Replace(modAppConfigSettings.EposUpdateSql, "?ID?", EposDoc.DIGITAL_ORDER_ID)
            NotifySql = Replace(NotifySql, "?CMP_ID?", EposDoc.DOCUMENT_FILE_INFO_ID)
            NotifySql = Replace(NotifySql, "?F_DOCNUMBER?", "")
            NotifySql = Replace(NotifySql, "?STATE?", 4) ' state 4 is translated to error state

            If OteDB.ExecuteInsertUpdate(NotifySql) Then
                Console.WriteLine("Notified the Database  of the UN-successful document commital! Document is now on error state")
                modStatistics.IncreaseErrorCounter()
                Return True
            Else
                Console.WriteLine("!!!Could not execute the Database Notify query ")
                Return False
            End If
        Catch ex As Exception
            Console.WriteLine("Not valid document data to set the document to error state!!!")
            modLogger.WriteToLog("Not valid document data to set the document to error state!!!")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try
    End Function

    Private Function SetDocumentToErrorStateDueNonData(digital_order_id As String, document_file_info_id As Integer) As Boolean
        Dim NotifySql As String
        Try
            NotifySql = Replace(modAppConfigSettings.EposUpdateSql, "?ID?", digital_order_id)
            NotifySql = Replace(NotifySql, "?CMP_ID?", document_file_info_id)
            NotifySql = Replace(NotifySql, "?F_DOCNUMBER?", "")
            NotifySql = Replace(NotifySql, "?STATE?", 4) ' state 4 is translated to error state

            If OteDB.ExecuteInsertUpdate(NotifySql) Then
                If modMain.g_DebugMode Then
                    Console.WriteLine("Notified the Database of the UN-successful document commital! Document is now on error state")
                End If
                modLogger.WriteToINSUFFICIENTDATALog("The current document is in ERROR state due to the lack of information!!!")
                modLogger.WriteToINSUFFICIENTDATALog("")
                modStatistics.IncreaseErrorCounter()
                Return True
            Else
                Console.WriteLine("!!!Could not execute the Database Notify query ")
                Return False
            End If
        Catch ex As Exception
            Console.WriteLine("Not valid document data to set the document to error state!!!")
            modLogger.WriteToLog("Not valid document data to set the document to error state!!!")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try
    End Function

    Private Function NotifyEposDB(EposDoc As clsEposDocument) As Boolean
        Dim NotifySql As String
        Try
            NotifySql = Replace(modAppConfigSettings.EposUpdateSql, "?ID?", EposDoc.DIGITAL_ORDER_ID)
            NotifySql = Replace(NotifySql, "?CMP_ID?", EposDoc.DOCUMENT_FILE_INFO_ID)
            NotifySql = Replace(NotifySql, "?F_DOCNUMBER?", ", F_DOCNUMBER = " & EposDoc.F_Docnumber)
            NotifySql = Replace(NotifySql, "?STATE?", 1) ' state 1 is translated to successfully commited

            If OteDB.ExecuteInsertUpdate(NotifySql) Then
                Console.WriteLine("Notified the Database  of the successful document commital!")
                Return True
            Else
                Console.WriteLine("!!!Could not execute the Database Notify query ")
                Return False
            End If

        Catch ex As Exception
            modLogger.WriteToLog("!!!Exception occured In NotifyEposDB Of modImport! Could not set document to success state! Specific exception message:")
            modLogger.WriteToLog(ex.Message)
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            Return False
        End Try
    End Function

    Private Function TimeIsUp(StopTime As String) As Boolean
        Try
            Dim dtStopTime As Date
            StopTime = StopTime & ":00"
            Debug.Print(StopTime)
            Debug.Print(DateTime.Now.ToString("HH:mm:ss"))
            If DateTime.Now.ToString("HH:mm:ss") > StopTime Then
                Debug.Print("The time is up, must stop!!!!!")
                Return True
            Else
                Debug.Print("Time is not up....")
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function



    Public Function getCounters() As Array
        Dim counterArray(2) As Integer
        'counterArray = New Array(3)
        counterArray(0) = modStatistics.getSuccessful
        counterArray(1) = modStatistics.getRepeat
        counterArray(2) = modStatistics.getError
        Return counterArray
    End Function


End Module

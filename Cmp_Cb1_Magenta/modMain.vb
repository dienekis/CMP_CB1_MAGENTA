Imports System.Diagnostics

Module modMain
    Friend global_NOT_EXISTS_COUNTER As Long
    Friend global_NO_EPOS_METADATA_COUNTER As Long
    Friend global_ZERO_SIZE_COUNTER As Long
    Friend g_DebugMode As Boolean
    Friend g_WaitMode As Boolean
    Friend g_AbnormalLoopExit As Boolean

    Sub Main()
        Console.Title = "B2C_RENEWALS DOCUMENT IMPORT FLOW"
        Dim p() As Process
        Dim CommandArgs() As String = Environment.GetCommandLineArgs

        modAppConfigSettings.ReadApplicationSettings()
        modLogger.generateProcessLog()
        modStatistics.ResetTimers()
        modStatistics.SetStartTime()
        modLogger.WriteToLog("Loading essential components...")
        modLogger.WriteToLog("Parsing command line arguments...")
        Console.WriteLine("Loading essential components...")
        Console.WriteLine("Parsing command line arguments...")

        'Checks if the application is already running. If so the app terminates otherwise it continues.
        Dim appName As String = System.IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location) ' get the name of the application from the filepath
        appName = appName.Replace(".exe", "") 'remove the .exe extension from file name
        'Console.WriteLine("AppName is: {0}", appName) 'prints the file name of the app
        p = Process.GetProcessesByName(appName)
        If p.Count > 1 Then
            'Console.WriteLine("ALREADY RUNS")
            Console.WriteLine("There is another instance of the application that is currently running.")
            Console.WriteLine("Therefore this one will terminate!")
            modLogger.WriteToLog("There is another instance of the application that is currently running.")
            modLogger.WriteToLog("Therefore this one will terminate!")
            modLogger.WriteToLog("----------------------------------------------------------------------------------------------------")
            modLogger.WriteToLog("END OF FILE")
            Exit Sub
        Else
            'Console.WriteLine("FIRST RUNS")
            Console.WriteLine("This is the only instance of the application that currently runs.")
            modLogger.WriteToLog("This is the only instance of the application that currently runs.")
        End If

        g_DebugMode = False
        g_WaitMode = False
        If CommandArgs.Length >= 2 Then
            If (CommandArgs.GetValue(1).ToString.ToLower = "debug") Then
                Console.WriteLine("The application detected a DEBUG command line argument. Initiating DEBUG Mode.")
                modLogger.WriteToLog("The application detected a DEBUG command line argument. Initiating DEBUG Mode.")
                g_DebugMode = True
            ElseIf (CommandArgs.GetValue(1).ToString.ToLower = "wait") Then
                Console.WriteLine("The application detected a WAIT command line argument. The console will remain open until you press a key.")
                modLogger.WriteToLog("The application detected a WAIT command line argument. The console will remain open until you press a key.")
                g_WaitMode = True
            Else
                Console.WriteLine("The application detected an argument different from DEBUG. DEBUG Mode will not be initiated")
                modLogger.WriteToLog("The application detected an argument different from DEBUG. DEBUG Mode will not be initiated")
            End If
        Else
            Console.WriteLine("The application detected no command line arguments. Resuming normal execution...")
            modLogger.WriteToLog("The application detected no command line arguments. Resuming normal execution...")
        End If
        ''
        g_AbnormalLoopExit = False
        'g_DebugMode = True 'For development set to true by default, remove after
        'g_WaitMode = True
        Debug.Print("Entered Main")
        Console.WriteLine("Entered Main")
        global_NOT_EXISTS_COUNTER = 0
        global_NO_EPOS_METADATA_COUNTER = 0

        ' Connection to FileNet has been moved under the fetched documents due to session time out
        ' modImport line:192

        modImport.ImportBatch()
        modLogger.WriteToLog("")
        Console.WriteLine("")

        If modFilenetHandler.FnLogoff() Then

        Else
            Console.WriteLine("Unsuccessful Filenet Lofoff")
            modLogger.WriteToLog("Unsuccessful Filenet Lofoff")
        End If

        modLogger.WriteToLog("")
        modStatistics.SetFinishTime()
        modLogger.FinalizeLog(modStatistics.getSuccessful, modStatistics.getRepeat, modStatistics.getError, modStatistics.GetRunTime, modStatistics.getzerosize, modStatistics.getDontexist, modStatistics.getInsuffisient, modStatistics.getTotalDocs, modImport.getInfuferror, modStatistics.getMetadata, modStatistics.getUnError)
        'Uncomment below if u want to keep the console from terminating
        'Console.WriteLine("Press any key to continue...")
        'Console.ReadLine()

        If (modMain.g_WaitMode) Then
            Console.WriteLine("Press any key to continue...")
            Console.ReadLine()
        End If
        '

    End Sub

End Module

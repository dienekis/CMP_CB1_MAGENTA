Imports System.Configuration
Imports System.Xml

Module modAppConfigSettings
    'Module for reading all settings stored in App.config
    Public LogToDB As Integer
    Public DbName As String
    Public DbUser As String
    Public DbPwd As String
    Public FnUser As String
    Public FnPwd As String
    Public Receiver As String
    Public Author As String
    Public vType As Integer
    Public Department As Integer
    Public Pages As Integer
    Public UpdateUserCode As Integer
    Public ClassName As String
    Public Path As String
    Public DocNaturalPosition As String
    Public SleepInterval As Integer
    Public StopTime As String
    Public ThreadId As Long
    Public RepeatNum As Integer
    Public SchemaBaseFolder As String
    Public OterevDBname As String
    Public OterevDBuser As String
    Public OterevDBpwd As String
    Public EposUpdateSql As String
    Public EposContractsSql As String
    Public EposFilesSql As String
    Public DocumentsImported As Long
    Public DocumentsFailed As Long
    Public RunPath As String
    Public Category As Long
    Public ErrorPath As String
    Public RepeatPath As String
    Public RowNum As String
    Public RowNumEnable As Integer
    Public View_mode As String
    Public DeleteDocuments As String
    Public DeleteFolderContents As String



    Sub New()
        'Reseting all properties to default (empty) values
        LogToDB = 1
        DbName = ""
        DbUser = ""
        DbPwd = ""
        FnUser = ""
        FnPwd = ""
        Receiver = ""
        vType = 0
        Department = 0
        UpdateUserCode = 0
        ClassName = ""
        Path = ""
        DocNaturalPosition = ""
        SleepInterval = 5000
        StopTime = ""
        ThreadId = 0
        RepeatNum = 0
        SchemaBaseFolder = ""
        OterevDBname = ""
        OterevDBuser = ""
        OterevDBpwd = ""
        EposUpdateSql = ""
        EposContractsSql = ""
        EposFilesSql = ""
        DocumentsImported = 0
        DocumentsFailed = 0
        RunPath = ""
        Category = 0
        RepeatPath = ""
        ErrorPath = ""
        DeleteDocuments = ""
        DeleteFolderContents = ""

    End Sub

    Public Function ReadApplicationSettings() As Boolean

        Try
            'Read all settings stored in the App.config
            LogToDB = ConfigurationSettings.AppSettings.Item("LogToDB")
            DbName = ConfigurationSettings.AppSettings.Item("DBname")
            DbUser = ConfigurationSettings.AppSettings.Item("DBuser")
            DbPwd = ConfigurationSettings.AppSettings.Item("DBpwd")
            FnUser = ConfigurationSettings.AppSettings.Item("FNuser")
            FnPwd = ConfigurationSettings.AppSettings.Item("FNpwd")
            SchemaBaseFolder = ConfigurationSettings.AppSettings.Item("SchemaBaseFolder")
            ClassName = ConfigurationSettings.AppSettings.Item("ClassName")
            Receiver = ConfigurationSettings.AppSettings.Item("Receiver")
            vType = ConfigurationSettings.AppSettings.Item("Type")
            Department = ConfigurationSettings.AppSettings.Item("Department")
            Category = ConfigurationSettings.AppSettings.Item("Category")
            UpdateUserCode = ConfigurationSettings.AppSettings.Item("UpdateUserCode")
            DocNaturalPosition = ConfigurationSettings.AppSettings.Item("DocNaturalPosition")
            Path = ConfigurationSettings.AppSettings.Item("Path")
            SleepInterval = ConfigurationSettings.AppSettings.Item("SleepInterval")
            StopTime = ConfigurationSettings.AppSettings.Item("StopTime")
            RepeatNum = ConfigurationSettings.AppSettings.Item("RepeatNumber")
            ThreadId = ConfigurationSettings.AppSettings.Item("ThreadID")
            RowNumEnable = ConfigurationSettings.AppSettings.Item("RowNumEnable")
            OterevDBname = ConfigurationSettings.AppSettings.Item("OterevDBname")
            OterevDBuser = ConfigurationSettings.AppSettings.Item("OterevDBuser")
            OterevDBpwd = ConfigurationSettings.AppSettings.Item("OterevDBpwd")
            EposUpdateSql = ConfigurationSettings.AppSettings.Item("SQLUpdate")
            EposFilesSql = ConfigurationSettings.AppSettings.Item("SQLFilesToImport")
            RowNum = ConfigurationSettings.AppSettings.Item("RowNumSQl")
            Pages = ConfigurationSettings.AppSettings.Item("Pages")
            Author = ConfigurationSettings.AppSettings.Item("Author")
            View_mode = ConfigurationSettings.AppSettings.Item("View_mode")
            DeleteDocuments = ConfigurationSettings.AppSettings.Item("DeleteDocuments")
            DeleteFolderContents = ConfigurationSettings.AppSettings.Item("DeleteFolderContents")

            'Make the necessary changes to the default values
            'ClassName = ClassName.Replace("%YYYY%", Date.Now.Year) ' Not Needed for Ote Documents Since the Class is only one and it isn't increasing for every year like Faxes.
            ThreadId += 1
            UpdateAppSettings("ThreadID", ThreadId)

            Return True
        Catch ex As Exception
            'Print exception to log file in case of exception
            Console.WriteLine("Failed to read application settings from App.config file!")
            Console.WriteLine(ex.GetBaseException().ToString)
            Console.WriteLine("")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
            Console.WriteLine("Cannot Procced with execution. Exiting....")
            Return False
        End Try

    End Function



    Public Sub UpdateAppSettings(ByVal KeyName As String, ByVal KeyValue As String)
        'Method used for parsing the App.config file and incrementing the ThreadID value on each execution of the program (Careful with the App.config location)
        Dim XmlDoc As New XmlDocument()
        Dim TestPath As String
        Dim appName As String = System.IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Try
            TestPath = AppDomain.CurrentDomain.BaseDirectory & appName & ".config"
            'TestPath = AppDomain.CurrentDomain.BaseDirectory
            'TestPath = TestPath.Replace("bin\Debug\", "App.config") 'Change to bin\Release\ or to target directory
            XmlDoc.Load(TestPath)

            Dim xmlNodetst As XmlNode
            xmlNodetst = XmlDoc.DocumentElement.SelectSingleNode("appSettings/add[@key=""" & KeyName & """]") 'Target the ThreadID node in the config file

            xmlNodetst.Attributes(1).Value = KeyValue 'Set new value

            XmlDoc.Save(TestPath) 'Save the changes

        Catch ex As Exception
            Console.WriteLine("!Exception occured while trying to parse the App.config xml file so as to set the current Thread ID")
            Console.WriteLine(ex.GetBaseException())
            Console.WriteLine("")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("")
        End Try
    End Sub

End Module


Public Class clsEposDocument
    'Each document to be imported will be treated as an object with its own properties
    Public FileName As String
    Public FilePath As String
    Public AbsoluteFilePath As String
    Public Extension As String
    Public F_Docnumber As Long
    Public DocDate As String
    Public Author As String
    Public Receiver As String
    Public Topic As String
    Public vType As Integer
    Public Category As String
    Public SubCategory As String
    Public SubCategoryInteger As Integer
    Public NetApp As Integer
    Public Department As Integer
    Public DocNaturalPosistion As String
    Public Pages As Integer
    Public DIGITAL_ORDER_ID As String
    Public DOCUMENT_FILE_INFO_ID As String
    Public BILLING_ACCOUNT_ID As String
    Public Customer_Code As String
    Public Shop_Code As String
    Public View_mode As String

    Public FileState As Integer
    'Public ORDER_NUM As String
    Public FolderNumber() As Long


    Private vFolders() As String

    Public Function FillObjectFolders(Folders As Object) As Boolean
        'USE this Function to store BSCS contract id and BSCS customer code AS Folder names!!!! -ADJUST FOR INPUT QUERY
        Try

            ReDim vFolders(UBound(Folders))
            vFolders = Folders
            Debug.Print(UBound(Folders))

            Return True
        Catch ex As Exception
            Console.WriteLine("[E]!!Exception occured in FillObjectFolders of clsEposDocument! Specific exception: ")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.GetBaseException.ToString)
            Console.WriteLine(ex.StackTrace)
            modLogger.WriteToLog("[E]!!Exception occured in FillObjectFolders of clsEposDocument! Specific exception: ")
            modLogger.WriteToLog(ex.GetBaseException.ToString)
            modLogger.WriteToLog(ex.StackTrace)
            Return False
        End Try
    End Function


    Public Function GetFoldersArray() As String()
        GetFoldersArray = vFolders
    End Function


End Class

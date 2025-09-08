Module modStatistics
    'Class used for documenting statistics on a per-execution basis
    Public SuccessfulCounter As Long
    Public RepeatCounter As Long
    Public ErrorCounter As Long
    Public StartTime As DateTime
    Public FinishTime As DateTime
    Public RunTime As TimeSpan
    Public InDontExist As Long
    Public zerosize As Long
    Public Insufficient As Long
    Public TotalDocProc As Long
    Public metadata As Long
    Public UnError As Long

    Public Sub ResetCounters() 'Initialize with all values=0
        SuccessfulCounter = 0
        RepeatCounter = 0
        ErrorCounter = 0
        InDontExist = 0
        Insufficient = 0
        zerosize = 0
        TotalDocProc = 0
        metadata = 0
        UnError = 0
    End Sub

    Public Sub increaseUnError()
        UnError += 1
    End Sub

    Public Sub increaseMetadata()
        metadata += 1
    End Sub

    Public Sub increasezerosize()
        zerosize += 1

    End Sub

    Public Sub IncreaseDontexistCounter()
        InDontExist += 1

    End Sub

    Public Sub IncreaseInsufficient()
        Insufficient += 1

    End Sub

    Public Sub IncreaseSuccessfulCounter()
        SuccessfulCounter += 1

    End Sub

    Public Sub IncreaseRepeatCounter()
        RepeatCounter += 1

    End Sub

    Public Sub IncreaseErrorCounter()
        ErrorCounter += 1

    End Sub

    Public Function getUnError()
        Return UnError
    End Function

    Public Function getMetadata()
        Return metadata
    End Function

    Public Function getSuccessful()
        Return SuccessfulCounter

    End Function

    Public Function getRepeat()
        Return RepeatCounter

    End Function

    Public Function getError()
        Return ErrorCounter

    End Function

    Public Function getDontexist()
        Return InDontExist

    End Function

    Public Function getInsuffisient()
        Return Insufficient

    End Function

    Public Function getzerosize()
        Return zerosize

    End Function

    Public Function setTotalDocs(docs As Long)
        TotalDocProc = docs

    End Function

    Public Function getTotalDocs()
        Return TotalDocProc
    End Function

    Public Sub ResetTimers()
        StartTime = New DateTime
        FinishTime = New DateTime
        RunTime = New TimeSpan

    End Sub

    Public Sub SetStartTime()
        StartTime = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")

    End Sub

    Public Sub SetFinishTime()
        FinishTime = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")

    End Sub

    Public Function GetRunTime()
        RunTime = FinishTime - StartTime
        GetRunTime = RunTime
    End Function

End Module

Imports System.IO
Imports System.ServiceProcess
Imports System.Threading
Public Class WindowsService
    Private Schedular As Timer
    Protected Overrides Sub OnStart(ByVal args() As String)
        WriteToFile("Windows Service started at " + Date.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        ScheduleService()
    End Sub
    Private Sub ScheduleService()
        Try
            Schedular = New Timer(New TimerCallback(AddressOf SchedularCallback))
            Dim intervalMode As Boolean = My.Settings.IntervalMode
            Dim dailyMode As Boolean = My.Settings.DailyMode
            Dim scheduledTime As Date = Date.MinValue
            If dailyMode Then
                WriteToFile(Convert.ToString("Windows Service Mode: Daily.") + " {0}")
                scheduledTime = Date.Parse(My.Settings.ScheduledTime)
                If Date.Now > scheduledTime Then
                    scheduledTime = scheduledTime.AddDays(1)
                End If
            End If
            If intervalMode Then
                WriteToFile(Convert.ToString("Windows Service Mode: Interval in minutes.") + " {0}")
                Dim intervalMinutes As Integer = Convert.ToInt32(My.Settings.IntervalMinutes)
                scheduledTime = Date.Now.AddMinutes(intervalMinutes)
                If Date.Now > scheduledTime Then
                    scheduledTime = scheduledTime.AddMinutes(intervalMinutes)
                End If
            End If
            Dim timeSpan As TimeSpan = scheduledTime.Subtract(Date.Now)
            Dim schedule As String = String.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds)
            WriteToFile((Convert.ToString("Windows Service scheduled to run after: ") & schedule) + " {0}")
            Dim dueTime As Integer = Convert.ToInt32(timeSpan.TotalMilliseconds)
            Schedular.Change(dueTime, Timeout.Infinite)
        Catch ex As Exception
            WriteToFile("Windows Service Error on: {0} " + ex.Message + ex.StackTrace)
            Using serviceController As New ServiceController("MultiRouteQuickbooksService")
                serviceController.[Stop]()
            End Using
        End Try
    End Sub
    Protected Overrides Sub OnStop()
        WriteToFile("Windows Service stopped at " + Date.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Schedular.Dispose()
    End Sub
    Private Sub SchedularCallback(e As Object)
        WriteToFile("Windows Service Log: " + Date.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        ScheduleService()
    End Sub
    Sub WriteToFile(text As String)
        If (Not Directory.Exists(My.Settings.LogFilePath)) Then
            Directory.CreateDirectory(My.Settings.LogFilePath)
        End If
        Dim path As String = My.Settings.LogFilePath & "\Log.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, Date.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub
End Class

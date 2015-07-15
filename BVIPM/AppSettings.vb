Public Class ApplicationAdministrator
    Public Event ExitApplication()

    Public Enum ReportTypeEnum
        Information = 1
        [Error] = 2
        ErrorNoShutdown = 3
        MessageBoxAndShutdown = 4
    End Enum

    Public Sub Report(ByVal ReportType As ReportTypeEnum, ByVal Message As String, Optional ByVal Caption As String = Nothing)
        Dim Coll As New Collection
        Dim SendEmailResults As Results
        Dim MessageHeader As String

        Select Case ReportType
            Case ReportTypeEnum.MessageBoxAndShutdown
                MessageBox.Show(Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                RaiseEvent ExitApplication()


            Case ReportTypeEnum.Information
                WriteToLogFile(Message)
            Case ReportTypeEnum.Error, ReportTypeEnum.ErrorNoShutdown
                WriteToLogFile(Message)

                ' ___ Screen capture
                Dim FileFullPath As String
                FileFullPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & "ScreenCapture_" & System.Environment.UserName & "_" & Date.Now.ToUniversalTime.AddHours(-5).ToString("yyyyMMddssff") & ".gif"
                Common.CaptureScreen(FileFullPath)

                ' ___ Send email to help desk, including log and screenshot
                Coll.Add(cEnviro.LogFileFullPath)
                Coll.Add(FileFullPath)

                '   Common.SendEmail("HelpDesk@benefitvision.com", System.Environment.UserName & "@benefitvision.com", "rbluestein@benefitvision.com", "TEST- DISREGARD - Program Manager error", "An error occurred in the execution of the Program Manager. Attached please find a copy of the application log and a screen shot.", Coll)
                SendEmailResults = Common.SendEmail("HelpDesk@benefitvision.com;rbluestein@benefitvision.com", System.Environment.UserName & "@benefitvision.com", "", "BVIPM.exe Program Manager error - " & System.Environment.UserName & "/" & System.Environment.MachineName & "  " & Date.Now.ToUniversalTime.AddHours(-5).ToString, "An error occurred in the execution of the Program Manager launcher. Attached please find a copy of the application log and a screen shot.", Coll)

                MessageHeader = "An error has occurred requiring the Program Manager to shut down. You may view the details of the problem in the log. "
                If SendEmailResults.Success Then
                    MessageHeader &= "The Program Manager has emailed a copy of the log to the help desk." & vbCrLf & vbCrLf
                Else
                    MessageHeader &= "The Program Manager was unable to email a copy of the log to the help desk." & vbCrLf & vbCrLf
                End If

                Dim FormMessage As New FrmMessage
                FormMessage.ShowDialog(MessageHeader & Message)

                If System.IO.File.Exists(FileFullPath) Then
                    System.IO.File.Delete(FileFullPath)
                End If

                If ReportType = ReportTypeEnum.Error Then
                    RaiseEvent ExitApplication()
                End If
        End Select
    End Sub

    Private Function ReadLogFile() As String
        Dim StreamReader As System.IO.StreamReader
        Dim FileText As String

        Try
            StreamReader = New System.IO.StreamReader(cEnviro.LogFileFullPath)
            FileText = StreamReader.ReadToEnd

            'Do While StreamReader.Peek() >= 0
            '    'Console.WriteLine(StreamReader.ReadLine())

            '    x = StreamReader.ReadLine()
            'Loop
            StreamReader.Close()
            Return FileText

        Catch
        End Try
    End Function

    Private Sub WriteToLogFile(ByVal Message As String)
        'Dim FileInfo As System.IO.FileInfo
        Dim StreamWriter As System.IO.StreamWriter

        Try
            'FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)
            StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
            'StreamWriter.Write("[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] " & Message & vbCrLf)
            StreamWriter.Write(Common.GetTimeStamp & Message & vbCrLf)
            StreamWriter.Close()
        Catch
            Try
                StreamWriter.Close()
            Catch
            End Try
        End Try
    End Sub
End Class

#Region " Results class "
Public Class Results
    Private cSuccess As Boolean
    Private cMessage As String
    Private cUpdateRequired As Boolean
    Private cReturnStr As String

    Public Property UpdateRequired() As Boolean
        Get
            Return cUpdateRequired
        End Get
        Set(ByVal Value As Boolean)
            cUpdateRequired = Value
        End Set
    End Property
    Public Property Success() As Boolean
        Get
            Return cSuccess
        End Get
        Set(ByVal Value As Boolean)
            cSuccess = Value
        End Set
    End Property
    Public Property Message() As String
        Get
            Return cMessage
        End Get
        Set(ByVal Value As String)
            cMessage = Value
        End Set
    End Property
    Public Property ReturnStr() As String
        Get
            Return cReturnStr
        End Get
        Set(ByVal Value As String)
            cReturnStr = Value
        End Set
    End Property

End Class
Public Class ResultsAppUpdate
    Inherits Results

    Private cShutdownApp As Boolean

    Public Property ShutdownApp() As Boolean
        Get
            Return cShutdownApp
        End Get
        Set(ByVal Value As Boolean)
            cShutdownApp = Value
        End Set
    End Property
End Class
#End Region

Public Class Common
    Public Shared Function GetTimeStamp() As String
        Return "[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] "
    End Function

    Public Shared Function ProcessFound(ByVal ProcessName As String) As Boolean
        Dim i As Integer

        Try
            Dim CurrentProcesses As Process() = Process.GetProcesses()
            For i = 0 To CurrentProcesses.GetUpperBound(0)
                Try
                    If CurrentProcesses(i).MainModule.FileName.ToLower = ProcessName.ToLower Then
                        Return True
                    End If
                Catch
                End Try
            Next
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Sub DisallowRunningMultipleInstances()
        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MessageBox.Show("Another instance of the program manager launcher is already running", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Application.Exit()
        End If
    End Sub

    Public Shared Sub CaptureScreen(ByVal FileFullPath As String)
        Dim SC As New ScreenCapture
        System.Threading.Thread.Sleep(1000)
        SC.CaptureScreenToFile(FileFullPath, System.Drawing.Imaging.ImageFormat.Gif)
        SC = Nothing
    End Sub

    Public Shared Function SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String, ByRef AttachmentColl As Collection) As Results
        Dim schema As String
        Dim MyResults As New Results
        Dim i As Integer

        Try

            Dim CDOConfig As New CDO.Configuration
            schema = "http://schemas.microsoft.com/cdo/configuration/"

            'CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2
            ' CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
            'CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "mail.benefitvision.com"
            'CDOConfig.Fields.Update()


            Dim iMsg As New CDO.Message
            iMsg.To = SendTo
            iMsg.From = From
            iMsg.CC = cc
            iMsg.Subject = Subject

            For i = 1 To AttachmentColl.Count
                iMsg.AddAttachment(AttachmentColl(i))
            Next

            iMsg.Configuration = CDOConfig
            iMsg.TextBody = TextBody

            iMsg.Send()

            ' ___ Clean up
            CDOConfig = Nothing
            iMsg = Nothing

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #2104: " & ex.Message
            Return MyResults
        End Try
    End Function


    'Public Sub GetAssemblyInfo()

    '    ' NOTE: REQUIRES
    '    '   Imports System.Reflection

    '    Dim CopyRight As String
    '    Dim Product As String
    '    Dim Description As String

    '    CopyRight = CType(AssemblyCopyrightAttribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly, GetType(AssemblyCopyrightAttribute)), AssemblyCopyrightAttribute).Copyright
    '    Product = CType(AssemblyProductAttribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly, GetType(AssemblyProductAttribute)), AssemblyProductAttribute).Product
    '    Description = CType(AssemblyDescriptionAttribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly, GetType(AssemblyDescriptionAttribute)), AssemblyDescriptionAttribute).Description

    '    'Dim objCopyright As AssemblyCopyrightAttribute = CType(AssemblyCopyrightAttribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly, GetType(AssemblyCopyrightAttribute)), AssemblyCopyrightAttribute)
    '    'Dim objProduct As AssemblyProductAttribute = CType(AssemblyProductAttribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly, GetType(AssemblyProductAttribute)), AssemblyProductAttribute)
    '    'Dim objDescription As AssemblyDescriptionAttribute = CType(AssemblyDescriptionAttribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly, GetType(AssemblyDescriptionAttribute)), AssemblyDescriptionAttribute)
    'End Sub

End Class







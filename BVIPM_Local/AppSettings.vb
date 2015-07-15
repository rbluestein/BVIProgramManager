Imports System.DirectoryServices

Public Module Main
    Public cAppSettings As New AppSettings
    Public cEnviro As New Enviro
    Public cAppAdmin As New ApplicationAdministrator
    Public cCommon As New Common
End Module

Public Class AppSettings
    Public Const SoundLevelCutoff As Integer = 12
    Public Const GoodSoundLevel As Single = 0.65
    Public Const FairSoundLevel As Single = 0.3
    Public Const MinScriptRecordingSeconds As Integer = 30
    Private Const cDBTimeout As Integer = 10
    Private cProductionModeInd As Boolean = True

    Public ReadOnly Property DBTimeout() As Integer
        Get
            Return cDBTimeout
        End Get
    End Property

    Public ReadOnly Property ProductionModeInd() As Boolean
        Get
            Return cProductionModeInd
        End Get
    End Property
End Class

Public Class Enums
    Public Enum TransitionTo
        Disabled = 0
        Ready = 1
        LaunchProg = 2
        Refresh = 3
        ExitApp = 4
    End Enum
    Public Enum LaunchSource
        Tree = 1
        List = 2
        Button = 3
    End Enum
End Class

Public Class ChangeStateArgs
    Private cTransitionTo As Enums.TransitionTo
    Private cLaunchArgs As LaunchArgs

    Public Property TransitionTo() As Enums.TransitionTo
        Get
            Return cTransitionTo
        End Get
        Set(ByVal Value As Enums.TransitionTo)
            cTransitionTo = Value
        End Set
    End Property
    Public Property LaunchArgs() As LaunchArgs
        Get
            Return cLaunchArgs
        End Get
        Set(ByVal Value As LaunchArgs)
            cLaunchArgs = Value
        End Set
    End Property
End Class

Public Class LaunchArgs
    Private cLaunchSource As Enums.LaunchSource
    Private cSelectedTreeNode As TreeNode
    Private cDTRowNum As Integer

    Public Property LaunchSource() As Enums.LaunchSource
        Get
            Return cLaunchSource
        End Get
        Set(ByVal Value As Enums.LaunchSource)
            cLaunchSource = Value
        End Set
    End Property
    Public Property SelectedTreeNode() As TreeNode
        Get
            Return cSelectedTreeNode
        End Get
        Set(ByVal Value As TreeNode)
            cSelectedTreeNode = Value
        End Set
    End Property
    Public Property DTRowNum() As Integer
        Get
            Return cDTRowNum
        End Get
        Set(ByVal Value As Integer)
            cDTRowNum = Value
        End Set
    End Property
End Class

Public Class ApplicationAdministrator
    Public Event ExitApplication()
    Private clblStatus As Label

    Public Enum ReportTypeEnum
        Information = 1
        InformationNoLog = 2
        [Error] = 3
        ErrorNoShutdown = 4
        MessageBoxAndShutdown = 5
    End Enum

    Public Sub SetlblStatus(ByRef lblStatus As Label)
        clblStatus = lblStatus
    End Sub

    Public Sub Report(ByVal ReportType As ReportTypeEnum, ByVal Message As String, Optional ByVal Caption As String = Nothing)
        Dim Coll As New Collection
        Dim SendEmailResults As Results
        Dim HeaderMessage As String

        Select Case ReportType
            Case ReportTypeEnum.MessageBoxAndShutdown
                MessageBox.Show(Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                RaiseEvent ExitApplication()

            Case ReportTypeEnum.Information
                clblStatus.Text = Message
                clblStatus.Refresh()
                WriteToLogFile(Message)

            Case ReportTypeEnum.InformationNoLog
                clblStatus.Text = Message
                clblStatus.Refresh()

            Case ReportTypeEnum.Error
                clblStatus.Text = "Error"
                clblStatus.Refresh()
                WriteToLogFile(Message & " ** ERROR FORCING APPLICATION SHUTDOWN **")

            Case ReportTypeEnum.ErrorNoShutdown
                clblStatus.Text = "Error"
                clblStatus.Refresh()
                WriteToLogFile(Message)

        End Select

        Select Case ReportType
            Case ReportTypeEnum.Error, ReportTypeEnum.ErrorNoShutdown


                ' ___ Screen capture
                Dim FileFullPath As String
                FileFullPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & "ScreenCapture_" & System.Environment.UserName & "_" & Date.Now.ToUniversalTime.AddHours(-5).ToString("yyyyMMddssff") & ".gif"
                Common.CaptureScreen(FileFullPath)

                ' ___ Send email to help desk, including log and screenshot
                Coll.Add(cEnviro.LogFileFullPath)
                Coll.Add(FileFullPath)

                SendEmailResults = Common.SendEmail("HelpDesk@benefitvision.com;rbluestein@benefitvision.com", System.Environment.UserName & "@benefitvision.com", "", "BVIPM_Local.exe Program Manager error - " & System.Environment.UserName & "/" & System.Environment.MachineName & "  " & Date.Now.ToUniversalTime.AddHours(-5).ToString, "An error occurred in the execution of the Program Manager. Attached please find a copy of the application log and a screen shot.", Coll)
                If SendEmailResults.Success Then
                    HeaderMessage = "An error has occurred requiring the Program Manager to shut down. You may view the details of the problem in the log. The Program Manager has emailed a copy of the log to the help desk." & vbCrLf & vbCrLf
                Else
                    HeaderMessage = "An error has occurred requiring the Program Manager to shut down. You may view the details of the problem in the log." & vbCrLf & vbCrLf
                End If

                Dim FormMessage As New FrmMessage
                FormMessage.ShowDialog(HeaderMessage & Message)

                If System.IO.File.Exists(FileFullPath) Then
                    System.IO.File.Delete(FileFullPath)
                End If

        End Select

        If ReportType = ReportTypeEnum.Error Then
            RaiseEvent ExitApplication()
        End If

    End Sub

    'Private Function ReadLogFile() As String
    '    Dim StreamReader As System.IO.StreamReader
    '    Dim FileText As String

    '    Try
    '        StreamReader = New System.IO.StreamReader(cEnviro.LogFileFullPath)
    '        FileText = StreamReader.ReadToEnd

    '        'Do While StreamReader.Peek() >= 0
    '        '    'Console.WriteLine(StreamReader.ReadLine())

    '        '    x = StreamReader.ReadLine()
    '        'Loop
    '        StreamReader.Close()
    '        Return FileText

    '    Catch
    '    End Try
    'End Function

    Private Sub WriteToLogFile(ByVal Message As String)
        'Dim FileInfo As System.IO.FileInfo
        Dim StreamWriter As System.IO.StreamWriter = Nothing

        Try
            'FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)
            StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
            'StreamWriter.Write("[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] " & Message & vbCrLf)
            StreamWriter.Write(Common.GetTimeStamp & Message & vbCrLf)
            StreamWriter.Close()
        Catch ex As Exception
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
    Private cReturnStr As String

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
Public Class AcesInstallResults
    Inherits Results

    Private cInstallationRequred As Boolean
    Private cInstallationPath As String

    Public Property InstallationRequred() As Boolean
        Get
            Return cInstallationRequred
        End Get
        Set(ByVal Value As Boolean)
            cInstallationRequred = Value
        End Set
    End Property
    Public Property InstallationPath() As String
        Get
            Return cInstallationPath
        End Get
        Set(ByVal Value As String)
            cInstallationPath = Value
        End Set
    End Property
End Class
#End Region

Public Class Common
    Public Event ExitApplication()

    Public Shared Sub GenerateError()
        Dim a, b, c As Integer
        a = b / c
    End Sub

    Public ReadOnly Property NetworkConnectErrorMessage() As String
        Get
            Return "The Program Manager local has not detected the network V: drive with the password you entered. Passwords are case-sensitive. Please try again or report this problem to your supervisor or IT support person. If no one is available, submit an issue to the HelpDesk as a last resort. You will not be able to work with the Program Manager until the network drive is correctly mapped."
        End Get
    End Property

    Public Function TestAndConnectToNetwork(ByVal Source As String) As Boolean
        Dim MyResults As New Results
        Dim DirInfo As System.IO.DirectoryInfo
        Dim Count As Integer
        Dim Success As Boolean
        Dim WshNetwork As Object
        Dim PasswordForm As Password

        Try

            ' ___ Able to connect.
            'DirInfo = New System.IO.DirectoryInfo(cEnviro.ProgMgrServerPath)
            'If DirInfo.Exists Then
            '    Success = True
            'Else

            DirInfo = New System.IO.DirectoryInfo(cEnviro.ProgMgrServerPath)
            If DirInfo.Exists Then
                Success = True
            Else
                DirInfo = New System.IO.DirectoryInfo("\\web\eprog\bvipm")
                If DirInfo.Exists Then
                    cEnviro.ProgMgrServerPath = "\\web\eprog\bvipm"
                    Success = True
                End If
            End If

            ' ___ Try to connect from Main
            If Not Success Then
                WshNetwork = CreateObject("WScript.Network")
                If Source = "main" Then
                    For Count = 1 To 3
                        'cEnviro.Password = InputBox("Enter password:", "Password", "")
                        PasswordForm = New Password
                        PasswordForm.ShowDialog()
                        cEnviro.Password = PasswordForm.UserPassword
                        PasswordForm.Dispose()

                        Try
                            If System.IO.Directory.Exists("V:") Then
                                Success = True
                            Else
                                WshNetwork.MapNetworkDrive("V:", "\\192.168.1.12\eprog", False, Environment.UserName, cEnviro.Password)
                                Success = True
                            End If
                            If Success Then
                                Exit For
                            End If
                        Catch
                        End Try
                    Next
                End If

            ElseIf Source = "launcher" Then

                ' ___ Try to connect from launcher (start with previously entered password)
                If cEnviro.Password <> Nothing Then
                    Try

                        If System.IO.Directory.Exists("V:") Then
                            Success = True
                        Else
                            WshNetwork.MapNetworkDrive("V:", "\\192.168.1.12\eprog", False, Environment.UserName, cEnviro.Password)
                            Success = True
                        End If

                    Catch
                    End Try
                End If

                If Not Success Then
                    For Count = 1 To 3
                        'cEnviro.Password = InputBox("Enter password:", "Password", "")
                        PasswordForm = New Password
                        PasswordForm.ShowDialog()
                        cEnviro.Password = PasswordForm.UserPassword
                        PasswordForm.Dispose()

                        Try
                            WshNetwork.MapNetworkDrive("V:", "\\192.168.1.12\eprog", False, Environment.UserName, cEnviro.Password)
                            Success = True
                            Exit For
                        Catch
                        End Try
                    Next
                End If
            End If

            Return Success

        Catch
            Return False
        End Try
    End Function

    Public Shared Function GetTimeStamp() As String
        Return "[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] "
    End Function

    Public Shared Sub ResizeControls(ByRef Form As Windows.Forms.Form, ByVal DevResFormWidth As Single, ByVal DevResFormHeight As Single)
        Dim Size As Size
        Dim DevResScreenWidth As Single = 1280
        Dim DevResScreenHeight As Single = 800
        Dim CurResScreenWidth As Single = Screen.PrimaryScreen.Bounds.Width
        Dim CurResScreenHeight As Single = Screen.PrimaryScreen.Bounds.Height
        Dim BottomMenuPixels As Single = 28
        Dim TopMargin As Single = 28
        Dim SideAndBottomMargin As Single = 3

        Dim CurResFormWidth As Single
        Dim CurResFormHeight As Single
        Dim CurPercent As Single

        ' ___ Will the form fit on the screen as is?
        If Form.Width < CurResScreenWidth AndAlso Form.Height < CurResScreenHeight - BottomMenuPixels Then
            Exit Sub
        End If

        ' ___ The form is bigger than the screen. Resize the form.

        ' ___ Determine the new form dimensions
        If Form.Width > CurResScreenWidth Then
            CurResFormWidth = CurResScreenWidth
        Else
            CurResFormWidth = Form.Width
        End If
        If Form.Height > CurResScreenHeight - BottomMenuPixels Then
            CurResFormHeight = CurResScreenHeight - BottomMenuPixels
        Else
            CurResFormHeight = Form.Height
        End If

        Size = New Size(CurResFormWidth, CurResFormHeight)
        Form.MinimumSize = Size
        Form.MaximumSize = Size
        Form.Size = Size

        Dim Ctl As Control
        For Each Ctl In Form.Controls

            ' ___ Position
            CurPercent = (Ctl.Left + SideAndBottomMargin) / DevResFormWidth
            Ctl.Left = (CurPercent * CurResFormWidth) - SideAndBottomMargin
            CurPercent = (Ctl.Top + TopMargin) / DevResFormHeight
            Ctl.Top = (CurPercent * CurResFormHeight) - TopMargin

            ' ___ Size
            CurPercent = Ctl.Width / DevResFormWidth
            Ctl.Width = CurPercent * CurResFormWidth
            CurPercent = Ctl.Height / DevResFormHeight
            Ctl.Height = CurPercent * CurResFormHeight
        Next
    End Sub

    Public Shared Function IsConnectedToNetwork() As Boolean
        Dim DirInfo As System.IO.DirectoryInfo

        Try
            DirInfo = New System.IO.DirectoryInfo(cEnviro.ProgMgrServerPath)
            If DirInfo.Exists Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function DisallowRunningMultipleInstances() As Results
        Dim MyResults As New Results
        Dim ElapsedTime As Integer
        Dim Limit As Integer
        Dim IsRunning As Boolean = True

        Try

            Limit = 15000
            Do
                If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
                    IsRunning = True
                Else
                    IsRunning = False
                End If
                If Not IsRunning Then
                    Exit Do
                End If
                Threading.Thread.Sleep(250)
                ElapsedTime += 250
                If ElapsedTime >= Limit Then
                    Exit Do
                End If
            Loop

            If IsRunning Then
                MessageBox.Show("Another instance of Program Manager is already running", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                RaiseEvent ExitApplication()
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #4828: " & ex.Message
            Return MyResults
        End Try

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


    'Public Function MultipleInstancesRunning() As Boolean
    '    Dim Results As Boolean
    '    If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then

    '        'MessageBox.Show _
    '        ' ("Another Instance of this process is already running", _
    '        '     "Multiple Instances Forbidden", _
    '        '      MessageBoxButtons.OK, _
    '        '     MessageBoxIcon.Exclamation)
    '        Results = True
    '        'Application.Exit()
    '    End If
    '    Return Results
    'End Function

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
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2
            ' CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "mail.benefitvision.com"

            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").Value = 1
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername").Value = "automail"
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword").Value = "$bambam2004#"

            CDOConfig.Fields.Update()

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
            iMsg.Attachments.DeleteAll()
            CDOConfig = Nothing
            iMsg = Nothing

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #704: " & ex.Message
            Return MyResults
        End Try
    End Function

End Class

'Public Function GetDataType(ByVal Value As Object) As String
'    Dim Results As String
'    Results = Value.GetType.UnderlyingSystemType.Name
'    Return Results
'End Function

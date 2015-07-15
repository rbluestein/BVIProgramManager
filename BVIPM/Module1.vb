Imports System.IO

Public Module Main
    Public cEnviro As New Enviro
    Public cAppAdmin As New ApplicationAdministrator
    Private cNetworkConnectErrorMessage As String

    Public Sub Main()
        Dim MyResults As New Results
        Dim DisallowRunningMultipleInstancesResults As Results
        Dim FilePresentResults As Results
        Dim EnviroInitResults As Results
        Dim IsUpdateRequiredResults As Results
        Dim UpdateApplicationResults As Results
        Dim UpdateComponentsResults As Results
        Dim LaunchApplicationResults As Results
        Dim RollingSuccess As Boolean = True
        AddHandler cAppAdmin.ExitApplication, AddressOf ExitApplication

        Try

            cNetworkConnectErrorMessage = "The Program Manager launcher has not detected the network V: drive with the password you entered. Passwords are case-sensitive. Please try again or report this problem to your supervisor or IT support person. If no one is available, submit an issue to the HelpDesk as a last resort. You will not be able to work with the Program Manager until the network drive is correctly mapped."


            ' ___ Disallow running multiple instances
            'If (Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1) OrElse (Process.GetProcessesByName("BVIPM_Local").Length > 0) Then
            '    MessageBox.Show("Another instance of the program manager launcher is already running", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '    ExitApplication()
            'End If

            DisallowRunningMultipleInstancesResults = DisallowRunningMultipleInstances()
            If Not DisallowRunningMultipleInstancesResults.Success Then
                RollingSuccess = False
                MyResults.Message = DisallowRunningMultipleInstancesResults.Message
            End If

            ' ___ Set the network path
            cEnviro.ServerPath = "V:\BVIPM"

            ' ___ Test network issues
            ''If RollingSuccess Then
            ''    TestNetworkIssuesResults = TestNetworkIssues()
            ''    If Not TestNetworkIssuesResults.Success Then
            ''        NetworkConnectErrorMessage = "The Program Manager has not detected the network V: drive. Please report this problem to your supervisor or IT support person. If no one is available, submit an issue to the HelpDesk as a last resort. You will not be able to work with the Program Manager until the network drive is correctly mapped."
            ''        cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.MessageBoxAndShutdown, NetworkConnectErrorMessage, "Network Error")
            ''    End If
            ''End If

            ' ___ Test the network connection
            If RollingSuccess Then
                If Not TestAndConnectToNetwork("main") Then
                    RollingSuccess = False
                    MyResults.Message = "BVIPM.exe is unable to establish contact with network directory"
                    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.MessageBoxAndShutdown, cNetworkConnectErrorMessage, "Network Error")
                End If
            End If

            If RollingSuccess Then
                FilePresentResults = AreNetworkFilesPresent()
                If Not FilePresentResults.Success Then
                    MyResults.Success = False
                    MyResults.Message = FilePresentResults.Message
                    RollingSuccess = False
                    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.MessageBoxAndShutdown, FilePresentResults.Message, "Network File Error")
                End If

            End If

            ' ___ Set environment
            If RollingSuccess Then
                EnviroInitResults = EnviroInit()
                If Not EnviroInitResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = EnviroInitResults.Message
                End If
            End If

            ' ___ Disallow running on network drive
            If Path.GetDirectoryName(Application.ExecutablePath) = cEnviro.ServerPath Then
                MessageBox.Show("Attempting to run application from network drive", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                ExitApplication()
            End If

            '' ___ Test network issues
            'If RollingSuccess Then
            '    TestNetworkIssuesResults = TestNetworkIssues()
            '    If Not TestNetworkIssuesResults.Success Then
            '        RollingSuccess = False
            '        MyResults.Message = TestNetworkIssuesResults.Message
            '    End If
            'End If

            ' ___ Record app launch
            ' cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "** Launching " & Application.ExecutablePath & " Version " & GetVersion(Application.ExecutablePath) & "  **")
            Try
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "** Launching " & Application.ExecutablePath & " Version " & cEnviro.FormattedFileVersion(cEnviro.LauncherLocalFileVersion) & "  **")
            Catch
            End Try

            ' ___ Is update required?
            If RollingSuccess Then
                IsUpdateRequiredResults = IsUpdateRequired()
                If Not IsUpdateRequiredResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = IsUpdateRequiredResults.Message
                End If
            End If

            ' ___ Update application
            If RollingSuccess AndAlso IsUpdateRequiredResults.UpdateRequired Then
                UpdateApplicationResults = UpdateApplication()
                If Not UpdateApplicationResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = UpdateApplicationResults.Message
                End If
            End If

            ' ___ Update components
            If RollingSuccess Then
                UpdateComponentsResults = UpdateComponents()
                If Not UpdateComponentsResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = UpdateComponentsResults.Message
                End If
            End If

            ' ___ Launch application
            If RollingSuccess Then
                LaunchApplicationResults = LaunchApplication()
                If Not LaunchApplicationResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = LaunchApplicationResults.Message
                End If
            End If

            If Not RollingSuccess Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #2130b: " & MyResults.Message)
            End If

            ExitApplication()

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #2130a: " & ex.Message)
        End Try
    End Sub

    Public Function DisallowRunningMultipleInstances() As Results
        Dim MyResults As New Results
        Dim ElapsedTime As Integer
        Dim Limit As Integer
        Dim IsRunning As Boolean = True

        Try

            Limit = 30000
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
                ExitApplication()
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #4828: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Function TestAndConnectToNetwork(ByVal Source As String) As Boolean
        Dim MyResults As New Results
        Dim DirInfo As System.IO.DirectoryInfo
        Dim Count As Integer
        Dim Success As Boolean
        Dim WshNetwork As Object
        Dim UserPassword As String
        Dim PasswordForm As Password

        Try


            DirInfo = New DirectoryInfo(cEnviro.ServerPath)
            If DirInfo.Exists Then
                Success = True
            Else
                DirInfo = New DirectoryInfo("\\web\eprog\bvipm")
                If DirInfo.Exists Then
                    cEnviro.ServerPath = "\\web\eprog\bvipm"
                    Success = True
                End If
            End If

            If Not Success Then
                WshNetwork = CreateObject("WScript.Network")
                For Count = 1 To 3
                    'UserPassword = InputBox("Enter password:", "Password", "")
                    PasswordForm = New Password
                    PasswordForm.ShowDialog()
                    UserPassword = PasswordForm.UserPassword
                    PasswordForm.Dispose()

                    Try

                        ' ___ 07/03/2012: 192.168.1.17
                        WshNetwork.MapNetworkDrive("V:", "\\web\eprog", False, Environment.UserName, UserPassword)
                        Success = True
                        Exit For
                    Catch
                    End Try
                Next
            End If
            Return Success
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function AreNetworkFilesPresent() As Results
        Dim MyResults As New Results
        Dim i As Integer
        Dim ErrorMsg As String
        Dim ErrorColl As New Collection

        Try

            If Not TestAndConnectToNetwork("main") Then
                MyResults.Success = False
                MyResults.Message = cNetworkConnectErrorMessage
                Return MyResults
            End If

            If Not System.IO.File.Exists(cEnviro.ServerPath & "\BVIPM.exe") Then
                ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\BVIPM.exe")
            End If
            If Not System.IO.File.Exists(cEnviro.ServerPath & "\BVIPM_Local.exe") Then
                ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\BVIPM_Local.exe")
            End If
            If Not System.IO.File.Exists(cEnviro.ServerPath & "\Interop.CDO.dll") Then
                ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\Interop.CDO.dll")
            End If
            If Not System.IO.File.Exists(cEnviro.ServerPath & "\Interop.ADODB.dll") Then
                ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\Interop.ADODB.dll")
            End If
            If Not System.IO.File.Exists(cEnviro.ServerPath & "\Interop.CDO.dll") Then
                ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\Interop.CDO.dll")
            End If

            If ErrorColl.Count = 0 Then
                MyResults.Success = True
            Else
                ErrorMsg = String.Empty
                For i = 1 To ErrorColl.Count
                    If i < ErrorColl.Count Then
                        ErrorMsg &= ErrorColl(i) & ", "
                    Else
                        ErrorMsg &= ErrorColl(i) & "."
                    End If
                Next
                MyResults.Success = False
                MyResults.Message = ErrorMsg
            End If

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #4849: " & ex.Message
            Return MyResults
        End Try
    End Function

    'Private Function TestNetworkIssues() As Results
    '    Dim MyResults As New Results
    '    Dim i As Integer
    '    Dim ErrorMsg As String
    '    Dim ErrorColl As New Collection
    '    Dim DirInfo As DirectoryInfo

    '    Try

    '        ' ___ Can the app connect to the network directory?
    '        Try
    '            ErrorMsg = "BVIPM.exe is unable to establish contact with network directory"
    '            DirInfo = New DirectoryInfo(cEnviro.ServerPath)
    '            If Not DirInfo.Exists Then
    '                ErrorColl.Add(ErrorMsg)
    '            End If
    '        Catch ex As Exception
    '            ErrorColl.Add(ErrorMsg)
    '        End Try

    '        ' ___ Are the files present on the network drive?
    '        Try
    '            If DirInfo.Exists Then

    '                If Not System.IO.File.Exists(cEnviro.ServerPath & "\BVIPM.exe") Then
    '                    ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\BVIPM.exe")
    '                End If
    '                If Not System.IO.File.Exists(cEnviro.ServerPath & "\BVIPM_Local.exe") Then
    '                    ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\BVIPM_Local.exe")
    '                End If
    '                If Not System.IO.File.Exists(cEnviro.ServerPath & "\Interop.CDO.dll") Then
    '                    ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\Interop.CDO.dll")
    '                End If
    '                If Not System.IO.File.Exists(cEnviro.ServerPath & "\Interop.ADODB.dll") Then
    '                    ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\Interop.ADODB.dll")
    '                End If
    '                If Not System.IO.File.Exists(cEnviro.ServerPath & "\Interop.CDO.dll") Then
    '                    ErrorColl.Add("Missing file: " & cEnviro.ServerPath & "\Interop.CDO.dll")
    '                End If

    '            End If

    '        Catch ex As Exception
    '            ErrorColl.Add("Unable to test for presence of network files")
    '        End Try

    '        If ErrorColl.Count = 0 Then
    '            MyResults.Success = True
    '        Else
    '            ErrorMsg = String.Empty
    '            For i = 1 To ErrorColl.Count
    '                If i < ErrorColl.Count Then
    '                    ErrorMsg &= ErrorColl(i) & ", "
    '                Else
    '                    ErrorMsg &= ErrorColl(i) & "."
    '                End If
    '            Next
    '            MyResults.Success = False
    '            MyResults.Message = ErrorMsg
    '        End If

    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #4849: " & ex.message
    '        Return MyResults
    '    End Try

    '    '    Dim DirInfo As New DirectoryInfo(Path.GetDirectoryName(Application.ExecutablePath))
    '    'Dim FileInfo As FileInfo()
    '    'FileInfo = DirInfo.GetFiles
    '    'For i = 0 To FileInfo.GetUpperBound(0)
    '    '    Name = FileInfo(i).Name
    '    '    'If Name.Length > TestString.Length Then
    '    '    '    If Name.Substring(0, TestString.Length).ToLower = TestString Then
    '    '    '        File.Delete(FileInfo(i).FullName)
    '    '    '    End If
    '    '    'End If
    '    'Next
    'End Function

    Private Function EnviroInit() As Results
        Dim MyResults As New Results

        Try

            cEnviro.UserId = Environment.UserName
            cEnviro.ComputerId = Environment.MachineName.ToUpper
            cEnviro.SiteId = "HBG"
            'Select Case cEnviro.ComputerId.Substring(0, 2)
            '    Case "HB" : cEnviro.SiteId = "HBG"
            '    Case "OK" : cEnviro.SiteId = "OKC"
            '    Case "LA" : cEnviro.SiteId = "LA"
            '    Case Else : cEnviro.SiteId = "HBG"
            'End Select

            cEnviro.LogFileFullPath = Path.GetDirectoryName(Application.ExecutablePath) & "\" & cEnviro.UserId & ".log"
            'cEnviro.ServerPath = "V:\BVIPM"

            If Not TestAndConnectToNetwork("main") Then
                MyResults.Success = False
                MyResults.Message = "Error #2131a: " & cNetworkConnectErrorMessage
                Return MyResults
            End If

            ' ___ Version
            Try
                'Dim LauncherLocalBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM.exe")
                'cEnviro.LauncherLocalFileVersion = LauncherLocalBuildInfo.FileVersion
                'Dim LocalMasterBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(cEnviro.ServerPath & "\BVIPM_Local.exe")
                'cEnviro.LocalMasterFileVersion = LocalMasterBuildInfo.FileVersion
                'Dim LocalLocalBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_Local.exe")
                'cEnviro.LocalLocalFileVersion = LocalLocalBuildInfo.FileVersion

                ' ___ This app
                Dim CDrive_Launcher_BuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM.exe")
                cEnviro.LauncherLocalFileVersion = CDrive_Launcher_BuildInfo.FileVersion

                ' ___ Local on this machine
                Try
                    Dim CDrive_Local_BuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_Local.exe")
                    cEnviro.LocalLocalFileVersion = CDrive_Local_BuildInfo.FileVersion
                Catch ' If it fails it means there's no local version so set the version to 0.
                    cEnviro.LocalLocalFileVersion = "0.0.0.0"
                End Try

                ' ____ Launcher on V drive
                Dim VDrive_Local_BuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(cEnviro.ServerPath & "\BVIPM_Local.exe")
                cEnviro.LocalMasterFileVersion = VDrive_Local_BuildInfo.FileVersion

            Catch ex As Exception
                MyResults.Success = False
                MyResults.Message = "#Error #2131b: " & ex.Message
                Return MyResults
            End Try

                MyResults.Success = True
                Return MyResults

            Catch ex As Exception
                MyResults.Success = False
                MyResults.Message = "Error #2131c: " & ex.Message
                Return MyResults
            End Try
    End Function

    Private Function UpdateComponents() As Results
        Dim MyResults As New Results

        Try

            If Not TestAndConnectToNetwork("main") Then
                MyResults.Success = False
                MyResults.Message = "Error 2135a: " & cNetworkConnectErrorMessage
                Return MyResults
            End If

            If Not System.IO.File.Exists(Path.GetDirectoryName(Application.ExecutablePath) & "\Interop.ADODB.dll") Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Downloading Interop.ADODB.dll.")
                System.IO.File.Copy(cEnviro.ServerPath & "\Interop.ADODB.dll", Path.GetDirectoryName(Application.ExecutablePath) & "\Interop.ADODB.dll.", True)
                Application.DoEvents()
            End If
            If Not System.IO.File.Exists(Path.GetDirectoryName(Application.ExecutablePath) & "\Interop.CDO.dll") Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Downloading Interop.CDO.dll.")
                System.IO.File.Copy(cEnviro.ServerPath & "\Interop.CDO.dll", Path.GetDirectoryName(Application.ExecutablePath) & "\Interop.CDO.dll", True)
                Application.DoEvents()
            End If
            If Not System.IO.File.Exists(Path.GetDirectoryName(Application.ExecutablePath) & "\ComputerAcid.ico") Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Downloading ComputerAcid.ico.")
                System.IO.File.Copy(cEnviro.ServerPath & "\ComputerAcid.ico", Path.GetDirectoryName(Application.ExecutablePath) & "\ComputerAcid.ico", True)
                Application.DoEvents()
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #2135b: " & ex.Message
            Return MyResults
        End Try
    End Function

    'Private Function GetVersion(ByVal FileFullPath) As String
    '    Dim BuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(FileFullPath)
    '    Return BuildInfo.FileMajorPart & "." & BuildInfo.FileMinorPart
    'End Function

    Private Function IsUpdateRequired() As Results
        Dim MyResults As New Results

        Try

            If cEnviro.LocalLocalFileVersion <> cEnviro.LocalMasterFileVersion Then
                MyResults.UpdateRequired = True
            End If

            'Dim MasterBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(cEnviro.ServerPath & "\BVIPM_LOCAL.exe")
            'Dim ThisBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_LOCAL.exe")

            'cEnviro.MasterVersion = MasterBuildInfo.FileMajorPart & "." & MasterBuildInfo.FileMinorPart
            'cEnviro.CurrentVersion = ThisBuildInfo.FileMajorPart & "." & ThisBuildInfo.FileMinorPart

            'If CInt(MasterBuildInfo.FileMajorPart) > CInt(ThisBuildInfo.FileMajorPart) Then
            '    MyResults.UpdateRequired = True
            'Else
            '    If CInt(MasterBuildInfo.FileMajorPart) = CInt(ThisBuildInfo.FileMajorPart) Then
            '        If CInt(MasterBuildInfo.FileMinorPart) > CInt(ThisBuildInfo.FileMinorPart) Then
            '            MyResults.UpdateRequired = True
            '        End If
            '    End If
            'End If

            'If CInt(MasterBuildInfo.FileMajorPart) <> CInt(ThisBuildInfo.FileMajorPart) Then
            '    MyResults.UpdateRequired = True
            'ElseIf CInt(MasterBuildInfo.FileMinorPart) <> CInt(ThisBuildInfo.FileMinorPart) Then
            '    MyResults.UpdateRequired = True
            'End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #2132: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Function UpdateApplication() As Results
        Dim MyResults As New Results
        Dim OKToCopy As Boolean
        Dim ProcessName As String
        Dim ElapsedTime As Integer

        Try

            If Not TestAndConnectToNetwork("main") Then
                MyResults.Success = False
                MyResults.Message = "Error #2133a: " & cNetworkConnectErrorMessage
                Return MyResults
            End If

            Try
                'cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Performing update of BVIPM_Local.exe from version " & cEnviro.CurrentVersion & " to " & cEnviro.MasterVersion)
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Performing update of BVIPM_Local.exe from version " & cEnviro.FormattedFileVersion(cEnviro.LocalLocalFileVersion) & " to " & cEnviro.FormattedFileVersion(cEnviro.LocalMasterFileVersion))
            Catch ex As Exception
                MyResults.Success = False
                MyResults.Message = "Error #2133b: " & ex.Message
                Return MyResults
            End Try

            ' ___ Allow Windows up to 30 seconds to close the calling instance of the launcher
            Try
                ProcessName = Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_LOCAL.exe"
                Do
                    If Common.ProcessFound(ProcessName) Then
                        System.Threading.Thread.Sleep(250)
                    Else
                        OKToCopy = True
                        Exit Do
                    End If
                    ElapsedTime += 250
                    If ElapsedTime >= 30000 Then
                        Exit Do
                    End If
                Loop
            Catch ex As Exception
                MyResults.Success = False
                MyResults.Message = "Error #2133c: " & ex.Message
                Return MyResults
            End Try


            If Not TestAndConnectToNetwork("main") Then
                MyResults.Success = False
                MyResults.Message = "Error #2133d: " & cNetworkConnectErrorMessage
                Return MyResults
            End If

            Try
                If OKToCopy Then
                    System.IO.File.Copy(cEnviro.ServerPath & "\BVIPM_LOCAL.exe", Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_LOCAL.exe", True)
                    Application.DoEvents()
                    MyResults.Success = True
                Else
                    MyResults.Success = False
                    MyResults.Message = "Error #2133e: " & "Application reached 30 second timeout waiting for BVIPM_Local.exe to free up resources."
                End If
            Catch ex As Exception
                MyResults.Success = False
                MyResults.Message = "Error #2133f: " & ex.Message
                Return MyResults
            End Try

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #2133g: " & ex.Message
            Return MyResults
        End Try

        ' Try now to create a shortcut on the desktop, assuming the new program exists.
        Dim wsh As Object = CreateObject("WScript.Shell")

        wsh = CreateObject("WScript.Shell")

        Dim MyShortcut, DesktopPath As Object

        ' Read desktop path using WshSpecialFolders object

        DesktopPath = wsh.SpecialFolders("AllUsersDesktop")

        ' Create a shortcut object on the desktop

        MyShortcut = wsh.CreateShortcut(DesktopPath & "\BVIPM.lnk")

        ' Set shortcut object properties and save it

        MyShortcut.TargetPath = Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM.exe"

        MyShortcut.WorkingDirectory = Path.GetDirectoryName(Application.ExecutablePath) & "\"

        MyShortcut.WindowStyle = 4

        'Use this next line to assign a icon other then the default icon for the exe

        MyShortcut.IconLocation = Path.GetDirectoryName(Application.ExecutablePath) & "\ComputerAcid.ico,0"

        'Save the shortcut

        MyShortcut.Save()

        Return MyResults


    End Function

    Private Function LaunchApplication() As Results
        Dim MyResults As New Results
        Dim ElapsedTime As Integer

        Try

            ' ___ Invoke timeout in the event Windows has not finished processing the download


            ElapsedTime = 0
            Do
                Try
                    Shell(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_LOCAL.exe", AppWinStyle.NormalFocus)
                    Exit Do
                Catch
                    System.Threading.Thread.Sleep(250)
                    ElapsedTime += 250
                End Try
                If ElapsedTime >= 30000 Then
                    MyResults.Success = False
                    MyResults.Message = "Error #2137: Timeout occurred between download and execution of BVIPM_Local.exe."
                End If
            Loop





            ' Shell(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_LOCAL.exe", AppWinStyle.NormalFocus)
            'Dim Process As System.Diagnostics.Process = New Process
            'Process.StartInfo.FileName = Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_LOCAL.exe"
            'Process.StartInfo.UseShellExecute = True
            'Process.Start()
            'Application.Exit()

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #2134: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Sub ExitApplication()
        '  Application.Exit()
        Environment.Exit(0)
    End Sub
End Module

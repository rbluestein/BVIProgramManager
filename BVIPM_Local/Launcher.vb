Imports System.IO

Public Class Launcher
    Private cDBase As DBase
    Private cLaunchExecute As New LaunchExecute
    Private cTestTrainProd As TestTrainProdEnum

    Public Enum TestTrainProdEnum
        Prod = 0
        Test = 1
        Training = 2
    End Enum

    Public Function LaunchApp(ByRef DBase As DBase, ByVal RowNum As Integer, ByVal TestTrainProd As TestTrainProdEnum, ByVal SourceIsLaunchLocal As Boolean) As Results
        Dim MyResults As New Results
        Dim CollectInfoResults As Results
        Dim ExecuteApplicationResults As New Results

        Try

            cDBase = DBase
            cTestTrainProd = TestTrainProd
            MyResults.Success = True

            ' ___ Collect the program information
            If MyResults.Success Then
                cLaunchExecute = New LaunchExecute
                CollectInfoResults = CollectInfo(RowNum)
                MyResults.Success = CollectInfoResults.Success
                If Not CollectInfoResults.Success Then
                    MyResults.Success = False
                    MyResults.Message = CollectInfoResults.Message
                End If
            End If

            ' ___ Execute the application
            If MyResults.Success Then
                ExecuteApplicationResults = ExecuteApplication(RowNum)
                If Not ExecuteApplicationResults.Success Then
                    MyResults.Success = False
                    MyResults.Message = ExecuteApplicationResults.Message
                End If
            End If

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #640: " & ex.Message
            Return MyResults
        End Try
    End Function

    Public Function CollectInfo(ByVal RowNum As Integer) As Results
        Dim MyResults As New Results
        Dim Pos As Integer

        Try

            cLaunchExecute.ServerPath = String.Empty
            cLaunchExecute.TableListingForLocalPath32 = String.Empty
            cLaunchExecute.TableListingForLocalPath64 = String.Empty
            cLaunchExecute.FileName = String.Empty
            cLaunchExecute.CommandLineArgs = String.Empty
            cLaunchExecute.Description = String.Empty

            ' ___ FileName / Command line args
            If Not IsDBNull(cEnviro.dtProgs.Rows(RowNum)("FileName")) Then
                cLaunchExecute.FileName = cEnviro.GetValue(RowNum, Progs.FileName)
                Pos = InStr(cLaunchExecute.FileName, ".exe")
                If cLaunchExecute.FileName.Length > Pos + 3 Then
                    cLaunchExecute.CommandLineArgs = cLaunchExecute.FileName.Substring(Pos + 4)
                    cLaunchExecute.FileName = cLaunchExecute.FileName.Substring(0, Pos + 3)
                End If
            End If

            ' ___ Description
            If Not IsDBNull(cEnviro.dtProgs.Rows(RowNum)("Description")) Then
                cLaunchExecute.Description = cEnviro.dtProgs.Rows(RowNum)("Description")
            End If

            ' ___ Type of program: web, local or server.
            If (Not IsDBNull(cEnviro.GetValue(RowNum, Progs.ProgramTypeCode))) AndAlso (cEnviro.GetValue(RowNum, Progs.ProgramTypeCode).ToLower = "web") Then
                cLaunchExecute.Location = LaunchExecute.LocationEnum.Web
            Else
                If cEnviro.GetValue(RowNum, Progs.RunLocal).IndexOfAny("Yy") > -1 Then
                    cLaunchExecute.Location = LaunchExecute.LocationEnum.Local
                Else
                    cLaunchExecute.Location = LaunchExecute.LocationEnum.Server
                End If
            End If

            ' ___ ServerPath and app
            If Not IsDBNull(cEnviro.GetValue(RowNum, Progs.ServerPath)) Then
                cLaunchExecute.ServerPath = StripSlash(cEnviro.GetValue(RowNum, Progs.ServerPath))
                If System.IO.Directory.Exists(cLaunchExecute.ServerPath) Then
                    cLaunchExecute.ServerPathExists = True
                    If System.IO.File.Exists(cLaunchExecute.ServerPath & "\" & cLaunchExecute.FileName) Then
                        cLaunchExecute.ServerProgExists = True
                    End If
                End If
            End If

            If cLaunchExecute.Location = LaunchExecute.LocationEnum.Local Then

                ' ___ LocalPath32
                If Not IsDBNull(cEnviro.GetValue(RowNum, Progs.LocalPath)) Then
                    cLaunchExecute.TableListingForLocalPath32 = StripSlash(cEnviro.GetValue(RowNum, Progs.LocalPath))
                    cLaunchExecute.TableHasListingForLocalPath32 = True
                    If cLaunchExecute.TableHasListingForLocalPath32 Then
                        If System.IO.Directory.Exists(cLaunchExecute.TableListingForLocalPath32) Then
                            cLaunchExecute.LocalPath32Exists = True
                            If System.IO.File.Exists(cLaunchExecute.TableListingForLocalPath32 & "\" & cLaunchExecute.FileName) Then
                                cLaunchExecute.LocalProg32Exists = True
                            End If
                        End If
                    End If
                End If

                ' ___ LocalPath64
                If Not IsDBNull(cEnviro.GetValue(RowNum, Progs.LocalPath64)) Then
                    cLaunchExecute.TableListingForLocalPath64 = StripSlash(cEnviro.GetValue(RowNum, Progs.LocalPath64))
                    cLaunchExecute.TableHasListingForLocalPath64 = True
                    If cLaunchExecute.TableHasListingForLocalPath64 Then
                        If System.IO.Directory.Exists(cLaunchExecute.TableListingForLocalPath64) Then
                            cLaunchExecute.LocalPath64Exists = True
                            If System.IO.File.Exists(cLaunchExecute.TableListingForLocalPath64 & "\" & cLaunchExecute.FileName) Then
                                cLaunchExecute.LocalProg64Exists = True
                            End If
                        End If
                    End If
                End If

                ' ___ Aces app? Aces installer?
                cLaunchExecute.AcesInd = cEnviro.GetValue(RowNum, Progs.AcesInd)
                If cLaunchExecute.AcesInd Then
                    If (Not IsDBNull(cEnviro.GetValue(RowNum, Progs.Installer))) AndAlso cEnviro.GetValue(RowNum, Progs.Installer).Length > 0 Then
                        cLaunchExecute.Installer = cEnviro.GetValue(RowNum, Progs.Installer)
                        If System.IO.File.Exists(cLaunchExecute.ServerPath & "\" & cLaunchExecute.Installer) Then
                            cLaunchExecute.AcesInstallerExists = True
                        End If
                    End If
                End If

            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #642: " & ex.Message
            Return MyResults
        End Try
    End Function

#Region " Launch application "
    Private Function ExecuteApplication(ByVal RowNum As Integer) As Results
        Dim MyResults As New Results
        Dim PerformVersionUpdateResults As AcesInstallResults
        Dim LaunchLocalResults As Results
        Dim AppFullPath As String = Nothing
        Dim FileInfo As FileInfo
        Dim InstallationRequred As Boolean
        Dim InstallationPath As String = Nothing

        Try

            MyResults.Success = True

            ' ___ Test/Reconnect to V drive
            If cLaunchExecute.Location <> LaunchExecute.LocationEnum.Local Then
                If Not cCommon.TestAndConnectToNetwork("launcher") Then
                    MyResults.Success = False
                    MyResults.Message = cCommon.NetworkConnectErrorMessage
                    Return MyResults
                End If
            End If

            If cLaunchExecute.Location = LaunchExecute.LocationEnum.Web Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Launching " & cEnviro.GetValue(RowNum, Progs.ClientID) & "-" & cLaunchExecute.Description & " in browser")
                AppFullPath = cEnviro.GetValue(RowNum, Progs.ServerPath)
                System.Diagnostics.Process.Start(cLaunchExecute.ServerPath)

            ElseIf cLaunchExecute.Location = LaunchExecute.LocationEnum.Local Then

                ' ___ Aces update
                If cLaunchExecute.AcesInd AndAlso cLaunchExecute.AcesInstallerExists Then
                    PerformVersionUpdateResults = PerformAcesUpdateIfNeeded(RowNum)
                    MyResults.Success = PerformVersionUpdateResults.Success
                    If PerformVersionUpdateResults.Success Then
                        InstallationRequred = PerformVersionUpdateResults.InstallationRequred
                        InstallationPath = PerformVersionUpdateResults.InstallationPath
                    Else
                        MyResults.Success = False
                        MyResults.Message = PerformVersionUpdateResults.Message
                        Return MyResults
                    End If
                End If

                ' ___ LaunchLocal
                LaunchLocalResults = LaunchLocal(RowNum, InstallationRequred, InstallationPath)
                If Not LaunchLocalResults.Success Then
                    MyResults.Success = False
                    MyResults.Message = LaunchLocalResults.Message
                    Return MyResults
                End If

            ElseIf cLaunchExecute.Location = LaunchExecute.LocationEnum.Server Then

                ' ___ Perform connection test
                FileInfo = New FileInfo(cLaunchExecute.ServerPath & "\" & cLaunchExecute.FileName)
                If Not FileInfo.Exists Then
                    MyResults.Success = False
                    MyResults.Message = "BVIPM_Local.exe is unable to establish contact with " & cLaunchExecute.ServerPath & "\" & cLaunchExecute.FileName
                    Return MyResults
                End If

                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Launching " & cLaunchExecute.FileName)
                AppFullPath = cLaunchExecute.ServerPath & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs
                System.IO.Directory.SetCurrentDirectory(cLaunchExecute.ServerPath)
                Shell(cLaunchExecute.ServerPath & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs, AppWinStyle.NormalFocus)
            End If

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #643: AppFullFile is " & AppFullPath & "  " & "Error message: " & ex.Message
            Return MyResults
        End Try
    End Function

    'Private Function LaunchLocal(ByVal RowNum As Integer) As Results
    '    Dim AppFullPath As String
    '    Dim MyResults As New Results
    '    Dim SuperviseInstallerResults As Results


    '    Try

    '        MyResults.Success = True

    '        If cLaunchExecute.LOCALPROGEXISTS Then
    '            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Launching " & cLaunchExecute.FileName & "  " & cLaunchExecute.Description)
    '            AppFullPath = cLaunchExecute.LOCALPATH & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs
    '            System.IO.Directory.SetCurrentDirectory(cLaunchExecute.LOCALPATH)
    '            Shell(cLaunchExecute.LOCALPATH & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs, AppWinStyle.NormalFocus)

    '        Else

    '            If Not Directory.Exists(cLaunchExecute.LOCALPATH) Then
    '                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Creating '" & cLaunchExecute.LOCALPATH)
    '                System.IO.Directory.CreateDirectory(cLaunchExecute.LOCALPATH)
    '            End If

    '            If cLaunchExecute.InstallerExists Then
    '                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Launching installer for " & cLaunchExecute.FileName & "  " & cLaunchExecute.Description)

    '                SuperviseInstallerResults = SuperviseInstaller()

    '                If SuperviseInstallerResults.Success Then
    '                    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Launching " & cLaunchExecute.FileName & "  " & cLaunchExecute.Description)
    '                    AppFullPath = cLaunchExecute.LOCALPATH & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs
    '                    System.IO.Directory.SetCurrentDirectory(cLaunchExecute.LOCALPATH)
    '                    Shell(cLaunchExecute.LOCALPATH & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs, AppWinStyle.NormalFocus)
    '                Else
    '                    MyResults.Success = False
    '                    MyResults.Message = SuperviseInstallerResults.Message
    '                End If

    '            ElseIf (Not cLaunchExecute.InstallerExists) AndAlso cLaunchExecute.ServerProgExists Then

    '                ' ___ Copy the program from the server and try executing it
    '                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Copying '" & cLaunchExecute.FileName & "' from server")
    '                System.IO.File.Copy(cLaunchExecute.ServerPath & "\" & cLaunchExecute.FileName, cLaunchExecute.LOCALPATH & "\" & cLaunchExecute.FileName, True)
    '                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Launching " & cEnviro.GetValue(RowNum, Progs.ProgramID))
    '                AppFullPath = cLaunchExecute.LOCALPATH & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs
    '                Shell(cLaunchExecute.LOCALPATH & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs, AppWinStyle.NormalFocus)
    '            Else
    '                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Unable to launch " & cLaunchExecute.FileName & ". Files missing from local machine and server")
    '            End If
    '        End If

    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #646: AppFullFile is " & AppFullPath & "  " & "Error message: " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    'Private Function LaunchLocal(ByVal RowNum As Integer, ByVal InstallationRequired As Boolean, ByVal InstallationPath As String) As Results
    '    Dim DirPath As String
    '    Dim MyResults As New Results

    '    Try

    '        ' ___ Handle launch of updated Aces app
    '        If cLaunchExecute.AcesInd AndAlso InstallationRequired Then
    '            DirPath = InstallationPath

    '        Else

    '            ' ___ App resides in 64 folder if compiled as 64 bit app.
    '            Select Case cEnviro.OSPlatformBitCount
    '                Case Enviro.OSPlatformBitCountEnum.BitCount32
    '                    DirPath = cLaunchExecute.TableLocalPath32

    '                Case Enviro.OSPlatformBitCountEnum.BitCount64
    '                    If cLaunchExecute.LocalPath64Exists AndAlso cLaunchExecute.LocalProg64Exists Then
    '                        DirPath = cLaunchExecute.TableLocalPath64
    '                    ElseIf cLaunchExecute.LocalPath32Exists AndAlso cLaunchExecute.LocalProg32Exists Then
    '                        DirPath = cLaunchExecute.TableLocalPath32
    '                    End If
    '            End Select

    '            If DirPath = Nothing Then
    '                If cLaunchExecute.ServerProgExists Then
    '                    DirPath = cLaunchExecute.TableLocalPath32
    '                    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Copying '" & cLaunchExecute.FileName & "' from server")
    '                    System.IO.File.Copy(cLaunchExecute.ServerPath & "\" & cLaunchExecute.FileName, DirPath & "\" & cLaunchExecute.FileName, True)
    '                Else
    '                    MyResults.Success = False
    '                    MyResults.Message = "Error #646a: AppFullFile is " & cLaunchExecute.TableLocalPath32 & "\ " & cLaunchExecute.FileName & ". No local and no server copies of the app."
    '                    Return MyResults
    '                End If
    '            End If

    '        End If


    '        System.IO.Directory.SetCurrentDirectory(DirPath)
    '        Shell(DirPath & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs, AppWinStyle.NormalFocus)

    '        MyResults.Success = True
    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #646b: AppFullFile is " & DirPath & "  " & "Error message: " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    Private Function LaunchLocal(ByVal RowNum As Integer, ByVal InstallationRequired As Boolean, ByVal InstallationPath As String) As Results
        Dim LaunchPath As String = Nothing
        Dim LaunchPathExists As Boolean
        Dim LocalProgExists As Boolean
        Dim MyResults As New Results
        Dim OKToLaunch As Boolean
        Dim AppFullPath As String = Nothing
        Dim Use32 As Boolean
        Dim Use64 As Boolean

        Try

            ' ___ Handle launch of updated Aces app
            If cLaunchExecute.AcesInd AndAlso InstallationRequired Then
                LaunchPath = InstallationPath
                OKToLaunch = True

            Else

                ' ___ Non-Aces
                If cEnviro.OSPlatformBitCount = Enviro.OSPlatformBitCountEnum.BitCount32 Then
                    Use32 = True
                Else
                    If cLaunchExecute.TableHasListingForLocalPath64 Then
                        Use64 = True
                    Else
                        Use32 = True
                    End If
                End If

                If Use32 Then
                    LaunchPath = cLaunchExecute.TableListingForLocalPath32
                    LaunchPathExists = cLaunchExecute.LocalPath32Exists
                    LocalProgExists = cLaunchExecute.LocalProg32Exists
                ElseIf Use64 Then
                    LaunchPath = cLaunchExecute.TableListingForLocalPath64
                    LaunchPathExists = cLaunchExecute.LocalPath64Exists
                    LocalProgExists = cLaunchExecute.LocalProg64Exists
                End If
                AppFullPath = LaunchPath & "\" & cLaunchExecute.FileName


                If LaunchPathExists And LocalProgExists Then
                    OKToLaunch = True
                Else
                    If cLaunchExecute.ServerPathExists AndAlso cLaunchExecute.ServerProgExists Then
                        If Not LaunchPathExists Then
                            Try
                                System.IO.Directory.CreateDirectory(LaunchPath)
                            Catch ex As Exception
                                MyResults.Success = False
                                MyResults.Message = "Error #646a: AppFullPath is " & AppFullPath & ". Unable to create local directory. " & ex.message
                                Return MyResults
                            End Try
                        End If
                        If Not LocalProgExists Then
                            Try
                                System.IO.File.Copy(cLaunchExecute.ServerPath & "\" & cLaunchExecute.FileName, LaunchPath & "\" & cLaunchExecute.FileName, True)
                            Catch ex As Exception
                                MyResults.Success = False
                                MyResults.Message = "Error #646b: AppFullPath is " & AppFullPath & ". Unable to copy file to local directory. " & ex.message
                                Return MyResults
                            End Try
                        End If
                        OKToLaunch = True
                    Else
                        MyResults.Success = False
                        MyResults.Message = "Error #646c: AppFullPath is " & AppFullPath & ". No local and no server copies of the app."
                        Return MyResults
                    End If
                End If

            End If

            Try
                System.IO.Directory.SetCurrentDirectory(LaunchPath)
                Shell(LaunchPath & "\" & cLaunchExecute.FileName & " " & cLaunchExecute.CommandLineArgs, AppWinStyle.NormalFocus)
            Catch ex As Exception
                MyResults.Success = False
                MyResults.Message = "Error #646d: AppFullPath is " & AppFullPath & ". Unable to launch application. " & ex.Message
                Return MyResults
            End Try

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #646e: AppFullPath is " & AppFullPath & "  " & "Error message: " & ex.Message
            Return MyResults
        End Try
    End Function
#End Region

#Region " Aces "
    Private Function PerformAcesUpdateIfNeeded(ByVal RowNum As Integer) As AcesInstallResults
        Dim MyResults As New AcesInstallResults
        Dim SuperviseInstallerResults As AcesInstallResults
        Dim ThisBuildInfo32 As FileVersionInfo = Nothing
        Dim ThisBuildInfo64 As FileVersionInfo = Nothing
        Dim ThisBuildFileVersion32 As String = Nothing
        Dim ThisBuildFileVersion64 As String = Nothing
        Dim Sql As String
        Dim QueryPack As DBase.QueryPack
        Dim dt As DataTable

        Try

            MyResults.Success = True

            ' ___ Get the file version on the local machine
            If cLaunchExecute.LocalProg32Exists Then
                ThisBuildInfo32 = FileVersionInfo.GetVersionInfo(cLaunchExecute.TableListingForLocalPath32 & "\" & cLaunchExecute.FileName)
                ThisBuildFileVersion32 = ThisBuildInfo32.FileVersion
            End If
            If cLaunchExecute.LocalProg64Exists Then
                ThisBuildInfo64 = FileVersionInfo.GetVersionInfo(cLaunchExecute.TableListingForLocalPath64 & "\" & cLaunchExecute.FileName)
                ThisBuildFileVersion64 = ThisBuildInfo64.FileVersion
            End If

            ' // 3/5/2011 all moved to hbg-sql
            '' ___ Get the current master file version
            'Select Case cTestTrainProd
            '    Case TestTrainProdEnum.Prod
            '        DBConnString = cEnviro.DBConnStringPrefix & cEnviro.SiteId & "-SQL"
            '    Case TestTrainProdEnum.Test
            '        DBConnString = cEnviro.DBConnStringPrefix & "HBG-TST"
            '    Case TestTrainProdEnum.Training
            '        DBConnString = cEnviro.DBConnStringPrefix & "TRAINING"
            'End Select

            Select Case cTestTrainProd
                Case TestTrainProdEnum.Test
                    Sql = "SELECT CurrVersion FROM Program WHERE ProgramID = 'ACES_TEST' AND Description = 'ACES VERSION PLACEHOLDER'"
                Case Else
                    Sql = "SELECT CurrVersion FROM Program WHERE ProgramID = 'ACES' AND Description = 'ACES VERSION PLACEHOLDER'"
            End Select

            QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString, False)
            If QueryPack.Success Then
                dt = QueryPack.dt
            Else
                MyResults.Success = False
                MyResults.Message = "Error #648 (PerformVersionUpdate): " & QueryPack.TechErrMsg
                Return MyResults
            End If
            cLaunchExecute.AcesVersion = dt.Rows(0)(0)

            ' ___ Run installer if local version is not current
            If (ThisBuildFileVersion32 <> cLaunchExecute.AcesVersion) And (ThisBuildFileVersion64 <> cLaunchExecute.AcesVersion) Then
                MyResults.InstallationRequred = True
                SuperviseInstallerResults = SuperviseInstaller()
                MyResults.Success = SuperviseInstallerResults.Success
                If MyResults.Success Then
                    MyResults.InstallationPath = SuperviseInstallerResults.InstallationPath
                Else
                    MyResults.Message = SuperviseInstallerResults.Message
                End If
            End If

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #649 (PerformVersionUpdate): " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Function SuperviseInstaller() As AcesInstallResults
        Dim MyResults As New AcesInstallResults
        Dim RunInstallerResults As AcesInstallResults
        Dim Message As String
        Dim Response1 As DialogResult
        Dim Response2 As DialogResult

        Try

            Do

                RunInstallerResults = RunInstaller()
                If RunInstallerResults.Success Then
                    MyResults.InstallationRequred = RunInstallerResults.InstallationRequred
                    MyResults.InstallationPath = RunInstallerResults.InstallationPath
                    Message = "This was a critical update.  Failure to install the update will result in system failure!" & vbCrLf & vbCrLf & "Did the installation COMPLETE SUCCESSFULLY?"
                    Response1 = MessageBox.Show(Message, "Critical update", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    Select Case Response1
                        Case DialogResult.Yes
                            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Enroller reporting successful installation.")
                            Exit Do
                        Case DialogResult.No
                            Message = "Would you like to retry the update installation?"
                            Response2 = MessageBox.Show(Message, "Critical update", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            Select Case Response2
                                Case DialogResult.Yes

                                    ' ___ Check the connection
                                    If Not cCommon.TestAndConnectToNetwork("launcher") Then
                                        RunInstallerResults.Success = False
                                        RunInstallerResults.Message = cCommon.NetworkConnectErrorMessage
                                        Exit Do
                                    End If

                                    ' Keep looping

                                Case DialogResult.No
                                    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Enroller reporting installation failure.")
                                    RunInstallerResults.Success = False
                                    RunInstallerResults.Message = "Enroller reporting installation failure."
                                    Exit Do
                            End Select
                    End Select
                End If

            Loop

            MyResults.Success = RunInstallerResults.Success
            If Not RunInstallerResults.Success Then
                MyResults.Message = RunInstallerResults.Message
            End If

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #650: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Function RunInstaller() As AcesInstallResults
        Dim MyResults As New AcesInstallResults
        Dim ProcessName As String
        Dim Counter As Integer
        Dim ThisBuildInfo As FileVersionInfo
        Dim ThisBuildFileVersion As String

        Try

            ' ___ Installer and ProcessName
            If Microsoft.VisualBasic.Right(cLaunchExecute.Installer, 3) = "msi" Then
                cLaunchExecute.Installer = "Setup.exe"
            End If
            ProcessName = cLaunchExecute.ServerPath & "\" & cLaunchExecute.Installer
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Starting ACES installation")

            ' //////////// Process A: Setup.exe ///////////////////////////////////////////////////////////////////////

            ' ___ Start process #1
            System.IO.Directory.SetCurrentDirectory(cLaunchExecute.ServerPath)
            Dim oProcess As System.Diagnostics.Process = New Process
            oProcess.StartInfo.FileName = ProcessName
            oProcess.StartInfo.UseShellExecute = False
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Launching process " & ProcessName)
            Application.DoEvents()
            oProcess.Start()

            ' ___ Loop #1: Wait for Setup.exe to start.
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Waiting for process to start.")
            Do
                System.Diagnostics.Debug.WriteLine("Loop #1: " & Counter.ToString)
                If Common.ProcessFound(ProcessName) Then
                    Exit Do
                End If
                System.Threading.Thread.Sleep(250)
                Counter += 250
            Loop
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Process started.")

            ' ___ Loop #2: Wait for Setup.exe to finish.
            Counter = 0
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Waiting for process to finish.")
            Application.DoEvents()
            Do
                System.Diagnostics.Debug.WriteLine("Loop #2: " & Counter.ToString)
                If Not Common.ProcessFound(ProcessName) Then
                    Exit Do
                End If
                System.Threading.Thread.Sleep(1000)
                Counter += 1000
            Loop
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Process finished.")

            ' ___ Loop #3: Wait until Main_App.exe is detected.
            Counter = 0
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Waiting to detect Main_App.exe.")
            Application.DoEvents()
            Do
                System.Diagnostics.Debug.WriteLine("Loop #3: " & Counter.ToString)

                ''''If cEnviro.OSPlatformBitCount = Enviro.OSPlatformBitCountEnum.BitCount64 AndAlso System.IO.Directory.Exists(cLaunchExecute.TableLocalPath64) Then
                ''''    If System.IO.File.Exists(cLaunchExecute.TableLocalPath64 & "\" & cLaunchExecute.FileName) Then
                ''''        'cLaunchExecute.LocalProg64Exists = True
                ''''        MyResults.InstallationPath = cLaunchExecute.TableLocalPath64
                ''''        Exit Do
                ''''    End If
                ''''End If

                ''''If System.IO.Directory.Exists(cLaunchExecute.TableLocalPath32) Then
                ''''    If System.IO.File.Exists(cLaunchExecute.TableLocalPath32 & "\" & cLaunchExecute.FileName) Then
                ''''        'cLaunchExecute.LocalProg32Exists = True
                ''''        MyResults.InstallationPath = cLaunchExecute.TableLocalPath32
                ''''        Exit Do
                ''''    End If
                ''''End If



                ' ___ First, check for 64-bit installation on a 64-bit system
                If cEnviro.OSPlatformBitCount = Enviro.OSPlatformBitCountEnum.BitCount64 Then
                    If System.IO.Directory.Exists(cLaunchExecute.TableListingForLocalPath64) AndAlso System.IO.File.Exists(cLaunchExecute.TableListingForLocalPath64 & "\" & cLaunchExecute.FileName) Then
                        ThisBuildInfo = FileVersionInfo.GetVersionInfo(cLaunchExecute.TableListingForLocalPath64 & "\" & cLaunchExecute.FileName)
                        ThisBuildFileVersion = ThisBuildInfo.FileVersion
                        If ThisBuildFileVersion = cLaunchExecute.AcesVersion Then
                            MyResults.InstallationPath = cLaunchExecute.TableListingForLocalPath64
                            Exit Do
                        End If
                    End If
                End If

                ' ___ Then check for a 32-bit installation on any system
                If System.IO.Directory.Exists(cLaunchExecute.TableListingForLocalPath32) AndAlso System.IO.File.Exists(cLaunchExecute.TableListingForLocalPath32 & "\" & cLaunchExecute.FileName) Then
                    ThisBuildInfo = FileVersionInfo.GetVersionInfo(cLaunchExecute.TableListingForLocalPath32 & "\" & cLaunchExecute.FileName)
                    ThisBuildFileVersion = ThisBuildInfo.FileVersion
                    If ThisBuildFileVersion = cLaunchExecute.AcesVersion Then
                        MyResults.InstallationPath = cLaunchExecute.TableListingForLocalPath32
                        Exit Do
                    End If
                End If


                System.Threading.Thread.Sleep(1000)
                Counter += 1000

            Loop
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Main_App.exe detected.")
            Application.DoEvents()

            'End If

            System.Diagnostics.Debug.WriteLine("Install ended: " & Date.Now.ToString)
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Installation is finished. Exiting ACES install.")
            Application.DoEvents()

            MyResults.InstallationRequred = True
            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #645: " & ex.Message
            Return MyResults
        End Try
    End Function

    'Private Sub RunInstaller2(ByVal ProcessName As String)

    ' DON'T DELETE THIS


    '    '' ___ Installer and ProcessName
    '    'Installer = cEnviro.GetValue(RowNum, Progs.Installer)
    '    'If Microsoft.VisualBasic.Right(Installer, 3) = "msi" Then
    '    '    Installer = "Setup.exe"
    '    'End If
    '    'ProcessName = cEnviro.GetValue(RowNum, Progs.ServerPath) & "\" & Installer
    '    ''Shell(cEnviro.dtProgs.Rows(RowNum)("ServerPath") & "\" & tInstaller, AppWinStyle.NormalFocus)

    '    ' ___ Start the process
    '    Dim oProcess As System.Diagnostics.Process = New Process
    '    'oProcess.StartInfo.FileName = cEnviro.GetValue(RowNum, Progs.ServerPath) & "\" & Installer
    '    oProcess.StartInfo.FileName = ProcessName
    '    'oProcess.StartInfo.Arguments = "/DB:" & Base.DBObj.ServerAddress & " /FN:" & VoiceFilename
    '    oProcess.StartInfo.UseShellExecute = False
    '    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Initiating installation.")
    '    oProcess.Start()

    '    ' ___ Loop until we find the proc.
    '    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Waiting for process to start.")
    '    Do
    '        System.Threading.Thread.Sleep(1000)
    '    Loop Until System.Diagnostics.Process.GetProcessesByName(ProcessName).GetUpperBound(0) <> -1
    '    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Installation is running.")

    '    ' ___ Loop until the process ends.
    '    Do
    '        System.Threading.Thread.Sleep(1000)
    '    Loop While System.Diagnostics.Process.GetProcessesByName(ProcessName).GetUpperBound(0) > -1
    '    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Installation has completed.")

    'End Sub


#End Region

#Region " Helpers "
    Private Function StripSlash(ByVal Value As String) As String
        Dim Results As String
        Results = Trim(Value)
        If Results.Length > 0 AndAlso Microsoft.VisualBasic.Right(Results, 1) = "\" Then
            Results = Results.Substring(0, Results.Length - 1)
        End If
        Return Results
    End Function
#End Region
End Class


'If cEnviro.CheckMSIEXEC_exe Then

'    ' //////////// Process B: msiexec.exe ///////////////////////////////////////////////////////////////////////

'    ' ___ 7/26/2010: Mike update the Aces installer to VS2008. It added a hond-off to a second process.

'    ' //////////// Process #2

'    ' ___ Start process #2
'    Counter = 0
'    ProcessName = "msiexec.exe"
'    ProcessName = "Windows\system32\msiexec.exe"
'    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Installer #2: " & ProcessName)
'    oProcess = New System.Diagnostics.Process
'    oProcess.StartInfo.FileName = ProcessName
'    oProcess.StartInfo.UseShellExecute = False
'    oProcess.Start()

'    ' ___ Loop #3: Wait for msiexec.exe to start.
'    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Waiting for process msiexec.exe to start.")
'    Do
'        If Common.ProcessFound(ProcessName) Then
'            Exit Do
'        End If
'        System.Threading.Thread.Sleep(250)

'        If Counter Mod 1000 = 0 Then
'            System.Diagnostics.Debug.WriteLine("Loop #3: " & Counter.ToString)
'        End If
'        Counter += 250
'    Loop

'    ' ___ Loop #4: Wait for msiexec.exe to finish.
'    Counter = 0
'    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Waiting for process msiexec.exe to finish.")
'    Do
'        If Not Common.ProcessFound(ProcessName) Then
'            Exit Do
'        End If
'        System.Threading.Thread.Sleep(1000)

'        If Counter Mod 10000 = 0 Then
'            System.Diagnostics.Debug.WriteLine("Loop #4: " & Counter.ToString)
'        End If
'        Counter += 1000
'    Loop
'End If


'If cEnviro.CheckMain_App_exe Then

' //////////// Now loop until Main_App.exe is detected  ///////////////////////////////////////////////////////////////////////
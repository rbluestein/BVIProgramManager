Imports System.IO
Imports System.DirectoryServices
'Imports Microsoft.Win32

Public Class FrmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents SSTab1 As System.Windows.Forms.TabControl
    Friend WithEvents tpAssistance As System.Windows.Forms.TabPage
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents lblUserLabel As System.Windows.Forms.Label
    Friend WithEvents lblUserText As System.Windows.Forms.Label
    Friend WithEvents lblSiteText As System.Windows.Forms.Label
    Friend WithEvents lblSiteLabel As System.Windows.Forms.Label
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents CheckTimer As System.Windows.Forms.Timer
    Friend WithEvents tpProduction As System.Windows.Forms.TabPage
    Friend WithEvents treTestProgs As System.Windows.Forms.TreeView
    Friend WithEvents btnTroubleCall As System.Windows.Forms.Button
    Friend WithEvents btnLaunchProg As System.Windows.Forms.Button
    Friend WithEvents btnViewLog As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents ScreenCapture As System.Windows.Forms.PictureBox
    Friend WithEvents treProductionProgs As System.Windows.Forms.TreeView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents tpTraining As System.Windows.Forms.TabPage
    Friend WithEvents treTrainingProgs As System.Windows.Forms.TreeView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmMain))
        Me.SSTab1 = New System.Windows.Forms.TabControl
        Me.tpProduction = New System.Windows.Forms.TabPage
        Me.treProductionProgs = New System.Windows.Forms.TreeView
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.tpAssistance = New System.Windows.Forms.TabPage
        Me.treTestProgs = New System.Windows.Forms.TreeView
        Me.tpTraining = New System.Windows.Forms.TabPage
        Me.treTrainingProgs = New System.Windows.Forms.TreeView
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.lblUserLabel = New System.Windows.Forms.Label
        Me.lblUserText = New System.Windows.Forms.Label
        Me.lblSiteText = New System.Windows.Forms.Label
        Me.lblSiteLabel = New System.Windows.Forms.Label
        Me.CheckTimer = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.btnTroubleCall = New System.Windows.Forms.Button
        Me.btnLaunchProg = New System.Windows.Forms.Button
        Me.btnViewLog = New System.Windows.Forms.Button
        Me.lblStatus = New System.Windows.Forms.Label
        Me.ScreenCapture = New System.Windows.Forms.PictureBox
        Me.SSTab1.SuspendLayout()
        Me.tpProduction.SuspendLayout()
        Me.tpAssistance.SuspendLayout()
        Me.tpTraining.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SSTab1
        '
        Me.SSTab1.Controls.Add(Me.tpProduction)
        Me.SSTab1.Controls.Add(Me.tpAssistance)
        Me.SSTab1.Controls.Add(Me.tpTraining)
        Me.SSTab1.Location = New System.Drawing.Point(0, 40)
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.SelectedIndex = 0
        Me.SSTab1.Size = New System.Drawing.Size(230, 364)
        Me.SSTab1.TabIndex = 0
        '
        'tpProduction
        '
        Me.tpProduction.Controls.Add(Me.treProductionProgs)
        Me.tpProduction.Location = New System.Drawing.Point(4, 22)
        Me.tpProduction.Name = "tpProduction"
        Me.tpProduction.Size = New System.Drawing.Size(222, 338)
        Me.tpProduction.TabIndex = 0
        Me.tpProduction.Text = "Production"
        '
        'treProductionProgs
        '
        Me.treProductionProgs.ImageList = Me.ImageList1
        Me.treProductionProgs.Location = New System.Drawing.Point(4, 4)
        Me.treProductionProgs.Name = "treProductionProgs"
        Me.treProductionProgs.Size = New System.Drawing.Size(212, 332)
        Me.treProductionProgs.TabIndex = 3
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'tpAssistance
        '
        Me.tpAssistance.Controls.Add(Me.treTestProgs)
        Me.tpAssistance.Location = New System.Drawing.Point(4, 22)
        Me.tpAssistance.Name = "tpAssistance"
        Me.tpAssistance.Size = New System.Drawing.Size(222, 338)
        Me.tpAssistance.TabIndex = 1
        Me.tpAssistance.Text = "   Test   "
        '
        'treTestProgs
        '
        Me.treTestProgs.ImageList = Me.ImageList1
        Me.treTestProgs.Location = New System.Drawing.Point(4, 4)
        Me.treTestProgs.Name = "treTestProgs"
        Me.treTestProgs.Size = New System.Drawing.Size(212, 332)
        Me.treTestProgs.TabIndex = 2
        '
        'tpTraining
        '
        Me.tpTraining.Controls.Add(Me.treTrainingProgs)
        Me.tpTraining.Location = New System.Drawing.Point(4, 22)
        Me.tpTraining.Name = "tpTraining"
        Me.tpTraining.Size = New System.Drawing.Size(222, 338)
        Me.tpTraining.TabIndex = 2
        Me.tpTraining.Text = "Training"
        '
        'treTrainingProgs
        '
        Me.treTrainingProgs.ImageList = Me.ImageList1
        Me.treTrainingProgs.Location = New System.Drawing.Point(4, 4)
        Me.treTrainingProgs.Name = "treTrainingProgs"
        Me.treTrainingProgs.Size = New System.Drawing.Size(212, 332)
        Me.treTrainingProgs.TabIndex = 3
        '
        'lblUserLabel
        '
        Me.lblUserLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserLabel.Location = New System.Drawing.Point(7, 4)
        Me.lblUserLabel.Name = "lblUserLabel"
        Me.lblUserLabel.Size = New System.Drawing.Size(32, 16)
        Me.lblUserLabel.TabIndex = 1
        Me.lblUserLabel.Text = "User:"
        '
        'lblUserText
        '
        Me.lblUserText.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserText.Location = New System.Drawing.Point(36, 4)
        Me.lblUserText.Name = "lblUserText"
        Me.lblUserText.Size = New System.Drawing.Size(180, 16)
        Me.lblUserText.TabIndex = 2
        Me.lblUserText.Text = "User:"
        '
        'lblSiteText
        '
        Me.lblSiteText.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSiteText.Location = New System.Drawing.Point(36, 20)
        Me.lblSiteText.Name = "lblSiteText"
        Me.lblSiteText.Size = New System.Drawing.Size(184, 16)
        Me.lblSiteText.TabIndex = 4
        Me.lblSiteText.Text = "Site"
        '
        'lblSiteLabel
        '
        Me.lblSiteLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSiteLabel.Location = New System.Drawing.Point(7, 20)
        Me.lblSiteLabel.Name = "lblSiteLabel"
        Me.lblSiteLabel.Size = New System.Drawing.Size(32, 16)
        Me.lblSiteLabel.TabIndex = 3
        Me.lblSiteLabel.Text = "Site:"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnRefresh)
        Me.Panel1.Controls.Add(Me.btnTroubleCall)
        Me.Panel1.Controls.Add(Me.btnLaunchProg)
        Me.Panel1.Controls.Add(Me.btnViewLog)
        Me.Panel1.Controls.Add(Me.lblStatus)
        Me.Panel1.Controls.Add(Me.ScreenCapture)
        Me.Panel1.Location = New System.Drawing.Point(0, 380)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(224, 84)
        Me.Panel1.TabIndex = 6
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(144, 24)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(81, 24)
        Me.btnRefresh.TabIndex = 21
        Me.btnRefresh.Text = "Refresh"
        '
        'btnTroubleCall
        '
        Me.btnTroubleCall.Location = New System.Drawing.Point(1, 24)
        Me.btnTroubleCall.Name = "btnTroubleCall"
        Me.btnTroubleCall.Size = New System.Drawing.Size(80, 24)
        Me.btnTroubleCall.TabIndex = 20
        Me.btnTroubleCall.Text = "Trouble Call"
        '
        'btnLaunchProg
        '
        Me.btnLaunchProg.Location = New System.Drawing.Point(1, 0)
        Me.btnLaunchProg.Name = "btnLaunchProg"
        Me.btnLaunchProg.Size = New System.Drawing.Size(225, 24)
        Me.btnLaunchProg.TabIndex = 19
        Me.btnLaunchProg.Text = "Launch Program"
        '
        'btnViewLog
        '
        Me.btnViewLog.Location = New System.Drawing.Point(80, 24)
        Me.btnViewLog.Name = "btnViewLog"
        Me.btnViewLog.Size = New System.Drawing.Size(64, 24)
        Me.btnViewLog.TabIndex = 18
        Me.btnViewLog.Text = "Log"
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(4, 50)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(208, 26)
        Me.lblStatus.TabIndex = 17
        '
        'ScreenCapture
        '
        Me.ScreenCapture.Location = New System.Drawing.Point(92, 36)
        Me.ScreenCapture.Name = "ScreenCapture"
        Me.ScreenCapture.Size = New System.Drawing.Size(56, 32)
        Me.ScreenCapture.TabIndex = 16
        Me.ScreenCapture.TabStop = False
        Me.ScreenCapture.Visible = False
        '
        'FrmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(226, 466)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lblSiteText)
        Me.Controls.Add(Me.lblSiteLabel)
        Me.Controls.Add(Me.lblUserText)
        Me.Controls.Add(Me.lblUserLabel)
        Me.Controls.Add(Me.SSTab1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(234, 700)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(234, 380)
        Me.Name = "FrmMain"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.SSTab1.ResumeLayout(False)
        Me.tpProduction.ResumeLayout(False)
        Me.tpAssistance.ResumeLayout(False)
        Me.tpTraining.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Declarations "
    Private cDBase As New DBase
    Private cLauncher As New Launcher
#End Region

#Region " Events "
    Private Sub ExitApplication()
        '  Application.Exit()
        Environment.Exit(0)
    End Sub


#End Region

    'Private Function ORIGINAL_TestNetworkConnection() As Results
    '    Dim MyResults As New Results
    '    Dim DirInfo As DirectoryInfo

    '    Try
    '        DirInfo = New DirectoryInfo(cEnviro.ProgMgrServerPath)
    '        If DirInfo.Exists Then
    '            MyResults.Success = True
    '        Else
    '            MyResults.Success = False
    '        End If

    '        Return MyResults
    '    Catch
    '        MyResults.Success = False
    '        Return MyResults
    '    End Try
    'End Function

    'Private Function Testie() As Results
    '    Dim MyResults As New Results
    '    Dim FileInfo As System.IO.FileInfo
    '    Dim FileDays As Integer
    '    Dim NowDays As Integer
    '    Dim LogFileFullPath As String = "C:\Apps\_ProductionApps\ProgramManager\BVIPM_Local\bin\rbluestein.log"

    '    Try

    '        ' ___ Delete the log file if it is more than 10 days old..
    '        If System.IO.File.Exists(LogFileFullPath) Then

    '            FileInfo = New System.IO.FileInfo(LogFileFullPath)
    '            FileDays = FileInfo.CreationTimeUtc.Subtract(#1/1/2007#).Days()
    '            NowDays = Date.Now.Subtract(#1/1/2007#).Days()
    '            If NowDays - FileDays > 9 Then
    '                FileInfo.Delete()
    '            End If

    '        End If

    '        MyResults.Success = True
    '        ' Return MyResults

    '        Try
    '            ' cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "** Launching BVI Program Manager  **")
    '        Catch
    '        End Try

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #4888: BVIPM_Local.exe is unable to clear log " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    'Private Sub GGetInfoFromRegistry()

    '    '___ Open Register Key
    '    Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion")

    '    '___ Get System Root e.g. C:\Windows 
    '    Dim SysRoot As String = rk.GetValue("SystemRoot").ToString()

    '    '___ Get driver where current Operatiing System is installed e.g. C:\
    '    Dim drive As String = New IO.DirectoryInfo(SysRoot).Root.ToString

    '    '___ You also can get other information (such as Version,Name) of current Operatiing System
    '    Dim OSName As String = rk.GetValue("ProductName").ToString() 'Get Operating System name
    '    Dim ServicePack As String = rk.GetValue("CSDVersion").ToString() 'Get Operating System ServicePack Number
    '    Dim SerialNo As String = rk.GetValue("ProductId").ToString() ' Get OS Serial Number
    '    rk.Close()
    'End Sub

    'Private Sub GetInfoFromRegistry()

    '    '___ Open Register Key
    '    Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion")

    '    Dim SysRoot As String = rk.GetValue("SystemRoot").ToString()         'System Root 
    '    Dim drive As String = New IO.DirectoryInfo(SysRoot).Root.ToString  ' OS root
    '    Dim OSName As String = rk.GetValue("ProductName").ToString()     'OS name
    '    Dim ServicePack As String = rk.GetValue("CSDVersion").ToString()  'OS ServicePack Number
    '    Dim SerialNo As String = rk.GetValue("ProductId").ToString()           ' OS Serial Number
    '    rk.Close()
    'End Sub

#Region " Form load events "
    Private Sub FrmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim MyResults As New Results
        Dim DisallowRunningMultipleInstancesResults As Results
        Dim UpdateLauncherResults As Results
        Dim ClearLogResults As Results
        Dim EnviroInitResults As Results
        Dim DBDataResults As Results
        Dim ProgsResults As Results
        Dim PopulateTreesResults As Results
        Dim RollingSuccess As Boolean = True
        AddHandler cCommon.ExitApplication, AddressOf ExitApplication
        AddHandler cAppAdmin.ExitApplication, AddressOf ExitApplication

        Try

            cAppAdmin.SetlblStatus(lblStatus)

            DisallowRunningMultipleInstancesResults = cCommon.DisallowRunningMultipleInstances()
            If Not DisallowRunningMultipleInstancesResults.Success Then
                RollingSuccess = False
                MyResults.Message = DisallowRunningMultipleInstancesResults.Message
            End If

            'Me.Text = GetPageText()

            Common.ResizeControls(Me, 234, 525)
            ' AddHandler cLauncher.UpdateStatusEvent, AddressOf UpdateStatus
            lblStatus.Font = New Font(treProductionProgs.Font.Name, 8, FontStyle.Regular)

            ' ___ Set the ServerPath
            cEnviro.ProgMgrServerPath = "V:\BVIPM"

            ' ___ Test the network connection
            If RollingSuccess Then
                ' TestNetworkConnectionResults = TestNetworkConnection()
                If Not cCommon.TestAndConnectToNetwork("main") Then
                    RollingSuccess = False
                    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.MessageBoxAndShutdown, cCommon.NetworkConnectErrorMessage, "Network Error")
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

            ' ___ Clear out  log
            If RollingSuccess Then
                ClearLogResults = ClearLog()
                If Not ClearLogResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = ClearLogResults.Message
                End If
            End If


            ' ___ Update the launcher
            If Not Enviro.DebugMode Then
                If RollingSuccess Then
                    UpdateLauncherResults = UpdateLauncher()
                    If Not UpdateLauncherResults.Success Then
                        RollingSuccess = False
                        MyResults.Message = UpdateLauncherResults.Message
                    End If
                End If
            End If

            ' ___ Record app launch
            Try
                Me.Text = "BVI Prog Mgr v" & cEnviro.FormattedFileVersion(cEnviro.LocalLocalFileVersion)
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "** Launching " & Application.ExecutablePath & " Version " & cEnviro.FormattedFileVersion(cEnviro.LocalLocalFileVersion) & "  **")
            Catch
            End Try

            '' ___ Update components
            'If RollingSuccess Then
            '    UpdateComponentsResults = UpdateComponents()
            '    If Not UpdateComponentsResults.Success Then
            '        RollingSuccess = False
            '        MyResults.Message = UpdateComponentsResults.Message
            '    End If
            'End If

            ' ___ Delete orphan screen capture files
            If RollingSuccess Then
                Try
                    DeleteOrphanScreenshots()
                Catch
                End Try
            End If

            ' ___ Read ini file
            'If RollingSuccess Then
            '    ReadIniResults = ReadIni()
            '    If Not ReadIniResults.Success Then
            '        RollingSuccess = False
            '        MyResults.Message = ReadIniResults.Message
            '    End If
            'End If

            ' ___ Get db data
            If RollingSuccess Then
                DBDataResults = GetDBData()
                If Not DBDataResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = DBDataResults.Message
                End If
            End If

            ' ___ Get progs
            If RollingSuccess Then
                ProgsResults = GetProgs()
                If Not ProgsResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = ProgsResults.Message
                End If
            End If

            ' ___ Populate trees
            If RollingSuccess Then
                If cEnviro.dtProgs.Rows.Count > 0 Then
                    PopulateTreesResults = PopulateTrees()
                    If Not PopulateTreesResults.Success Then
                        RollingSuccess = False
                        MyResults.Message = PopulateTreesResults.Message
                    End If
                End If
            End If

            If RollingSuccess Then
                lblUserText.Text = IIf(cEnviro.UserFirstName = Nothing, String.Empty, cEnviro.UserFirstName) & " " & IIf(cEnviro.UserLastName = Nothing, String.Empty, cEnviro.UserLastName)
                lblSiteText.Text = cEnviro.SiteId
                If cEnviro.dtProgs.Rows.Count = 0 Then
                    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.InformationNoLog, "No apps available")
                Else
                    cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.InformationNoLog, "Please select a program to launch")
                End If
            Else
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #330b: " & MyResults.Message)
            End If

            ' lblStatus.Text = "Four score and seven years ago our fathers brought forth on this continent a new nation conceived in liberty and dedicated to the proposition that men were created equal."

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #330a: " & ex.Message)
        End Try
    End Sub

    'Private Function GetPageText() As String

    '    If cEnviro.DebugMode Then
    '        Return cEnviro.DebugModeVersion
    '    End If

    '    Dim FileVersion As String
    '    Dim LauncherLocalBuildInfo As FileVersionInfo
    '    LauncherLocalBuildInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_Local.exe")
    '    FileVersion = LauncherLocalBuildInfo.FileVersion
    '    FileVersion = cEnviro.FormattedFileVersion(FileVersion)
    '    Return "BVI Prog Mgr " & FileVersion
    'End Function

    Private Function ClearLog() As Results
        Dim MyResults As New Results
        Dim FileInfo As System.IO.FileInfo
        Dim FileDays As Integer
        Dim NowDays As Integer

        Try

            '   cEnviro.LogFileFullPath = "C:\Apps\_ProductionApps\ProgramManager\BVIPM_Local\bin\rbluestein.log"

            ' ___ Delete the log file if it is more than 10 days old..
            If System.IO.File.Exists(cEnviro.LogFileFullPath) Then

                FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)
                FileDays = FileInfo.CreationTimeUtc.Subtract(#1/1/2007#).Days()
                NowDays = Date.Now.Subtract(#1/1/2007#).Days()

                If NowDays - FileDays > 9 Then
                    Dim FileStream As New System.IO.FileStream(cEnviro.LogFileFullPath, FileMode.Create, FileAccess.Write, FileShare.None)
                    'Dim Writer As New System.IO.StreamWriter(FileStream)
                    'Writer.WriteLine("")
                    'Writer.Flush()
                    'Writer.Close()
                    FileStream.Close()
                    FileInfo.CreationTime = Date.Now.ToUniversalTime.AddHours(-5)
                End If

            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #4888: BVIPM_Local.exe is unable to clear log " & ex.Message
            Return MyResults
        End Try
    End Function

    'Private Function OldClearLog() As Results
    '    Dim MyResults As New Results
    '    Dim FileInfo As System.IO.FileInfo
    '    Dim FileDays As Integer
    '    Dim NowDays As Integer

    '    Try

    '        ' ___ Delete the log file if it is more than 10 days old..
    '        If System.IO.File.Exists(cEnviro.LogFileFullPath) Then

    '            FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)
    '            FileDays = FileInfo.CreationTimeUtc.Subtract(#1/1/2007#).Days()
    '            NowDays = Date.Now.Subtract(#1/1/2007#).Days()
    '            If NowDays - FileDays > 9 Then

    '                Do
    '                    FileInfo.Delete()
    '                    Application.DoEvents()
    '                    Threading.Thread.Sleep(1000)
    '                    If Not System.IO.File.Exists(cEnviro.LogFileFullPath) Then
    '                        Exit Do
    '                    End If
    '                Loop

    '            End If

    '        End If

    '        MyResults.Success = True
    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #4888: BVIPM_Local.exe is unable to clear log " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    Private Function UpdateLauncher() As Results
        Dim MyResults As New Results
        Dim CompareVersionsResults As Results
        Dim DownloadFileResults As Results
        Dim UpdateRequired As Boolean

        Try

            ' ___ First determine whether the app can connect to the network directory
            If Not cCommon.TestAndConnectToNetwork("main") Then
                MyResults.Success = False
                MyResults.Message = "Error 4832a: " & cCommon.NetworkConnectErrorMessage
                Return MyResults
            End If

            ' ___ Compare versions
            CompareVersionsResults = CompareVersions(UpdateRequired)
            If Not CompareVersionsResults.Success Then
                MyResults.Success = False
                MyResults.Message = CompareVersionsResults.Message
                Return MyResults
            End If

            ' ___ Update as appropriate
            If UpdateRequired Then
                DownloadFileResults = DownloadFile()
                If Not DownloadFileResults.Success Then
                    MyResults.Success = False
                    MyResults.Message = DownloadFileResults.Message
                    Return MyResults
                End If
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #4832b: " & ex.Message
            Return MyResults
        End Try
    End Function

    'Private Function ContactNetwork() As Results
    '    Dim MyResults As New Results
    '    Dim DirInfo As System.IO.DirectoryInfo

    '    Try

    '        DirInfo = New System.IO.DirectoryInfo(cEnviro.ProgMgrServerPath)
    '        If Not DirInfo.Exists Then
    '            MyResults.Success = False
    '            MyResults.Message = "BVIPM_Local.exe is unable to establish contact with network directory"
    '            Return MyResults
    '        End If

    '        MyResults.Success = True
    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "BVIPM_Local.exe is unable to establish contact with network directory"
    '        Return MyResults
    '    End Try
    'End Function

    Public Function CompareVersions(ByRef UpdateRequired As Boolean) As Results
        Dim MyResults As New Results

        Try

            If cEnviro.LauncherLocalFileVersion <> cEnviro.LauncherMasterFileVersion Then
                UpdateRequired = True
            End If

            ''Dim MasterBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(cEnviro.ProgMgrServerPath & "\BVIPM.exe")
            ''Dim ThisBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM.exe")

            ''cEnviro.MasterFileVersion = MasterBuildInfo.FileVersion
            ''cEnviro.ThisFileVersion = ThisBuildInfo.FileVersion


            'cEnviro.MasterVersion = MasterBuildInfo.FileMajorPart & "." & MasterBuildInfo.FileMinorPart
            'cEnviro.CurrentVersion = ThisBuildInfo.FileMajorPart & "." & ThisBuildInfo.FileMinorPart

            'If CInt(MasterBuildInfo.FileMajorPart) > CInt(ThisBuildInfo.FileMajorPart) Then
            '    UpdateRequired = True
            'Else
            '    If CInt(MasterBuildInfo.FileMajorPart) = CInt(ThisBuildInfo.FileMajorPart) Then
            '        If CInt(MasterBuildInfo.FileMinorPart) > CInt(ThisBuildInfo.FileMinorPart) Then
            '            UpdateRequired = True
            '        End If
            '    End If
            'End If

            'If CInt(MasterBuildInfo.FileMajorPart) <> CInt(ThisBuildInfo.FileMajorPart) Then
            '    UpdateRequired = True
            'ElseIf CInt(MasterBuildInfo.FileMinorPart) <> CInt(ThisBuildInfo.FileMinorPart) Then
            '    UpdateRequired = True
            'End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #647: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Function DownloadFile()
        Dim MyResults As New Results
        Dim ProcessName As String
        Dim ElapsedTime As Integer
        Dim OKToDownload As Boolean

        Try

            ' ___ Allow Windows up to 30 seconds to close the calling instance of the launcher
            ProcessName = Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM.exe"
            Do
                If Common.ProcessFound(ProcessName) Then
                    System.Threading.Thread.Sleep(250)
                Else
                    OKToDownload = True
                    Exit Do
                End If
                ElapsedTime += 250
                If ElapsedTime >= 30000 Then
                    Exit Do
                End If
            Loop

            If Not OKToDownload Then
                MyResults.Success = False
                MyResults.Message = "Error #653a: BVIPM.exe timeout elapsed prior to BVIPM.exe download."
                Return MyResults
            End If

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #653b: " & ex.Message
            Return MyResults
        End Try


        ' ___ Proceed with download
        If Not cCommon.TestAndConnectToNetwork("main") Then
            MyResults.Success = False
            MyResults.Message = "Error #653c: " & cCommon.NetworkConnectErrorMessage
            Return MyResults
        End If

        Try

            If OKToDownload Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Performing update of BVIPM.exe from version " & cEnviro.FormattedFileVersion(cEnviro.LauncherLocalFileVersion) & " to " & cEnviro.FormattedFileVersion(cEnviro.LauncherMasterFileVersion))
                Application.DoEvents()
                System.IO.File.Copy(cEnviro.ProgMgrServerPath & "\BVIPM.exe", Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM.exe", True)
                Application.DoEvents()

                'ElapsedTime = 0
                'Do
                '    Try
                '        Shell(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM.exe", AppWinStyle.NormalFocus)
                '        Exit Do
                '    Catch
                '        System.Threading.Thread.Sleep(250)
                '        ElapsedTime += 250
                '    End Try
                '    If ElapsedTime >= 30000 Then
                '        MyResults.Success = False
                '        MyResults.Message = "Error #651: Timeout occurred between download and execution of BVIPM.exe."
                '        Return MyResults
                '    End If
                'Loop

                'ExitApplication()

            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #653d: Update of launcher failed. " & ex.message
            'cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Error #648: Unable to perform update of BVIPM.exe from version " & cEnviro.FormattedFileVersion(cEnviro.LauncherLocalFileVersion) & " to " & cEnviro.FormattedFileVersion(cEnviro.LauncherMasterFileVersion))
            Return MyResults
        End Try
    End Function

    'Private Function IsProcessFound(ByVal ProcessName As String) As Boolean
    '    Dim i As Integer

    '    Try
    '        Dim CurrentProcesses As Process() = Process.GetProcesses()
    '        For i = 0 To CurrentProcesses.GetUpperBound(0)
    '            Try
    '                If CurrentProcesses(i).MainModule.FileName.ToLower = ProcessName.ToLower Then
    '                    Return True
    '                End If
    '            Catch
    '            End Try
    '        Next
    '        Return False
    '    Catch ex As Exception
    '        Return False
    '    End Try
    'End Function

    'Private Function GetVersion(ByVal FileFullPath) As String
    '    Dim BuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(FileFullPath)
    '    Return BuildInfo.FileMajorPart & "." & BuildInfo.FileMinorPart
    'End Function

    'Private Function UpdateComponents() As Results
    ''     moved to bvipm.exe
    '    Dim MyResults As New Results

    '    Try

    '        If Not System.IO.File.Exists(Path.GetDirectoryName(Application.ExecutablePath) & "\Interop.ADODB.dll") Then
    '            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Downloading Interop.ADODB.dll.")
    '            System.IO.File.Copy(cAppSettings.ServerPath & "\Interop.ADODB.dll", Path.GetDirectoryName(Application.ExecutablePath) & "\Interop.ADODB.dll.", True)
    '        End If
    '        If Not System.IO.File.Exists(Path.GetDirectoryName(Application.ExecutablePath) & "\Interop.CDO.dll") Then
    '            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Downloading Interop.CDO.dll.")
    '            System.IO.File.Copy(cAppSettings.ServerPath & "\Interop.CDO.dll", Path.GetDirectoryName(Application.ExecutablePath) & "\Interop.CDO.dll", True)
    '        End If

    '        MyResults.Success = True
    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #341: " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    Private Sub DeleteOrphanScreenshots()
        Dim i As Integer
        Dim TestString As String
        Dim Name As String

        TestString = "ScreenCapture_" & Environment.UserName & "_"
        TestString = TestString.ToLower
        Dim DirInfo As New DirectoryInfo(Path.GetDirectoryName(Application.ExecutablePath))
        Dim FileInfo As FileInfo()
        FileInfo = DirInfo.GetFiles
        For i = 0 To FileInfo.GetUpperBound(0)
            Name = FileInfo(i).Name
            If Name.Length > TestString.Length Then
                If Name.Substring(0, TestString.Length).ToLower = TestString Then
                    File.Delete(FileInfo(i).FullName)
                End If
            End If
        Next
    End Sub

    Public Function EnviroInit() As Results
        Dim MyResults As New Results
        Dim OSName As String

        Try

            '' ___ OSPlatformBitCount
            'If Environment.OSVersion.Platform = PlatformID.Win32NT Then
            '    cEnviro.OSPlatformBitCount = 32
            'Else
            '    cEnviro.OSPlatformBitCount = 64
            'End If

            '___ Open Register Key        
            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion")
            OSName = rk.GetValue("ProductName").ToString().ToLower
            If InStr(OSName, "xp") > 0 Then
                cEnviro.OSPlatformBitCount = Enviro.OSPlatformBitCountEnum.BitCount32
            ElseIf InStr(OSName, "windows 7") > 0 Then
                cEnviro.OSPlatformBitCount = Enviro.OSPlatformBitCountEnum.BitCount64
            Else
                cEnviro.OSPlatformBitCount = Enviro.OSPlatformBitCountEnum.BitCount32
            End If
            rk.Close()















            cEnviro.UserId = Environment.UserName
            cEnviro.ComputerId = Environment.MachineName.ToUpper
            cEnviro.SiteId = "HBG"
            'Select Case cEnviro.ComputerId.Substring(0, 2)
            '    Case "HB" : cEnviro.SiteId = "HBG"
            '    Case "OK" : cEnviro.SiteId = "OKC"
            '    Case Else : cEnviro.SiteId = "HBG"
            'End Select
            'cEnviro.ProgMgrServerPath = "V:\BVIPM"
            cEnviro.LogFileFullPath = Path.GetDirectoryName(Application.ExecutablePath) & "\" & cEnviro.UserId & ".log"

            ' ___ Version
            If Not Enviro.DebugMode Then

                ' ___ First determine whether the app can connect to the network directory
                If Not cCommon.TestAndConnectToNetwork("main") Then
                    MyResults.Success = False
                    MyResults.Message = "Error 331a: " & cCommon.NetworkConnectErrorMessage
                    Return MyResults
                End If

                Dim LauncherMasterBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(cEnviro.ProgMgrServerPath & "\BVIPM.exe")
                cEnviro.LauncherMasterFileVersion = LauncherMasterBuildInfo.FileVersion
                Dim LauncherLocalBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM.exe")
                cEnviro.LauncherLocalFileVersion = LauncherLocalBuildInfo.FileVersion
                Dim LocalLocalBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.GetDirectoryName(Application.ExecutablePath) & "\BVIPM_Local.exe")
                cEnviro.LocalLocalFileVersion = LocalLocalBuildInfo.FileVersion
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #331b: " & ex.Message
            Return MyResults
        End Try
    End Function

    Public Function GetDBData() As Results
        Dim MyResults As New Results
        Dim QueryPack As DBase.QueryPack
        Dim dt As DataTable
        Dim Sql As String

        Try

            '' ___ Is this a test user?
            'Sql = "SELECT COUNT (*)  FROM ProgramConfig WHERE ProgramID = 'BVIProgramManager' AND ParamName = 'TestUser' AND ParamValue = '" & cEnviro.UserId & "' AND ActiveInd = 'Y'"
            'QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString, False)
            'If QueryPack.Success Then
            '    dt = QueryPack.dt
            '    MyResults.Success = True
            'Else
            '    MyResults.Success = False
            '    MyResults.Message = "Error #335 (GetDBData: get test user): " & QueryPack.TechErrMsg
            '    Return MyResults
            'End If

            'If dt.Rows(0)(0) > 0 Then
            '    cEnviro.TestMode = True
            'End If

            '' ___ If running in test mode, get the test server.
            'If cEnviro.TestMode Then
            '    Sql = "SELECT ParamValue FROM ProgramConfig WHERE ProgramID = 'BVIProgramManager' AND ParamName = 'TestServer' AND ActiveInd = 'Y'"
            '    QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString, False)
            '    If QueryPack.Success Then
            '        dt = QueryPack.dt
            '        MyResults.Success = True
            '    Else
            '        MyResults.Success = False
            '        MyResults.Message = "Error #336 (GetDBData: get test server): " & QueryPack.TechErrMsg
            '        Return MyResults
            '    End If

            '    If dt.Rows.Count = 0 Then
            '        MyResults.Success = False
            '        MyResults.Message = "Error #337: GetDBData: Test user listed but test server not listed in ProgramConfig"
            '        Return MyResults
            '    Else
            '        cEnviro.TestServer = dt.Rows(0)("ParamValue")
            '    End If
            'End If

            ' ___ Get user name
            'Sql = "SELECT FirstName, LastName FROM BVIUser WHERE StatusCode = 'A' AND UserID = '" & cEnviro.UserId & "'"
            Sql = "SELECT FirstName, LastName FROM UserManagement..Users WHERE StatusCode = 'ACTIVE' AND UserID = '" & cEnviro.UserId & "'"
            QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString, True)
            If QueryPack.Success Then
                dt = QueryPack.dt
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Message = "Error #338 (GetDBData: get user name): " & QueryPack.TechErrMsg
                Return MyResults
            End If

            If dt.Rows.Count = 0 Then
                MyResults.Success = False
                MyResults.Message = "Error #339: GetDBData: get user name. User record not returned from UserManagement."
                Return MyResults
            Else
                cEnviro.UserFirstName = dt.Rows(0)("FirstName").ToText
                cEnviro.UserLastName = dt.Rows(0)("LastName").ToText
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #332: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Function GetProgs() As Results
        Dim MyResults As New Results
        Dim Sql As String
        Dim QueryPack As DBase.QueryPack

        Try

            Sql = "SELECT * FROM Program WHERE PMAccessInd='Y' ORDER BY ClientID, Description"
            QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString, False)
            If QueryPack.Success Then
                cEnviro.dtProgs = QueryPack.dt
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Message = "Error #340 (GetDBData): " & QueryPack.TechErrMsg
                Return MyResults
            End If

            cEnviro.dtProgs.TableName = "Progs"

            'Select ProgramID, [Description], ProgramTypeCode, ClientID, CurrVersion, [FileName], ServerPath, LocalPath, PMAccessInd, RunLocal, Installer 

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #333: " & ex.Message
            Return MyResults
        End Try
    End Function


    'Private Function GetProgs() As Results
    '    Dim i As Integer
    '    Dim MyResults As New Results
    '    Dim Sql As String
    '    Dim QueryPack As DBase.QueryPack
    '    Dim dtAll As DataTable
    '    Dim dtProd As DataTable
    '    Dim dtTrain As DataTable
    '    Dim ItemArray As Object
    '    Dim dr As DataRow

    '    Try

    '        ' ___ Test
    '        Sql = "SELECT * FROM Program WHERE PMAccessInd='Y' AND CharIndex('TEST', ProgramID) > 0 ORDER BY ClientID, Description"
    '        QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString(Launcher.TestTrainProdEnum.Test), False)
    '        If QueryPack.Success Then
    '            dtAll = QueryPack.dt
    '            MyResults.Success = True
    '        Else
    '            MyResults.Success = False
    '            MyResults.Message = "Error #340 (GetDBData): " & QueryPack.TechErrMsg
    '            Return MyResults
    '        End If

    '        ' ___ Prod
    '        Sql = "SELECT * FROM Program WHERE PMAccessInd='Y' AND (CharIndex('TEST', ProgramID) = 0 AND CharIndex('TRAIN', ProgramID) = 0) ORDER BY ClientID, Description"
    '        QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString(Launcher.TestTrainProdEnum.Prod), False)
    '        If QueryPack.Success Then
    '            dtProd = QueryPack.dt
    '            MyResults.Success = True
    '        Else
    '            MyResults.Success = False
    '            MyResults.Message = "Error #340 (GetDBData): " & QueryPack.TechErrMsg
    '            Return MyResults
    '        End If

    '        For i = 0 To dtProd.Rows.Count - 1
    '            ItemArray = dtProd.Rows(i).ItemArray
    '            dr = dtAll.NewRow
    '            dr.ItemArray = ItemArray
    '            dtAll.Rows.Add(dr)
    '        Next

    '        ' ___ Train
    '        Sql = "SELECT * FROM Program WHERE PMAccessInd='Y' AND CharIndex('TRAIN', ProgramID) > 0 ORDER BY ClientID, Description"
    '        QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString(Launcher.TestTrainProdEnum.Training), False)
    '        If QueryPack.Success Then
    '            dtTrain = QueryPack.dt
    '            MyResults.Success = True
    '        Else
    '            MyResults.Success = False
    '            MyResults.Message = "Error #340 (GetDBData): " & QueryPack.TechErrMsg
    '            Return MyResults
    '        End If

    '        For i = 0 To dtProd.Rows.Count - 1
    '            ItemArray = dtTrain.Rows(i).ItemArray
    '            dr = dtAll.NewRow
    '            dr.ItemArray = ItemArray
    '            dtAll.Rows.Add(dr)
    '        Next

    '        cEnviro.dtProgs = dtAll
    '        cEnviro.dtProgs.TableName = "Progs"

    '        'Select ProgramID, [Description], ProgramTypeCode, ClientID, CurrVersion, [FileName], ServerPath, LocalPath, PMAccessInd, RunLocal, Installer 

    '        MyResults.Success = True
    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #333: " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    Private Function PopulateTrees() As Results
        Dim MyResults As New Results
        Dim i As Integer
        Dim ProductionParentNode As TreeNode = Nothing
        Dim ProductionChildNode As TreeNode = Nothing
        Dim ProductionCurClient As String = String.Empty
        Dim TestParentNode As TreeNode = Nothing
        Dim TestChildNode As TreeNode = Nothing
        Dim TrainParentNode As TreeNode = Nothing
        Dim TrainChildNode As TreeNode = Nothing
        Dim ProgramID As String

        Try
            'treProgs.Font = New Font(fnt.Name, 8, FontStyle.Regular)
            treProductionProgs.Nodes.Clear()
            treTestProgs.Nodes.Clear()
            treTrainingProgs.Nodes.Clear()

            For i = 0 To cEnviro.dtProgs.Rows.Count - 1
                ProgramID = cEnviro.GetValue(i, Progs.ProgramID).Trim.ToUpper
                If InStr(ProgramID, "TEST") > 0 Then
                    PopulateOtherTrees(treTestProgs, i, TestParentNode, TestChildNode)
                ElseIf InStr(ProgramID, "TRAIN") > 0 Then
                    PopulateOtherTrees(treTrainingProgs, i, TrainParentNode, TrainChildNode)
                Else
                    PopulateProductionTree(i, ProductionParentNode, ProductionChildNode, ProductionCurClient)
                End If
            Next

            treTestProgs.ExpandAll()
            treTrainingProgs.ExpandAll()

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #334: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Sub PopulateProductionTree(ByVal i As Integer, ByRef ParentNode As TreeNode, ByRef ChildNode As TreeNode, ByRef CurClient As String)
        Dim NewClient As String

        NewClient = cEnviro.GetValue(i, Progs.ClientID).Trim

        If NewClient <> CurClient Then
            CurClient = NewClient

            ParentNode = New TreeNode(CurClient)
            ParentNode.NodeFont = New Font(treProductionProgs.Font.Name, 8, FontStyle.Regular)
            ParentNode.ImageIndex = 0
            treProductionProgs.Nodes.Add(ParentNode)

        End If

        ChildNode = New TreeNode(cEnviro.GetValue(i, Progs.Description))
        ChildNode.ImageIndex = GetImageIndex(cEnviro.GetValue(i, Progs.ProgramTypeCode), cEnviro.GetValue(i, Progs.RunLocal))
        ChildNode.Tag = i
        ParentNode.Nodes.Add(ChildNode)
    End Sub

    Private Sub PopulateOtherTrees(ByRef Tree As TreeView, ByVal i As Integer, ByRef ParentNode As TreeNode, ByRef ChildNode As TreeNode)
        ChildNode = New TreeNode(cEnviro.GetValue(i, Progs.Description))
        ChildNode.ImageIndex = GetImageIndex(cEnviro.GetValue(i, Progs.ProgramTypeCode), cEnviro.GetValue(i, Progs.RunLocal))
        ChildNode.Tag = i
        Tree.Nodes.Add(ChildNode)
        ' treTestProgs.Nodes.Add(ChildNode)
    End Sub

    Private Function GetImageIndex(ByVal ProgramTypeCode As String, ByVal RunLocal As String) As Integer
        If ProgramTypeCode.ToLower = "web" Then
            Return 1
        Else
            If RunLocal.ToLower = "y" Then
                Return 2
            Else
                Return 3
            End If
        End If
    End Function
#End Region

#Region " Launch program "
    Private Sub treProductionProgs_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles treProductionProgs.DoubleClick
        Dim ProcessTreeSelectResults As Results

        Try
            If treProductionProgs.SelectedNode Is Nothing Then
                MessageBox.Show("Please select a program to launch", "Launch prog", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            Else
                ProcessTreeSelectResults = ProcessTreeSelect(sender.SelectedNode)
            End If

            If Not ProcessTreeSelectResults.Success Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, ProcessTreeSelectResults.Message)
            End If

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #361: " & ex.Message)
        End Try
    End Sub

    Private Sub treTrainingProgs_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles treTrainingProgs.DoubleClick
        Dim ProcessTreeSelectResults As Results

        Try

            If treTrainingProgs.SelectedNode Is Nothing Then
                MessageBox.Show("Please select a program to launch", "Launch prog", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            Else
                ProcessTreeSelectResults = ProcessTreeSelect(sender.SelectedNode)
            End If

            If Not ProcessTreeSelectResults.Success Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, ProcessTreeSelectResults.Message)
            End If

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #362: " & ex.Message)
        End Try
    End Sub

    Private Sub treTestProgs_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles treTestProgs.DoubleClick
        Dim ProcessTreeSelectResults As Results

        Try

            If treTestProgs.SelectedNode Is Nothing Then
                MessageBox.Show("Please select a program to launch", "Launch prog", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            Else
                ProcessTreeSelectResults = ProcessTreeSelect(sender.SelectedNode)
            End If

            If Not ProcessTreeSelectResults.Success Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, ProcessTreeSelectResults.Message)
            End If

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #362: " & ex.Message)
        End Try
    End Sub

    Private Sub btnLaunchProg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLaunchProg.Click
        Dim ProcessTreeSelectResults As Results = Nothing

        Try
            Dim ChangeStateArgs As New ChangeStateArgs
            Dim LaunchArgs As New LaunchArgs
            Dim CurTree As TreeView = Nothing

            'If SSTab1.SelectedIndex = 0 Then
            '    CurTree = treProductionProgs
            'Else
            '    CurTree = treTestProgs
            'End If

            Select Case SSTab1.SelectedIndex
                Case 0
                    CurTree = treProductionProgs
                Case 1
                    CurTree = treTestProgs
                Case 2
                    CurTree = treTrainingProgs
            End Select

            If CurTree.SelectedNode Is Nothing Then
                MessageBox.Show("Please select a program to launch", "Launch prog", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            Else
                ProcessTreeSelectResults = ProcessTreeSelect(CurTree.SelectedNode)
            End If

            If Not ProcessTreeSelectResults.Success Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, ProcessTreeSelectResults.Message)
            End If

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error# 360: " & ProcessTreeSelectResults.Message)
        End Try
    End Sub

    Private Function ProcessTreeSelect(ByRef Node As TreeNode) As Results
        Dim MyResults As New Results
        Dim ChangeStateResults As New Results

        Try

            Dim ChangeStateArgs As New ChangeStateArgs
            Dim LaunchArgs As New LaunchArgs
            ChangeStateArgs.LaunchArgs = LaunchArgs
            ChangeStateArgs.TransitionTo = Enums.TransitionTo.LaunchProg
            LaunchArgs.LaunchSource = Enums.LaunchSource.Tree
            LaunchArgs.SelectedTreeNode = Node
            ChangeStateResults = ChangeState(ChangeStateArgs)
            MyResults.Success = ChangeStateResults.Success
            If Not ChangeStateResults.Success Then
                MyResults.Message = ChangeStateResults.Message
            End If

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #363: " & ex.Message
            Return MyResults
        End Try
    End Function
#End Region

#Region " Trouble call "
    Private Sub btnTroubleCall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTroubleCall.Click
        Dim FormTroubleCall As New FrmTroubleCall
        cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Clicked trouble call")
        FormTroubleCall.ShowDialog(Me, SSTab1.SelectedIndex)


        'Dim SourceIsProduction As Boolean
        'Dim CurTree As TreeView

        'If SSTab1.SelectedIndex = 0 Then
        '    SourceIsProduction = True
        '    CurTree = treProductionProgs
        'Else
        '    CurTree = treTestProgs
        'End If

        'If CurTree.SelectedNode Is Nothing Then
        '    MessageBox.Show("Please select an application before submitting a trouble report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand)
        'Else
        '    FormTroubleCall.ShowDialog(Me, SourceIsProduction, CurTree.SelectedNode.Text)
        'End If
    End Sub
#End Region

#Region " ChangeState "
    Private Function ChangeState(ByRef ChangeStateArgs As ChangeStateArgs) As Results
        Dim MyResults As New Results
        Dim RollingSuccess As Boolean = True
        Static IsActive As Boolean

        If IsActive Then
            MyResults.Success = True
            Return MyResults
        End If

        Try

            IsActive = True

            Select Case ChangeStateArgs.TransitionTo
                Case Enums.TransitionTo.Disabled

                Case Enums.TransitionTo.LaunchProg
                    Dim LaunchResults As New Results
                    Me.Cursor = Cursors.WaitCursor
                    If ChangeStateArgs.LaunchArgs.LaunchSource = Enums.LaunchSource.Tree Then
                        If Not ChangeStateArgs.LaunchArgs.SelectedTreeNode.Tag Is Nothing Then
                            ' LaunchResults = cLauncher.LaunchApp(ChangeStateArgs.LaunchArgs.SelectedTreeNode.Tag, False)
                            LaunchResults = cLauncher.LaunchApp(cDBase, ChangeStateArgs.LaunchArgs.SelectedTreeNode.Tag, CType(SSTab1.SelectedIndex, Launcher.TestTrainProdEnum), False)
                            If Not LaunchResults.Success Then
                                RollingSuccess = False
                                MyResults.Message = LaunchResults.Message
                            Else
                                treProductionProgs.SelectedImageIndex = ChangeStateArgs.LaunchArgs.SelectedTreeNode.ImageIndex
                            End If
                        End If
                    ElseIf ChangeStateArgs.LaunchArgs.LaunchSource = Enums.LaunchSource.List Then
                        ' LaunchResults = cLauncher.LaunchApp(ChangeStateArgs.LaunchArgs.DTRowNum, CType(SSTab1.SelectedIndex, Launcher.TestProdEnum), False)
                        LaunchResults = cLauncher.LaunchApp(cDBase, ChangeStateArgs.LaunchArgs.SelectedTreeNode.Tag, CType(SSTab1.SelectedIndex, Launcher.TestTrainProdEnum), False)
                        If Not LaunchResults.Success Then
                            RollingSuccess = False
                            MyResults.Message = LaunchResults.Message
                        End If
                    End If
                    Me.Cursor = Cursors.Default

                Case Enums.TransitionTo.Ready
                    'Case Enums.TransitionTo.Refresh
                    '    Dim ProgsResults As New Results
                    '    Dim PopulateListsResults As New Results

                    '    ProgsResults = GetProgs()
                    '    If Not ProgsResults.Success Then
                    '        RollingSuccess = False
                    '        MyResults.Message = ProgsResults.Message
                    '    End If

                    '    If RollingSuccess Then
                    '        If cEnviro.dtProgs.Rows.Count > 0 Then
                    '            PopulateListsResults = PopulateLists()
                    '            If Not PopulateListsResults.Success Then
                    '                RollingSuccess = False
                    '                MyResults.Message = PopulateListsResults.Message
                    '            End If
                    '        Else
                    '            UpdateStatus("No apps available.")
                    '        End If
                    '    End If

                    ' Case Enums.TransitionTo.ExitApp
                    '  Application.Exit()

            End Select

            IsActive = False

            If RollingSuccess Then
                MyResults.Success = True
            Else
                MyResults.Success = False
            End If

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #652: " & ex.Message
            Return MyResults
        End Try
    End Function
#End Region

#Region " Unused "
    'Private Sub Testie()
    '    Dim MyResults As New Results
    '    Dim QueryPack As DBase.QueryPack
    '    Dim dt As DataTable
    '    Dim Sql As String

    '    ' ___ Is the user running the app in test mode?
    '    '  Sql = "SELECT ServerPath,  LocalPath, FileName, ProgramTypeCode, RunLocal from Program WHERE PMAccessInd='Y' ORDER BY ClientID, Description"
    '    Sql = "SELECT ProgramID, ProgramTypeCode, CurrVersion, FileName, ServerPath, LocalPath, RunLocal, Installer from Program WHERE PMAccessInd='Y' ORDER BY ClientID, Description"
    '    ' Dim ConnString As String = "Provider=SQLOLEDB.1;User ID=BVI_SQL_SERVER;Password=noisivtifeneb;Persist Security Info=True;ConnectionTimeout=10;Initial Catalog=BVI;Data Source=HBG-SQL"
    '    QueryPack = cDBase.GetDTSourceIsText(Sql, cEnviro.DBConnString, False)
    '    If QueryPack.Success Then
    '        dt = QueryPack.dt

    '        Dim Excel As New Excel
    '        Excel.ExportToExcel(dt)

    '        MyResults.Success = True
    '    Else
    '        MyResults.Success = False
    '        MyResults.Message = "Error in GetDBData. " & QueryPack.TechErrMsg
    '    End If
    'End Sub

    'Private Function ReadIni() As Results
    '    Dim MyResults As New Results
    '    Dim ReadStream As StreamReader
    '    Dim Line As String
    '    Dim Key As String
    '    Dim Value As String
    '    Dim LineArr As Object

    '    Try

    '        MyResults.Success = False
    '        cEnviro.DefBackColor = Color.FromArgb(16756059)

    '        If File.Exists(cEnviro.IniFileFullPath) Then
    '            ReadStream = New StreamReader(cEnviro.IniFileFullPath, FileMode.Open)
    '            Line = ReadStream.ReadLine
    '            If Line.Length > 0 AndAlso InStr(1, Line, ":") > 0 Then
    '                LineArr = Split(Line, ":")
    '                If LineArr.GetUpperBound(0) = 6 Then
    '                    Me.Top = LineArr(0)
    '                    Me.Left = LineArr(1)
    '                    Me.Height = LineArr(2)
    '                    Me.Width = LineArr(3)
    '                    Me.BackColor = LineArr(4)
    '                    SSTab1.BackColor = LineArr(4)
    '                    If LineArr(5) = "HideSTAT" Then
    '                        'menu_StatusWindow_Click()
    '                    End If
    '                    If LineArr(6) = "Tree" Then
    '                        Do While ReadStream.Peek() >= 0
    '                            Line = ReadStream.ReadLine
    '                            cEnviro.ExpandList &= cEnviro.ExpandList & ":" & Line & ":"
    '                        Loop
    '                        'menu_ViewStyle_Tree_Click()
    '                    Else
    '                        'menu_ViewStyle_List_Click()
    '                    End If
    '                End If
    '            End If
    '            ReadStream.Close()
    '        Else
    '            'Me.BackColor = cEnviro.DefBackColor
    '            'SSTab1.BackColor = cEnviro.DefBackColor
    '        End If

    '        Me.Top = 1
    '        If Me.Left < 1 Then
    '            Me.Left = 1
    '        End If
    '        'SSTab1.Tab = 0

    '        MyResults.Success = True
    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #332: " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    'Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim ChangeStateArgs As New ChangeStateArgs
    '    ChangeStateArgs.TransitionTo = Enums.TransitionTo.ExitApp
    '    ChangeState(ChangeStateArgs)
    'End Sub


    'Private Sub mnuSubmitTroubleCall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim FormTroubleCall As New FrmTroubleCall
    '    FormTroubleCall.ShowDialog(Me)
    'End Sub

    'Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim sb As New System.Text.StringBuilder

    '    sb.Append("BVI Program Manager " & cAppSettings.CurrentVersion & vbCrLf)
    '    sb.Append("User: " & cEnviro.UserFirstName & " " & cEnviro.UserLastName & vbCrLf)
    '    sb.Append("Site: " & cEnviro.SiteId & vbCrLf)
    '    If cEnviro.TestMode Then
    '        sb.Append("Mode: Test" & vbCrLf)
    '        sb.Append("Server: " & cEnviro.TestServer & vbCrLf)
    '    Else
    '        sb.Append("Mode: Live" & vbCrLf)
    '        sb.Append("Server: " & cEnviro.SiteId & "-SQL" & vbCrLf)
    '    End If


    '    MessageBox.Show(sb.ToString, "About", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

    '    'ErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
    '    'Private Sub menu_About_Click()
    '    '        SUBNAME = "menu_About_Click"
    '    '       On Error GoTo Err

    '    '        Dim MSG As String
    '    '        MSG = MSG & AppTitle & Chr(13) & Chr(13)
    '    '        'MSG = MSG & "  created: " & App.EXEName & Chr(13) & Chr(13)
    '    '        MSG = MSG & "User: " & UserID & Chr(13)
    '    '        MSG = MSG & "Site: " & SiteID & "/" & CompID & Chr(13)
    '    '        MSG = MSG & "Mode: " & IIf(TESTMODE, "Test", "Live") & Chr(13)
    '    '        MSG = MSG & "Server: " & IIf(TESTMODE, TESTServer, SiteID & "-SQL") & Chr(13)
    '    '        MsgBox(MSG, vbInformation)

    '    'Err:
    '    '        If Err.Number <> 0 Then Call LOG("ERROR [" & SUBNAME & "]:" & Err.Number & "-" & Err.Description, True)
    '    '    End Sub

    'End Sub


    'Private Sub treProductionProgs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles treProductionProgs.Click
    '    Dim i As Integer
    '    i = treProductionProgs.SelectedNode.Tag
    '    treProductionProgs.SelectedNode.ImageIndex = GetImageIndex(cEnviro.GetValue(i, Progs.ProgramTypeCode), cEnviro.GetValue(i, Progs.RunLocal))
    'End Sub

    'Sub UpdateStatus(ByVal Msg As String, Optional ByVal ShowMsgBox As Boolean = False)

    '    Msg = Date.Now.ToString("hh:mm:ss tt: ") & Msg

    '    If lblStatus.Text.Length = 0 Then
    '        lblStatus.Text = Msg
    '    Else
    '        lblStatus.Text &= vbCrLf & Msg
    '    End If
    '    ' lblStatus.SelectionStart = lblStatus.Text.Length
    '    'lblStatus.ScrollToCaret()
    '    lblStatus.Refresh()
    '    If ShowMsgBox Then
    '        MessageBox.Show(Msg, Nothing, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
    '    End If
    'End Sub

    'Private Sub ProcessListSelect(ByVal DTRowNum As Integer)
    '    Dim ChangeStateArgs As New ChangeStateArgs
    '    Dim LaunchArgs As New LaunchArgs
    '    ChangeStateArgs.LaunchArgs = LaunchArgs
    '    ChangeStateArgs.TransitionTo = Enums.TransitionTo.LaunchProg
    '    LaunchArgs.LaunchSource = Enums.LaunchSource.List
    '    LaunchArgs.DTRowNum = DTRowNum
    '    ChangeState(ChangeStateArgs)
    'End Sub
#End Region

    Private Sub btnViewLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewLog.Click
        Try
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Clicked view log")
            Shell("notepad.exe " & cEnviro.LogFileFullPath, AppWinStyle.NormalFocus)
        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #370: " & ex.Message)
        End Try
    End Sub

    Private Sub FrmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        SSTab1.Height = Me.Height - 152
        treProductionProgs.Height = SSTab1.Height - 32
        treTestProgs.Height = SSTab1.Height - 32
        treTrainingProgs.Height = SSTab1.Height - 32
        Panel1.Top = SSTab1.Top + SSTab1.Height
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Dim MyResults As New Results
        Dim ProgsResults As New Results
        Dim PopulateTreesResults As New Results
        Dim RollingSuccess As Boolean = True

        Try

            ' ___ Get progs
            ProgsResults = GetProgs()
            If Not ProgsResults.Success Then
                RollingSuccess = False
                MyResults.Message = ProgsResults.Message
            End If

            ' ___ Populate trees
            If RollingSuccess Then
                If cEnviro.dtProgs.Rows.Count > 0 Then
                    PopulateTreesResults = PopulateTrees()
                    If Not PopulateTreesResults.Success Then
                        RollingSuccess = False
                        MyResults.Message = PopulateTreesResults.Message
                    End If
                End If
            End If

            If Not RollingSuccess Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #430: " & MyResults.Message)
            End If

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #430: " & ex.Message)
        End Try
    End Sub

    Private Sub FrmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "** Closing " & Application.ExecutablePath & " Version " & cEnviro.FormattedFileVersion(cEnviro.LocalLocalFileVersion) & "  **")
        Catch
        End Try
    End Sub
End Class

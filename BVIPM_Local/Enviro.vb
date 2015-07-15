Public Class Enviro
    Private cUserId As String
    Private cPassword As String
    Private cComputerId As String
    Private cSiteId As String
    Private cDBConnString As String
    Private cDBConnStringPrefix As String
    Private cTestMode As Boolean
    Private cTestServer As String
    Private cUserFirstName As String
    Private cUserLastName As String
    Private cdtProgs As DataTable
    Private cFileNameArg As String
    Private cdtIRTS As DataTable
    Private cAnyRecordsReturned As Boolean
    Private cLogFileFullPath As String
    Private cProgMgrServerPath As String
    Private cLauncherMasterFileVersion As String
    Private cLauncherLocalFileVersion As String
    Private cLocalLocalFileVersion As String
    Private cOSPlatformBitCount As OSPlatformBitCountEnum

    Public Const DebugMode As Boolean = True
    Public Const cDebugModeVersion As String = "Test v100"

    'Public Const CheckMSIEXEC_exe As Boolean = False
    'Public Const CheckMain_App_exe As Boolean = True

    Public Enum OSPlatformBitCountEnum
        BitCount32 = 1
        BitCount64 = 2
    End Enum

    Public Sub New()
        cDBConnStringPrefix = "Provider=SQLOLEDB.1;User ID=BVI_SQL_SERVER;Password=noisivtifeneb;Persist Security Info=True;ConnectionTimeout=@DBTimeout;Initial Catalog=BVI;Data Source="
        cDBConnStringPrefix = Replace(cDBConnStringPrefix, "@DBTimeout", cAppSettings.DBTimeout)
    End Sub

    Public Function FormattedFileVersion(ByVal FileVersion As String) As String
        If DebugMode Then
            Return cDebugModeVersion
        Else
            Dim Box As Object
            Box = Split(FileVersion, ".")

            If Box(1) = "0" AndAlso Box(2) = "0" Then
                Return Box(0) & "." & Box(1)
            Else
                Return Box(0) & "." & Box(1) & Box(2)
            End If
        End If
    End Function

    Public Function GetValue(ByVal RowNum As Integer, ByVal FldName As String) As Object
        Return cdtProgs.Rows(RowNum)(FldName)
    End Function

    Public Sub SetValue(ByVal RowNum As String, ByVal FldName As String, ByVal Value As String)
        cdtProgs.Rows(RowNum)(FldName) = Value
    End Sub

    Public Property DBConnStringPrefix() As String
        Get
            Return cDBConnStringPrefix
        End Get
        Set(ByVal Value As String)
            cDBConnStringPrefix = Value
        End Set
    End Property
    'Public Property CurrentVersion() As String
    '    Get
    '        Return cCurrentVersion
    '    End Get
    '    Set(ByVal Value As String)
    '        cCurrentVersion = Value
    '    End Set
    'End Property
    'Public Property MasterVersion() As String
    '    Get
    '        Return cMasterVersion
    '    End Get
    '    Set(ByVal Value As String)
    '        cMasterVersion = Value
    '    End Set
    'End Property

    Public Property LauncherMasterFileVersion() As String
        Get
            Return cLauncherMasterFileVersion
        End Get
        Set(ByVal Value As String)
            cLauncherMasterFileVersion = Value
        End Set
    End Property
    Public Property LauncherLocalFileVersion() As String
        Get
            Return cLauncherLocalFileVersion
        End Get
        Set(ByVal Value As String)
            cLauncherLocalFileVersion = Value
        End Set
    End Property

    Public Property LocalLocalFileVersion() As String
        Get
            Return cLocalLocalFileVersion
        End Get
        Set(ByVal Value As String)
            cLocalLocalFileVersion = Value
        End Set
    End Property

    Public Property AnyRecordsReturned() As Boolean
        Get
            Return cAnyRecordsReturned
        End Get
        Set(ByVal Value As Boolean)
            cAnyRecordsReturned = Value
        End Set
    End Property

    Public Property ProgMgrServerPath() As String
        Get
            Return cProgMgrServerPath
        End Get
        Set(ByVal Value As String)
            cProgMgrServerPath = Value
        End Set
    End Property
    Public Property FileNameArg() As String
        Get
            Return cFileNameArg
        End Get
        Set(ByVal Value As String)
            cFileNameArg = Value
        End Set
    End Property

    Public Property dtProgs() As DataTable
        Get
            Return cdtProgs
        End Get
        Set(ByVal Value As DataTable)
            cdtProgs = Value
        End Set
    End Property


    Public Property UserFirstName() As String
        Get
            Return cUserFirstName
        End Get
        Set(ByVal Value As String)
            cUserFirstName = Value
        End Set
    End Property
    Public Property UserLastName() As String
        Get
            Return cUserLastName
        End Get
        Set(ByVal Value As String)
            cUserLastName = Value
        End Set
    End Property

    Public Property TestServer() As String
        Get
            Return cTestServer
        End Get
        Set(ByVal Value As String)
            cTestServer = Value
        End Set
    End Property

    Public Property TestMode() As Boolean
        Get
            Return cTestMode
        End Get
        Set(ByVal Value As Boolean)
            cTestMode = Value
        End Set
    End Property
    Public ReadOnly Property IRTSConnString() As String
        Get
            Return cDBConnStringPrefix & "HBG-SQL"
        End Get
    End Property
    Public ReadOnly Property DBConnString() As String
        Get
            Return cDBConnStringPrefix & cSiteId & "-SQL"
        End Get
    End Property
    Public ReadOnly Property DBConnString(ByVal TestProdTrain As Launcher.TestTrainProdEnum) As String
        Get
            Select Case TestProdTrain
                Case Launcher.TestTrainProdEnum.Test
                    Return cDBConnStringPrefix & cSiteId & "-TST"
                Case Launcher.TestTrainProdEnum.Prod
                    Return cDBConnStringPrefix & cSiteId & "-SQL"
                Case Launcher.TestTrainProdEnum.Training
                    Return cDBConnStringPrefix & cSiteId & "-TRAINING"
            End Select
        End Get
    End Property

    Public Property UserId() As String
        Get
            Return cUserId
        End Get
        Set(ByVal Value As String)
            cUserId = Value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return cPassword
        End Get
        Set(ByVal Value As String)
            cPassword = Value
        End Set
    End Property
    Public Property ComputerId() As String
        Get
            Return cComputerId
        End Get
        Set(ByVal Value As String)
            cComputerId = Value
        End Set
    End Property
    Public Property SiteId() As String
        Get
            Return cSiteId
        End Get
        Set(ByVal Value As String)
            cSiteId = Value
        End Set
    End Property
    Public Property dtIRTS() As DataTable
        Get
            Return cdtIRTS
        End Get
        Set(ByVal Value As DataTable)
            cdtIRTS = Value
        End Set
    End Property
    Public Property LogFileFullPath() As String
        Get
            Return cLogFileFullPath
        End Get
        Set(ByVal Value As String)
            cLogFileFullPath = Value
        End Set
    End Property
    Public Property OSPlatformBitCount() As OSPlatformBitCountEnum
        Get
            Return cOSPlatformBitCount
        End Get
        Set(ByVal Value As OSPlatformBitCountEnum)
            cOSPlatformBitCount = Value
        End Set
    End Property
End Class

Public Class Progs
    Public Const ProgramID As String = "ProgramID"
    Public Const Description As String = "Description"
    Public Const ProgramTypeCode As String = "ProgramTypeCode"
    Public Const AcesInd As String = "AcesInd"
    Public Const ClientID As String = "ClientID"
    Public Const CurrVersion As String = "CurrVersion"
    Public Const FileName As String = "FileName"
    Public Const ServerPath As String = "ServerPath"
    Public Const LocalPath As String = "LocalPath"
    Public Const LocalPath64 As String = "LocalPath64"
    Public Const PMAccessInd As String = "PMAccessInd"
    Public Const RunLocal As String = "RunLocal"
    Public Const Installer As String = "Installer"
End Class

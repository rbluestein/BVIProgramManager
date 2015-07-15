Public Class Enviro
    Private cUserId As String
    Private cComputerId As String
    Private cSiteId As String
    Private cDBConnStringPrefix As String
    Private cLogFileFullPath As String
    Private cServerPath As String
    Private Const cDBTimeout As Integer = 10
    'Private cCurrentVersion As String
    'Private cMasterVersion As String

    Private cLauncherLocalFileVersion As String
    Private cLocalMasterFileVersion As String
    Private cLocalLocalFileVersion As String

    Public Sub New()
        cDBConnStringPrefix = "Provider=SQLOLEDB.1;User ID=BVI_SQL_SERVER;Password=noisivtifeneb;Persist Security Info=True;ConnectionTimeout=@DBTimeout;Initial Catalog=BVI;Data Source="
        cDBConnStringPrefix = Replace(cDBConnStringPrefix, "@DBTimeout", DBTimeout)
    End Sub

    Public Function FormattedFileVersion(ByVal FileVersion As String) As String
        Dim Box As Object
        Box = Split(FileVersion, ".")

        If Box(1) = "0" AndAlso Box(2) = "0" Then
            Return Box(0) & "." & Box(1)
        Else
            Return Box(0) & "." & Box(1) & Box(2)
        End If
    End Function

    Public Property LauncherLocalFileVersion() As String
        Get
            Return cLauncherLocalFileVersion
        End Get
        Set(ByVal Value As String)
            cLauncherLocalFileVersion = Value
        End Set
    End Property

    Public Property LocalMasterFileVersion() As String
        Get
            Return cLocalMasterFileVersion
        End Get
        Set(ByVal Value As String)
            cLocalMasterFileVersion = Value
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
    Public ReadOnly Property DBTimeout() As Integer
        Get
            Return cDBTimeout
        End Get
    End Property

    Public Property ServerPath() As String
        Get
            Return cServerPath
        End Get
        Set(ByVal Value As String)
            cServerPath = Value
        End Set
    End Property
    Public Property UserId() As String
        Get
            Return cUserId
        End Get
        Set(ByVal Value As String)
            cUserId = Value
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
    Public Property LogFileFullPath() As String
        Get
            Return cLogFileFullPath
        End Get
        Set(ByVal Value As String)
            cLogFileFullPath = Value
        End Set
    End Property
End Class
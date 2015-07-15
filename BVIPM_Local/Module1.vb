Module Module1

    Public Class LaunchExecute
        Private cServerPath As String

        Private cTableListingForLocalPath32 As String
        Private cTableListingForLocalPath64 As String
        Private cTableHasListingForLocalPath32 As Boolean
        Private cTableHasListingForLocalPath64 As Boolean
        Private cLocalPath32Exists As Boolean
        Private cLocalPath64Exists As Boolean
        Private cLocalProg32Exists As Boolean
        Private cLocalProg64Exists As Boolean

        Private cFileName As String
        Private cCommandLineArgs As String
        Private cLocation As LocationEnum
        Private cServerPathExists As Boolean
        Private cServerProgExists As Boolean
        Private cInstaller As String
        Private cAcesInstallerExists As Boolean
        Private cDescription As String
        Private cAcesInd As Boolean
        Private cAcesVersion As String

        Public Enum LocationEnum
            Web = 1
            Local = 2
            Server = 3
        End Enum

        Public Property AcesInd() As Boolean
            Get
                Return cAcesInd
            End Get
            Set(ByVal Value As Boolean)
                cAcesInd = Value
            End Set
        End Property
        Public Property Description() As String
            Get
                Return cDescription
            End Get
            Set(ByVal Value As String)
                cDescription = Value
            End Set
        End Property
        Public Property Installer() As String
            Get
                Return cInstaller
            End Get
            Set(ByVal Value As String)
                cInstaller = Value
            End Set
        End Property
        Public Property AcesInstallerExists() As Boolean
            Get
                Return cAcesInstallerExists
            End Get
            Set(ByVal Value As Boolean)
                cAcesInstallerExists = Value
            End Set
        End Property
        Public Property TableListingForLocalPath32() As String
            Get
                Return cTableListingForLocalPath32
            End Get
            Set(ByVal Value As String)
                cTableListingForLocalPath32 = Value
            End Set
        End Property
        Public Property TableListingForLocalPath64() As String
            Get
                Return cTableListingForLocalPath64
            End Get
            Set(ByVal Value As String)
                cTableListingForLocalPath64 = Value
            End Set
        End Property
        Public Property LocalPath32Exists() As Boolean
            Get
                Return cLocalPath32Exists
            End Get
            Set(ByVal Value As Boolean)
                cLocalPath32Exists = Value
            End Set
        End Property
        Public Property LocalPath64Exists() As Boolean
            Get
                Return cLocalPath64Exists
            End Get
            Set(ByVal Value As Boolean)
                cLocalPath64Exists = Value
            End Set
        End Property
        Public Property LocalProg32Exists() As Boolean
            Get
                Return cLocalProg32Exists
            End Get
            Set(ByVal Value As Boolean)
                cLocalProg32Exists = Value
            End Set
        End Property
        Public Property LocalProg64Exists() As Boolean
            Get
                Return cLocalProg64Exists
            End Get
            Set(ByVal Value As Boolean)
                cLocalProg64Exists = Value
            End Set
        End Property

        Public Property TableHasListingForLocalPath32() As Boolean
            Get
                Return cTableHasListingForLocalPath32
            End Get
            Set(ByVal Value As Boolean)
                cTableHasListingForLocalPath32 = Value
            End Set
        End Property
        Public Property TableHasListingForLocalPath64() As Boolean
            Get
                Return cTableHasListingForLocalPath64
            End Get
            Set(ByVal Value As Boolean)
                cTableHasListingForLocalPath64 = Value
            End Set
        End Property







        Public Property ServerPathExists() As Boolean
            Get
                Return cServerPathExists
            End Get
            Set(ByVal Value As Boolean)
                cServerPathExists = Value
            End Set
        End Property
        Public Property ServerProgExists() As Boolean
            Get
                Return cServerProgExists
            End Get
            Set(ByVal Value As Boolean)
                cServerProgExists = Value
            End Set
        End Property
        Public Property Location() As LocationEnum
            Get
                Return cLocation
            End Get
            Set(ByVal Value As LocationEnum)
                cLocation = Value
            End Set
        End Property
        Public Property ServerPath() As String
            Get
                Return cServerPath
            End Get
            Set(ByVal Value As String)
                cServerPath = Value
            End Set
        End Property


        Public Property FileName() As String
            Get
                Return cFileName
            End Get
            Set(ByVal Value As String)
                cFileName = Value
            End Set
        End Property
        Public Property CommandLineArgs() As String
            Get
                Return cCommandLineArgs
            End Get
            Set(ByVal Value As String)
                cCommandLineArgs = Value
            End Set
        End Property
        Public Property AcesVersion() As String
            Get
                Return cAcesVersion
            End Get
            Set(ByVal Value As String)
                cAcesVersion = Value
            End Set
        End Property
    End Class

End Module

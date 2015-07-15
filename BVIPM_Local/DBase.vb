Imports System.Threading

Public Class DBase
    Public Function GetDTSql(ByRef dt As DataTable) As DataTable
        Dim Row As Integer
        Dim Col As Integer
        Dim dt2 As New DataTable
        Dim dr2 As DataRow

        For Col = 0 To dt.Columns.Count - 1
            dt2.Columns.Add(dt.Columns(Col).ColumnName, GetType(DataTypeSQL))
        Next

        For Row = 0 To dt.Rows.Count - 1
            dr2 = dt2.NewRow
            For Col = 0 To dt.Columns.Count - 1
                Select Case dt.Columns(Col).DataType.ToString.ToLower
                    Case "system.guid"
                        Dim VarcharSql As New VarcharSQL(dt.Rows(Row)(Col).ToString)
                        Dim DataTypeSql As DataTypeSQL = VarcharSql
                        dr2(Col) = DataTypeSql
                    Case "system.string"
                        Dim VarcharSql As New VarcharSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = VarcharSql
                        dr2(Col) = DataTypeSql
                    Case "system.int64", "system.int32", "system.int16"
                        Dim IntSql As New IntSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = IntSql
                        dr2(Col) = DataTypeSql
                    Case "system.boolean"
                        Dim BitSql As New BitSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = BitSql
                        dr2(Col) = DataTypeSql
                    Case "system.datetime"
                        Dim DateTimeSql As New DateTimeSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = DateTimeSql
                        dr2(Col) = DataTypeSql
                    Case "system.double", "system.single"
                        Dim FloatSql As New FloatSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = FloatSql
                        dr2(Col) = DataTypeSql
                    Case "system.decimal"
                        Dim MoneySql As New MoneySQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = MoneySql
                        dr2(Col) = DataTypeSql
                    Case Else
                        Throw New Exception("Unprocessed data type in DBase.GetDTSql -- " & dt.Columns(Col).DataType.ToString.ToLower & ".")
                End Select
            Next
            dt2.Rows.Add(dr2)
        Next
        Return dt2
    End Function


    Public Function ConnStringCondition(ByVal Input As String) As String
        Dim i As Integer
        Dim Params As Object
        Dim KeyValue As Object
        Dim Coll As New Collection
        Dim Output As String = Nothing

        Params = Input.Split(";")
        For i = 0 To Params.GetUpperBound(0)
            KeyValue = Params(i).Split("=")
            Select Case KeyValue(0).ToLower
                Case "provider"
                Case "connectiontimeout"
                    Coll.Add("Connection Timeout=" & KeyValue(1))
                Case Else
                    Coll.Add(KeyValue(0) & "=" & KeyValue(1))
            End Select
        Next

        For i = 1 To Coll.Count
            If i = 1 Then
                Output = Coll(i) & ";"
            ElseIf i = Coll.Count Then
                Output &= Coll(i)
            Else
                Output &= Coll(i) & ";"
            End If
        Next

        Return Output
    End Function

    Public Function GetDTSourceIsText(ByVal Sql As String, ByVal DBConnString As String, ByVal SqlTable As Boolean) As DBase.QueryPack
        Dim sb As New System.Text.StringBuilder
        Dim QueryPack As New DBase.QueryPack
        Dim DataAdapter As System.Data.SqlClient.SqlDataAdapter
        Dim dt As New DataTable
        DBConnString = ConnStringCondition(DBConnString)

        Dim SqlCmd As New SqlClient.SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New Data.SqlClient.SqlConnection(DBConnString)
        DataAdapter = New System.Data.SqlClient.SqlDataAdapter(SqlCmd)
        Try
            DataAdapter.Fill(dt)
            QueryPack.Success = True
            If SqlTable Then
                dt = GetDTSql(dt)
            End If
            QueryPack.dt = dt
        Catch ex As Exception
            QueryPack.Success = False
            QueryPack.TechErrMsg = ex.Message
        End Try
        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return QueryPack
    End Function

    'Public Function GetDTAllDBQueryPack(ByVal Sql As String, ByVal DBConnString As String, ByVal Timeout As Integer) As QueryPack
    '    cSql = Sql
    '    cDBConnString = DBConnString
    '    cTimeout = Timeout

    '    cTimerThread = New Thread(AddressOf Me.LaunchTimerThread)
    '    cTimerThread.Name = "TimerThread"
    '    cQueryThread = New Thread(AddressOf Me.LaunchQueryThread)
    '    cQueryThread.Name = "QueryThread"

    '    '   cTimerThread.Start()
    '    cQueryThread.Start()

    '    Do
    '        Thread.Sleep(100)
    '    Loop Until cDone

    '    Return cQueryPack
    'End Function

    'Private Sub LaunchTimerThread()
    '    'Dim Current As Thread = Thread.CurrentThread
    '    cTimerThread.Sleep(cTimeout * 1000)
    '    ThreadLanding("Timer")
    'End Sub

    'Public Sub LaunchQueryThread()
    '    cQueryPack = GetDTSqlText(cSql, cDBConnString)
    '    ThreadLanding("QueryThread")
    'End Sub

    'Private Sub ThreadLanding(ByVal ThreadName As String)
    '    If ThreadName = "Timer" Then
    '        If cQueryThread.IsAlive Then
    '            cQueryThread.Abort()
    '            cQueryPack.Success = False
    '            cQueryPack.TechErrMsg = "Timeout occurred"
    '        End If
    '    ElseIf ThreadName = "QueryThread" Then
    '        If cTimerThread.IsAlive Then
    '            cTimerThread.Abort()
    '            cDataReturned = True
    '        End If
    '    End If
    '    cDone = True
    'End Sub

    Public Class QueryPack
        Private cReturnDataTable As Boolean
        Private cReturnDataSet As Boolean
        Private cSuccess As Boolean
        Private cGenErrMsg As String
        Private cTechErrMsg As String
        Private cdt As DataTable

        Public Property Success() As Boolean
            Get
                Return cSuccess
            End Get
            Set(ByVal Value As Boolean)
                cSuccess = Value
            End Set
        End Property

        Public ReadOnly Property GenErrMsg() As String
            Get
                Return cGenErrMsg
            End Get
        End Property
        Public Property TechErrMsg() As String
            Get
                Return cTechErrMsg
            End Get
            Set(ByVal Value As String)
                cTechErrMsg = Value
            End Set
        End Property
        Public Property dt() As DataTable
            Get
                Return cdt
            End Get
            Set(ByVal Value As DataTable)
                cdt = Value
            End Set
        End Property
    End Class

End Class

#Region " SQL Server DataTypes "
Public Class DataTypeSQL
End Class

Public Class DateTimeSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value

    Public ReadOnly Property SqlOut() As String
        Get
            If IsDBNull(cValue) Then
                Return "null"
            Else
                Return "'" & ToText & "'"
            End If
        End Get
    End Property
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                Return CType(cValue, DateTime)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, DateTime)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                Return CType(cValue, DateTime)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class IntSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public ReadOnly Property SqlOut() As String
        Get
            If IsDBNull(cValue) Then
                Return "null"
            Else
                Return ToText
            End If
        End Get
    End Property
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                Return CType(cValue, System.Int64)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, System.Int64)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                Return CType(cValue, System.Int64)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class VarcharSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public ReadOnly Property SqlOut() As String
        Get
            If IsDBNull(cValue) Then
                Return "null"
            Else
                Return "'" & ToText & "'"
            End If
        End Get
    End Property
    Public Property Value()
        Get
            If IsDBNull(cValue) OrElse cValue = Nothing Then
                Return DBNull.Value
            Else
                Return CType(cValue, System.String)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) OrElse Value = Nothing Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, System.String)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) OrElse cValue = Nothing Then
                Return String.Empty
            Else
                Return CType(cValue, System.String)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class BitSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public ReadOnly Property SqlOut() As String
        Get
            If IsDBNull(cValue) Then
                Return "null"
            Else
                If cValue = 0 Then
                    Return "0"
                Else
                    Return "1"
                End If
            End If
        End Get
    End Property
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                Return CType(cValue, System.Int64)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                If Value Then
                    cValue = 1
                Else
                    cValue = 0
                End If
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                Return CType(cValue, System.Int64)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class MoneySQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public ReadOnly Property SqlOut() As String
        Get
            If IsDBNull(cValue) Then
                Return "null"
            Else
                Return ToText
            End If
        End Get
    End Property
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                'Return CType(cValue, System.Decimal)
                Return FormatNumber(CType(cValue, System.Decimal), 2)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, System.Decimal)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                'Return CType(cValue, System.Decimal)
                Return FormatNumber(CType(cValue, System.Decimal), 2)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class FloatSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public ReadOnly Property SqlOut() As String
        Get
            If IsDBNull(cValue) Then
                Return "null"
            Else
                Return ToText
            End If
        End Get
    End Property
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                Return CType(cValue, System.Double)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, System.Double)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                Return CType(cValue, System.Double)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class
#End Region



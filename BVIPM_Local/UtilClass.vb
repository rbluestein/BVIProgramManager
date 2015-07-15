Public Class UtilClass
    Public Function DTFindRow(ByVal DT As DataTable, ByVal FldName As String, ByVal FldValue As Object, Optional ByVal IgnoreCase As Boolean = True) As Integer
        Dim i As Integer

        If DT.Rows.Count = 0 Then
            Return -1
        Else
            For i = 0 To DT.Columns.Count - 1
                If IsStrEqual(DT.Columns(i).ColumnName, FldName) Then
                    Exit For
                End If
            Next
            If i = DT.Columns.Count Then
                'Dim errorobj As New ErrorClass
                'errorobj.SimpleHandleError(EnumClass.ErrorSourceUserAppIOItem.AppError, "[UtilClass.DTFindRow] Invalid field name.", False)
            End If
            For i = 0 To DT.Rows.Count - 1
                If Not IsDBNull(DT.Rows(i)(FldName)) Then
                    If IsStrEqual(DT.Rows(i)(FldName), FldValue, IgnoreCase) Then
                        Return i
                    End If
                End If
            Next
        End If
        Return -1
    End Function

    Public Function IsStrEqual(ByVal FirstValue As String, ByVal SecondValue As String, Optional ByVal IgnoreCase As Boolean = True) As Boolean
        Dim Output As Integer
        Dim Results As Boolean
        Output = String.Compare(Trim(FirstValue), Trim(SecondValue), IgnoreCase)
        If Output = 0 Then
            Results = True
        End If
        Return Results
    End Function

    Public Function DTSort(ByRef dt As DataTable, ByVal Filter As String, ByVal Sort As String, ByVal NewTable As Boolean) As DataTable
        Return DTSort(dt, Filter, Sort, NewTable, True)
    End Function


    Public Function DTSort(ByRef dt As DataTable, ByVal Filter As String, ByVal Sort As String, ByVal NewTable As Boolean, ByVal Ascending As Boolean) As DataTable
        'strExpr = "id > 5"
        ' Sort descending by CompanyName column.
        'strSort = "name DESC"
        ' Use the Select method to find all rows matching the filter.

        ' To filter out the null first row in a DS DT:
        ' Filter = "CUST_ID > ''"
        '.FilterOn("Holiday = 'xmas'")

        Dim i As Integer
        Dim Row As Integer
        Dim SortedRows() As DataRow
        Dim NewDT As New DataTable
        Dim dr As DataRow
        Dim CompoundSort As Boolean

        Try

            ' Use the Select method to find all rows matching the filter.
            ' ===========================================================
            If Sort = Nothing Then
                SortedRows = dt.Select(Filter)
            Else

                If InStr(Sort, " ") > 0 Then
                    CompoundSort = True
                End If

                If CompoundSort Then
                    SortedRows = dt.Select(Filter, Sort)
                Else
                    If Ascending Then
                        SortedRows = dt.Select(Filter, Sort & " ASC")
                    Else
                        SortedRows = dt.Select(Filter, Sort & " DESC")
                    End If
                End If



            End If

            ' Add the columns to the new datatable.
            ' =====================================
            For i = 0 To dt.Columns.Count - 1
                NewDT.Columns.Add(CloneColumn(dt.Columns(i)))
            Next

            ' Write the data to the new datatable.
            ' ====================================
            For Row = 0 To SortedRows.GetUpperBound(0)
                Dim ItemArray(dt.Columns.Count - 1) As Object
                ItemArray = SortedRows(Row).ItemArray

                dr = NewDT.NewRow()
                NewDT.Rows.Add(dr)
                NewDT.Rows(Row).ItemArray = ItemArray

                'For Col = 0 To Output.Columns.Count - 1
                '    OutputRow(Col) = SortedRows(Row)(Col)
                'Next
            Next

            If NewTable Then
                Return NewDT
            Else

                dt.Rows.Clear()
                For Row = 0 To NewDT.Rows.Count - 1
                    Dim ItemArray(dt.Columns.Count - 1) As Object
                    ItemArray = NewDT.Rows(Row).ItemArray

                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dt.Rows(Row).ItemArray = itemarray
                Next
                Return dt
            End If



            ''Output = dt.Copy
            ''Output.Rows.Clear()

            'For Row = 0 To SortedRows.GetUpperBound(0)
            '    OutputRow = Output.NewRow()

            '    Dim ItemArray(dt.Columns.Count - 1) As Object
            '    ItemArray = SortedRows(Row).ItemArray
            '    Output.Rows.Add(OutputRow)
            '    Output.Rows(Row).ItemArray = ItemArray


            '    'For Col = 0 To Output.Columns.Count - 1
            '    '    OutputRow(Col) = SortedRows(Row)(Col)
            '    'Next
            'Next

            'If NewTable Then
            '    Return Output
            'Else
            '    dt = Output
            '    Return dt
            'End If

        Catch ex As Exception
            'Dim ErrorObj As New ErrorClass
            'ErrorObj.SimpleHandleError(EnumClass.ErrorSourceUserAppIOItem.AppError, "[UtilMod.DTSort] Error" & vbCrLf & ex.Message, False)
        End Try
    End Function

    Public Function CloneColumn(ByVal SourceColumn As DataColumn) As DataColumn
        Dim ColumnName As String
        Dim ColumnType As Type
        Dim NewColumn As DataColumn

        ColumnName = SourceColumn.ColumnName
        ColumnType = SourceColumn.DataType
        NewColumn = New DataColumn(ColumnName, ColumnType)
        Return NewColumn
    End Function


    Public Function IsNumericRB(ByVal Value As Object) As Boolean
        Try
            If Value = 0 Then
                Return True
            End If
        Catch
        End Try
        If IsDBNull(Value) Then
            Return False
        ElseIf Value = Nothing Then
            Return False
        Else
            If Len(Value) > 0 And IsNumeric(Value) Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Function FormatStringRB(ByVal DataSubType As EnumClass.DataSubType) As String
        Dim Results As String = String.Empty
        Select Case DataSubType
            Case EnumClass.DataSubType.Boolean
                Results = String.Empty

            Case EnumClass.DataSubType.Currency
                Results = "##,###.00"

            Case EnumClass.DataSubType.Date
                Results = "dd-MMM-yyyy"

            Case EnumClass.DataSubType.Percent
                Results = "##,###.000"

            Case EnumClass.DataSubType.SearchCurrency
                Results = "##,###.000"
        End Select
        Return Results
    End Function
End Class

Public Class EnumClass
    Public Enum DataSubType
        [NotApplicable] = 0
        [String] = 1
        [Percent] = 2
        Currency = 3
        [Integer] = 4
        [Boolean] = 5
        [Date] = 6
        SearchCurrency = 7
    End Enum
End Class
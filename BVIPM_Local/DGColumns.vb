Imports System.Reflection

Public Class TextBoxColumnExtended
    Inherits DataGridTextBoxColumn

    Public UtilObj As New UtilClass

    Private cIsBoolean As Boolean
    Private cTrueValue As String
    Private cFalseValue As String
    Private cUseAlternatingRowColor As Boolean
    Private cAlternatingBackColor As Color
    Private cFormat As String

    Public Sub New(ByVal MappingName As String)
        MyBase.New()
        Me.MappingName = MappingName
    End Sub

    Public Sub New(ByVal MappingName As String, _
           ByVal Width As Integer, _
           ByVal Alignment As HorizontalAlignment, _
           ByVal [ReadOnly] As Boolean, _
           ByVal HeaderText As String, _
           ByVal NullText As String, _
           ByVal DataSubType As EnumClass.DataSubType, _
           Optional ByVal FormatString As String = Nothing)

        Me.New(MappingName)

        'If IsEnumEqual(DataSubType, EnumClass.DataSubType.Currency) Then
        '    'cFormat = "c"
        '    cFormat = UtilObj.FormatStringRB(DataSubType)
        'End If
        'Me.Format = cFormat

        If Not FormatString = Nothing Then
            Me.Format = FormatString
        Else
            Me.Format = UtilObj.FormatStringRB(DataSubType)
        End If


        Me.Alignment = Alignment
        Me.Width = Width
        Me.ReadOnly = [ReadOnly]
        Me.HeaderText = HeaderText
        Me.NullText = NullText
        cIsBoolean = False
    End Sub

    Public Sub New(ByVal MappingName As String, _
                    ByVal Width As Integer, _
                    ByVal Alignment As HorizontalAlignment, _
                    ByVal [ReadOnly] As Boolean, _
                    ByVal HeaderText As String, _
                    ByVal NullText As String, _
                    ByVal DataSubType As EnumClass.DataSubType, _
                    ByVal IsBoolean As Boolean, _
                    ByVal TrueValue As String, _
                    ByVal FalseValue As String, _
                    Optional ByVal FormatString As String = Nothing)

        Me.New(MappingName)
        'Me.Format = UtilObj.FormatStringRB(DataSubType)

        If Not FormatString = Nothing Then
            Me.Format = FormatString
        Else
            Me.Format = UtilObj.FormatStringRB(DataSubType)
        End If

        Me.Alignment = Alignment
        Me.Width = Width
        Me.ReadOnly = [ReadOnly]
        Me.HeaderText = HeaderText
        Me.NullText = NullText
        cIsBoolean = IsBoolean
        cTrueValue = TrueValue
        cFalseValue = FalseValue
    End Sub

    Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal GridRow As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)
        Dim Value As String
        Dim BoolValue As Boolean

        Try

            If Not cIsBoolean Then
                MyBase.Paint(g, bounds, source, GridRow, backBrush, foreBrush, False)
                Exit Sub
            End If

            Try
                If cIsBoolean Then
                    BoolValue = CType(GetColumnValueAtRow([source], GridRow), Boolean)
                    If BoolValue Then
                        Value = cTrueValue
                    Else
                        Value = cFalseValue
                    End If
                Else
                    Value = getcolumnvalueatrow([source], GridRow)
                End If
            Catch
                Value = IIf(IsDBNull(getcolumnvalueatrow([source], GridRow)), String.Empty, getcolumnvalueatrow([source], GridRow))
            End Try
            g.FillRectangle(backBrush, bounds)
            bounds.Offset(0, 2)
            bounds.Height -= 2
            g.DrawString(Value, Me.DataGridTableStyle.DataGrid.Font, foreBrush, RectangleF.FromLTRB(bounds.X, bounds.Y, bounds.Right, bounds.Bottom))


        Catch ex As Exception
            'Dim ErrorObj As New ErrorClass
            'ErrorObj.SimpleHandleError(EnumClass.ErrorSourceUserAppIOItem.AppError, "[DataGridEditedTextBoxColumnExtended.Paint], Error" & vbCrLf & ex.Message, False)
        End Try
    End Sub

    Protected Overrides Sub Abort(ByVal rowNum As Integer)
    End Sub

    Protected Overrides Function Commit(ByVal dataSource As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer) As Boolean
    End Function

    Protected Overrides Function GetMinimumHeight() As Integer
    End Function

    Protected Overrides Function GetPreferredHeight(ByVal g As System.Drawing.Graphics, ByVal value As Object) As Integer
    End Function

    Protected Overrides Function GetPreferredSize(ByVal g As System.Drawing.Graphics, ByVal value As Object) As System.Drawing.Size
    End Function

    Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle, ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer)
    End Sub

    Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle, ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal alignToRight As Boolean)
    End Sub
End Class

Public Class MultiLineColumn
    Inherits DataGridTextBoxColumn

    Private mTxtAlign As HorizontalAlignment
    Private mDrawTxt As New StringFormat
    Private mbAdjustHeight As Boolean = True
    Private m_intPreEditHeight As Integer
    Private m_rownum As Integer
    Dim WithEvents dg As DataGrid
    Private arHeights As ArrayList

    Public Sub New(ByVal MappingName As String)
        MyBase.New()
        Me.MappingName = MappingName
        mTxtAlign = HorizontalAlignment.Left
        'DrawFormat.Alignment = StringAlignment.Near
        'MyBase.TextBox.TextAlign = HAlignment
        'MyBase.TextBox.Multiline = AdjustHeight
    End Sub

    Public Sub New(ByVal MappingName As String, _
               ByVal Width As Integer, _
               ByVal Alignment As HorizontalAlignment, _
               ByVal [ReadOnly] As Boolean, _
               ByVal HeaderText As String, _
               ByVal NullText As String, _
               ByVal DataSubType As EnumClass.DataSubType)

        Me.New(MappingName)
        Me.MappingName = MappingName
        Me.Alignment = Alignment
        Me.Width = Width
        Me.ReadOnly = [ReadOnly]
        Me.HeaderText = HeaderText
        Me.NullText = NullText
    End Sub

    Private Sub GetHeightList()
        Dim mi As MethodInfo = dg.GetType().GetMethod("get_DataGridRows", BindingFlags.FlattenHierarchy Or BindingFlags.IgnoreCase Or BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static)
        Dim dgra As Array = CType(mi.Invoke(Me.dg, Nothing), Array)
        arHeights = New ArrayList
        Dim dgRowHeight As Object
        For Each dgRowHeight In dgra
            If dgRowHeight.ToString().EndsWith("DataGridRelationshipRow") = True Then
                arHeights.Add(dgRowHeight)
            End If
        Next
    End Sub

    Public Sub New()
        mTxtAlign = HorizontalAlignment.Left
        mDrawTxt.Alignment = StringAlignment.Near
        Me.ReadOnly = True
    End Sub

    Protected Overloads Overrides Sub Edit(ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal bounds As System.Drawing.Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
        MyBase.Edit(source, rowNum, bounds, [readOnly], instantText, cellIsVisible)
        Me.TextBox.TextAlign = mTxtAlign
        Me.TextBox.Multiline = mbAdjustHeight
        If rowNum = source.Count - 1 Then
            GetHeightList()
        End If
    End Sub

    Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle, ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As System.Drawing.Brush, ByVal foreBrush As System.Drawing.Brush, ByVal alignToRight As Boolean)
        Static bPainted As Boolean = False
        If Not bPainted Then
            dg = Me.DataGridTableStyle.DataGrid
            GetHeightList()
        End If

        'clear the cell
        g.FillRectangle(backBrush, bounds)

        'draw the value
        Dim s As String = Me.GetColumnValueAtRow([source], rowNum).ToString()
        Dim r As New RectangleF(bounds.X, bounds.Y, bounds.Width, bounds.Height)
        r.Inflate(0, -1)

        ' get the height column should be
        Dim sDraw As SizeF = g.MeasureString(s, Me.TextBox.Font, Me.Width, mDrawTxt)
        Dim h As Integer = sDraw.Height + 15
        If mbAdjustHeight Then
            Try
                Dim pi As PropertyInfo = arHeights(rowNum).GetType().GetProperty("Height")

                ' get current height
                Dim curHeight As Integer = pi.GetValue(arHeights(rowNum), Nothing)

                ' adjust height
                If h > curHeight Then
                    pi.SetValue(arHeights(rowNum), h, Nothing)
                    Dim sz As Size = dg.Size
                    dg.Size = New Size(sz.Width - 1, sz.Height - 1)
                    dg.Size = sz
                End If
            Catch

                ' something wrong leave default height
                GetHeightList()
            End Try
        End If

        g.DrawString(s, MyBase.TextBox.Font, foreBrush, r, mDrawTxt)
        bPainted = True
    End Sub

    Public Property DataAlignment() As HorizontalAlignment
        Get
            Return mTxtAlign
        End Get
        Set(ByVal Value As HorizontalAlignment)
            mTxtAlign = Value
            If mTxtAlign = HorizontalAlignment.Center Then
                mDrawTxt.Alignment = StringAlignment.Center
            ElseIf mTxtAlign = HorizontalAlignment.Right Then
                mDrawTxt.Alignment = StringAlignment.Far
            Else
                mDrawTxt.Alignment = StringAlignment.Near
            End If
        End Set
    End Property

    Public Property AutoAdjustHeight() As Boolean
        Get
            Return mbAdjustHeight
        End Get
        Set(ByVal Value As Boolean)
            mbAdjustHeight = Value
            Try
                dg.Invalidate()
            Catch
            End Try
        End Set
    End Property
End Class

'Public Class oldMultilineColumn
'    Inherits DataGridTextBoxColumn

'    Private HAlignment As HorizontalAlignment
'    Private DrawFormat As New StringFormat
'    Private AdjustHeight As Boolean = True
'    Private dg As DataGrid
'    Private Heights As ArrayList

'    Public Property DataAlignment() As HorizontalAlignment
'        Get
'            Return HAlignment
'        End Get
'        Set(ByVal Value As HorizontalAlignment)
'            HAlignment = Value
'            If HAlignment = HorizontalAlignment.Center Then
'                DrawFormat.Alignment = StringAlignment.Center
'            ElseIf HAlignment = HorizontalAlignment.Right Then
'                DrawFormat.Alignment = StringAlignment.Far
'            Else
'                DrawFormat.Alignment = StringAlignment.Near
'            End If
'        End Set
'    End Property

'    Public Property AutoAdjustHeight() As Boolean
'        Get
'            Return AdjustHeight
'        End Get
'        Set(ByVal Value As Boolean)
'            AdjustHeight = Value
'            dg.Invalidate()
'        End Set
'    End Property

'    Public Sub New(ByVal MappingName As String)
'        MyBase.new()
'        Me.MappingName = MappingName
'        HAlignment = HorizontalAlignment.Left
'        DrawFormat.Alignment = StringAlignment.Near
'        MyBase.TextBox.TextAlign = HAlignment
'        MyBase.TextBox.Multiline = AdjustHeight
'    End Sub

'    Public Sub New(ByVal MappingName As String, _
'                   ByVal Width As Integer, _
'                   ByVal Alignment As HorizontalAlignment, _
'                   ByVal [ReadOnly] As Boolean, _
'                   ByVal HeaderText As String, _
'                   ByVal NullText As String, _
'                   ByVal DataSubType As EnumClass.DataSubType)

'        Me.New(MappingName)
'        Me.Alignment = Alignment
'        Me.Width = Width
'        Me.ReadOnly = [ReadOnly]
'        Me.HeaderText = HeaderText
'        Me.NullText = NullText
'    End Sub


'    Private Sub FillHeightArrayList()
'        Dim mi As MethodInfo = _
'           dg.GetType().GetMethod("get_DataGridRows", _
'           BindingFlags.FlattenHierarchy Or _
'           BindingFlags.IgnoreCase Or _
'           BindingFlags.Instance Or _
'           BindingFlags.NonPublic Or _
'           BindingFlags.Public Or _
'           BindingFlags.Static)
'        Dim dgRowArray As Array = CType(mi.Invoke(Me.dg, Nothing), Array)
'        Heights = New ArrayList
'        Dim dgRowHeight As Object
'        'For Each dgRowHeight In dgRowArray
'        '    If dgRowHeight.ToString().EndsWith _
'        '    ("DataGridRelationshipRow") = True _
'        '    Then
'        '        Heights.Add(dgRowHeight)
'        '    End If
'        'Next
'        For Each dgRowHeight In dgRowArray
'            If dgRowHeight.ToString().EndsWith("DataGridRelationshipRow") = True Then
'                Heights.Add(dgRowHeight)
'            End If
'        Next
'    End Sub

'    Protected Overloads Overrides Sub Edit(ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, _
'                                           ByVal bounds As System.Drawing.Rectangle, _
'                                           ByVal [readOnly] As Boolean, _
'                                           ByVal instantText As String, _
'                                           ByVal cellIsVisible As Boolean)

'        '  If IsEditable(source, rowNum) Then
'        MyBase.Edit(source, rowNum, bounds, [readOnly], instantText, cellIsVisible)
'        MyBase.TextBox.TextAlign = HAlignment
'        MyBase.TextBox.Multiline = AdjustHeight
'        '  End If

'    End Sub

'    Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, _
'                                        ByVal bounds As System.Drawing.Rectangle, _
'                                        ByVal source As System.Windows.Forms.CurrencyManager, _
'                                        ByVal rowNum As Integer, _
'                                        ByVal backBrush As System.Drawing.Brush, _
'                                        ByVal foreBrush As System.Drawing.Brush, _
'                                        ByVal alignToRight As Boolean)

'        Static bPainted As Boolean = False
'        If Not bPainted Then
'            dg = Me.DataGridTableStyle.DataGrid
'            FillHeightArrayList()
'        End If

'        'clear the cell
'        g.FillRectangle(backBrush, bounds)

'        'draw the value
'        Dim o As Object = Me.GetColumnValueAtRow([source], rowNum)
'        Dim s As String
'        If IsDBNull(o) Then
'            s = Me.NullText
'        Else
'            s = CStr(Me.GetColumnValueAtRow([source], rowNum))
'        End If

'        Dim r As New RectangleF(bounds.X, bounds.Y, bounds.Width, bounds.Height)
'        r.Inflate(0, -1)

'        ' get the height column should be
'        Dim sDraw As SizeF = g.MeasureString(s, Me.TextBox.Font, Me.Width, DrawFormat)
'        Dim h As Integer = CInt(sDraw.Height + 15)

'        If AdjustHeight Then
'            FillHeightArrayList()
'            Dim pi As PropertyInfo = Heights(rowNum).GetType().GetProperty("Height")
'            Dim curHeight As Integer = CInt(pi.GetValue(Heights(rowNum), Nothing))
'            If h > curHeight Then
'                pi.SetValue(Heights(rowNum), h, Nothing)


'                '' copied from class1
'                'Dim sz As Size = dg.Size
'                ''   dg.Size = New Size(sz.Width - 1, sz.Height - 1)
'                'dg.Size = New Size(sz.Width - 1, 300)
'                'dg.Size = sz


'            End If
'        End If

'        g.DrawString(s, MyBase.TextBox.Font, foreBrush, r, DrawFormat)
'        bPainted = True
'    End Sub

'    'Protected Overrides Sub Abort(ByVal rowNum As Integer)
'    'End Sub

'    'Protected Overrides Function Commit(ByVal dataSource As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer) As Boolean
'    'End Function

'    'Protected Overloads Overrides Sub Edit(ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal bounds As System.Drawing.Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
'    'End Sub

'    'Protected Overrides Function GetMinimumHeight() As Integer
'    'End Function

'    'Protected Overrides Function GetPreferredHeight(ByVal g As System.Drawing.Graphics, ByVal value As Object) As Integer
'    'End Function

'    'Protected Overrides Function GetPreferredSize(ByVal g As System.Drawing.Graphics, ByVal value As Object) As System.Drawing.Size
'    'End Function

'    'Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle, ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer)
'    'End Sub

'    'Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle, ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal alignToRight As Boolean)
'    'End Sub

'    Protected Overrides Sub Abort(ByVal rowNum As Integer)

'    End Sub

'    Protected Overrides Function Commit(ByVal dataSource As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer) As Boolean

'    End Function

'    Protected Overrides Function GetMinimumHeight() As Integer

'    End Function

'    Protected Overrides Function GetPreferredHeight(ByVal g As System.Drawing.Graphics, ByVal value As Object) As Integer

'    End Function

'    Protected Overrides Function GetPreferredSize(ByVal g As System.Drawing.Graphics, ByVal value As Object) As System.Drawing.Size

'    End Function

'    Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle, ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer)

'    End Sub

'    Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle, ByVal source As System.Windows.Forms.CurrencyManager, ByVal rowNum As Integer, ByVal alignToRight As Boolean)

'    End Sub
'End Class

'Public Class MessedUp
'    Inherits DataGridTextBoxColumn

'    Private mTxtAlign As HorizontalAlignment
'    Private mDrawTxt As New StringFormat
'    Private mbAdjustHeight As Boolean = True
'    Private m_intPreEditHeight As Integer
'    Private m_rownum As Integer
'    Dim WithEvents dg As DataGrid
'    Private arHeights As ArrayList

'    Private Sub GetHeightList()
'        Dim mi As MethodInfo = dg.GetType().GetMethod("get_DataGridRows", BindingFlags.FlattenHierarchy Or BindingFlags.IgnoreCase Or BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static)
'        Dim dgra As Array = CType(mi.Invoke(Me.dg, Nothing), Array)
'        arHeights = New ArrayList
'        Dim dgRowHeight As Object
'        For Each dgRowHeight In dgra
'            If dgRowHeight.ToString().EndsWith("DataGridRelationshipRow") = True Then
'                arHeights.Add(dgRowHeight)
'            End If
'        Next
'    End Sub

'    Public Sub New()
'        mTxtAlign = HorizontalAlignment.Left
'        mDrawTxt.Alignment = StringAlignment.Near
'        Me.ReadOnly = True

'        'Me.TextBox.WordWrap = True
'        'Me.TextBox.Multiline = True
'    End Sub

'    Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, _
'    ByVal bounds As System.Drawing.Rectangle, ByVal source As System.Windows.Forms.CurrencyManager, _
'    ByVal rowNum As Integer, ByVal backBrush As System.Drawing.Brush, _
'    ByVal foreBrush As System.Drawing.Brush, ByVal alignToRight As Boolean)

'        Static bPainted As Boolean = False
'        If Not bPainted Then
'            dg = Me.DataGridTableStyle.DataGrid
'            GetHeightList()
'        End If

'        'clear the cell
'        g.FillRectangle(backBrush, bounds)

'        'draw the value
'        Dim s As String = Me.GetColumnValueAtRow([source], rowNum).ToString()
'        Dim r As New RectangleF(bounds.X, bounds.Y, bounds.Width, bounds.Height)
'        r.Inflate(0, -1)

'        ' get the height column should be
'        Dim sDraw As SizeF = g.MeasureString(s, Me.TextBox.Font, Me.Width, mDrawTxt)
'        Dim h As Integer = sDraw.Height + 15
'        If mbAdjustHeight Then

'            Try
'                Dim pi As PropertyInfo = arHeights(rowNum).GetType().GetProperty("Height")

'                ' get current height
'                Dim curHeight As Integer = pi.GetValue(arHeights(rowNum), Nothing)

'                ' adjust height
'                If h > curHeight Then
'                    pi.SetValue(arHeights(rowNum), h, Nothing)
'                    'pi.SetValue(arHeights(rowNum), 27, Nothing)
'                    Dim sz As Size = dg.Size
'                    dg.Size = New Size(sz.Width - 1, sz.Height - 1)
'                    dg.Size = sz
'                End If

'            Catch
'                ' something wrong leave default height
'                GetHeightList()
'            End Try
'        End If

'        g.DrawString(s, MyBase.TextBox.Font, foreBrush, r, mDrawTxt)
'        bPainted = True
'    End Sub

'    Public Property DataAlignment() As HorizontalAlignment
'        Get
'            Return mTxtAlign
'        End Get
'        Set(ByVal Value As HorizontalAlignment)
'            mTxtAlign = Value
'            If mTxtAlign = HorizontalAlignment.Center Then
'                mDrawTxt.Alignment = StringAlignment.Center
'            ElseIf mTxtAlign = HorizontalAlignment.Right Then
'                mDrawTxt.Alignment = StringAlignment.Far
'            Else
'                mDrawTxt.Alignment = StringAlignment.Near
'            End If
'        End Set
'    End Property
'End Class

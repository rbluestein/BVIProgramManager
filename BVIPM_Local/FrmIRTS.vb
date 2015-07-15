Public Class FrmIRTS
    Inherits System.Windows.Forms.Form

    ' Response from DevExpress to my query whether its datagrid has multiline textbox columns. 
    'Hi Robert,
    'Yes, we have. The product is called XtraGrid. Its included in the VS2002/VS2003 compatible evaluation install available from: 
    'http://www.devexpress.com/Downloads/NET/DXperience/
    'Thanks,
    'Max


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
    Friend WithEvents dgIRTS As System.Windows.Forms.DataGrid
    Friend WithEvents lblHeader As System.Windows.Forms.Label
    Friend WithEvents lblSeparator As System.Windows.Forms.Label
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents lblTroubleCall As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnGo As System.Windows.Forms.Button
    Friend WithEvents cboClientID As System.Windows.Forms.ComboBox
    Friend WithEvents lblCategory As System.Windows.Forms.Label
    Friend WithEvents cboCategory As System.Windows.Forms.ComboBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents lblOpened As System.Windows.Forms.Label
    Friend WithEvents cboOpened As System.Windows.Forms.ComboBox
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents lblIRID As System.Windows.Forms.Label
    Friend WithEvents lblClientID As System.Windows.Forms.Label
    Friend WithEvents txtIrID As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmIRTS))
        Me.dgIRTS = New System.Windows.Forms.DataGrid
        Me.lblHeader = New System.Windows.Forms.Label
        Me.lblSeparator = New System.Windows.Forms.Label
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.lblTroubleCall = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.lblDescription = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.btnGo = New System.Windows.Forms.Button
        Me.cboClientID = New System.Windows.Forms.ComboBox
        Me.lblIRID = New System.Windows.Forms.Label
        Me.lblClientID = New System.Windows.Forms.Label
        Me.lblCategory = New System.Windows.Forms.Label
        Me.cboCategory = New System.Windows.Forms.ComboBox
        Me.lblStatus = New System.Windows.Forms.Label
        Me.cboStatus = New System.Windows.Forms.ComboBox
        Me.lblOpened = New System.Windows.Forms.Label
        Me.cboOpened = New System.Windows.Forms.ComboBox
        Me.lblProgress = New System.Windows.Forms.Label
        Me.txtIrID = New System.Windows.Forms.TextBox
        CType(Me.dgIRTS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgIRTS
        '
        Me.dgIRTS.AllowSorting = False
        Me.dgIRTS.CaptionVisible = False
        Me.dgIRTS.DataMember = ""
        Me.dgIRTS.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgIRTS.Location = New System.Drawing.Point(16, 168)
        Me.dgIRTS.Name = "dgIRTS"
        Me.dgIRTS.PreferredRowHeight = 0
        Me.dgIRTS.ReadOnly = True
        Me.dgIRTS.RowHeaderWidth = 0
        Me.dgIRTS.Size = New System.Drawing.Size(936, 432)
        Me.dgIRTS.TabIndex = 0
        '
        'lblHeader
        '
        Me.lblHeader.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblHeader.Location = New System.Drawing.Point(0, 0)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(984, 88)
        Me.lblHeader.TabIndex = 1
        '
        'lblSeparator
        '
        Me.lblSeparator.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeparator.Location = New System.Drawing.Point(0, 88)
        Me.lblSeparator.Name = "lblSeparator"
        Me.lblSeparator.Size = New System.Drawing.Size(968, 2)
        Me.lblSeparator.TabIndex = 2
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(32, 7)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(64, 73)
        Me.PictureBox3.TabIndex = 5
        Me.PictureBox3.TabStop = False
        '
        'lblTroubleCall
        '
        Me.lblTroubleCall.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblTroubleCall.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTroubleCall.Location = New System.Drawing.Point(136, 24)
        Me.lblTroubleCall.Name = "lblTroubleCall"
        Me.lblTroubleCall.Size = New System.Drawing.Size(544, 40)
        Me.lblTroubleCall.TabIndex = 6
        Me.lblTroubleCall.Text = "Issue Report and Tracking System"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(872, 608)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "Close"
        '
        'lblDescription
        '
        Me.lblDescription.Location = New System.Drawing.Point(656, 103)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(208, 16)
        Me.lblDescription.TabIndex = 11
        Me.lblDescription.Text = "Search Description"
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(656, 120)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(208, 20)
        Me.txtDescription.TabIndex = 12
        Me.txtDescription.Text = ""
        '
        'btnGo
        '
        Me.btnGo.Location = New System.Drawing.Point(872, 119)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.TabIndex = 13
        Me.btnGo.Text = "Go"
        '
        'cboClientID
        '
        Me.cboClientID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClientID.Location = New System.Drawing.Point(144, 120)
        Me.cboClientID.Name = "cboClientID"
        Me.cboClientID.Size = New System.Drawing.Size(121, 21)
        Me.cboClientID.TabIndex = 14
        '
        'lblIRID
        '
        Me.lblIRID.Location = New System.Drawing.Point(16, 103)
        Me.lblIRID.Name = "lblIRID"
        Me.lblIRID.Size = New System.Drawing.Size(72, 16)
        Me.lblIRID.TabIndex = 15
        Me.lblIRID.Text = "IRT#"
        '
        'lblClientID
        '
        Me.lblClientID.Location = New System.Drawing.Point(144, 103)
        Me.lblClientID.Name = "lblClientID"
        Me.lblClientID.Size = New System.Drawing.Size(72, 16)
        Me.lblClientID.TabIndex = 17
        Me.lblClientID.Text = "ClientID"
        '
        'lblCategory
        '
        Me.lblCategory.Location = New System.Drawing.Point(272, 103)
        Me.lblCategory.Name = "lblCategory"
        Me.lblCategory.Size = New System.Drawing.Size(72, 16)
        Me.lblCategory.TabIndex = 19
        Me.lblCategory.Text = "Category"
        '
        'cboCategory
        '
        Me.cboCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCategory.Location = New System.Drawing.Point(272, 120)
        Me.cboCategory.Name = "cboCategory"
        Me.cboCategory.Size = New System.Drawing.Size(121, 21)
        Me.cboCategory.TabIndex = 18
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(400, 103)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(72, 16)
        Me.lblStatus.TabIndex = 21
        Me.lblStatus.Text = "Status"
        '
        'cboStatus
        '
        Me.cboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatus.Location = New System.Drawing.Point(400, 120)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(121, 21)
        Me.cboStatus.TabIndex = 20
        '
        'lblOpened
        '
        Me.lblOpened.Location = New System.Drawing.Point(528, 103)
        Me.lblOpened.Name = "lblOpened"
        Me.lblOpened.Size = New System.Drawing.Size(72, 16)
        Me.lblOpened.TabIndex = 23
        Me.lblOpened.Text = "Opened"
        '
        'cboOpened
        '
        Me.cboOpened.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboOpened.Location = New System.Drawing.Point(528, 120)
        Me.cboOpened.Name = "cboOpened"
        Me.cboOpened.Size = New System.Drawing.Size(121, 21)
        Me.cboOpened.TabIndex = 22
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(16, 608)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(336, 23)
        Me.lblProgress.TabIndex = 24
        '
        'txtIrID
        '
        Me.txtIrID.Location = New System.Drawing.Point(16, 120)
        Me.txtIrID.Name = "txtIrID"
        Me.txtIrID.Size = New System.Drawing.Size(120, 20)
        Me.txtIrID.TabIndex = 25
        Me.txtIrID.Text = ""
        '
        'FrmIRTS
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(968, 638)
        Me.Controls.Add(Me.txtIrID)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.lblOpened)
        Me.Controls.Add(Me.cboOpened)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.cboStatus)
        Me.Controls.Add(Me.lblCategory)
        Me.Controls.Add(Me.cboCategory)
        Me.Controls.Add(Me.lblClientID)
        Me.Controls.Add(Me.lblIRID)
        Me.Controls.Add(Me.cboClientID)
        Me.Controls.Add(Me.btnGo)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lblTroubleCall)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.lblSeparator)
        Me.Controls.Add(Me.lblHeader)
        Me.Controls.Add(Me.dgIRTS)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(976, 672)
        Me.MinimumSize = New System.Drawing.Size(976, 672)
        Me.Name = "FrmIRTS"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "IRTS"
        CType(Me.dgIRTS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cDBase As New DBase
    Private cFilter As New Filter(cDBase)

#Region " Load events "
    Private Sub IRTS_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PrepareFilterResults As New Results

        Try

            Common.ResizeControls(Me, 976, 672)
            PrepareFilterResults = PrepareFilter()

            If Not PrepareFilterResults.Success Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, PrepareFilterResults.Message)
            End If

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #820: " & ex.Message)
        End Try
    End Sub

    Private Function PrepareFilter() As Results
        Dim i As Integer
        Dim MyResults As New Results

        Try

            cFilter.Add("irID", txtIrID, Filter.DataTypeEnum.Numeric, Filter.MatchTypeEnum.MatchExact, False)
            cFilter.Add("ClientID", cboClientID, Filter.DataTypeEnum.String, Filter.MatchTypeEnum.MatchLike, False)
            cFilter.Add("Category", cboCategory, Filter.DataTypeEnum.String, Filter.MatchTypeEnum.MatchExact, False)
            cFilter.Add("Status", cboStatus, Filter.DataTypeEnum.String, Filter.MatchTypeEnum.MatchExact, False)
            cFilter.AddSpecial("DateOpened", cboOpened, Filter.DataTypeEnum.String, "cast(datepart(yyyy, DateOpened) as varchar(4)) >= '|'")
            cFilter.Add("Description", txtDescription, Filter.DataTypeEnum.String, Filter.MatchTypeEnum.MatchLike, True)

            cFilter.PopulateDropdown(cboClientID, "ClientID")
            'cFilter.PopulateDropdown(cboAssignedTo, "AssignedTo")
            ' cFilter.PopulateDropdown(cboCategory, "Category")

            cboCategory.Items.Add(New ListItem("Enrollment App", "Enrollment App"))
            cboCategory.Items.Add(New ListItem("Confirmation Stateme", "Confirmation Statement"))
            cboCategory.Items.Add(New ListItem("DataBase", "DataBase"))

            cboStatus.Items.Add(New ListItem("", ""))
            cboStatus.Items.Add(New ListItem("O", "Open"))
            cboStatus.Items.Add(New ListItem("C", "Closed"))
            cboStatus.Items.Add(New ListItem("L", "Limbo"))
            cboStatus.Items.Add(New ListItem("R", "Review"))

            cboOpened.Items.Add(New ListItem("", ""))
            For i = Date.Now.Year To 2003 Step -1
                cboOpened.Items.Add(New ListItem(i.ToString, i.ToString))
            Next

            MyResults.Success = True
            MyResults.Message = "Error in PrepareFilter"
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #821: " & ex.Message
            Return MyResults
        End Try
    End Function
#End Region

#Region " Run report "
    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
        Dim PerformQueryResults As Results
        Try

            PerformQueryResults = PerformQuery()
            If Not PerformQueryResults.Success Then
                cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, PerformQueryResults.Message)
            End If

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #830: " & ex.Message)
        End Try
    End Sub

    Private Function PerformQuery() As Results
        Dim MyResults As New Results
        Dim GetDataResults As Results
        Dim PopulateDatagridResults As Results
        Dim RollingSuccess As Boolean = True

        Try

            Me.Cursor = Cursors.WaitCursor
            lblProgress.Text = "Query in progress..."
            lblProgress.Refresh()

            ' ___ Get data
            GetDataResults = GetData()
            If Not GetDataResults.Success Then
                RollingSuccess = False
                MyResults.Message = GetDataResults.Message
            End If

            If RollingSuccess Then

                ' ___ Populate DataGrid
                PopulateDatagridResults = ConfigureDatagrid()
                If Not PopulateDatagridResults.Success Then
                    RollingSuccess = False
                    MyResults.Message = PopulateDatagridResults.Message
                End If

            End If

            If RollingSuccess Then
                MyResults.Success = True
            Else
                MyResults.Success = False
            End If

            If cEnviro.dtIRTS.Rows.Count = 0 Then
                lblProgress.Text = "No records returned"
            Else
                lblProgress.Text = String.Empty
            End If

            Me.Cursor = Cursors.Default

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #831: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Function GetData() As Results
        Dim i As Integer
        Dim WhereClauseResults As New Results
        Dim MyResults As New Results
        Dim Sql As New System.Text.StringBuilder
        Dim QueryPack As DBase.QueryPack
        Dim WhereColl As New Collection

        Try

            Sql.Append("SELECT ClientID, irID, Description, Category, Status, Priority, AssignedTo, DateOpened, DateClosed, ")
            Sql.Append("LastUpdate = UserID + char(13) + char(10) +  Cast(datepart(m, ChangeDate) as varchar(2)) + '/' + cast(datepart(d, ChangeDate) as varchar(2)) + '/' + cast(datepart(yyyy, ChangeDate) as varchar(4))")
            Sql.Append(" FROM IRTS ")
            WhereClauseResults = cFilter.GetWhereClause
            If Not WhereClauseResults.Success Then
                MyResults.Success = False
                MyResults.Message = WhereClauseResults.Message
                Return MyResults
            End If
            Sql.Append(WhereClauseResults.ReturnStr)

            Sql.Append(" ORDER BY irID")

            '--FROM IRTS  where description like '%EXPLICIT%'  --  WHERE irID = 2271
            'FROM IRTS where ClientID like '%bvi%'

            QueryPack = cDBase.GetDTSourceIsText(Sql.ToString, cEnviro.IRTSConnString, False)
            If QueryPack.Success Then
                cEnviro.dtIRTS = QueryPack.dt
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Message = "Error in GetDBData. " & QueryPack.TechErrMsg
                Return MyResults
            End If

            cEnviro.dtIRTS.TableName = "IRTS"

            ' ___ Replace tilde with apostrope in Description field
            If cEnviro.dtIRTS.Rows.Count > 0 Then
                For i = 0 To cEnviro.dtIRTS.Rows.Count - 1
                    If Not IsDBNull(cEnviro.dtIRTS.Rows(i)("Description")) Then
                        cEnviro.dtIRTS.Rows(i)("Description") = Replace(cEnviro.dtIRTS.Rows(i)("Description"), "''", "'")
                        cEnviro.dtIRTS.Rows(i)("Description") = Replace(cEnviro.dtIRTS.Rows(i)("Description"), "~", "'")
                    End If
                Next
            End If


            'If cEnviro.dtIRTS.Rows.Count > 0 Then
            '    cEnviro.AnyRecordsReturned = True
            '    Dim dr As DataRow
            '    dr = cEnviro.dtIRTS.NewRow
            '    cEnviro.dtIRTS.Rows.Add(dr)
            '    cEnviro.dtIRTS.Rows(cEnviro.dtIRTS.Rows.Count - 1)("Description") = "End of report"
            '    For i = 1 To cEnviro.dtIRTS.Rows.Count
            '        cEnviro.dtIRTS.Rows(cEnviro.dtIRTS.Rows.Count - 1)("AssignedTo") &= vbCrLf
            '    Next
            'End If

            'Dim Col As New DataColumn("RecNo", GetType(System.Int64))
            'cEnviro.dtIRTS.Columns.Add(Col)
            'For i = 1 To cEnviro.dtIRTS.Rows.Count
            '    cEnviro.dtIRTS.Rows(i - 1)("RecNo") = i
            'Next

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #832: " & ex.Message
            Return MyResults
        End Try
    End Function
#End Region



#Region " ConfigureDatagrid "
    Public Function ConfigureDatagrid() As Results
        Dim MyResults As New Results

        Try

            Dim dgTS As New DataGridTableStyle
            dgTS.GridColumnStyles.Clear()
            dgTS.MappingName = cEnviro.dtIRTS.TableName
            dgTS.AllowSorting = False

            Dim fnt As Font
            fnt = dgIRTS.Font
            dgIRTS.Font = New Font(fnt.Name, 7, FontStyle.Regular)

            ' dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("RecNo", 30, HorizontalAlignment.Left, True, "RecNo", String.Empty, EnumClass.DataSubType.Integer))
            dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("irID", 30, HorizontalAlignment.Left, True, "IRT#", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("ClientID", 65, HorizontalAlignment.Left, True, "Client", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New MultiLineColumn("Description", 410, HorizontalAlignment.Left, True, "Description", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("Category", 85, HorizontalAlignment.Left, True, "Category", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("Status", 40, HorizontalAlignment.Left, True, "Status", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("Priority", 20, HorizontalAlignment.Left, True, "Pr", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New MultiLineColumn("AssignedTo", 70, HorizontalAlignment.Left, True, "Assigned To", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New MultiLineColumn)
            dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("DateOpened", 60, HorizontalAlignment.Left, True, "Opened", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("DateClosed", 60, HorizontalAlignment.Left, True, "Closed", String.Empty, EnumClass.DataSubType.String))
            dgTS.GridColumnStyles.Add(New TextBoxColumnExtended("LastUpdate", 85, HorizontalAlignment.Left, True, "Last Update", String.Empty, EnumClass.DataSubType.String))

            'For i = 1 To Coll.Count
            '    dgTS.GridColumnStyles.Add(Coll(i))
            'Next

            dgTS.AlternatingBackColor = Color.LightBlue

            If dgIRTS.TableStyles.Count > 0 Then
                dgIRTS.TableStyles.Clear()
            End If

            dgIRTS.TableStyles.Add(dgTS)
            dgIRTS.DataSource = cEnviro.dtIRTS

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #840: " & ex.Message
            Return MyResults
        End Try
    End Function
#End Region


    'Private Sub ResizeControls()
    '    Dim i As Integer
    '    Dim Size As Size
    '    Dim DevResScreenWidth As Single = 1280
    '    Dim DevResScreenHeight As Single = 800
    '    Dim DevResFormWidth As Single = 976
    '    Dim DevResFormHeight As Single = 672
    '    Dim CurResScreenWidth As Single = Screen.PrimaryScreen.Bounds.Width
    '    Dim CurResScreenHeight As Single = Screen.PrimaryScreen.Bounds.Height
    '    Dim BottomMenuPixels As Single = 28
    '    Dim TopMargin As Single = 28
    '    Dim SideAndBottomMargin As Single = 3

    '    Dim CurResFormWidth As Single
    '    Dim CurResFormHeight As Single
    '    Dim CurPercent As Single

    '    ' ___ Will the form fit on the screen as is?
    '    If Me.Width < CurResScreenWidth AndAlso Me.Height < CurResScreenHeight - BottomMenuPixels Then
    '        Exit Sub
    '    End If

    '    ' ___ The form is bigger than the screen. Resize the form.

    '    ' ___ Determine the new form dimensions
    '    If Me.Width > CurResScreenWidth Then
    '        CurResFormWidth = CurResScreenWidth
    '    Else
    '        CurResFormWidth = Me.Width
    '    End If
    '    If Me.Height > CurResScreenHeight - BottomMenuPixels Then
    '        CurResFormHeight = CurResScreenHeight - BottomMenuPixels
    '    Else
    '        CurResFormHeight = Me.Height
    '    End If

    '    Size = New Size(CurResFormWidth, CurResFormHeight)
    '    Me.MinimumSize = Size
    '    Me.MaximumSize = Size
    '    Me.Size = Size

    '    Dim Ctl As Control
    '    For Each Ctl In Me.Controls
    '        ' ___ Position
    '        'CurPercent = Ctl.Left / DevResFormWidth
    '        'Ctl.Left = CurPercent * CurResFormWidth
    '        'CurPercent = Ctl.Top / DevResFormHeight
    '        'Ctl.Top = CurPercent * CurResFormHeight

    '        '' ___ Size
    '        'CurPercent = Ctl.Width / DevResFormWidth
    '        'Ctl.Width = CurPercent * CurResFormWidth
    '        'CurPercent = Ctl.Height / DevResFormHeight
    '        'Ctl.Height = CurPercent * CurResFormHeight

    '        ' ___ Position
    '        CurPercent = (Ctl.Left + SideAndBottomMargin) / DevResFormWidth
    '        Ctl.Left = (CurPercent * CurResFormWidth) - SideAndBottomMargin
    '        CurPercent = (Ctl.Top + TopMargin) / DevResFormHeight
    '        Ctl.Top = (CurPercent * CurResFormHeight) - TopMargin

    '        ' ___ Size
    '        CurPercent = Ctl.Width / DevResFormWidth
    '        Ctl.Width = CurPercent * CurResFormWidth
    '        CurPercent = Ctl.Height / DevResFormHeight
    '        Ctl.Height = CurPercent * CurResFormHeight
    '    Next

    '    ' ___ Position of form on screen
    '    Me.StartPosition = FormStartPosition.CenterScreen

    'End Sub

    'Private Sub ResizeControls2(ByRef Ctl As Control, ByRef CurRes As Collection, ByRef DevRes As Collection)
    '    Dim CurLeftPercent As Single
    '    Dim CurWidthPercent As Single
    '    Dim CurTopPercent As Single
    '    Dim CurHeightPercent As Single

    '    Select Case Dimension
    '        Case "x"
    '            CurLeftPercent = Ctl.Left / Me.Width
    '            Ctl.Left = CurLeftPercent * Me.Width

    '            CurWidthPercent = Ctl.Width / Me.Width
    '            Ctl.Width = CurWidthPercent * Me.Width

    '        Case "y"
    '            CurTopPercent = Ctl.Top / Me.Height
    '            Ctl.Top = CurTopPercent * Me.Height

    '            CurHeightPercent = Ctl.Height / Me.Height
    '            Ctl.Height = CurHeightPercent * Me.Height
    '    End Select

    'End Sub

    Public Class ListItem
        Private cValue As String
        Private cText As String

        Public Sub New(ByVal Value As String, ByVal Text As String)
            cValue = Value
            cText = Text
        End Sub

        Public ReadOnly Property Value() As String
            Get
                Return cValue
            End Get
        End Property
        Public ReadOnly Property Text() As String
            Get
                Return cText
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return cText
        End Function
    End Class

    Public Class Filter
        Private cDBase As DBase
        Private cColl As New Collection

        Public Enum MatchTypeEnum
            MatchExact = 1
            MatchLike = 2
            MatchGTEqual = 3
        End Enum

        Public Enum DataTypeEnum
            [String] = 1
            Numeric = 2
        End Enum

        'Public Const MatchExact As String = "MatchExact"
        'Public Const MatchLike As String = "MatchLike"
        'Public Const MatchGTEqual As String = "MatchGTEqual"

        Public Sub New(ByRef DBase As DBase)
            cDBase = DBase
        End Sub

        Public Sub Add(ByVal FldName As String, ByVal Ctl As Control, ByVal DataType As DataTypeEnum, ByVal MatchType As MatchTypeEnum, ByVal ConvertTildes As Boolean)
            cColl.Add(New FilterItem(FldName, Ctl, DataType, MatchType, ConvertTildes))
        End Sub

        Public Sub AddSpecial(ByVal FldName As String, ByVal Ctl As Control, ByVal DataType As DataTypeEnum, ByVal WhereStr As String)
            cColl.Add(New FilterItem(FldName, Ctl, True, DataType, WhereStr))
        End Sub

        Public Function PopulateDropdown(ByRef Dropdown As ComboBox, ByVal Value As String)
            Dim i, j As Integer
            Dim MyResults As New Results
            Dim Sql As String
            Dim QueryPack As DBase.QueryPack
            Dim dt As DataTable
            Dim Coll As New Collection
            Dim Box As Object
            Dim RawList(0) As Object
            Dim CompValue As String = String.Empty

            Try

                Sql = "SELECT DISTINCT " & Value & " FROM Irts ORDER BY " & Value
                QueryPack = cDBase.GetDTSourceIsText(Sql.ToString, cEnviro.IRTSConnString, False)
                If QueryPack.Success Then
                    MyResults.Success = True
                Else
                    MyResults.Success = False
                    MyResults.Message = "Error in GetDBData. " & QueryPack.TechErrMsg
                    Return MyResults
                End If

                dt = QueryPack.dt

                ' ___ Get a list of values
                For i = 0 To dt.Rows.Count - 1
                    If InStr(dt.Rows(i)(0), ",") Then

                        Box = Split(dt.Rows(i)(0), ",")
                        For j = 0 To Box.GetUpperBound(0)
                            ReDim Preserve RawList(RawList.GetUpperBound(0) + 1)
                            RawList(RawList.GetUpperBound(0)) = Trim(Box(j))
                        Next

                    Else
                        If (Not IsDBNull(dt.Rows(i)(0))) AndAlso dt.Rows(i)(0).length > 0 Then
                            ReDim Preserve RawList(RawList.GetUpperBound(0) + 1)
                            RawList(RawList.GetUpperBound(0)) = Trim(dt.Rows(i)(0))
                        End If
                    End If
                Next

                ' ___ Sort the list
                RawList.Sort(RawList)

                ' ___ Add unique values to dropdown
                Dropdown.Items.Add(New ListItem("", ""))
                For i = 0 To RawList.GetUpperBound(0)
                    If RawList(i) <> CompValue Then
                        Dim li As New ListItem(RawList(i), RawList(i))
                        Dropdown.Items.Add(li)
                        CompValue = RawList(i)
                    End If
                Next

            Catch ex As Exception
                Throw New Exception("Error in Filter: PopulateDropdown for " & Value)
            End Try
        End Function

        Public Function GetWhereClause() As Results
            Dim MyResults As New Results
            Dim i As Integer
            Dim WhereColl As New Collection
            Dim FilterItem As FilterItem
            Dim FilterText As String = Nothing
            Dim Sql As New System.Text.StringBuilder
            Dim ItemAdded As Boolean
            Dim WhereStrAdjust As String

            Try

                For i = 1 To cColl.Count
                    FilterItem = cColl(i)
                    ItemAdded = False

                    If FilterItem.IsTextbox AndAlso Trim(FilterItem.Textbox.TextLength > 0) Then
                        FilterText = Trim(FilterItem.Textbox.Text)
                        ItemAdded = True
                    ElseIf FilterItem.IsDropdown AndAlso FilterItem.Dropdown.SelectedIndex > -1 AndAlso FilterItem.Dropdown.SelectedItem.value.length > 0 Then
                        FilterText = FilterItem.Dropdown.SelectedItem.value
                        ItemAdded = True
                    End If

                    If ItemAdded Then

                        If FilterItem.ConvertTildes Then
                            FilterText = Replace(FilterText, "'", "~")
                        End If

                        If FilterItem.IsSpecial Then
                            WhereStrAdjust = Replace(FilterItem.WhereStr, "|", FilterItem.Dropdown.SelectedItem.value)
                            WhereColl.Add(WhereStrAdjust)
                        Else
                            If FilterItem.DataType = DataTypeEnum.String Then
                                Select Case FilterItem.MatchType
                                    Case MatchTypeEnum.MatchExact
                                        WhereColl.Add(FilterItem.FldName & " = '" & FilterText & "'")
                                    Case MatchTypeEnum.MatchLike
                                        WhereColl.Add(FilterItem.FldName & " like '%" & FilterText & "%' ")
                                    Case MatchTypeEnum.MatchGTEqual
                                        WhereColl.Add(FilterItem.FldName & " >= '" & FilterText & "'")
                                End Select
                            ElseIf FilterItem.DataType = DataTypeEnum.Numeric Then
                                Select Case FilterItem.MatchType
                                    Case MatchTypeEnum.MatchExact
                                        WhereColl.Add(FilterItem.FldName & " = " & FilterText)
                                    Case MatchTypeEnum.MatchGTEqual
                                        WhereColl.Add(FilterItem.FldName & " >= " & FilterText)
                                End Select
                            End If
                        End If
                    End If

                Next

                Sql.Append(" ")
                If WhereColl.Count > 0 Then
                    For i = 1 To WhereColl.Count
                        If i = 1 Then
                            Sql.Append(" WHERE " & WhereColl(1))
                        Else
                            Sql.Append(" AND " & WhereColl(i))
                        End If
                    Next
                End If

                MyResults.Success = True
                MyResults.ReturnStr = Sql.ToString
                Return MyResults

            Catch ex As Exception
                MyResults.Success = False
                MyResults.Message = "Error in Filter:GetWhereClause: " & ex.Message
                Return MyResults
            End Try

        End Function

        Public Class FilterItem
            Private cFldName As String
            Private cIsTextbox As Boolean
            Private cIsDropdown As Boolean
            Private cTextbox As Textbox
            Private cDropdown As ComboBox
            Private cMatchType As MatchTypeEnum
            Private cIsSpecial As Boolean
            Private cDataType As DataTypeEnum
            Private cWhereStr As String
            Private cConvertTildes As Boolean

            Public Sub New(ByVal FldName As String, ByRef Ctl As Control, ByVal DataType As DataTypeEnum, ByVal MatchType As MatchTypeEnum, ByVal ConvertTildes As Boolean)
                If TypeOf (Ctl) Is TextBox Then
                    cTextbox = Ctl
                    cIsTextbox = True
                ElseIf TypeOf (Ctl) Is ComboBox Then
                    cDropdown = Ctl
                    cIsDropdown = True
                End If
                cFldName = FldName
                cDataType = DataType
                cMatchType = MatchType
                cConvertTildes = ConvertTildes
            End Sub

            Public Sub New(ByVal FldName As String, ByRef Ctl As Control, ByVal IsSpecial As Boolean, ByVal DataType As DataTypeEnum, ByVal WhereStr As String)
                If TypeOf (Ctl) Is TextBox Then
                    cTextbox = Ctl
                    cIsTextbox = True
                ElseIf TypeOf (Ctl) Is ComboBox Then
                    cDropdown = Ctl
                    cIsDropdown = True
                End If
                cFldName = FldName
                cDataType = DataType
                cIsSpecial = IsSpecial
                cWhereStr = WhereStr
                cConvertTildes = False
            End Sub

            Public ReadOnly Property DataType() As DataTypeEnum
                Get
                    Return cDataType
                End Get
            End Property
            Public ReadOnly Property FldName() As String
                Get
                    Return cFldName
                End Get
            End Property
            Public ReadOnly Property IsTextbox() As Boolean
                Get
                    Return cIsTextbox
                End Get
            End Property
            Public ReadOnly Property IsDropdown() As Boolean
                Get
                    Return cIsDropdown
                End Get
            End Property
            Public ReadOnly Property Dropdown() As ComboBox
                Get
                    Return cDropdown
                End Get
            End Property
            Public ReadOnly Property Textbox() As Textbox
                Get
                    Return cTextbox
                End Get
            End Property
            Public ReadOnly Property MatchType() As String
                Get
                    Return cMatchType
                End Get
            End Property
            Public ReadOnly Property IsSpecial() As Boolean
                Get
                    Return cIsSpecial
                End Get
            End Property
            Public ReadOnly Property WhereStr() As String
                Get
                    Return cWhereStr
                End Get
            End Property
            Public ReadOnly Property ConvertTildes() As Boolean
                Get
                    Return cConvertTildes
                End Get
            End Property
        End Class
    End Class

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    'Private Sub Datagrid1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgIRTS.MouseUp
    '    Return

    '    Dim Row As Integer
    '    Dim NewRecordNum As String
    '    Dim Grid As DataGrid
    '    Dim HitInfo As DataGrid.HitTestInfo

    '    Try

    '        Dim cm As CurrencyManager = CType(Me.BindingContext(dgIRTS.DataSource, dgIRTS.DataMember), CurrencyManager)
    '        CType(cm.List, DataView).AllowNew = False

    '        Grid = CType(sender, DataGrid)
    '        HitInfo = Grid.HitTest(e.X, e.Y)
    '        If HitInfo.Row = -1 Then
    '            ConfigureDatagrid()
    '            PerformQuery()
    '        Else
    '            Return
    '        End If

    '    Catch ex As Exception
    '        'Dim ErrorObj As New ErrorClass
    '        'ErrorObj.SimpleHandleError(EnumClass.ErrorSourceUserAppIOItem.AppError, "[DatagridRB.MouseUp] Error", False)
    '    End Try

    'End Sub
End Class

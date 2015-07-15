Imports CDO
Imports System.Drawing.Imaging

' Microsoft CDO 1.21 Library

' good
' http://www.kdkeys.net/forums/thread/704.aspx

Public Class FrmTroubleCall
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
    Friend WithEvents lblHeader As System.Windows.Forms.Label
    Friend WithEvents lblSeparator As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents lblTroubleCall As System.Windows.Forms.Label
    Friend WithEvents lblViewIRTS As System.Windows.Forms.Label
    Friend WithEvents btnViewIRTS As System.Windows.Forms.Button
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents btnSubmit As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblToCaption As System.Windows.Forms.Label
    Friend WithEvents lblccCaption As System.Windows.Forms.Label
    Friend WithEvents lblSubjCaption As System.Windows.Forms.Label
    Friend WithEvents lblToText As System.Windows.Forms.Label
    Friend WithEvents lblSubjText As System.Windows.Forms.Label
    Friend WithEvents ddClientID As System.Windows.Forms.ComboBox
    Friend WithEvents lblClientCaption As System.Windows.Forms.Label
    Friend WithEvents llViewAttachment As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnScreenshot As System.Windows.Forms.Button
    Friend WithEvents txtccText As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmTroubleCall))
        Me.lblHeader = New System.Windows.Forms.Label
        Me.lblSeparator = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.lblTroubleCall = New System.Windows.Forms.Label
        Me.btnSubmit = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblViewIRTS = New System.Windows.Forms.Label
        Me.btnViewIRTS = New System.Windows.Forms.Button
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.lblToCaption = New System.Windows.Forms.Label
        Me.lblccCaption = New System.Windows.Forms.Label
        Me.lblSubjCaption = New System.Windows.Forms.Label
        Me.lblToText = New System.Windows.Forms.Label
        Me.lblSubjText = New System.Windows.Forms.Label
        Me.ddClientID = New System.Windows.Forms.ComboBox
        Me.lblClientCaption = New System.Windows.Forms.Label
        Me.llViewAttachment = New System.Windows.Forms.LinkLabel
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnScreenshot = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtccText = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'lblHeader
        '
        Me.lblHeader.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblHeader.Location = New System.Drawing.Point(0, 0)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(616, 88)
        Me.lblHeader.TabIndex = 0
        '
        'lblSeparator
        '
        Me.lblSeparator.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeparator.Location = New System.Drawing.Point(0, 89)
        Me.lblSeparator.Name = "lblSeparator"
        Me.lblSeparator.Size = New System.Drawing.Size(616, 2)
        Me.lblSeparator.TabIndex = 1
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(40, 16)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(64, 56)
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(24, 6)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(80, 80)
        Me.PictureBox3.TabIndex = 4
        Me.PictureBox3.TabStop = False
        '
        'lblTroubleCall
        '
        Me.lblTroubleCall.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblTroubleCall.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTroubleCall.Location = New System.Drawing.Point(112, 24)
        Me.lblTroubleCall.Name = "lblTroubleCall"
        Me.lblTroubleCall.Size = New System.Drawing.Size(240, 40)
        Me.lblTroubleCall.TabIndex = 5
        Me.lblTroubleCall.Text = "Submit Trouble Call"
        '
        'btnSubmit
        '
        Me.btnSubmit.Location = New System.Drawing.Point(536, 308)
        Me.btnSubmit.Name = "btnSubmit"
        Me.btnSubmit.Size = New System.Drawing.Size(72, 23)
        Me.btnSubmit.TabIndex = 7
        Me.btnSubmit.Text = "Submit"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(456, 308)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(72, 23)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        '
        'lblViewIRTS
        '
        Me.lblViewIRTS.Location = New System.Drawing.Point(8, 352)
        Me.lblViewIRTS.Name = "lblViewIRTS"
        Me.lblViewIRTS.Size = New System.Drawing.Size(504, 16)
        Me.lblViewIRTS.TabIndex = 10
        Me.lblViewIRTS.Text = "Would you like to view a list of outstanding IRTS items? You may find information" & _
        " about this issue."
        '
        'btnViewIRTS
        '
        Me.btnViewIRTS.Location = New System.Drawing.Point(536, 352)
        Me.btnViewIRTS.Name = "btnViewIRTS"
        Me.btnViewIRTS.Size = New System.Drawing.Size(72, 23)
        Me.btnViewIRTS.TabIndex = 11
        Me.btnViewIRTS.Text = "View IRTS"
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(8, 184)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDescription.Size = New System.Drawing.Size(600, 112)
        Me.txtDescription.TabIndex = 13
        Me.txtDescription.Text = ""
        '
        'lblToCaption
        '
        Me.lblToCaption.Location = New System.Drawing.Point(8, 96)
        Me.lblToCaption.Name = "lblToCaption"
        Me.lblToCaption.Size = New System.Drawing.Size(32, 16)
        Me.lblToCaption.TabIndex = 14
        Me.lblToCaption.Text = "To:"
        Me.lblToCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblccCaption
        '
        Me.lblccCaption.Location = New System.Drawing.Point(8, 116)
        Me.lblccCaption.Name = "lblccCaption"
        Me.lblccCaption.Size = New System.Drawing.Size(32, 16)
        Me.lblccCaption.TabIndex = 15
        Me.lblccCaption.Text = "cc:"
        Me.lblccCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSubjCaption
        '
        Me.lblSubjCaption.Location = New System.Drawing.Point(8, 137)
        Me.lblSubjCaption.Name = "lblSubjCaption"
        Me.lblSubjCaption.Size = New System.Drawing.Size(32, 16)
        Me.lblSubjCaption.TabIndex = 16
        Me.lblSubjCaption.Text = "Subj:"
        Me.lblSubjCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblToText
        '
        Me.lblToText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblToText.Location = New System.Drawing.Point(48, 96)
        Me.lblToText.Name = "lblToText"
        Me.lblToText.Size = New System.Drawing.Size(560, 16)
        Me.lblToText.TabIndex = 17
        Me.lblToText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSubjText
        '
        Me.lblSubjText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSubjText.Location = New System.Drawing.Point(48, 138)
        Me.lblSubjText.Name = "lblSubjText"
        Me.lblSubjText.Size = New System.Drawing.Size(560, 16)
        Me.lblSubjText.TabIndex = 19
        Me.lblSubjText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ddClientID
        '
        Me.ddClientID.Location = New System.Drawing.Point(48, 158)
        Me.ddClientID.Name = "ddClientID"
        Me.ddClientID.Size = New System.Drawing.Size(152, 21)
        Me.ddClientID.TabIndex = 20
        '
        'lblClientCaption
        '
        Me.lblClientCaption.Location = New System.Drawing.Point(0, 159)
        Me.lblClientCaption.Name = "lblClientCaption"
        Me.lblClientCaption.Size = New System.Drawing.Size(40, 16)
        Me.lblClientCaption.TabIndex = 21
        Me.lblClientCaption.Text = "Client:"
        Me.lblClientCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'llViewAttachment
        '
        Me.llViewAttachment.Location = New System.Drawing.Point(448, 168)
        Me.llViewAttachment.Name = "llViewAttachment"
        Me.llViewAttachment.Size = New System.Drawing.Size(152, 16)
        Me.llViewAttachment.TabIndex = 22
        Me.llViewAttachment.TabStop = True
        Me.llViewAttachment.Text = "View Original Attachment"
        Me.llViewAttachment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(0, 344)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(616, 2)
        Me.Label1.TabIndex = 23
        '
        'btnScreenshot
        '
        Me.btnScreenshot.Location = New System.Drawing.Point(376, 308)
        Me.btnScreenshot.Name = "btnScreenshot"
        Me.btnScreenshot.Size = New System.Drawing.Size(72, 23)
        Me.btnScreenshot.TabIndex = 24
        Me.btnScreenshot.Text = "Screenshot"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 300)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(368, 40)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "The Submit Trouble Call form takes a screenshot when it loads. You may attach an " & _
        "additional screenshot each time you click the Screenshot button. Maximum of thre" & _
        "e additional screenshots."
        '
        'txtccText
        '
        Me.txtccText.Location = New System.Drawing.Point(48, 115)
        Me.txtccText.Name = "txtccText"
        Me.txtccText.Size = New System.Drawing.Size(560, 20)
        Me.txtccText.TabIndex = 26
        Me.txtccText.Text = ""
        '
        'FrmTroubleCall
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(616, 382)
        Me.Controls.Add(Me.txtccText)
        Me.Controls.Add(Me.btnScreenshot)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.llViewAttachment)
        Me.Controls.Add(Me.lblClientCaption)
        Me.Controls.Add(Me.ddClientID)
        Me.Controls.Add(Me.lblSubjText)
        Me.Controls.Add(Me.lblToText)
        Me.Controls.Add(Me.lblSubjCaption)
        Me.Controls.Add(Me.lblccCaption)
        Me.Controls.Add(Me.lblToCaption)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.btnViewIRTS)
        Me.Controls.Add(Me.lblViewIRTS)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSubmit)
        Me.Controls.Add(Me.lblTroubleCall)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblSeparator)
        Me.Controls.Add(Me.lblHeader)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(624, 416)
        Me.MinimumSize = New System.Drawing.Size(624, 416)
        Me.Name = "FrmTroubleCall"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "TroubleCall"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cFormMain As Form
    Private cSelectedTab As Integer
    Private cScreenCaptureColl As New Collection

    Public Overloads Function ShowDialog(ByRef FormMain As Form, ByVal SelectedTab As Integer) As System.Windows.Forms.DialogResult
        cFormMain = FormMain
        cSelectedTab = SelectedTab
        MyBase.ShowDialog()
    End Function

    Private Sub FrmTroubleCall_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        Dim DistinctClientID As String = String.Empty
        Dim CurClientID As String

        Try

            Common.ResizeControls(Me, 400, 488)
            If cSelectedTab = 0 Then
                lblToText.Text = "HelpDesk@benefitvision.com;" & cEnviro.SiteId & "_Supervisors@benefitvision.com"
            Else
                lblToText.Text = "CRM@benefitvision.com;" & cEnviro.SiteId & "_Supervisors@benefitvision.com"
            End If
            txtccText.Text = System.Environment.UserName & "@benefitvision.com;"
            lblSubjText.Text = "Trouble Call - " & cEnviro.UserId & "/" & cEnviro.ComputerId
            If cSelectedTab = 1 Then
                lblSubjText.Text &= " - TEST"
            ElseIf cSelectedTab = 2 Then
                lblSubjText.Text &= " - TRAINING"
            End If

            txtDescription.Text = "EmpID:" & vbCrLf & "Last 4 of SSN:" & vbCrLf & "Birth Date:" & vbCrLf & "Last Name:" & vbCrLf & "First Name:" & vbCrLf & "Division/Plant (if applicable):" & vbCrLf & "Phone #:" & vbCrLf & "Description:"

            'cScreenCaptureFileName = "ScreenCapture_" & System.Environment.UserName & "_" & Date.Now.ToUniversalTime.AddHours(-5).ToString("yyyyMMddssff") & ".gif"
            For i = 0 To cEnviro.dtProgs.Rows.Count - 1
                CurClientID = cEnviro.GetValue(i, Progs.ClientID).Trim
                If CurClientID <> DistinctClientID Then
                    DistinctClientID = CurClientID
                    ddClientID.Items.Add(DistinctClientID)
                End If
            Next

            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.InformationNoLog, "Capturing screen")
            CaptureScreen()

        Catch ex As Exception
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Error, "Error #920: " & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Clicked cancel trouble call")
        DeleteScreenCaptures()
        Me.Close()
    End Sub

    Private Sub btnViewIRTS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewIRTS.Click
        Dim FormIRTS As New FrmIRTS
        FormIRTS.ShowDialog()

        ' RaiseEvent UpdateStatusEvent("Launching " & cEnviro.GetValue(RowNum, Progs.ClientID) & "-" & cEnviro.GetValue(RowNum, Progs.Description) & " in browser", False)
        'Try
        '    System.Diagnostics.Process.Start("https://netserver.benefitvision.com/client")
        'Catch ex As Exception
        '    Stop
        'End Try
    End Sub

    Private Sub CaptureScreen()
        Dim FileFullPath As String

        FileFullPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "\ScreenCapture_" & System.Environment.UserName & "_" & Date.Now.ToString("yyyyMMddssff") & ".gif"
        cScreenCaptureColl.Add(FileFullPath)
        cFormMain.Visible = False
        Me.Visible = False
        Common.CaptureScreen(FileFullPath)
        If cScreenCaptureColl.Count = 4 Then
            btnScreenshot.Enabled = False
        End If
        Me.Visible = True
        cFormMain.Visible = True
    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim i As Integer
        Dim Coll As New Collection
        Dim SendEmailResults As New Results
        Dim Box As String() = Nothing
        Dim CCText As String
        Dim BadCC As Boolean = False

        Try
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Clicked submit trouble report")

            ' ___ Test for semicolons in cc
            CCText = txtccText.Text

            If CCText.Length > 0 Then

                Box = Split(CCText, "@")
                If Box.GetUpperBound(0) = 0 Then
                    BadCC = True
                Else
                    For i = 1 To Box.GetUpperBound(0) - 1
                        If InStr(Box(i), ";") = 0 Then
                            BadCC = True
                            Exit For
                        End If
                    Next
                End If

            End If

            If BadCC Then
                MessageBox.Show("Invalid cc entry. Format is: a@b.com;x@y.com.")
                Exit Sub
            End If

            ' ___ Test for domain
            If CCText.Length > 0 AndAlso Box.GetUpperBound(0) > 0 Then
                For i = 1 To Box.GetUpperBound(0)
                    If Box(i).Length < 17 OrElse Box(i).ToLower.Substring(0, 17) <> "benefitvision.com" Then
                        BadCC = True
                        Exit For
                    End If
                Next
            End If
            If BadCC Then
                MessageBox.Show("Invalid cc entry. Email can be sent to recipients @benefitvision.com only.")
                Exit Sub
            End If

            If ddClientID.SelectedItem = String.Empty Then
                MessageBox.Show("Please select a client.")
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            Coll.Add(cEnviro.LogFileFullPath)
            For i = 1 To cScreenCaptureColl.Count
                Coll.Add(cScreenCaptureColl(i))
            Next

            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.Information, "Sending email")

            SendEmailResults = Common.SendEmail(lblToText.Text, System.Environment.UserName & "@benefitvision.com", txtccText.Text, lblSubjText.Text, txtDescription.Text, Coll)
            DeleteScreenCaptures()
            Me.Cursor = Cursors.Default
            If SendEmailResults.Success Then
                MessageBox.Show("An email has been sent.")
            Else
                MessageBox.Show("Unable to send email. " & SendEmailResults.Message)
            End If
            Me.Close()

        Catch ex As Exception
            MessageBox.Show("Unable to send email.")
            Try
                DeleteScreenCaptures()
            Catch
            End Try
            Me.Close()
        End Try

        'Dim FileFullPath As String
        'Dim Subject As String
        'Dim schema As String

        'Try

        '    If ddClientID.SelectedItem = String.Empty Then
        '        MessageBox.Show("Please select a client.")
        '        Exit Sub
        '    End If

        '    FileFullPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & cFileName
        '    Dim CDOConfig As New CDO.Configuration
        '    schema = "http://schemas.microsoft.com/cdo/configuration/"
        '    CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2
        '    ' CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
        '    CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "mail.benefitvision.com"
        '    CDOConfig.Fields.Update()

        '    Dim iMsg As New CDO.Message
        '    iMsg.To = lblToText.Text
        '    iMsg.From = System.Environment.UserName & "@benefitvision.com"
        '    iMsg.CC = lblccText.Text
        '    iMsg.Subject = lblSubjText.Text
        '    iMsg.AddAttachment(FileFullPath)

        '    iMsg.Configuration = CDOConfig
        '    iMsg.TextBody = txtDescription.Text

        '    '   iMsg.Send()

        '    ' ___ Clean up
        '    CDOConfig = Nothing
        '    iMsg = Nothing
        '    System.IO.File.Delete(cFileName)

        '    MessageBox.Show("An email has been sent.")
        '    Me.Close()

        'Catch ex As Exception
        '    Try
        '        If System.IO.File.Exists(cFileName) Then
        '            System.IO.File.Delete(cFileName)
        '        End If
        '    Catch
        '    End Try
        'End Try
        ''cErrorLog.Write("Error in FrmTroubleCall_btnSubmitClick: " & ex.Message, True)
    End Sub

    Private Sub llViewAttachment_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llViewAttachment.LinkClicked
        Dim FormScreenCaptureView As New FrmScreenCaptureView
        FormScreenCaptureView.ShowDialog(cScreenCaptureColl(1))
    End Sub

    Private Sub ddClientID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddClientID.SelectedIndexChanged
        Dim Client As String
        Client = ddClientID.SelectedItem
        '  lblSubjText.Text = Client & " Trouble Call - " & cEnviro.UserId & "/" & cEnviro.ComputerId & "  " & Date.Now.ToUniversalTime.AddHours(-5).ToString & "  " & IIf(Client.ToUpper = "TRAINING", "TEST", String.Empty)
        lblSubjText.Text = Client & " " & lblSubjText.Text & "  " & Date.Now.ToUniversalTime.AddHours(-5).ToString
    End Sub

    Private Sub FrmTroubleCall_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            cAppAdmin.Report(ApplicationAdministrator.ReportTypeEnum.InformationNoLog, "Closing trouble report form")
            DeleteScreenCaptures()
        Catch
        End Try
    End Sub

    Private Sub DeleteScreenCaptures()
        Dim i As Integer
        If cScreenCaptureColl.Count > 0 Then
            For i = 1 To cScreenCaptureColl.Count
                If System.IO.File.Exists(cScreenCaptureColl(i)) Then
                    Try
                        System.IO.File.Delete(cScreenCaptureColl(i))
                    Catch
                    End Try
                End If
            Next
        End If
    End Sub

    Private Sub btnScreenshot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScreenshot.Click
        CaptureScreen()
    End Sub
End Class

Public Class FrmMessage
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
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lblHeaderText As System.Windows.Forms.Label
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblHeader = New System.Windows.Forms.Label
        Me.lblSeparator = New System.Windows.Forms.Label
        Me.lblHeaderText = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.txtMessage = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'lblHeader
        '
        Me.lblHeader.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblHeader.Location = New System.Drawing.Point(0, 0)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(451, 47)
        Me.lblHeader.TabIndex = 2
        '
        'lblSeparator
        '
        Me.lblSeparator.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeparator.Location = New System.Drawing.Point(0, 48)
        Me.lblSeparator.Name = "lblSeparator"
        Me.lblSeparator.Size = New System.Drawing.Size(451, 2)
        Me.lblSeparator.TabIndex = 3
        '
        'lblHeaderText
        '
        Me.lblHeaderText.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblHeaderText.Font = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeaderText.Location = New System.Drawing.Point(8, 10)
        Me.lblHeaderText.Name = "lblHeaderText"
        Me.lblHeaderText.Size = New System.Drawing.Size(248, 32)
        Me.lblHeaderText.TabIndex = 7
        Me.lblHeaderText.Text = "An error has occurred"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(352, 212)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(8, 64)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.ReadOnly = True
        Me.txtMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtMessage.Size = New System.Drawing.Size(424, 136)
        Me.txtMessage.TabIndex = 10
        Me.txtMessage.Text = ""
        '
        'FrmMessage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(443, 246)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lblHeaderText)
        Me.Controls.Add(Me.lblSeparator)
        Me.Controls.Add(Me.lblHeader)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(451, 280)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(451, 280)
        Me.Name = "FrmMessage"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Error"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cMessage As String

    Public Overloads Function ShowDialog(ByVal Message As String) As Windows.Forms.DialogResult
        cMessage = Message
        MyBase.ShowDialog()
    End Function

    Private Sub FrmMessagl_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Error"
        lblHeaderText.Text = "An error has occurred"
        txtMessage.Text = cMessage
        txtMessage.SelectionStart = txtMessage.Text.Length
        txtMessage.ScrollToCaret()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class

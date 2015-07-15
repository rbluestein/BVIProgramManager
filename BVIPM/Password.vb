Public Class Password
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents txtUserPassword As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtUserPassword = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(112, 88)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 23)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "OK"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(248, 32)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Program Manager will try to connect to the BVI network drive. Please enter your p" & _
        "assword:"
        '
        'txtUserPassword
        '
        Me.txtUserPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserPassword.Location = New System.Drawing.Point(16, 48)
        Me.txtUserPassword.Name = "txtUserPassword"
        Me.txtUserPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtUserPassword.Size = New System.Drawing.Size(248, 20)
        Me.txtUserPassword.TabIndex = 0
        Me.txtUserPassword.Text = ""
        '
        'Password
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(280, 118)
        Me.Controls.Add(Me.txtUserPassword)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(288, 152)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(288, 152)
        Me.Name = "Password"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Password required"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cUserPassword As String = String.Empty

    Public Overloads Function ShowDialog() As System.Windows.Forms.DialogResult
        MyBase.ShowDialog()
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        CloseForm()
    End Sub

    Public ReadOnly Property UserPassword() As String
        Get
            Return cUserPassword
        End Get
    End Property

    Private Sub txtUserPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUserPassword.KeyPress
        If Asc(e.KeyChar) = 13 Then
            CloseForm()
        End If
    End Sub

    Private Sub CloseForm()
        cUserPassword = Trim(txtUserPassword.Text)
        Me.Hide()
    End Sub
End Class
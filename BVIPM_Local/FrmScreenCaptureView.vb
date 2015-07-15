Public Class FrmScreenCaptureView
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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmScreenCaptureView))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(112, 72)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'FrmScreenCaptureView
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.PictureBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmScreenCaptureView"
        Me.Text = "This is the screen capture"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cFileName As String

    Public Overloads Function ShowDialog(ByVal FileName As String) As System.Windows.Forms.DialogResult
        cFileName = FileName
        MyBase.ShowDialog()
    End Function

    Private Sub ScreenCaptureView_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = (Screen.PrimaryScreen.Bounds.Height - 28)
        Me.Top = Me.Height * 0.1
        Me.Left = Me.Width * 0.1
        PictureBox1.Width = Me.Width - 6
        PictureBox1.Height = Me.Height - 39
        PictureBox1.Left = 0
        PictureBox1.Top = 0
        PictureBox1.Image = Image.FromFile(cFileName)
    End Sub

End Class

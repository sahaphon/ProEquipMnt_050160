<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLogin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin))
Me.GroupBox1 = New System.Windows.Forms.GroupBox
Me.cboUser = New System.Windows.Forms.ComboBox
Me.txtPass = New System.Windows.Forms.TextBox
Me.lklLogin = New System.Windows.Forms.LinkLabel
Me.PictureBox1 = New System.Windows.Forms.PictureBox
Me.Label1 = New System.Windows.Forms.Label
Me.Label2 = New System.Windows.Forms.Label
Me.lklClose = New System.Windows.Forms.LinkLabel
Me.GroupBox1.SuspendLayout()
CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
Me.SuspendLayout()
'
'GroupBox1
'
Me.GroupBox1.BackColor = System.Drawing.Color.White
Me.GroupBox1.Controls.Add(Me.cboUser)
Me.GroupBox1.Controls.Add(Me.txtPass)
Me.GroupBox1.Controls.Add(Me.lklLogin)
Me.GroupBox1.Controls.Add(Me.PictureBox1)
Me.GroupBox1.Controls.Add(Me.Label1)
Me.GroupBox1.Controls.Add(Me.Label2)
Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.GroupBox1.ForeColor = System.Drawing.Color.Green
Me.GroupBox1.Location = New System.Drawing.Point(12, 25)
Me.GroupBox1.Name = "GroupBox1"
Me.GroupBox1.Size = New System.Drawing.Size(311, 168)
Me.GroupBox1.TabIndex = 3
Me.GroupBox1.TabStop = False
Me.GroupBox1.Text = "EQUIPMENTS PROGRAM"
'
'cboUser
'
Me.cboUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cboUser.ForeColor = System.Drawing.Color.Navy
Me.cboUser.FormattingEnabled = True
Me.cboUser.Location = New System.Drawing.Point(28, 58)
Me.cboUser.MaxLength = 10
Me.cboUser.Name = "cboUser"
Me.cboUser.Size = New System.Drawing.Size(123, 23)
Me.cboUser.TabIndex = 22
'
'txtPass
'
Me.txtPass.AcceptsReturn = True
Me.txtPass.BackColor = System.Drawing.SystemColors.Window
Me.txtPass.Cursor = System.Windows.Forms.Cursors.IBeam
Me.txtPass.Font = New System.Drawing.Font("Webdings", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
Me.txtPass.ForeColor = System.Drawing.Color.Navy
Me.txtPass.ImeMode = System.Windows.Forms.ImeMode.Disable
Me.txtPass.Location = New System.Drawing.Point(28, 100)
Me.txtPass.MaxLength = 4
Me.txtPass.Name = "txtPass"
Me.txtPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(61)
Me.txtPass.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.txtPass.Size = New System.Drawing.Size(124, 22)
Me.txtPass.TabIndex = 23
'
'lklLogin
'
Me.lklLogin.ActiveLinkColor = System.Drawing.Color.Red
Me.lklLogin.AutoSize = True
Me.lklLogin.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lklLogin.LinkColor = System.Drawing.Color.Blue
Me.lklLogin.Location = New System.Drawing.Point(59, 139)
Me.lklLogin.Name = "lklLogin"
Me.lklLogin.Size = New System.Drawing.Size(57, 16)
Me.lklLogin.TabIndex = 3
Me.lklLogin.TabStop = True
Me.lklLogin.Text = "LOG IN"
'
'PictureBox1
'
Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
Me.PictureBox1.Image = Global.ProEquipMnt.My.Resources.Resources.compizconfig
Me.PictureBox1.Location = New System.Drawing.Point(171, 22)
Me.PictureBox1.Name = "PictureBox1"
Me.PictureBox1.Size = New System.Drawing.Size(129, 123)
Me.PictureBox1.TabIndex = 1
Me.PictureBox1.TabStop = False
'
'Label1
'
Me.Label1.AutoSize = True
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label1.ForeColor = System.Drawing.Color.DimGray
Me.Label1.Location = New System.Drawing.Point(24, 45)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(65, 13)
Me.Label1.TabIndex = 25
Me.Label1.Text = "UserName"
'
'Label2
'
Me.Label2.AutoSize = True
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label2.ForeColor = System.Drawing.Color.DimGray
Me.Label2.Location = New System.Drawing.Point(25, 88)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(64, 13)
Me.Label2.TabIndex = 26
Me.Label2.Text = "PassWord"
'
'lklClose
'
Me.lklClose.AutoSize = True
Me.lklClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lklClose.ForeColor = System.Drawing.Color.Red
Me.lklClose.LinkColor = System.Drawing.Color.Red
Me.lklClose.Location = New System.Drawing.Point(275, 9)
Me.lklClose.Name = "lklClose"
Me.lklClose.Size = New System.Drawing.Size(47, 13)
Me.lklClose.TabIndex = 4
Me.lklClose.TabStop = True
Me.lklClose.Text = "CLOSE"
'
'frmLogin
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.BackColor = System.Drawing.Color.White
Me.ClientSize = New System.Drawing.Size(334, 211)
Me.Controls.Add(Me.lklClose)
Me.Controls.Add(Me.GroupBox1)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
Me.KeyPreview = True
Me.Name = "frmLogin"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "EQUIPMENTS PROGRAM"
Me.GroupBox1.ResumeLayout(False)
Me.GroupBox1.PerformLayout()
CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
Me.ResumeLayout(False)
Me.PerformLayout()

End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboUser As System.Windows.Forms.ComboBox
    Public WithEvents txtPass As System.Windows.Forms.TextBox
    Friend WithEvents lklLogin As System.Windows.Forms.LinkLabel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lklClose As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class

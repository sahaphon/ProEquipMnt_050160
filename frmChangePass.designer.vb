<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChangePass
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
Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChangePass))
Me.PictureBox2 = New System.Windows.Forms.PictureBox
Me.lblUsrLogIn = New System.Windows.Forms.Label
Me.btnCancel = New System.Windows.Forms.Button
Me.Label1 = New System.Windows.Forms.Label
Me.Label5 = New System.Windows.Forms.Label
Me.txtPass = New System.Windows.Forms.TextBox
Me.GroupBox1 = New System.Windows.Forms.GroupBox
Me.btnSave = New System.Windows.Forms.Button
CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
Me.GroupBox1.SuspendLayout()
Me.SuspendLayout()
'
'PictureBox2
'
Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
Me.PictureBox2.Image = Global.ProEquipMnt.My.Resources.Resources.Registration
Me.PictureBox2.Location = New System.Drawing.Point(14, 29)
Me.PictureBox2.Name = "PictureBox2"
Me.PictureBox2.Size = New System.Drawing.Size(39, 33)
Me.PictureBox2.TabIndex = 54
Me.PictureBox2.TabStop = False
'
'lblUsrLogIn
'
Me.lblUsrLogIn.AutoSize = True
Me.lblUsrLogIn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblUsrLogIn.ForeColor = System.Drawing.Color.Blue
Me.lblUsrLogIn.Location = New System.Drawing.Point(111, 36)
Me.lblUsrLogIn.Name = "lblUsrLogIn"
Me.lblUsrLogIn.Size = New System.Drawing.Size(78, 18)
Me.lblUsrLogIn.TabIndex = 40
Me.lblUsrLogIn.Text = "XXXXXXX"
Me.lblUsrLogIn.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'btnCancel
'
Me.btnCancel.BackColor = System.Drawing.SystemColors.ButtonFace
Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnCancel.ForeColor = System.Drawing.Color.Black
Me.btnCancel.Location = New System.Drawing.Point(161, 141)
Me.btnCancel.Name = "btnCancel"
Me.btnCancel.Size = New System.Drawing.Size(76, 30)
Me.btnCancel.TabIndex = 56
Me.btnCancel.Text = "ยกเลิก"
Me.btnCancel.UseVisualStyleBackColor = False
'
'Label1
'
Me.Label1.AutoSize = True
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label1.ForeColor = System.Drawing.Color.Black
Me.Label1.Location = New System.Drawing.Point(15, 73)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(87, 18)
Me.Label1.TabIndex = 39
Me.Label1.Text = "รหัสผ่านใหม่ :"
Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label5
'
Me.Label5.AutoSize = True
Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label5.ForeColor = System.Drawing.Color.Black
Me.Label5.Location = New System.Drawing.Point(51, 34)
Me.Label5.Name = "Label5"
Me.Label5.Size = New System.Drawing.Size(59, 18)
Me.Label5.TabIndex = 38
Me.Label5.Text = "ผู้ใช้งาน :"
Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'txtPass
'
Me.txtPass.AcceptsReturn = True
Me.txtPass.BackColor = System.Drawing.SystemColors.Window
Me.txtPass.Cursor = System.Windows.Forms.Cursors.IBeam
Me.txtPass.Font = New System.Drawing.Font("Webdings", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
Me.txtPass.ForeColor = System.Drawing.Color.BlueViolet
Me.txtPass.ImeMode = System.Windows.Forms.ImeMode.Disable
Me.txtPass.Location = New System.Drawing.Point(109, 73)
Me.txtPass.MaxLength = 4
Me.txtPass.Name = "txtPass"
Me.txtPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(61)
Me.txtPass.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.txtPass.Size = New System.Drawing.Size(93, 22)
Me.txtPass.TabIndex = 37
Me.txtPass.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'GroupBox1
'
Me.GroupBox1.Controls.Add(Me.PictureBox2)
Me.GroupBox1.Controls.Add(Me.lblUsrLogIn)
Me.GroupBox1.Controls.Add(Me.Label1)
Me.GroupBox1.Controls.Add(Me.Label5)
Me.GroupBox1.Controls.Add(Me.txtPass)
Me.GroupBox1.Location = New System.Drawing.Point(12, 9)
Me.GroupBox1.Name = "GroupBox1"
Me.GroupBox1.Size = New System.Drawing.Size(226, 120)
Me.GroupBox1.TabIndex = 57
Me.GroupBox1.TabStop = False
Me.GroupBox1.Text = "ระบุข้อมูล"
'
'btnSave
'
Me.btnSave.BackColor = System.Drawing.SystemColors.ButtonFace
Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnSave.ForeColor = System.Drawing.SystemColors.ControlText
Me.btnSave.Location = New System.Drawing.Point(79, 141)
Me.btnSave.Name = "btnSave"
Me.btnSave.Size = New System.Drawing.Size(76, 30)
Me.btnSave.TabIndex = 55
Me.btnSave.Text = "บันทึก"
Me.btnSave.UseVisualStyleBackColor = False
'
'frmChangePass
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(250, 180)
Me.Controls.Add(Me.btnCancel)
Me.Controls.Add(Me.GroupBox1)
Me.Controls.Add(Me.btnSave)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
Me.Name = "frmChangePass"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "เปลี่ยนรหัสผ่านผู้ใช้งาน"
CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
Me.GroupBox1.ResumeLayout(False)
Me.GroupBox1.PerformLayout()
Me.ResumeLayout(False)

End Sub
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents lblUsrLogIn As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents txtPass As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWipImp
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
Me.gpbTopic = New System.Windows.Forms.GroupBox
Me.lblCurentRd = New System.Windows.Forms.Label
Me.btnExit = New System.Windows.Forms.Button
Me.btnCancel = New System.Windows.Forms.Button
Me.btnImport = New System.Windows.Forms.Button
Me.pgbSta = New System.Windows.Forms.ProgressBar
Me.gpbTopic.SuspendLayout()
Me.SuspendLayout()
'
'gpbTopic
'
Me.gpbTopic.Controls.Add(Me.lblCurentRd)
Me.gpbTopic.Controls.Add(Me.btnExit)
Me.gpbTopic.Controls.Add(Me.btnCancel)
Me.gpbTopic.Controls.Add(Me.btnImport)
Me.gpbTopic.Controls.Add(Me.pgbSta)
Me.gpbTopic.Font = New System.Drawing.Font("BrowalliaUPC", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.gpbTopic.ForeColor = System.Drawing.Color.Navy
Me.gpbTopic.Location = New System.Drawing.Point(10, 11)
Me.gpbTopic.Name = "gpbTopic"
Me.gpbTopic.Size = New System.Drawing.Size(302, 123)
Me.gpbTopic.TabIndex = 1
Me.gpbTopic.TabStop = False
Me.gpbTopic.Text = "กำลังนำเข้าข้อมูล...."
'
'lblCurentRd
'
Me.lblCurentRd.BackColor = System.Drawing.Color.Black
Me.lblCurentRd.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblCurentRd.ForeColor = System.Drawing.Color.Yellow
Me.lblCurentRd.Location = New System.Drawing.Point(15, 34)
Me.lblCurentRd.Name = "lblCurentRd"
Me.lblCurentRd.Size = New System.Drawing.Size(64, 29)
Me.lblCurentRd.TabIndex = 39
Me.lblCurentRd.Text = "0"
Me.lblCurentRd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'btnExit
'
Me.btnExit.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnExit.ForeColor = System.Drawing.Color.Black
Me.btnExit.Location = New System.Drawing.Point(204, 76)
Me.btnExit.Name = "btnExit"
Me.btnExit.Size = New System.Drawing.Size(83, 31)
Me.btnExit.TabIndex = 3
Me.btnExit.Text = "ออก"
Me.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
Me.btnExit.UseVisualStyleBackColor = True
'
'btnCancel
'
Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnCancel.Enabled = False
Me.btnCancel.ForeColor = System.Drawing.Color.Black
Me.btnCancel.Location = New System.Drawing.Point(98, 76)
Me.btnCancel.Name = "btnCancel"
Me.btnCancel.Size = New System.Drawing.Size(83, 31)
Me.btnCancel.TabIndex = 2
Me.btnCancel.Text = "ยกเลิก"
Me.btnCancel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
Me.btnCancel.UseVisualStyleBackColor = True
'
'btnImport
'
Me.btnImport.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnImport.ForeColor = System.Drawing.Color.Black
Me.btnImport.Location = New System.Drawing.Point(15, 76)
Me.btnImport.Name = "btnImport"
Me.btnImport.Size = New System.Drawing.Size(83, 31)
Me.btnImport.TabIndex = 1
Me.btnImport.Text = "นำเข้า"
Me.btnImport.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
Me.btnImport.UseVisualStyleBackColor = True
'
'pgbSta
'
Me.pgbSta.ForeColor = System.Drawing.Color.Red
Me.pgbSta.Location = New System.Drawing.Point(78, 34)
Me.pgbSta.Name = "pgbSta"
Me.pgbSta.Size = New System.Drawing.Size(209, 29)
Me.pgbSta.TabIndex = 0
'
'frmWipImp
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(323, 141)
Me.ControlBox = False
Me.Controls.Add(Me.gpbTopic)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
Me.MaximizeBox = False
Me.MinimizeBox = False
Me.Name = "frmWipImp"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "ฟอร์มนำเข้าข้อมูลลูกค้า"
Me.gpbTopic.ResumeLayout(False)
Me.ResumeLayout(False)

End Sub
    Friend WithEvents gpbTopic As System.Windows.Forms.GroupBox
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents pgbSta As System.Windows.Forms.ProgressBar
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents lblCurentRd As System.Windows.Forms.Label
End Class

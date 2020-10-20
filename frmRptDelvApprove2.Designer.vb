<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptDelvApprove2
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
Me.spnRpt = New System.Windows.Forms.SplitContainer
Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
Me.lblComplete = New System.Windows.Forms.Label
Me.gpb2 = New System.Windows.Forms.GroupBox
Me.lblApp2 = New System.Windows.Forms.Label
Me.Label7 = New System.Windows.Forms.Label
Me.btnAcp2 = New System.Windows.Forms.Button
Me.Label2 = New System.Windows.Forms.Label
Me.gpb4 = New System.Windows.Forms.GroupBox
Me.lblApp4 = New System.Windows.Forms.Label
Me.Label4 = New System.Windows.Forms.Label
Me.btnAcp4 = New System.Windows.Forms.Button
Me.Label6 = New System.Windows.Forms.Label
Me.gpb3 = New System.Windows.Forms.GroupBox
Me.lblApp3 = New System.Windows.Forms.Label
Me.Label3 = New System.Windows.Forms.Label
Me.btnAcp3 = New System.Windows.Forms.Button
Me.lblPicname = New System.Windows.Forms.Label
Me.btnFeed = New System.Windows.Forms.Button
Me.spnRpt.Panel1.SuspendLayout()
Me.spnRpt.Panel2.SuspendLayout()
Me.spnRpt.SuspendLayout()
Me.gpb2.SuspendLayout()
Me.gpb4.SuspendLayout()
Me.gpb3.SuspendLayout()
Me.SuspendLayout()
'
'spnRpt
'
Me.spnRpt.Dock = System.Windows.Forms.DockStyle.Fill
Me.spnRpt.Location = New System.Drawing.Point(0, 0)
Me.spnRpt.Name = "spnRpt"
'
'spnRpt.Panel1
'
Me.spnRpt.Panel1.Controls.Add(Me.CrystalReportViewer1)
'
'spnRpt.Panel2
'
Me.spnRpt.Panel2.Controls.Add(Me.lblComplete)
Me.spnRpt.Panel2.Controls.Add(Me.gpb2)
Me.spnRpt.Panel2.Controls.Add(Me.gpb4)
Me.spnRpt.Panel2.Controls.Add(Me.gpb3)
Me.spnRpt.Panel2.Controls.Add(Me.btnFeed)
Me.spnRpt.Size = New System.Drawing.Size(1028, 722)
Me.spnRpt.SplitterDistance = 835
Me.spnRpt.TabIndex = 2
'
'CrystalReportViewer1
'
Me.CrystalReportViewer1.ActiveViewIndex = -1
Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.CrystalReportViewer1.DisplayGroupTree = False
Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
Me.CrystalReportViewer1.SelectionFormula = ""
Me.CrystalReportViewer1.Size = New System.Drawing.Size(835, 722)
Me.CrystalReportViewer1.TabIndex = 1
Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
'
'lblComplete
'
Me.lblComplete.BackColor = System.Drawing.Color.Yellow
Me.lblComplete.Location = New System.Drawing.Point(42, 421)
Me.lblComplete.Name = "lblComplete"
Me.lblComplete.Size = New System.Drawing.Size(115, 23)
Me.lblComplete.TabIndex = 13
Me.lblComplete.Visible = False
'
'gpb2
'
Me.gpb2.Controls.Add(Me.lblApp2)
Me.gpb2.Controls.Add(Me.Label7)
Me.gpb2.Controls.Add(Me.btnAcp2)
Me.gpb2.Controls.Add(Me.Label2)
Me.gpb2.Location = New System.Drawing.Point(21, 25)
Me.gpb2.Name = "gpb2"
Me.gpb2.Size = New System.Drawing.Size(144, 79)
Me.gpb2.TabIndex = 12
Me.gpb2.TabStop = False
'
'lblApp2
'
Me.lblApp2.BackColor = System.Drawing.SystemColors.Control
Me.lblApp2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblApp2.ForeColor = System.Drawing.Color.Red
Me.lblApp2.Image = Global.ProEquipMnt.My.Resources.Resources.Chk
Me.lblApp2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblApp2.Location = New System.Drawing.Point(8, 14)
Me.lblApp2.Name = "lblApp2"
Me.lblApp2.Size = New System.Drawing.Size(130, 37)
Me.lblApp2.TabIndex = 8
Me.lblApp2.Text = "เซ็นต์อนุมัติแล้ว"
Me.lblApp2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblApp2.Visible = False
'
'Label7
'
Me.Label7.AutoSize = True
Me.Label7.Location = New System.Drawing.Point(25, 29)
Me.Label7.Name = "Label7"
Me.Label7.Size = New System.Drawing.Size(0, 13)
Me.Label7.TabIndex = 11
'
'btnAcp2
'
Me.btnAcp2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.btnAcp2.ForeColor = System.Drawing.Color.Green
Me.btnAcp2.Location = New System.Drawing.Point(37, 18)
Me.btnAcp2.Name = "btnAcp2"
Me.btnAcp2.Size = New System.Drawing.Size(71, 32)
Me.btnAcp2.TabIndex = 0
Me.btnAcp2.Text = "อนุมัติ"
Me.btnAcp2.UseVisualStyleBackColor = True
'
'Label2
'
Me.Label2.AutoSize = True
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label2.ForeColor = System.Drawing.Color.Blue
Me.Label2.Location = New System.Drawing.Point(8, 53)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(128, 16)
Me.Label2.TabIndex = 2
Me.Label2.Text = "หน.ส่วนเทคนิคอุปกรณ์"
'
'gpb4
'
Me.gpb4.Controls.Add(Me.lblApp4)
Me.gpb4.Controls.Add(Me.Label4)
Me.gpb4.Controls.Add(Me.btnAcp4)
Me.gpb4.Controls.Add(Me.Label6)
Me.gpb4.Location = New System.Drawing.Point(21, 219)
Me.gpb4.Name = "gpb4"
Me.gpb4.Size = New System.Drawing.Size(144, 79)
Me.gpb4.TabIndex = 10
Me.gpb4.TabStop = False
'
'lblApp4
'
Me.lblApp4.BackColor = System.Drawing.SystemColors.Control
Me.lblApp4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblApp4.ForeColor = System.Drawing.Color.Red
Me.lblApp4.Image = Global.ProEquipMnt.My.Resources.Resources.Chk
Me.lblApp4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblApp4.Location = New System.Drawing.Point(8, 16)
Me.lblApp4.Name = "lblApp4"
Me.lblApp4.Size = New System.Drawing.Size(130, 37)
Me.lblApp4.TabIndex = 8
Me.lblApp4.Text = "เซ็นต์อนุมัติแล้ว"
Me.lblApp4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblApp4.Visible = False
'
'Label4
'
Me.Label4.AutoSize = True
Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label4.ForeColor = System.Drawing.Color.Blue
Me.Label4.Location = New System.Drawing.Point(22, 56)
Me.Label4.Name = "Label4"
Me.Label4.Size = New System.Drawing.Size(114, 16)
Me.Label4.TabIndex = 2
Me.Label4.Text = "ผจก.แผนก / รับของ"
'
'btnAcp4
'
Me.btnAcp4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.btnAcp4.ForeColor = System.Drawing.Color.Green
Me.btnAcp4.Location = New System.Drawing.Point(37, 19)
Me.btnAcp4.Name = "btnAcp4"
Me.btnAcp4.Size = New System.Drawing.Size(77, 34)
Me.btnAcp4.TabIndex = 3
Me.btnAcp4.Text = "อนุมัติ"
Me.btnAcp4.UseVisualStyleBackColor = True
'
'Label6
'
Me.Label6.AutoSize = True
Me.Label6.Location = New System.Drawing.Point(25, 29)
Me.Label6.Name = "Label6"
Me.Label6.Size = New System.Drawing.Size(0, 13)
Me.Label6.TabIndex = 11
'
'gpb3
'
Me.gpb3.Controls.Add(Me.lblApp3)
Me.gpb3.Controls.Add(Me.Label3)
Me.gpb3.Controls.Add(Me.btnAcp3)
Me.gpb3.Controls.Add(Me.lblPicname)
Me.gpb3.Location = New System.Drawing.Point(21, 121)
Me.gpb3.Name = "gpb3"
Me.gpb3.Size = New System.Drawing.Size(144, 79)
Me.gpb3.TabIndex = 9
Me.gpb3.TabStop = False
'
'lblApp3
'
Me.lblApp3.BackColor = System.Drawing.SystemColors.Control
Me.lblApp3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblApp3.ForeColor = System.Drawing.Color.Red
Me.lblApp3.Image = Global.ProEquipMnt.My.Resources.Resources.Chk
Me.lblApp3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblApp3.Location = New System.Drawing.Point(6, 14)
Me.lblApp3.Name = "lblApp3"
Me.lblApp3.Size = New System.Drawing.Size(130, 37)
Me.lblApp3.TabIndex = 8
Me.lblApp3.Text = "เซ็นต์อนุมัติแล้ว"
Me.lblApp3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblApp3.Visible = False
'
'Label3
'
Me.Label3.AutoSize = True
Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label3.ForeColor = System.Drawing.Color.Blue
Me.Label3.Location = New System.Drawing.Point(3, 54)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(140, 16)
Me.Label3.TabIndex = 2
Me.Label3.Text = "ผจก.แผนกเทคนิคอุปกรณ์"
'
'btnAcp3
'
Me.btnAcp3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.btnAcp3.ForeColor = System.Drawing.Color.Green
Me.btnAcp3.Location = New System.Drawing.Point(31, 18)
Me.btnAcp3.Name = "btnAcp3"
Me.btnAcp3.Size = New System.Drawing.Size(74, 34)
Me.btnAcp3.TabIndex = 12
Me.btnAcp3.Text = "อนุมัติ"
Me.btnAcp3.UseVisualStyleBackColor = True
'
'lblPicname
'
Me.lblPicname.AutoSize = True
Me.lblPicname.Location = New System.Drawing.Point(25, 29)
Me.lblPicname.Name = "lblPicname"
Me.lblPicname.Size = New System.Drawing.Size(0, 13)
Me.lblPicname.TabIndex = 11
'
'btnFeed
'
Me.btnFeed.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.btnFeed.ForeColor = System.Drawing.Color.Blue
Me.btnFeed.Location = New System.Drawing.Point(2, 309)
Me.btnFeed.Name = "btnFeed"
Me.btnFeed.Size = New System.Drawing.Size(18, 81)
Me.btnFeed.TabIndex = 8
Me.btnFeed.UseVisualStyleBackColor = True
'
'frmRptDelvApprove2
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(1028, 722)
Me.Controls.Add(Me.spnRpt)
Me.Name = "frmRptDelvApprove2"
Me.Text = "รายงานโอนอุปกรณ์ลงผลิต"
Me.spnRpt.Panel1.ResumeLayout(False)
Me.spnRpt.Panel2.ResumeLayout(False)
Me.spnRpt.ResumeLayout(False)
Me.gpb2.ResumeLayout(False)
Me.gpb2.PerformLayout()
Me.gpb4.ResumeLayout(False)
Me.gpb4.PerformLayout()
Me.gpb3.ResumeLayout(False)
Me.gpb3.PerformLayout()
Me.ResumeLayout(False)

End Sub
    Friend WithEvents spnRpt As System.Windows.Forms.SplitContainer
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents gpb2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblApp2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnAcp2 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents gpb4 As System.Windows.Forms.GroupBox
    Friend WithEvents lblApp4 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnAcp4 As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents gpb3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblApp3 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnAcp3 As System.Windows.Forms.Button
    Friend WithEvents lblPicname As System.Windows.Forms.Label
    Friend WithEvents btnFeed As System.Windows.Forms.Button
    Friend WithEvents lblComplete As System.Windows.Forms.Label
End Class

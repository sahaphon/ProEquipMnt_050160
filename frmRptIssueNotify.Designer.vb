<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptIssueNotify
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
Me.GroupBox2 = New System.Windows.Forms.GroupBox
Me.gpb1 = New System.Windows.Forms.GroupBox
Me.lblApp1 = New System.Windows.Forms.Label
Me.Label1 = New System.Windows.Forms.Label
Me.Label8 = New System.Windows.Forms.Label
Me.btnAcp1 = New System.Windows.Forms.Button
Me.GroupBox1 = New System.Windows.Forms.GroupBox
Me.gpb3 = New System.Windows.Forms.GroupBox
Me.lblApp3 = New System.Windows.Forms.Label
Me.Label3 = New System.Windows.Forms.Label
Me.btnAcp3 = New System.Windows.Forms.Button
Me.lblPicname = New System.Windows.Forms.Label
Me.btnFeed = New System.Windows.Forms.Button
Me.spnRpt.Panel1.SuspendLayout()
Me.spnRpt.Panel2.SuspendLayout()
Me.spnRpt.SuspendLayout()
Me.GroupBox2.SuspendLayout()
Me.gpb1.SuspendLayout()
Me.GroupBox1.SuspendLayout()
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
Me.spnRpt.Panel2.Controls.Add(Me.GroupBox2)
Me.spnRpt.Panel2.Controls.Add(Me.GroupBox1)
Me.spnRpt.Panel2.Controls.Add(Me.btnFeed)
Me.spnRpt.Size = New System.Drawing.Size(1028, 722)
Me.spnRpt.SplitterDistance = 835
Me.spnRpt.TabIndex = 0
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
Me.lblComplete.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
Me.lblComplete.Location = New System.Drawing.Point(21, 294)
Me.lblComplete.Name = "lblComplete"
Me.lblComplete.Size = New System.Drawing.Size(147, 23)
Me.lblComplete.TabIndex = 14
Me.lblComplete.Visible = False
'
'GroupBox2
'
Me.GroupBox2.BackColor = System.Drawing.SystemColors.Control
Me.GroupBox2.Controls.Add(Me.gpb1)
Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.GroupBox2.Location = New System.Drawing.Point(15, 14)
Me.GroupBox2.Name = "GroupBox2"
Me.GroupBox2.Size = New System.Drawing.Size(162, 113)
Me.GroupBox2.TabIndex = 13
Me.GroupBox2.TabStop = False
Me.GroupBox2.Text = "ส่วนผู้แจ้งปัญหา"
'
'gpb1
'
Me.gpb1.Controls.Add(Me.lblApp1)
Me.gpb1.Controls.Add(Me.Label1)
Me.gpb1.Controls.Add(Me.Label8)
Me.gpb1.Controls.Add(Me.btnAcp1)
Me.gpb1.Location = New System.Drawing.Point(9, 19)
Me.gpb1.Name = "gpb1"
Me.gpb1.Size = New System.Drawing.Size(144, 79)
Me.gpb1.TabIndex = 10
Me.gpb1.TabStop = False
'
'lblApp1
'
Me.lblApp1.BackColor = System.Drawing.SystemColors.Control
Me.lblApp1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblApp1.ForeColor = System.Drawing.Color.Red
Me.lblApp1.Image = Global.ProEquipMnt.My.Resources.Resources.Chk
Me.lblApp1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblApp1.Location = New System.Drawing.Point(6, 8)
Me.lblApp1.Name = "lblApp1"
Me.lblApp1.Size = New System.Drawing.Size(130, 37)
Me.lblApp1.TabIndex = 8
Me.lblApp1.Text = "เซ็นต์อนุมัติแล้ว"
Me.lblApp1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblApp1.Visible = False
'
'Label1
'
Me.Label1.AutoSize = True
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label1.ForeColor = System.Drawing.Color.Blue
Me.Label1.Location = New System.Drawing.Point(5, 53)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(137, 16)
Me.Label1.TabIndex = 2
Me.Label1.Text = "ผจก.(แผนกที่แจ้งปัญหา)"
'
'Label8
'
Me.Label8.AutoSize = True
Me.Label8.Location = New System.Drawing.Point(25, 29)
Me.Label8.Name = "Label8"
Me.Label8.Size = New System.Drawing.Size(0, 13)
Me.Label8.TabIndex = 11
'
'btnAcp1
'
Me.btnAcp1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.btnAcp1.ForeColor = System.Drawing.Color.Green
Me.btnAcp1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.btnAcp1.Location = New System.Drawing.Point(33, 10)
Me.btnAcp1.Name = "btnAcp1"
Me.btnAcp1.Size = New System.Drawing.Size(72, 32)
Me.btnAcp1.TabIndex = 1
Me.btnAcp1.Text = "อนุมัติ"
Me.btnAcp1.UseVisualStyleBackColor = True
'
'GroupBox1
'
Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
Me.GroupBox1.Controls.Add(Me.gpb3)
Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.GroupBox1.Location = New System.Drawing.Point(15, 142)
Me.GroupBox1.Name = "GroupBox1"
Me.GroupBox1.Size = New System.Drawing.Size(162, 112)
Me.GroupBox1.TabIndex = 12
Me.GroupBox1.TabStop = False
Me.GroupBox1.Text = "ส่วนผู้รับเเจ้ง"
'
'gpb3
'
Me.gpb3.Controls.Add(Me.lblApp3)
Me.gpb3.Controls.Add(Me.Label3)
Me.gpb3.Controls.Add(Me.btnAcp3)
Me.gpb3.Controls.Add(Me.lblPicname)
Me.gpb3.Location = New System.Drawing.Point(6, 19)
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
Me.lblApp3.Location = New System.Drawing.Point(7, 14)
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
Me.Label3.Location = New System.Drawing.Point(18, 53)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(110, 16)
Me.Label3.TabIndex = 2
Me.Label3.Text = "ผจก.เทคนิคอุปกรณ์"
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
Me.btnFeed.Location = New System.Drawing.Point(3, 359)
Me.btnFeed.Name = "btnFeed"
Me.btnFeed.Size = New System.Drawing.Size(18, 81)
Me.btnFeed.TabIndex = 8
Me.btnFeed.UseVisualStyleBackColor = True
'
'frmRptIssueNotify
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(1028, 722)
Me.Controls.Add(Me.spnRpt)
Me.Name = "frmRptIssueNotify"
Me.Text = "ฟอร์มรายงานการแจ้้งปัญหาอุปกรณ์"
Me.spnRpt.Panel1.ResumeLayout(False)
Me.spnRpt.Panel2.ResumeLayout(False)
Me.spnRpt.ResumeLayout(False)
Me.GroupBox2.ResumeLayout(False)
Me.gpb1.ResumeLayout(False)
Me.gpb1.PerformLayout()
Me.GroupBox1.ResumeLayout(False)
Me.gpb3.ResumeLayout(False)
Me.gpb3.PerformLayout()
Me.ResumeLayout(False)

End Sub
    Friend WithEvents spnRpt As System.Windows.Forms.SplitContainer
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnFeed As System.Windows.Forms.Button
    Friend WithEvents gpb1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblApp1 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnAcp1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents gpb3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblApp3 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnAcp3 As System.Windows.Forms.Button
    Friend WithEvents lblPicname As System.Windows.Forms.Label
    Friend WithEvents lblComplete As System.Windows.Forms.Label
End Class

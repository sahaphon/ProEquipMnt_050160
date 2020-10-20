<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptDelvApprove
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
Me.gpbInputname = New System.Windows.Forms.GroupBox
Me.lnkCancle = New System.Windows.Forms.LinkLabel
Me.lnkSave = New System.Windows.Forms.LinkLabel
Me.txtName = New System.Windows.Forms.TextBox
Me.Label5 = New System.Windows.Forms.Label
Me.gpb1 = New System.Windows.Forms.GroupBox
Me.Label1 = New System.Windows.Forms.Label
Me.Label8 = New System.Windows.Forms.Label
Me.btnAcp1 = New System.Windows.Forms.Button
Me.btnFeed = New System.Windows.Forms.Button
Me.lblApp1 = New System.Windows.Forms.Label
Me.spnRpt.Panel1.SuspendLayout()
Me.spnRpt.Panel2.SuspendLayout()
Me.spnRpt.SuspendLayout()
Me.gpbInputname.SuspendLayout()
Me.gpb1.SuspendLayout()
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
Me.spnRpt.Panel2.Controls.Add(Me.gpbInputname)
Me.spnRpt.Panel2.Controls.Add(Me.gpb1)
Me.spnRpt.Panel2.Controls.Add(Me.btnFeed)
Me.spnRpt.Size = New System.Drawing.Size(1028, 722)
Me.spnRpt.SplitterDistance = 835
Me.spnRpt.TabIndex = 1
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
'gpbInputname
'
Me.gpbInputname.Controls.Add(Me.lnkCancle)
Me.gpbInputname.Controls.Add(Me.lnkSave)
Me.gpbInputname.Controls.Add(Me.txtName)
Me.gpbInputname.Controls.Add(Me.Label5)
Me.gpbInputname.ForeColor = System.Drawing.Color.Blue
Me.gpbInputname.Location = New System.Drawing.Point(10, 105)
Me.gpbInputname.Name = "gpbInputname"
Me.gpbInputname.Size = New System.Drawing.Size(173, 139)
Me.gpbInputname.TabIndex = 13
Me.gpbInputname.TabStop = False
Me.gpbInputname.Text = "ผู้จัดของ"
Me.gpbInputname.Visible = False
'
'lnkCancle
'
Me.lnkCancle.AutoSize = True
Me.lnkCancle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lnkCancle.ForeColor = System.Drawing.Color.Red
Me.lnkCancle.LinkColor = System.Drawing.Color.Red
Me.lnkCancle.Location = New System.Drawing.Point(127, 113)
Me.lnkCancle.Name = "lnkCancle"
Me.lnkCancle.Size = New System.Drawing.Size(44, 13)
Me.lnkCancle.TabIndex = 2
Me.lnkCancle.TabStop = True
Me.lnkCancle.Text = "ยกเลิก"
'
'lnkSave
'
Me.lnkSave.AutoSize = True
Me.lnkSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lnkSave.LinkColor = System.Drawing.Color.Green
Me.lnkSave.Location = New System.Drawing.Point(88, 113)
Me.lnkSave.Name = "lnkSave"
Me.lnkSave.Size = New System.Drawing.Size(36, 13)
Me.lnkSave.TabIndex = 2
Me.lnkSave.TabStop = True
Me.lnkSave.Text = "ตกลง"
'
'txtName
'
Me.txtName.Location = New System.Drawing.Point(10, 36)
Me.txtName.Multiline = True
Me.txtName.Name = "txtName"
Me.txtName.Size = New System.Drawing.Size(152, 71)
Me.txtName.TabIndex = 1
'
'Label5
'
Me.Label5.AutoSize = True
Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
Me.Label5.Location = New System.Drawing.Point(7, 20)
Me.Label5.Name = "Label5"
Me.Label5.Size = New System.Drawing.Size(29, 13)
Me.Label5.TabIndex = 0
Me.Label5.Text = "ชื่อ : "
'
'gpb1
'
Me.gpb1.Controls.Add(Me.lblApp1)
Me.gpb1.Controls.Add(Me.Label1)
Me.gpb1.Controls.Add(Me.Label8)
Me.gpb1.Controls.Add(Me.btnAcp1)
Me.gpb1.Location = New System.Drawing.Point(21, 12)
Me.gpb1.Name = "gpb1"
Me.gpb1.Size = New System.Drawing.Size(144, 79)
Me.gpb1.TabIndex = 11
Me.gpb1.TabStop = False
'
'Label1
'
Me.Label1.AutoSize = True
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label1.ForeColor = System.Drawing.Color.Blue
Me.Label1.Location = New System.Drawing.Point(48, 52)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(48, 16)
Me.Label1.TabIndex = 2
Me.Label1.Text = "ผู้บันทึก"
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
'lblApp1
'
Me.lblApp1.BackColor = System.Drawing.SystemColors.Control
Me.lblApp1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblApp1.ForeColor = System.Drawing.Color.Red
Me.lblApp1.Image = Global.ProEquipMnt.My.Resources.Resources.Chk
Me.lblApp1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblApp1.Location = New System.Drawing.Point(8, 8)
Me.lblApp1.Name = "lblApp1"
Me.lblApp1.Size = New System.Drawing.Size(130, 37)
Me.lblApp1.TabIndex = 8
Me.lblApp1.Text = "เซ็นต์อนุมัติแล้ว"
Me.lblApp1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblApp1.Visible = False
'
'frmRptDelvApprove
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(1028, 722)
Me.Controls.Add(Me.spnRpt)
Me.Name = "frmRptDelvApprove"
Me.Text = "รายงานการโอนอุปกรณ์ลงผลิต"
Me.spnRpt.Panel1.ResumeLayout(False)
Me.spnRpt.Panel2.ResumeLayout(False)
Me.spnRpt.ResumeLayout(False)
Me.gpbInputname.ResumeLayout(False)
Me.gpbInputname.PerformLayout()
Me.gpb1.ResumeLayout(False)
Me.gpb1.PerformLayout()
Me.ResumeLayout(False)

End Sub
    Friend WithEvents spnRpt As System.Windows.Forms.SplitContainer
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnFeed As System.Windows.Forms.Button
    Friend WithEvents gpbInputname As System.Windows.Forms.GroupBox
    Friend WithEvents lnkCancle As System.Windows.Forms.LinkLabel
    Friend WithEvents lnkSave As System.Windows.Forms.LinkLabel
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents gpb1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblApp1 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnAcp1 As System.Windows.Forms.Button
End Class

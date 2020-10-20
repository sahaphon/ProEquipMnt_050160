<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptIssueReceive
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
Me.Label16 = New System.Windows.Forms.Label
Me.txtWantDate = New System.Windows.Forms.TextBox
Me.mskWantDate = New System.Windows.Forms.MaskedTextBox
Me.Label13 = New System.Windows.Forms.Label
Me.Label15 = New System.Windows.Forms.Label
Me.txtWanttime = New System.Windows.Forms.TextBox
Me.txtFxissue = New System.Windows.Forms.TextBox
Me.GroupBox1 = New System.Windows.Forms.GroupBox
Me.gpb2 = New System.Windows.Forms.GroupBox
Me.Label7 = New System.Windows.Forms.Label
Me.btnAcp = New System.Windows.Forms.Button
Me.Label2 = New System.Windows.Forms.Label
Me.btnFeed = New System.Windows.Forms.Button
Me.gpbRecvNotify = New System.Windows.Forms.GroupBox
Me.lblApp = New System.Windows.Forms.Label
Me.spnRpt.Panel1.SuspendLayout()
Me.spnRpt.Panel2.SuspendLayout()
Me.spnRpt.SuspendLayout()
Me.GroupBox1.SuspendLayout()
Me.gpb2.SuspendLayout()
Me.gpbRecvNotify.SuspendLayout()
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
Me.spnRpt.Panel2.Controls.Add(Me.gpbRecvNotify)
Me.spnRpt.Panel2.Controls.Add(Me.GroupBox1)
Me.spnRpt.Panel2.Controls.Add(Me.btnFeed)
Me.spnRpt.Size = New System.Drawing.Size(1028, 722)
Me.spnRpt.SplitterDistance = 705
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
Me.CrystalReportViewer1.Size = New System.Drawing.Size(705, 722)
Me.CrystalReportViewer1.TabIndex = 1
Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
'
'Label16
'
Me.Label16.AutoSize = True
Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label16.ForeColor = System.Drawing.Color.Black
Me.Label16.Location = New System.Drawing.Point(22, 27)
Me.Label16.Name = "Label16"
Me.Label16.Size = New System.Drawing.Size(55, 16)
Me.Label16.TabIndex = 182
Me.Label16.Text = "การแก้ไข :"
Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'txtWantDate
'
Me.txtWantDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtWantDate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtWantDate.Location = New System.Drawing.Point(170, 182)
Me.txtWantDate.MaxLength = 10
Me.txtWantDate.Name = "txtWantDate"
Me.txtWantDate.Size = New System.Drawing.Size(149, 29)
Me.txtWantDate.TabIndex = 1
Me.txtWantDate.Text = "__/__/____"
Me.txtWantDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'mskWantDate
'
Me.mskWantDate.BackColor = System.Drawing.Color.SlateBlue
Me.mskWantDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.mskWantDate.ForeColor = System.Drawing.Color.White
Me.mskWantDate.Location = New System.Drawing.Point(170, 182)
Me.mskWantDate.Mask = "99/99/9999"
Me.mskWantDate.Name = "mskWantDate"
Me.mskWantDate.Size = New System.Drawing.Size(149, 29)
Me.mskWantDate.TabIndex = 181
Me.mskWantDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'Label13
'
Me.Label13.AutoSize = True
Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label13.ForeColor = System.Drawing.Color.Black
Me.Label13.Location = New System.Drawing.Point(129, 230)
Me.Label13.Name = "Label13"
Me.Label13.Size = New System.Drawing.Size(34, 16)
Me.Label13.TabIndex = 179
Me.Label13.Text = "เวลา :"
Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label15
'
Me.Label15.AutoSize = True
Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label15.ForeColor = System.Drawing.Color.Black
Me.Label15.Location = New System.Drawing.Point(46, 187)
Me.Label15.Name = "Label15"
Me.Label15.Size = New System.Drawing.Size(115, 16)
Me.Label15.TabIndex = 180
Me.Label15.Text = "กำหนดให้เสร็จภายใน  :"
Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'txtWanttime
'
Me.txtWanttime.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtWanttime.ForeColor = System.Drawing.Color.Navy
Me.txtWanttime.Location = New System.Drawing.Point(170, 225)
Me.txtWanttime.MaxLength = 5
Me.txtWanttime.Name = "txtWanttime"
Me.txtWanttime.Size = New System.Drawing.Size(109, 29)
Me.txtWanttime.TabIndex = 2
'
'txtFxissue
'
Me.txtFxissue.Location = New System.Drawing.Point(21, 49)
Me.txtFxissue.Multiline = True
Me.txtFxissue.Name = "txtFxissue"
Me.txtFxissue.Size = New System.Drawing.Size(298, 118)
Me.txtFxissue.TabIndex = 0
'
'GroupBox1
'
Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
Me.GroupBox1.Controls.Add(Me.gpb2)
Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.GroupBox1.Location = New System.Drawing.Point(41, 12)
Me.GroupBox1.Name = "GroupBox1"
Me.GroupBox1.Size = New System.Drawing.Size(162, 124)
Me.GroupBox1.TabIndex = 12
Me.GroupBox1.TabStop = False
Me.GroupBox1.Text = "ส่วนผู้รับเเจ้ง"
'
'gpb2
'
Me.gpb2.Controls.Add(Me.lblApp)
Me.gpb2.Controls.Add(Me.Label7)
Me.gpb2.Controls.Add(Me.btnAcp)
Me.gpb2.Controls.Add(Me.Label2)
Me.gpb2.Location = New System.Drawing.Point(9, 19)
Me.gpb2.Name = "gpb2"
Me.gpb2.Size = New System.Drawing.Size(144, 79)
Me.gpb2.TabIndex = 11
Me.gpb2.TabStop = False
'
'Label7
'
Me.Label7.AutoSize = True
Me.Label7.Location = New System.Drawing.Point(25, 29)
Me.Label7.Name = "Label7"
Me.Label7.Size = New System.Drawing.Size(0, 13)
Me.Label7.TabIndex = 11
'
'btnAcp
'
Me.btnAcp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.btnAcp.ForeColor = System.Drawing.Color.Green
Me.btnAcp.Location = New System.Drawing.Point(37, 18)
Me.btnAcp.Name = "btnAcp"
Me.btnAcp.Size = New System.Drawing.Size(71, 32)
Me.btnAcp.TabIndex = 0
Me.btnAcp.Text = "อนุมัติ"
Me.btnAcp.UseVisualStyleBackColor = True
'
'Label2
'
Me.Label2.AutoSize = True
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label2.ForeColor = System.Drawing.Color.Blue
Me.Label2.Location = New System.Drawing.Point(3, 54)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(138, 16)
Me.Label2.TabIndex = 2
Me.Label2.Text = "ผู้รับแจ้ง(เทคนิคอุปกรณ์)"
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
'gpbRecvNotify
'
Me.gpbRecvNotify.Controls.Add(Me.txtWantDate)
Me.gpbRecvNotify.Controls.Add(Me.Label16)
Me.gpbRecvNotify.Controls.Add(Me.txtWanttime)
Me.gpbRecvNotify.Controls.Add(Me.txtFxissue)
Me.gpbRecvNotify.Controls.Add(Me.Label15)
Me.gpbRecvNotify.Controls.Add(Me.mskWantDate)
Me.gpbRecvNotify.Controls.Add(Me.Label13)
Me.gpbRecvNotify.Location = New System.Drawing.Point(40, 142)
Me.gpbRecvNotify.Name = "gpbRecvNotify"
Me.gpbRecvNotify.Size = New System.Drawing.Size(341, 280)
Me.gpbRecvNotify.TabIndex = 183
Me.gpbRecvNotify.TabStop = False
Me.gpbRecvNotify.Text = "สำหรับผู้รับเเจ้ง"
Me.gpbRecvNotify.Visible = False
'
'lblApp
'
Me.lblApp.BackColor = System.Drawing.SystemColors.Control
Me.lblApp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblApp.ForeColor = System.Drawing.Color.Red
Me.lblApp.Image = Global.ProEquipMnt.My.Resources.Resources.Chk
Me.lblApp.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblApp.Location = New System.Drawing.Point(8, 13)
Me.lblApp.Name = "lblApp"
Me.lblApp.Size = New System.Drawing.Size(130, 37)
Me.lblApp.TabIndex = 8
Me.lblApp.Text = "เซ็นต์อนุมัติแล้ว"
Me.lblApp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblApp.Visible = False
'
'frmRptIssueReceive
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(1028, 722)
Me.Controls.Add(Me.spnRpt)
Me.Name = "frmRptIssueReceive"
Me.Text = "ฟอร์มรายงานการเเจ้งปัญหาอุปกรณ์"
Me.spnRpt.Panel1.ResumeLayout(False)
Me.spnRpt.Panel2.ResumeLayout(False)
Me.spnRpt.ResumeLayout(False)
Me.GroupBox1.ResumeLayout(False)
Me.gpb2.ResumeLayout(False)
Me.gpb2.PerformLayout()
Me.gpbRecvNotify.ResumeLayout(False)
Me.gpbRecvNotify.PerformLayout()
Me.ResumeLayout(False)

End Sub
    Friend WithEvents spnRpt As System.Windows.Forms.SplitContainer
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents gpb2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblApp As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnAcp As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFeed As System.Windows.Forms.Button
    Friend WithEvents txtFxissue As System.Windows.Forms.TextBox
    Friend WithEvents txtWantDate As System.Windows.Forms.TextBox
    Friend WithEvents mskWantDate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtWanttime As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents gpbRecvNotify As System.Windows.Forms.GroupBox
End Class

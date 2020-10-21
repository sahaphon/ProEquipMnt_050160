<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptEqpTrnsf
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
        Me.Viewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.cboEqpid = New System.Windows.Forms.ComboBox()
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker()
        Me.dtpStart = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gpbMain = New System.Windows.Forms.GroupBox()
        Me.ChkAllEqp = New System.Windows.Forms.CheckBox()
        Me.gpbSub = New System.Windows.Forms.GroupBox()
        Me.ChkTime = New System.Windows.Forms.CheckBox()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.gpbBtn = New System.Windows.Forms.GroupBox()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnCancle = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.gpbMain.SuspendLayout()
        Me.gpbSub.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gpbBtn.SuspendLayout()
        Me.SuspendLayout()
        '
        'Viewer1
        '
        Me.Viewer1.ActiveViewIndex = -1
        Me.Viewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Viewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Viewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Viewer1.Location = New System.Drawing.Point(0, 0)
        Me.Viewer1.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.Viewer1.Name = "Viewer1"
        Me.Viewer1.SelectionFormula = ""
        Me.Viewer1.ShowCloseButton = False
        Me.Viewer1.ShowGotoPageButton = False
        Me.Viewer1.ShowGroupTreeButton = False
        Me.Viewer1.ShowLogo = False
        Me.Viewer1.Size = New System.Drawing.Size(1366, 1388)
        Me.Viewer1.TabIndex = 0
        Me.Viewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        Me.Viewer1.ToolPanelWidth = 400
        Me.Viewer1.ViewTimeSelectionFormula = ""
        '
        'cboEqpid
        '
        Me.cboEqpid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEqpid.FormattingEnabled = True
        Me.cboEqpid.Location = New System.Drawing.Point(12, 67)
        Me.cboEqpid.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.cboEqpid.Name = "cboEqpid"
        Me.cboEqpid.Size = New System.Drawing.Size(360, 33)
        Me.cboEqpid.TabIndex = 14
        '
        'dtpEnd
        '
        Me.dtpEnd.Location = New System.Drawing.Point(74, 144)
        Me.dtpEnd.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.dtpEnd.Name = "dtpEnd"
        Me.dtpEnd.Size = New System.Drawing.Size(298, 31)
        Me.dtpEnd.TabIndex = 11
        '
        'dtpStart
        '
        Me.dtpStart.Location = New System.Drawing.Point(74, 88)
        Me.dtpStart.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(298, 31)
        Me.dtpStart.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(18, 144)
        Me.Label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(52, 30)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "ถึง :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(12, 92)
        Me.Label2.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 30)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "จาก :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Blue
        Me.Label5.Location = New System.Drawing.Point(18, 23)
        Me.Label5.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(273, 37)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "ข้อมูลตามรหัสอุปกรณ์"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(12, 33)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(252, 37)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "ข้อมูลตามช่วงเวลา  "
        '
        'gpbMain
        '
        Me.gpbMain.Controls.Add(Me.ChkAllEqp)
        Me.gpbMain.Controls.Add(Me.cboEqpid)
        Me.gpbMain.Controls.Add(Me.Label5)
        Me.gpbMain.Location = New System.Drawing.Point(12, 23)
        Me.gpbMain.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.gpbMain.Name = "gpbMain"
        Me.gpbMain.Padding = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.gpbMain.Size = New System.Drawing.Size(618, 138)
        Me.gpbMain.TabIndex = 17
        Me.gpbMain.TabStop = False
        '
        'ChkAllEqp
        '
        Me.ChkAllEqp.AutoSize = True
        Me.ChkAllEqp.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ChkAllEqp.ForeColor = System.Drawing.Color.Blue
        Me.ChkAllEqp.Location = New System.Drawing.Point(404, 22)
        Me.ChkAllEqp.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.ChkAllEqp.Name = "ChkAllEqp"
        Me.ChkAllEqp.Size = New System.Drawing.Size(194, 41)
        Me.ChkAllEqp.TabIndex = 15
        Me.ChkAllEqp.Text = "เลือกทั้งหมด"
        Me.ChkAllEqp.UseVisualStyleBackColor = True
        '
        'gpbSub
        '
        Me.gpbSub.Controls.Add(Me.ChkTime)
        Me.gpbSub.Controls.Add(Me.dtpStart)
        Me.gpbSub.Controls.Add(Me.Label1)
        Me.gpbSub.Controls.Add(Me.Label2)
        Me.gpbSub.Controls.Add(Me.Label3)
        Me.gpbSub.Controls.Add(Me.dtpEnd)
        Me.gpbSub.Location = New System.Drawing.Point(12, 173)
        Me.gpbSub.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.gpbSub.Name = "gpbSub"
        Me.gpbSub.Padding = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.gpbSub.Size = New System.Drawing.Size(618, 252)
        Me.gpbSub.TabIndex = 18
        Me.gpbSub.TabStop = False
        '
        'ChkTime
        '
        Me.ChkTime.AutoSize = True
        Me.ChkTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ChkTime.ForeColor = System.Drawing.Color.Blue
        Me.ChkTime.Location = New System.Drawing.Point(404, 29)
        Me.ChkTime.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.ChkTime.Name = "ChkTime"
        Me.ChkTime.Size = New System.Drawing.Size(194, 41)
        Me.ChkTime.TabIndex = 15
        Me.ChkTime.Text = "เลือกทั้งหมด"
        Me.ChkTime.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.gpbBtn)
        Me.SplitContainer1.Panel1.Controls.Add(Me.gpbMain)
        Me.SplitContainer1.Panel1.Controls.Add(Me.gpbSub)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Viewer1)
        Me.SplitContainer1.Size = New System.Drawing.Size(2040, 1388)
        Me.SplitContainer1.SplitterDistance = 666
        Me.SplitContainer1.SplitterWidth = 8
        Me.SplitContainer1.TabIndex = 20
        '
        'gpbBtn
        '
        Me.gpbBtn.Controls.Add(Me.btnExit)
        Me.gpbBtn.Controls.Add(Me.btnCancle)
        Me.gpbBtn.Controls.Add(Me.btnOK)
        Me.gpbBtn.Location = New System.Drawing.Point(12, 437)
        Me.gpbBtn.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.gpbBtn.Name = "gpbBtn"
        Me.gpbBtn.Padding = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.gpbBtn.Size = New System.Drawing.Size(618, 144)
        Me.gpbBtn.TabIndex = 20
        Me.gpbBtn.TabStop = False
        '
        'btnExit
        '
        Me.btnExit.Image = Global.ProEquipMnt.My.Resources.Resources.exit_winxp1
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(440, 37)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(158, 77)
        Me.btnExit.TabIndex = 12
        Me.btnExit.Text = "ปิด"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnCancle
        '
        Me.btnCancle.Image = Global.ProEquipMnt.My.Resources.Resources._Erase
        Me.btnCancle.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancle.Location = New System.Drawing.Point(182, 37)
        Me.btnCancle.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.btnCancle.Name = "btnCancle"
        Me.btnCancle.Size = New System.Drawing.Size(158, 77)
        Me.btnCancle.TabIndex = 12
        Me.btnCancle.Text = "ยกเลิก"
        Me.btnCancle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnCancle.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Image = Global.ProEquipMnt.My.Resources.Resources.Registration
        Me.btnOK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOK.Location = New System.Drawing.Point(14, 37)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(156, 77)
        Me.btnOK.TabIndex = 13
        Me.btnOK.Text = "ตกลง"
        Me.btnOK.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'frmRptEqpTrnsf
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2040, 1388)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.Name = "frmRptEqpTrnsf"
        Me.Text = "รายงานใบโอนอุปกรณ์ลงผลิต"
        Me.gpbMain.ResumeLayout(False)
        Me.gpbMain.PerformLayout()
        Me.gpbSub.ResumeLayout(False)
        Me.gpbSub.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.gpbBtn.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Viewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents cboEqpid As System.Windows.Forms.ComboBox
    Friend WithEvents btnCancle As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents dtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gpbMain As System.Windows.Forms.GroupBox
    Friend WithEvents gpbSub As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents ChkAllEqp As System.Windows.Forms.CheckBox
    Friend WithEvents ChkTime As System.Windows.Forms.CheckBox
    Friend WithEvents gpbBtn As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
End Class

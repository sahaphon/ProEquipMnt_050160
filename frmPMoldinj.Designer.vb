<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPMoldinj
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgvSize = New System.Windows.Forms.DataGridView()
        Me.Column16 = New System.Windows.Forms.DataGridViewLinkColumn()
        Me.Column17 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column18 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgvSelectSize = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column19 = New System.Windows.Forms.DataGridViewLinkColumn()
        Me.btnPrntCancel = New System.Windows.Forms.Button()
        Me.btnPrntPrevw = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.chkAll = New System.Windows.Forms.CheckBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmbOptPrint = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.dgvSize, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSelectSize, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvSize
        '
        Me.dgvSize.AllowUserToAddRows = False
        Me.dgvSize.AllowUserToDeleteRows = False
        Me.dgvSize.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.dgvSize.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvSize.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSize.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvSize.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSize.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column16, Me.Column17, Me.Column18})
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.Info
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvSize.DefaultCellStyle = DataGridViewCellStyle6
        Me.dgvSize.Location = New System.Drawing.Point(17, 81)
        Me.dgvSize.Name = "dgvSize"
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSize.RowHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvSize.RowHeadersVisible = False
        Me.dgvSize.Size = New System.Drawing.Size(372, 520)
        Me.dgvSize.TabIndex = 56
        '
        'Column16
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Column16.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column16.HeaderText = ""
        Me.Column16.LinkColor = System.Drawing.Color.Black
        Me.Column16.Name = "Column16"
        '
        'Column17
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.Blue
        Me.Column17.DefaultCellStyle = DataGridViewCellStyle4
        Me.Column17.HeaderText = "SIZE"
        Me.Column17.Name = "Column17"
        Me.Column17.Width = 80
        '
        'Column18
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.Color.Blue
        Me.Column18.DefaultCellStyle = DataGridViewCellStyle5
        Me.Column18.HeaderText = "กรุ๊ป SIZE"
        Me.Column18.Name = "Column18"
        Me.Column18.Width = 120
        '
        'dgvSelectSize
        '
        Me.dgvSelectSize.AllowUserToAddRows = False
        Me.dgvSelectSize.AllowUserToDeleteRows = False
        Me.dgvSelectSize.AllowUserToResizeRows = False
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.dgvSelectSize.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle8
        Me.dgvSelectSize.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSelectSize.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle9
        Me.dgvSelectSize.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSelectSize.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8, Me.Column19})
        Me.dgvSelectSize.Location = New System.Drawing.Point(406, 81)
        Me.dgvSelectSize.Name = "dgvSelectSize"
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSelectSize.RowHeadersDefaultCellStyle = DataGridViewCellStyle13
        Me.dgvSelectSize.RowHeadersVisible = False
        Me.dgvSelectSize.Size = New System.Drawing.Size(372, 520)
        Me.dgvSelectSize.TabIndex = 57
        '
        'DataGridViewTextBoxColumn7
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle10.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.Color.Blue
        Me.DataGridViewTextBoxColumn7.DefaultCellStyle = DataGridViewCellStyle10
        Me.DataGridViewTextBoxColumn7.HeaderText = "SIZE"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.Width = 80
        '
        'DataGridViewTextBoxColumn8
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.Color.Blue
        Me.DataGridViewTextBoxColumn8.DefaultCellStyle = DataGridViewCellStyle11
        Me.DataGridViewTextBoxColumn8.HeaderText = "กรุ๊ป SIZE"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.Width = 120
        '
        'Column19
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Column19.DefaultCellStyle = DataGridViewCellStyle12
        Me.Column19.HeaderText = ""
        Me.Column19.Name = "Column19"
        Me.Column19.Width = 80
        '
        'btnPrntCancel
        '
        Me.btnPrntCancel.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnPrntCancel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPrntCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnPrntCancel.ForeColor = System.Drawing.Color.Red
        Me.btnPrntCancel.Image = Global.ProEquipMnt.My.Resources.Resources.Cancel_16x16
        Me.btnPrntCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrntCancel.Location = New System.Drawing.Point(709, 609)
        Me.btnPrntCancel.Name = "btnPrntCancel"
        Me.btnPrntCancel.Size = New System.Drawing.Size(69, 39)
        Me.btnPrntCancel.TabIndex = 59
        Me.btnPrntCancel.Text = "ยกเลิก"
        Me.btnPrntCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPrntCancel.UseVisualStyleBackColor = False
        '
        'btnPrntPrevw
        '
        Me.btnPrntPrevw.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnPrntPrevw.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPrntPrevw.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnPrntPrevw.ForeColor = System.Drawing.Color.Blue
        Me.btnPrntPrevw.Image = Global.ProEquipMnt.My.Resources.Resources.printer_large
        Me.btnPrntPrevw.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrntPrevw.Location = New System.Drawing.Point(632, 609)
        Me.btnPrntPrevw.Name = "btnPrntPrevw"
        Me.btnPrntPrevw.Size = New System.Drawing.Size(71, 39)
        Me.btnPrntPrevw.TabIndex = 58
        Me.btnPrntPrevw.Text = "พิมพ์"
        Me.btnPrntPrevw.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPrntPrevw.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(403, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 15)
        Me.Label2.TabIndex = 60
        Me.Label2.Text = "SIZE ที่เลือก"
        '
        'chkAll
        '
        Me.chkAll.AutoSize = True
        Me.chkAll.BackColor = System.Drawing.Color.Transparent
        Me.chkAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.chkAll.ForeColor = System.Drawing.Color.Blue
        Me.chkAll.Location = New System.Drawing.Point(17, 55)
        Me.chkAll.Name = "chkAll"
        Me.chkAll.Size = New System.Drawing.Size(96, 22)
        Me.chkAll.TabIndex = 61
        Me.chkAll.Text = "(ทุก Size)"
        Me.chkAll.UseVisualStyleBackColor = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.ProEquipMnt.My.Resources.Resources.printer_large
        Me.PictureBox2.Location = New System.Drawing.Point(15, 16)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(39, 33)
        Me.PictureBox2.TabIndex = 63
        Me.PictureBox2.TabStop = False
        '
        'cmbOptPrint
        '
        Me.cmbOptPrint.BackColor = System.Drawing.Color.Black
        Me.cmbOptPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cmbOptPrint.ForeColor = System.Drawing.Color.Red
        Me.cmbOptPrint.FormattingEnabled = True
        Me.cmbOptPrint.Location = New System.Drawing.Point(174, 17)
        Me.cmbOptPrint.Name = "cmbOptPrint"
        Me.cmbOptPrint.Size = New System.Drawing.Size(214, 32)
        Me.cmbOptPrint.TabIndex = 62
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Cordia New", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(60, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 26)
        Me.Label1.TabIndex = 64
        Me.Label1.Text = "รหัสอุปกรณ์ :"
        '
        'frmPMoldinj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(796, 655)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmbOptPrint)
        Me.Controls.Add(Me.chkAll)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnPrntCancel)
        Me.Controls.Add(Me.btnPrntPrevw)
        Me.Controls.Add(Me.dgvSelectSize)
        Me.Controls.Add(Me.dgvSize)
        Me.MaximumSize = New System.Drawing.Size(812, 693)
        Me.MinimumSize = New System.Drawing.Size(812, 693)
        Me.Name = "frmPMoldinj"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ตัวเลือกการพิมพ์"
        CType(Me.dgvSize, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSelectSize, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvSize As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSelectSize As System.Windows.Forms.DataGridView
    Friend WithEvents btnPrntCancel As System.Windows.Forms.Button
    Friend WithEvents btnPrntPrevw As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkAll As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmbOptPrint As System.Windows.Forms.ComboBox
    Friend WithEvents Column16 As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents Column17 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column18 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column19 As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class

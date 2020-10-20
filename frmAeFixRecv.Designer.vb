<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAeFixRecv
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
Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAeFixRecv))
Me.gpbHead = New System.Windows.Forms.GroupBox
Me.lblFix_id = New System.Windows.Forms.Label
Me.txtBegin = New System.Windows.Forms.TextBox
Me.Label4 = New System.Windows.Forms.Label
Me.cmbType = New System.Windows.Forms.ComboBox
Me.txtEqp_id = New System.Windows.Forms.TextBox
Me.Label1 = New System.Windows.Forms.Label
Me.mskBegin = New System.Windows.Forms.MaskedTextBox
Me.lblDocTopic = New System.Windows.Forms.Label
Me.Label14 = New System.Windows.Forms.Label
Me.txtEqpnm = New System.Windows.Forms.TextBox
Me.txtRemark = New System.Windows.Forms.TextBox
Me.lblAmount = New System.Windows.Forms.Label
Me.lblAmt = New System.Windows.Forms.Label
Me.Label8 = New System.Windows.Forms.Label
Me.Label3 = New System.Windows.Forms.Label
Me.Label16 = New System.Windows.Forms.Label
Me.txtRef = New System.Windows.Forms.TextBox
Me.btnExit = New System.Windows.Forms.Button
Me.btnSaveData = New System.Windows.Forms.Button
Me.lblComplete = New System.Windows.Forms.Label
Me.Gpb2 = New System.Windows.Forms.GroupBox
Me.txtFxPrice = New System.Windows.Forms.TextBox
Me.mskFxPrice = New System.Windows.Forms.MaskedTextBox
Me.txtRecvAmt = New System.Windows.Forms.TextBox
Me.mskRecvAmt = New System.Windows.Forms.MaskedTextBox
Me.txtSetQty = New System.Windows.Forms.TextBox
Me.mskSetQty = New System.Windows.Forms.MaskedTextBox
Me.Label7 = New System.Windows.Forms.Label
Me.Label2 = New System.Windows.Forms.Label
Me.txtDueDate = New System.Windows.Forms.TextBox
Me.mskDueDate = New System.Windows.Forms.MaskedTextBox
Me.txtRecvDate = New System.Windows.Forms.TextBox
Me.txtFixdate = New System.Windows.Forms.TextBox
Me.txtFixDetail = New System.Windows.Forms.TextBox
Me.txtIssue = New System.Windows.Forms.TextBox
Me.txtRmk = New System.Windows.Forms.TextBox
Me.txtSupp = New System.Windows.Forms.TextBox
Me.Label17 = New System.Windows.Forms.Label
Me.Label6 = New System.Windows.Forms.Label
Me.Label11 = New System.Windows.Forms.Label
Me.Label13 = New System.Windows.Forms.Label
Me.Label5 = New System.Windows.Forms.Label
Me.Label9 = New System.Windows.Forms.Label
Me.Label15 = New System.Windows.Forms.Label
Me.Label12 = New System.Windows.Forms.Label
Me.Label31 = New System.Windows.Forms.Label
Me.Label19 = New System.Windows.Forms.Label
Me.Label10 = New System.Windows.Forms.Label
Me.Label34 = New System.Windows.Forms.Label
Me.txtRecvBy = New System.Windows.Forms.TextBox
Me.txtFixnm = New System.Windows.Forms.TextBox
Me.txtPr = New System.Windows.Forms.TextBox
Me.Label35 = New System.Windows.Forms.Label
Me.txtSize = New System.Windows.Forms.TextBox
Me.mskRecvDate = New System.Windows.Forms.MaskedTextBox
Me.Label25 = New System.Windows.Forms.Label
Me.mskFixdate = New System.Windows.Forms.MaskedTextBox
Me.Label18 = New System.Windows.Forms.Label
Me.gpbHead.SuspendLayout()
Me.Gpb2.SuspendLayout()
Me.SuspendLayout()
'
'gpbHead
'
Me.gpbHead.Controls.Add(Me.lblFix_id)
Me.gpbHead.Controls.Add(Me.txtBegin)
Me.gpbHead.Controls.Add(Me.Label4)
Me.gpbHead.Controls.Add(Me.cmbType)
Me.gpbHead.Controls.Add(Me.txtEqp_id)
Me.gpbHead.Controls.Add(Me.Label1)
Me.gpbHead.Controls.Add(Me.mskBegin)
Me.gpbHead.Controls.Add(Me.lblDocTopic)
Me.gpbHead.Controls.Add(Me.Label14)
Me.gpbHead.Controls.Add(Me.txtEqpnm)
Me.gpbHead.Controls.Add(Me.txtRemark)
Me.gpbHead.Controls.Add(Me.lblAmount)
Me.gpbHead.Controls.Add(Me.lblAmt)
Me.gpbHead.Controls.Add(Me.Label8)
Me.gpbHead.Controls.Add(Me.Label18)
Me.gpbHead.Controls.Add(Me.Label3)
Me.gpbHead.Controls.Add(Me.Label16)
Me.gpbHead.Location = New System.Drawing.Point(12, 12)
Me.gpbHead.Name = "gpbHead"
Me.gpbHead.Size = New System.Drawing.Size(882, 209)
Me.gpbHead.TabIndex = 160
Me.gpbHead.TabStop = False
'
'lblFix_id
'
Me.lblFix_id.BackColor = System.Drawing.Color.Black
Me.lblFix_id.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.lblFix_id.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblFix_id.ForeColor = System.Drawing.Color.Yellow
Me.lblFix_id.Location = New System.Drawing.Point(78, 25)
Me.lblFix_id.Name = "lblFix_id"
Me.lblFix_id.Size = New System.Drawing.Size(135, 27)
Me.lblFix_id.TabIndex = 152
Me.lblFix_id.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'txtBegin
'
Me.txtBegin.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtBegin.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtBegin.Location = New System.Drawing.Point(766, 168)
Me.txtBegin.MaxLength = 5
Me.txtBegin.Name = "txtBegin"
Me.txtBegin.Size = New System.Drawing.Size(110, 29)
Me.txtBegin.TabIndex = 70
Me.txtBegin.Text = "__/__/____"
Me.txtBegin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
Me.txtBegin.Visible = False
'
'Label4
'
Me.Label4.AutoSize = True
Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label4.ForeColor = System.Drawing.Color.Black
Me.Label4.Location = New System.Drawing.Point(7, 29)
Me.Label4.Name = "Label4"
Me.Label4.Size = New System.Drawing.Size(68, 16)
Me.Label4.TabIndex = 66
Me.Label4.Text = "รหัสส่งซ่อม :"
Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'cmbType
'
Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
Me.cmbType.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.cmbType.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.cmbType.FormattingEnabled = True
Me.cmbType.Location = New System.Drawing.Point(317, 23)
Me.cmbType.Name = "cmbType"
Me.cmbType.Size = New System.Drawing.Size(275, 28)
Me.cmbType.TabIndex = 3
'
'txtEqp_id
'
Me.txtEqp_id.AcceptsReturn = True
Me.txtEqp_id.BackColor = System.Drawing.Color.Black
Me.txtEqp_id.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtEqp_id.ForeColor = System.Drawing.Color.Yellow
Me.txtEqp_id.Location = New System.Drawing.Point(78, 66)
Me.txtEqp_id.MaxLength = 12
Me.txtEqp_id.Name = "txtEqp_id"
Me.txtEqp_id.Size = New System.Drawing.Size(157, 31)
Me.txtEqp_id.TabIndex = 0
Me.txtEqp_id.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'Label1
'
Me.Label1.AutoSize = True
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label1.ForeColor = System.Drawing.Color.Black
Me.Label1.Location = New System.Drawing.Point(244, 72)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(99, 16)
Me.Label1.TabIndex = 52
Me.Label1.Text = "รายละเอียดอุปกรณ์ :"
Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'mskBegin
'
Me.mskBegin.BackColor = System.Drawing.Color.SlateBlue
Me.mskBegin.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.mskBegin.ForeColor = System.Drawing.Color.White
Me.mskBegin.Location = New System.Drawing.Point(766, 168)
Me.mskBegin.Mask = "99/99/9999"
Me.mskBegin.Name = "mskBegin"
Me.mskBegin.Size = New System.Drawing.Size(110, 29)
Me.mskBegin.TabIndex = 71
Me.mskBegin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
Me.mskBegin.Visible = False
'
'lblDocTopic
'
Me.lblDocTopic.AutoSize = True
Me.lblDocTopic.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblDocTopic.ForeColor = System.Drawing.Color.Black
Me.lblDocTopic.Location = New System.Drawing.Point(6, 72)
Me.lblDocTopic.Name = "lblDocTopic"
Me.lblDocTopic.Size = New System.Drawing.Size(69, 16)
Me.lblDocTopic.TabIndex = 52
Me.lblDocTopic.Text = "รหัสอุปกรณ์ :"
Me.lblDocTopic.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label14
'
Me.Label14.AutoSize = True
Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label14.ForeColor = System.Drawing.Color.Black
Me.Label14.Location = New System.Drawing.Point(231, 29)
Me.Label14.Name = "Label14"
Me.Label14.Size = New System.Drawing.Size(85, 16)
Me.Label14.TabIndex = 68
Me.Label14.Text = "ประเภทอุปกรณ์ :"
Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'txtEqpnm
'
Me.txtEqpnm.BackColor = System.Drawing.Color.Black
Me.txtEqpnm.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtEqpnm.ForeColor = System.Drawing.Color.Yellow
Me.txtEqpnm.Location = New System.Drawing.Point(344, 65)
Me.txtEqpnm.MaxLength = 150
Me.txtEqpnm.Name = "txtEqpnm"
Me.txtEqpnm.Size = New System.Drawing.Size(445, 31)
Me.txtEqpnm.TabIndex = 2
'
'txtRemark
'
Me.txtRemark.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtRemark.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtRemark.Location = New System.Drawing.Point(78, 160)
Me.txtRemark.MaxLength = 150
Me.txtRemark.Name = "txtRemark"
Me.txtRemark.Size = New System.Drawing.Size(368, 29)
Me.txtRemark.TabIndex = 4
'
'lblAmount
'
Me.lblAmount.BackColor = System.Drawing.Color.Black
Me.lblAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblAmount.ForeColor = System.Drawing.Color.Yellow
Me.lblAmount.Location = New System.Drawing.Point(78, 112)
Me.lblAmount.Name = "lblAmount"
Me.lblAmount.Size = New System.Drawing.Size(72, 32)
Me.lblAmount.TabIndex = 151
Me.lblAmount.Text = "0"
Me.lblAmount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lblAmt
'
Me.lblAmt.BackColor = System.Drawing.Color.Black
Me.lblAmt.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblAmt.ForeColor = System.Drawing.Color.Yellow
Me.lblAmt.Location = New System.Drawing.Point(344, 113)
Me.lblAmt.Name = "lblAmt"
Me.lblAmt.Size = New System.Drawing.Size(128, 32)
Me.lblAmt.TabIndex = 151
Me.lblAmt.Text = "0.00"
Me.lblAmt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'Label8
'
Me.Label8.AutoSize = True
Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label8.ForeColor = System.Drawing.Color.Black
Me.Label8.Location = New System.Drawing.Point(13, 166)
Me.Label8.Name = "Label8"
Me.Label8.Size = New System.Drawing.Size(61, 16)
Me.Label8.TabIndex = 64
Me.Label8.Text = "หมายเหตุ  :"
Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label3
'
Me.Label3.AutoSize = True
Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label3.ForeColor = System.Drawing.Color.Black
Me.Label3.Location = New System.Drawing.Point(220, 121)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(118, 16)
Me.Label3.TabIndex = 64
Me.Label3.Text = "ราคาอุปกรณ์รวม(บาท) :"
Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label16
'
Me.Label16.AutoSize = True
Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label16.ForeColor = System.Drawing.Color.Black
Me.Label16.Location = New System.Drawing.Point(8, 120)
Me.Label16.Name = "Label16"
Me.Label16.Size = New System.Drawing.Size(66, 16)
Me.Label16.TabIndex = 66
Me.Label16.Text = "รวมส่งซ่อม :"
Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'txtRef
'
Me.txtRef.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtRef.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtRef.Location = New System.Drawing.Point(904, 163)
Me.txtRef.MaxLength = 12
Me.txtRef.Name = "txtRef"
Me.txtRef.Size = New System.Drawing.Size(116, 29)
Me.txtRef.TabIndex = 171
Me.txtRef.Visible = False
'
'btnExit
'
Me.btnExit.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
Me.btnExit.Location = New System.Drawing.Point(904, 78)
Me.btnExit.Name = "btnExit"
Me.btnExit.Size = New System.Drawing.Size(117, 49)
Me.btnExit.TabIndex = 170
Me.btnExit.Text = "ออกจากหน้าจอ"
Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.btnExit.UseVisualStyleBackColor = True
'
'btnSaveData
'
Me.btnSaveData.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnSaveData.Image = CType(resources.GetObject("btnSaveData.Image"), System.Drawing.Image)
Me.btnSaveData.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.btnSaveData.Location = New System.Drawing.Point(904, 18)
Me.btnSaveData.Name = "btnSaveData"
Me.btnSaveData.Size = New System.Drawing.Size(117, 49)
Me.btnSaveData.TabIndex = 169
Me.btnSaveData.Text = "บันทึกข้อมูล"
Me.btnSaveData.TextAlign = System.Drawing.ContentAlignment.MiddleRight
Me.btnSaveData.UseVisualStyleBackColor = True
'
'lblComplete
'
Me.lblComplete.BackColor = System.Drawing.Color.Gold
Me.lblComplete.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblComplete.ForeColor = System.Drawing.Color.GhostWhite
Me.lblComplete.Location = New System.Drawing.Point(904, 130)
Me.lblComplete.Name = "lblComplete"
Me.lblComplete.Size = New System.Drawing.Size(117, 30)
Me.lblComplete.TabIndex = 168
Me.lblComplete.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblComplete.Visible = False
'
'Gpb2
'
Me.Gpb2.BackColor = System.Drawing.SystemColors.Control
Me.Gpb2.Controls.Add(Me.txtFxPrice)
Me.Gpb2.Controls.Add(Me.mskFxPrice)
Me.Gpb2.Controls.Add(Me.txtRecvAmt)
Me.Gpb2.Controls.Add(Me.mskRecvAmt)
Me.Gpb2.Controls.Add(Me.txtSetQty)
Me.Gpb2.Controls.Add(Me.mskSetQty)
Me.Gpb2.Controls.Add(Me.Label7)
Me.Gpb2.Controls.Add(Me.Label2)
Me.Gpb2.Controls.Add(Me.txtDueDate)
Me.Gpb2.Controls.Add(Me.mskDueDate)
Me.Gpb2.Controls.Add(Me.txtRecvDate)
Me.Gpb2.Controls.Add(Me.txtFixdate)
Me.Gpb2.Controls.Add(Me.txtFixDetail)
Me.Gpb2.Controls.Add(Me.txtIssue)
Me.Gpb2.Controls.Add(Me.txtRmk)
Me.Gpb2.Controls.Add(Me.txtSupp)
Me.Gpb2.Controls.Add(Me.Label17)
Me.Gpb2.Controls.Add(Me.Label6)
Me.Gpb2.Controls.Add(Me.Label11)
Me.Gpb2.Controls.Add(Me.Label13)
Me.Gpb2.Controls.Add(Me.Label5)
Me.Gpb2.Controls.Add(Me.Label9)
Me.Gpb2.Controls.Add(Me.Label15)
Me.Gpb2.Controls.Add(Me.Label12)
Me.Gpb2.Controls.Add(Me.Label31)
Me.Gpb2.Controls.Add(Me.Label19)
Me.Gpb2.Controls.Add(Me.Label10)
Me.Gpb2.Controls.Add(Me.Label34)
Me.Gpb2.Controls.Add(Me.txtRecvBy)
Me.Gpb2.Controls.Add(Me.txtFixnm)
Me.Gpb2.Controls.Add(Me.txtPr)
Me.Gpb2.Controls.Add(Me.Label35)
Me.Gpb2.Controls.Add(Me.txtSize)
Me.Gpb2.Controls.Add(Me.mskRecvDate)
Me.Gpb2.Controls.Add(Me.Label25)
Me.Gpb2.Controls.Add(Me.mskFixdate)
Me.Gpb2.Location = New System.Drawing.Point(12, 288)
Me.Gpb2.Name = "Gpb2"
Me.Gpb2.Size = New System.Drawing.Size(981, 422)
Me.Gpb2.TabIndex = 175
Me.Gpb2.TabStop = False
'
'txtFxPrice
'
Me.txtFxPrice.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtFxPrice.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtFxPrice.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
Me.txtFxPrice.Location = New System.Drawing.Point(817, 277)
Me.txtFxPrice.MaxLength = 100
Me.txtFxPrice.Name = "txtFxPrice"
Me.txtFxPrice.Size = New System.Drawing.Size(112, 29)
Me.txtFxPrice.TabIndex = 191
Me.txtFxPrice.Text = "0.00"
Me.txtFxPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
'
'mskFxPrice
'
Me.mskFxPrice.BackColor = System.Drawing.Color.Purple
Me.mskFxPrice.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.mskFxPrice.ForeColor = System.Drawing.Color.White
Me.mskFxPrice.InsertKeyMode = System.Windows.Forms.InsertKeyMode.Insert
Me.mskFxPrice.Location = New System.Drawing.Point(817, 278)
Me.mskFxPrice.Mask = "###,##0.00"
Me.mskFxPrice.Name = "mskFxPrice"
Me.mskFxPrice.Size = New System.Drawing.Size(112, 29)
Me.mskFxPrice.SkipLiterals = False
Me.mskFxPrice.TabIndex = 192
Me.mskFxPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'txtRecvAmt
'
Me.txtRecvAmt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtRecvAmt.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtRecvAmt.ForeColor = System.Drawing.Color.Green
Me.txtRecvAmt.Location = New System.Drawing.Point(631, 280)
Me.txtRecvAmt.MaxLength = 100
Me.txtRecvAmt.Name = "txtRecvAmt"
Me.txtRecvAmt.Size = New System.Drawing.Size(69, 29)
Me.txtRecvAmt.TabIndex = 182
Me.txtRecvAmt.Text = "0"
Me.txtRecvAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'mskRecvAmt
'
Me.mskRecvAmt.BackColor = System.Drawing.Color.Purple
Me.mskRecvAmt.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.mskRecvAmt.ForeColor = System.Drawing.SystemColors.Window
Me.mskRecvAmt.Location = New System.Drawing.Point(631, 280)
Me.mskRecvAmt.Mask = "99.9"
Me.mskRecvAmt.Name = "mskRecvAmt"
Me.mskRecvAmt.Size = New System.Drawing.Size(69, 29)
Me.mskRecvAmt.TabIndex = 183
Me.mskRecvAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'txtSetQty
'
Me.txtSetQty.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtSetQty.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtSetQty.ForeColor = System.Drawing.Color.Green
Me.txtSetQty.Location = New System.Drawing.Point(293, 53)
Me.txtSetQty.MaxLength = 100
Me.txtSetQty.Name = "txtSetQty"
Me.txtSetQty.Size = New System.Drawing.Size(69, 29)
Me.txtSetQty.TabIndex = 182
Me.txtSetQty.Text = "0"
Me.txtSetQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'mskSetQty
'
Me.mskSetQty.BackColor = System.Drawing.Color.Purple
Me.mskSetQty.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.mskSetQty.ForeColor = System.Drawing.SystemColors.Window
Me.mskSetQty.Location = New System.Drawing.Point(293, 53)
Me.mskSetQty.Mask = "99.9"
Me.mskSetQty.Name = "mskSetQty"
Me.mskSetQty.Size = New System.Drawing.Size(69, 29)
Me.mskSetQty.TabIndex = 183
Me.mskSetQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'Label7
'
Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label7.ForeColor = System.Drawing.Color.Red
Me.Label7.Location = New System.Drawing.Point(4, 234)
Me.Label7.Name = "Label7"
Me.Label7.Size = New System.Drawing.Size(181, 28)
Me.Label7.TabIndex = 175
Me.Label7.Text = "รายละเอียดรับเข้า"
'
'Label2
'
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
Me.Label2.Location = New System.Drawing.Point(5, 11)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(181, 28)
Me.Label2.TabIndex = 175
Me.Label2.Text = "ข้อมูลส่งซ่อม"
'
'txtDueDate
'
Me.txtDueDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtDueDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtDueDate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtDueDate.Location = New System.Drawing.Point(75, 139)
Me.txtDueDate.MaxLength = 5
Me.txtDueDate.Name = "txtDueDate"
Me.txtDueDate.Size = New System.Drawing.Size(149, 29)
Me.txtDueDate.TabIndex = 11
Me.txtDueDate.Text = "__/__/____"
Me.txtDueDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'mskDueDate
'
Me.mskDueDate.BackColor = System.Drawing.Color.SlateBlue
Me.mskDueDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.mskDueDate.ForeColor = System.Drawing.Color.White
Me.mskDueDate.Location = New System.Drawing.Point(75, 139)
Me.mskDueDate.Mask = "99/99/9999"
Me.mskDueDate.Name = "mskDueDate"
Me.mskDueDate.Size = New System.Drawing.Size(149, 29)
Me.mskDueDate.TabIndex = 174
Me.mskDueDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'txtRecvDate
'
Me.txtRecvDate.AcceptsReturn = True
Me.txtRecvDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtRecvDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtRecvDate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtRecvDate.Location = New System.Drawing.Point(74, 277)
Me.txtRecvDate.MaxLength = 5
Me.txtRecvDate.Name = "txtRecvDate"
Me.txtRecvDate.Size = New System.Drawing.Size(149, 29)
Me.txtRecvDate.TabIndex = 9
Me.txtRecvDate.Text = "__/__/____"
Me.txtRecvDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'txtFixdate
'
Me.txtFixdate.AcceptsReturn = True
Me.txtFixdate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtFixdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtFixdate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtFixdate.Location = New System.Drawing.Point(74, 96)
Me.txtFixdate.MaxLength = 5
Me.txtFixdate.Name = "txtFixdate"
Me.txtFixdate.Size = New System.Drawing.Size(149, 29)
Me.txtFixdate.TabIndex = 9
Me.txtFixdate.Text = "__/__/____"
Me.txtFixdate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'txtFixDetail
'
Me.txtFixDetail.AcceptsReturn = True
Me.txtFixDetail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtFixDetail.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtFixDetail.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtFixDetail.Location = New System.Drawing.Point(131, 328)
Me.txtFixDetail.MaxLength = 150
Me.txtFixDetail.Multiline = True
Me.txtFixDetail.Name = "txtFixDetail"
Me.txtFixDetail.Size = New System.Drawing.Size(569, 78)
Me.txtFixDetail.TabIndex = 12
'
'txtIssue
'
Me.txtIssue.AcceptsReturn = True
Me.txtIssue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtIssue.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtIssue.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtIssue.Location = New System.Drawing.Point(554, 96)
Me.txtIssue.MaxLength = 150
Me.txtIssue.Multiline = True
Me.txtIssue.Name = "txtIssue"
Me.txtIssue.Size = New System.Drawing.Size(399, 64)
Me.txtIssue.TabIndex = 12
'
'txtRmk
'
Me.txtRmk.AcceptsReturn = True
Me.txtRmk.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtRmk.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtRmk.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtRmk.Location = New System.Drawing.Point(74, 179)
Me.txtRmk.MaxLength = 150
Me.txtRmk.Name = "txtRmk"
Me.txtRmk.Size = New System.Drawing.Size(265, 29)
Me.txtRmk.TabIndex = 13
'
'txtSupp
'
Me.txtSupp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtSupp.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtSupp.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtSupp.Location = New System.Drawing.Point(746, 53)
Me.txtSupp.MaxLength = 50
Me.txtSupp.Name = "txtSupp"
Me.txtSupp.Size = New System.Drawing.Size(207, 29)
Me.txtSupp.TabIndex = 8
'
'Label17
'
Me.Label17.AutoSize = True
Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label17.ForeColor = System.Drawing.Color.Black
Me.Label17.Location = New System.Drawing.Point(22, 58)
Me.Label17.Name = "Label17"
Me.Label17.Size = New System.Drawing.Size(46, 16)
Me.Label17.TabIndex = 81
Me.Label17.Text = "SIZE  :"
Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'Label6
'
Me.Label6.AutoSize = True
Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label6.ForeColor = System.Drawing.Color.Black
Me.Label6.Location = New System.Drawing.Point(653, 59)
Me.Label6.Name = "Label6"
Me.Label6.Size = New System.Drawing.Size(86, 16)
Me.Label6.TabIndex = 172
Me.Label6.Text = "ชื่อร้านที่ส่งซ่อม :"
Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label11
'
Me.Label11.AutoSize = True
Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label11.ForeColor = System.Drawing.Color.Red
Me.Label11.Location = New System.Drawing.Point(238, 285)
Me.Label11.Name = "Label11"
Me.Label11.Size = New System.Drawing.Size(57, 16)
Me.Label11.TabIndex = 74
Me.Label11.Text = "ผู้รับเข้า :"
Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label13
'
Me.Label13.AutoSize = True
Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label13.ForeColor = System.Drawing.Color.Red
Me.Label13.Location = New System.Drawing.Point(7, 333)
Me.Label13.Name = "Label13"
Me.Label13.Size = New System.Drawing.Size(123, 16)
Me.Label13.TabIndex = 64
Me.Label13.Text = "รายละเอียดการซ่อม  :"
Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label5
'
Me.Label5.AutoSize = True
Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label5.ForeColor = System.Drawing.Color.Black
Me.Label5.Location = New System.Drawing.Point(230, 103)
Me.Label5.Name = "Label5"
Me.Label5.Size = New System.Drawing.Size(54, 16)
Me.Label5.TabIndex = 74
Me.Label5.Text = "ผู้ส่งซ่อม :"
Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label9
'
Me.Label9.AutoSize = True
Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label9.ForeColor = System.Drawing.Color.Black
Me.Label9.Location = New System.Drawing.Point(482, 101)
Me.Label9.Name = "Label9"
Me.Label9.Size = New System.Drawing.Size(69, 16)
Me.Label9.TabIndex = 64
Me.Label9.Text = "ปัญหาที่พบ  :"
Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label15
'
Me.Label15.AutoSize = True
Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label15.ForeColor = System.Drawing.Color.Red
Me.Label15.Location = New System.Drawing.Point(717, 284)
Me.Label15.Name = "Label15"
Me.Label15.Size = New System.Drawing.Size(94, 16)
Me.Label15.TabIndex = 155
Me.Label15.Text = "ค่าซ่อม (บาท)  :"
Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label12
'
Me.Label12.AutoSize = True
Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label12.ForeColor = System.Drawing.Color.Red
Me.Label12.Location = New System.Drawing.Point(501, 286)
Me.Label12.Name = "Label12"
Me.Label12.Size = New System.Drawing.Size(126, 16)
Me.Label12.TabIndex = 155
Me.Label12.Text = "จำนวนรับเข้า (Set)  :"
Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label31
'
Me.Label31.AutoSize = True
Me.Label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label31.ForeColor = System.Drawing.Color.Black
Me.Label31.Location = New System.Drawing.Point(7, 184)
Me.Label31.Name = "Label31"
Me.Label31.Size = New System.Drawing.Size(61, 16)
Me.Label31.TabIndex = 64
Me.Label31.Text = "หมายเหตุ  :"
Me.Label31.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label19
'
Me.Label19.AutoSize = True
Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label19.ForeColor = System.Drawing.Color.Black
Me.Label19.Location = New System.Drawing.Point(181, 59)
Me.Label19.Name = "Label19"
Me.Label19.Size = New System.Drawing.Size(109, 16)
Me.Label19.TabIndex = 155
Me.Label19.Text = "จำนวนส่งซ่อม(Set)  :"
Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label10
'
Me.Label10.AutoSize = True
Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label10.ForeColor = System.Drawing.Color.Red
Me.Label10.Location = New System.Drawing.Point(5, 284)
Me.Label10.Name = "Label10"
Me.Label10.Size = New System.Drawing.Size(73, 16)
Me.Label10.TabIndex = 74
Me.Label10.Text = "วันที่รับเข้า :"
Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label34
'
Me.Label34.AutoSize = True
Me.Label34.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label34.ForeColor = System.Drawing.Color.Black
Me.Label34.Location = New System.Drawing.Point(5, 102)
Me.Label34.Name = "Label34"
Me.Label34.Size = New System.Drawing.Size(68, 16)
Me.Label34.TabIndex = 74
Me.Label34.Text = "วันที่ส่งซ่อม :"
Me.Label34.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'txtRecvBy
'
Me.txtRecvBy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtRecvBy.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtRecvBy.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtRecvBy.Location = New System.Drawing.Point(298, 280)
Me.txtRecvBy.MaxLength = 20
Me.txtRecvBy.Name = "txtRecvBy"
Me.txtRecvBy.Size = New System.Drawing.Size(188, 29)
Me.txtRecvBy.TabIndex = 10
'
'txtFixnm
'
Me.txtFixnm.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtFixnm.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtFixnm.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtFixnm.Location = New System.Drawing.Point(287, 98)
Me.txtFixnm.MaxLength = 20
Me.txtFixnm.Name = "txtFixnm"
Me.txtFixnm.Size = New System.Drawing.Size(188, 29)
Me.txtFixnm.TabIndex = 10
'
'txtPr
'
Me.txtPr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtPr.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtPr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtPr.Location = New System.Drawing.Point(437, 53)
Me.txtPr.MaxLength = 20
Me.txtPr.Name = "txtPr"
Me.txtPr.Size = New System.Drawing.Size(205, 29)
Me.txtPr.TabIndex = 7
'
'Label35
'
Me.Label35.AutoSize = True
Me.Label35.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label35.ForeColor = System.Drawing.Color.Black
Me.Label35.Location = New System.Drawing.Point(4, 145)
Me.Label35.Name = "Label35"
Me.Label35.Size = New System.Drawing.Size(64, 16)
Me.Label35.TabIndex = 74
Me.Label35.Text = "วันที่นัดเข้า :"
Me.Label35.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'txtSize
'
Me.txtSize.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtSize.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtSize.Location = New System.Drawing.Point(74, 53)
Me.txtSize.MaxLength = 12
Me.txtSize.Name = "txtSize"
Me.txtSize.Size = New System.Drawing.Size(97, 29)
Me.txtSize.TabIndex = 5
Me.txtSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'mskRecvDate
'
Me.mskRecvDate.BackColor = System.Drawing.Color.SlateBlue
Me.mskRecvDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.mskRecvDate.ForeColor = System.Drawing.Color.White
Me.mskRecvDate.Location = New System.Drawing.Point(74, 278)
Me.mskRecvDate.Mask = "99/99/9999"
Me.mskRecvDate.Name = "mskRecvDate"
Me.mskRecvDate.Size = New System.Drawing.Size(149, 29)
Me.mskRecvDate.TabIndex = 174
Me.mskRecvDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'Label25
'
Me.Label25.AutoSize = True
Me.Label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label25.ForeColor = System.Drawing.Color.Black
Me.Label25.Location = New System.Drawing.Point(372, 60)
Me.Label25.Name = "Label25"
Me.Label25.Size = New System.Drawing.Size(60, 16)
Me.Label25.TabIndex = 81
Me.Label25.Text = "เลขที่ PR :"
Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'mskFixdate
'
Me.mskFixdate.BackColor = System.Drawing.Color.SlateBlue
Me.mskFixdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.mskFixdate.ForeColor = System.Drawing.Color.White
Me.mskFixdate.Location = New System.Drawing.Point(74, 96)
Me.mskFixdate.Mask = "99/99/9999"
Me.mskFixdate.Name = "mskFixdate"
Me.mskFixdate.Size = New System.Drawing.Size(149, 29)
Me.mskFixdate.TabIndex = 174
Me.mskFixdate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'Label18
'
Me.Label18.AutoSize = True
Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label18.ForeColor = System.Drawing.Color.Black
Me.Label18.Location = New System.Drawing.Point(153, 123)
Me.Label18.Name = "Label18"
Me.Label18.Size = New System.Drawing.Size(38, 16)
Me.Label18.TabIndex = 64
Me.Label18.Text = "SET."
Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'frmAeFixRecv
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(1028, 722)
Me.Controls.Add(Me.Gpb2)
Me.Controls.Add(Me.txtRef)
Me.Controls.Add(Me.btnExit)
Me.Controls.Add(Me.btnSaveData)
Me.Controls.Add(Me.lblComplete)
Me.Controls.Add(Me.gpbHead)
Me.Name = "frmAeFixRecv"
Me.Text = "ฟอร์มแก้ไขรับเข้าส่งซ่อม"
Me.gpbHead.ResumeLayout(False)
Me.gpbHead.PerformLayout()
Me.Gpb2.ResumeLayout(False)
Me.Gpb2.PerformLayout()
Me.ResumeLayout(False)
Me.PerformLayout()

End Sub
    Friend WithEvents gpbHead As System.Windows.Forms.GroupBox
    Friend WithEvents lblFix_id As System.Windows.Forms.Label
    Friend WithEvents txtBegin As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents txtEqp_id As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents mskBegin As System.Windows.Forms.MaskedTextBox
    Friend WithEvents lblDocTopic As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtEqpnm As System.Windows.Forms.TextBox
    Friend WithEvents txtRemark As System.Windows.Forms.TextBox
    Friend WithEvents lblAmount As System.Windows.Forms.Label
    Friend WithEvents lblAmt As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtRef As System.Windows.Forms.TextBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnSaveData As System.Windows.Forms.Button
    Friend WithEvents lblComplete As System.Windows.Forms.Label
    Friend WithEvents Gpb2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSetQty As System.Windows.Forms.TextBox
    Friend WithEvents mskSetQty As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDueDate As System.Windows.Forms.TextBox
    Friend WithEvents mskDueDate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtFixdate As System.Windows.Forms.TextBox
    Friend WithEvents txtIssue As System.Windows.Forms.TextBox
    Friend WithEvents txtRmk As System.Windows.Forms.TextBox
    Friend WithEvents txtSupp As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents txtFixnm As System.Windows.Forms.TextBox
    Friend WithEvents txtPr As System.Windows.Forms.TextBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents txtSize As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents mskFixdate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtRecvAmt As System.Windows.Forms.TextBox
    Friend WithEvents mskRecvAmt As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtRecvDate As System.Windows.Forms.TextBox
    Friend WithEvents txtFixDetail As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtRecvBy As System.Windows.Forms.TextBox
    Friend WithEvents mskRecvDate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtFxPrice As System.Windows.Forms.TextBox
    Friend WithEvents mskFxPrice As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
End Class

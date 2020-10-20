<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAeNotifyIssue
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
Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAeNotifyIssue))
Me.lblDocTopic = New System.Windows.Forms.Label
Me.gpbNotify = New System.Windows.Forms.GroupBox
Me.cboDepto = New System.Windows.Forms.ComboBox
Me.txtNeedtime = New System.Windows.Forms.TextBox
Me.txtName = New System.Windows.Forms.TextBox
Me.Label19 = New System.Windows.Forms.Label
Me.txtDocid = New System.Windows.Forms.TextBox
Me.picEqp = New System.Windows.Forms.PictureBox
Me.txtSizeQty = New System.Windows.Forms.TextBox
Me.mskSizeQty = New System.Windows.Forms.MaskedTextBox
Me.lblPicName = New System.Windows.Forms.Label
Me.btnDelEqp1 = New System.Windows.Forms.Button
Me.btnEditEqp1 = New System.Windows.Forms.Button
Me.lblPicPath = New System.Windows.Forms.Label
Me.txtNeedDate = New System.Windows.Forms.TextBox
Me.mskNeedDate = New System.Windows.Forms.MaskedTextBox
Me.txtBegin = New System.Windows.Forms.TextBox
Me.mskBegin = New System.Windows.Forms.MaskedTextBox
Me.txtCause = New System.Windows.Forms.TextBox
Me.txtIssue = New System.Windows.Forms.TextBox
Me.txtRemark = New System.Windows.Forms.TextBox
Me.txtEqpnm = New System.Windows.Forms.TextBox
Me.txtSize = New System.Windows.Forms.TextBox
Me.txtShoe = New System.Windows.Forms.TextBox
Me.txtOrder = New System.Windows.Forms.TextBox
Me.txtFrom = New System.Windows.Forms.TextBox
Me.cboGroup = New System.Windows.Forms.ComboBox
Me.cboDepfrom = New System.Windows.Forms.ComboBox
Me.Label14 = New System.Windows.Forms.Label
Me.Label6 = New System.Windows.Forms.Label
Me.Label10 = New System.Windows.Forms.Label
Me.Label8 = New System.Windows.Forms.Label
Me.Label17 = New System.Windows.Forms.Label
Me.Label9 = New System.Windows.Forms.Label
Me.Label7 = New System.Windows.Forms.Label
Me.Label12 = New System.Windows.Forms.Label
Me.Label5 = New System.Windows.Forms.Label
Me.Label18 = New System.Windows.Forms.Label
Me.Label4 = New System.Windows.Forms.Label
Me.Label1 = New System.Windows.Forms.Label
Me.Label3 = New System.Windows.Forms.Label
Me.Label2 = New System.Windows.Forms.Label
Me.Label11 = New System.Windows.Forms.Label
Me.gpbReceive = New System.Windows.Forms.GroupBox
Me.txtWantDate = New System.Windows.Forms.TextBox
Me.mskWantDate = New System.Windows.Forms.MaskedTextBox
Me.txtFxIssue = New System.Windows.Forms.TextBox
Me.Label13 = New System.Windows.Forms.Label
Me.Label15 = New System.Windows.Forms.Label
Me.Label16 = New System.Windows.Forms.Label
Me.txtWanttime = New System.Windows.Forms.TextBox
Me.txtRef = New System.Windows.Forms.TextBox
Me.lblComplete = New System.Windows.Forms.Label
Me.btnExit = New System.Windows.Forms.Button
Me.btnSaveData = New System.Windows.Forms.Button
Me.gpbNotify.SuspendLayout()
CType(Me.picEqp, System.ComponentModel.ISupportInitialize).BeginInit()
Me.gpbReceive.SuspendLayout()
Me.SuspendLayout()
'
'lblDocTopic
'
Me.lblDocTopic.AutoSize = True
Me.lblDocTopic.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblDocTopic.ForeColor = System.Drawing.Color.Black
Me.lblDocTopic.Location = New System.Drawing.Point(18, 73)
Me.lblDocTopic.Name = "lblDocTopic"
Me.lblDocTopic.Size = New System.Drawing.Size(73, 16)
Me.lblDocTopic.TabIndex = 53
Me.lblDocTopic.Text = "เรียน (แผนก) :"
Me.lblDocTopic.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'gpbNotify
'
Me.gpbNotify.BackColor = System.Drawing.SystemColors.Control
Me.gpbNotify.Controls.Add(Me.cboDepto)
Me.gpbNotify.Controls.Add(Me.txtNeedtime)
Me.gpbNotify.Controls.Add(Me.txtName)
Me.gpbNotify.Controls.Add(Me.Label19)
Me.gpbNotify.Controls.Add(Me.txtDocid)
Me.gpbNotify.Controls.Add(Me.picEqp)
Me.gpbNotify.Controls.Add(Me.txtSizeQty)
Me.gpbNotify.Controls.Add(Me.mskSizeQty)
Me.gpbNotify.Controls.Add(Me.lblPicName)
Me.gpbNotify.Controls.Add(Me.btnDelEqp1)
Me.gpbNotify.Controls.Add(Me.btnEditEqp1)
Me.gpbNotify.Controls.Add(Me.lblPicPath)
Me.gpbNotify.Controls.Add(Me.txtNeedDate)
Me.gpbNotify.Controls.Add(Me.mskNeedDate)
Me.gpbNotify.Controls.Add(Me.txtBegin)
Me.gpbNotify.Controls.Add(Me.mskBegin)
Me.gpbNotify.Controls.Add(Me.txtCause)
Me.gpbNotify.Controls.Add(Me.txtIssue)
Me.gpbNotify.Controls.Add(Me.txtRemark)
Me.gpbNotify.Controls.Add(Me.txtEqpnm)
Me.gpbNotify.Controls.Add(Me.txtSize)
Me.gpbNotify.Controls.Add(Me.txtShoe)
Me.gpbNotify.Controls.Add(Me.txtOrder)
Me.gpbNotify.Controls.Add(Me.txtFrom)
Me.gpbNotify.Controls.Add(Me.cboGroup)
Me.gpbNotify.Controls.Add(Me.cboDepfrom)
Me.gpbNotify.Controls.Add(Me.Label14)
Me.gpbNotify.Controls.Add(Me.Label6)
Me.gpbNotify.Controls.Add(Me.Label10)
Me.gpbNotify.Controls.Add(Me.Label8)
Me.gpbNotify.Controls.Add(Me.Label17)
Me.gpbNotify.Controls.Add(Me.Label9)
Me.gpbNotify.Controls.Add(Me.Label7)
Me.gpbNotify.Controls.Add(Me.Label12)
Me.gpbNotify.Controls.Add(Me.Label5)
Me.gpbNotify.Controls.Add(Me.Label18)
Me.gpbNotify.Controls.Add(Me.Label4)
Me.gpbNotify.Controls.Add(Me.Label1)
Me.gpbNotify.Controls.Add(Me.Label3)
Me.gpbNotify.Controls.Add(Me.Label2)
Me.gpbNotify.Controls.Add(Me.Label11)
Me.gpbNotify.Controls.Add(Me.lblDocTopic)
Me.gpbNotify.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.gpbNotify.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.gpbNotify.Location = New System.Drawing.Point(12, 12)
Me.gpbNotify.Name = "gpbNotify"
Me.gpbNotify.Size = New System.Drawing.Size(898, 586)
Me.gpbNotify.TabIndex = 54
Me.gpbNotify.TabStop = False
Me.gpbNotify.Text = "สำหรับผู้แจ้งปัญหา"
'
'cboDepto
'
Me.cboDepto.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.cboDepto.ForeColor = System.Drawing.Color.Navy
Me.cboDepto.FormattingEnabled = True
Me.cboDepto.Location = New System.Drawing.Point(97, 72)
Me.cboDepto.Name = "cboDepto"
Me.cboDepto.Size = New System.Drawing.Size(209, 26)
Me.cboDepto.TabIndex = 187
'
'txtNeedtime
'
Me.txtNeedtime.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtNeedtime.ForeColor = System.Drawing.Color.Navy
Me.txtNeedtime.Location = New System.Drawing.Point(353, 502)
Me.txtNeedtime.MaxLength = 5
Me.txtNeedtime.Name = "txtNeedtime"
Me.txtNeedtime.Size = New System.Drawing.Size(112, 29)
Me.txtNeedtime.TabIndex = 14
Me.txtNeedtime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'txtName
'
Me.txtName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtName.ForeColor = System.Drawing.Color.Navy
Me.txtName.Location = New System.Drawing.Point(353, 110)
Me.txtName.Name = "txtName"
Me.txtName.Size = New System.Drawing.Size(124, 26)
Me.txtName.TabIndex = 3
'
'Label19
'
Me.Label19.AutoSize = True
Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label19.ForeColor = System.Drawing.Color.Black
Me.Label19.Location = New System.Drawing.Point(313, 114)
Me.Label19.Name = "Label19"
Me.Label19.Size = New System.Drawing.Size(33, 16)
Me.Label19.TabIndex = 186
Me.Label19.Text = "ผู้แจ้ง"
Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'txtDocid
'
Me.txtDocid.BackColor = System.Drawing.Color.Black
Me.txtDocid.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtDocid.ForeColor = System.Drawing.Color.Yellow
Me.txtDocid.Location = New System.Drawing.Point(97, 37)
Me.txtDocid.Name = "txtDocid"
Me.txtDocid.Size = New System.Drawing.Size(145, 29)
Me.txtDocid.TabIndex = 0
Me.txtDocid.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'picEqp
'
Me.picEqp.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
Me.picEqp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
Me.picEqp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.picEqp.Cursor = System.Windows.Forms.Cursors.Hand
Me.picEqp.ImageLocation = ""
Me.picEqp.Location = New System.Drawing.Point(575, 37)
Me.picEqp.Name = "picEqp"
Me.picEqp.Size = New System.Drawing.Size(313, 276)
Me.picEqp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
Me.picEqp.TabIndex = 184
Me.picEqp.TabStop = False
'
'txtSizeQty
'
Me.txtSizeQty.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtSizeQty.ForeColor = System.Drawing.Color.Green
Me.txtSizeQty.Location = New System.Drawing.Point(500, 338)
Me.txtSizeQty.MaxLength = 100
Me.txtSizeQty.Name = "txtSizeQty"
Me.txtSizeQty.Size = New System.Drawing.Size(69, 29)
Me.txtSizeQty.TabIndex = 9
Me.txtSizeQty.Text = "0"
Me.txtSizeQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'mskSizeQty
'
Me.mskSizeQty.BackColor = System.Drawing.Color.Purple
Me.mskSizeQty.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.mskSizeQty.ForeColor = System.Drawing.SystemColors.Window
Me.mskSizeQty.Location = New System.Drawing.Point(500, 338)
Me.mskSizeQty.Mask = "99.9"
Me.mskSizeQty.Name = "mskSizeQty"
Me.mskSizeQty.Size = New System.Drawing.Size(69, 29)
Me.mskSizeQty.TabIndex = 183
Me.mskSizeQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'lblPicName
'
Me.lblPicName.BackColor = System.Drawing.Color.DimGray
Me.lblPicName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblPicName.ForeColor = System.Drawing.Color.GhostWhite
Me.lblPicName.Location = New System.Drawing.Point(581, 249)
Me.lblPicName.Name = "lblPicName"
Me.lblPicName.Size = New System.Drawing.Size(300, 30)
Me.lblPicName.TabIndex = 181
Me.lblPicName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblPicName.Visible = False
'
'btnDelEqp1
'
Me.btnDelEqp1.BackColor = System.Drawing.SystemColors.Control
Me.btnDelEqp1.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnDelEqp1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.btnDelEqp1.ForeColor = System.Drawing.Color.Navy
Me.btnDelEqp1.Image = CType(resources.GetObject("btnDelEqp1.Image"), System.Drawing.Image)
Me.btnDelEqp1.ImageAlign = System.Drawing.ContentAlignment.TopLeft
Me.btnDelEqp1.Location = New System.Drawing.Point(822, 322)
Me.btnDelEqp1.Name = "btnDelEqp1"
Me.btnDelEqp1.Size = New System.Drawing.Size(66, 34)
Me.btnDelEqp1.TabIndex = 179
Me.btnDelEqp1.Text = "ลบรูป"
Me.btnDelEqp1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
Me.btnDelEqp1.UseVisualStyleBackColor = False
'
'btnEditEqp1
'
Me.btnEditEqp1.BackColor = System.Drawing.SystemColors.Control
Me.btnEditEqp1.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnEditEqp1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.btnEditEqp1.ForeColor = System.Drawing.Color.Navy
Me.btnEditEqp1.Image = CType(resources.GetObject("btnEditEqp1.Image"), System.Drawing.Image)
Me.btnEditEqp1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.btnEditEqp1.Location = New System.Drawing.Point(700, 322)
Me.btnEditEqp1.Name = "btnEditEqp1"
Me.btnEditEqp1.Size = New System.Drawing.Size(116, 34)
Me.btnEditEqp1.TabIndex = 178
Me.btnEditEqp1.Text = "เพิ่ม / แก้ไขรูป"
Me.btnEditEqp1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
Me.btnEditEqp1.UseVisualStyleBackColor = False
'
'lblPicPath
'
Me.lblPicPath.BackColor = System.Drawing.Color.Gray
Me.lblPicPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblPicPath.ForeColor = System.Drawing.Color.GhostWhite
Me.lblPicPath.Location = New System.Drawing.Point(581, 219)
Me.lblPicPath.Name = "lblPicPath"
Me.lblPicPath.Size = New System.Drawing.Size(300, 30)
Me.lblPicPath.TabIndex = 180
Me.lblPicPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblPicPath.Visible = False
'
'txtNeedDate
'
Me.txtNeedDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtNeedDate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtNeedDate.Location = New System.Drawing.Point(139, 502)
Me.txtNeedDate.MaxLength = 10
Me.txtNeedDate.Name = "txtNeedDate"
Me.txtNeedDate.Size = New System.Drawing.Size(149, 29)
Me.txtNeedDate.TabIndex = 13
Me.txtNeedDate.Text = "__/__/____"
Me.txtNeedDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'mskNeedDate
'
Me.mskNeedDate.BackColor = System.Drawing.Color.SlateBlue
Me.mskNeedDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.mskNeedDate.ForeColor = System.Drawing.Color.White
Me.mskNeedDate.Location = New System.Drawing.Point(139, 502)
Me.mskNeedDate.Mask = "99/99/9999"
Me.mskNeedDate.Name = "mskNeedDate"
Me.mskNeedDate.Size = New System.Drawing.Size(149, 29)
Me.mskNeedDate.TabIndex = 176
Me.mskNeedDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'txtBegin
'
Me.txtBegin.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtBegin.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtBegin.Location = New System.Drawing.Point(778, 389)
Me.txtBegin.MaxLength = 5
Me.txtBegin.Name = "txtBegin"
Me.txtBegin.Size = New System.Drawing.Size(110, 29)
Me.txtBegin.TabIndex = 72
Me.txtBegin.Text = "__/__/____"
Me.txtBegin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
Me.txtBegin.Visible = False
'
'mskBegin
'
Me.mskBegin.BackColor = System.Drawing.Color.SlateBlue
Me.mskBegin.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.mskBegin.ForeColor = System.Drawing.Color.White
Me.mskBegin.Location = New System.Drawing.Point(778, 389)
Me.mskBegin.Mask = "99/99/9999"
Me.mskBegin.Name = "mskBegin"
Me.mskBegin.Size = New System.Drawing.Size(110, 29)
Me.mskBegin.TabIndex = 73
Me.mskBegin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
Me.mskBegin.Visible = False
'
'txtCause
'
Me.txtCause.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtCause.ForeColor = System.Drawing.Color.Navy
Me.txtCause.Location = New System.Drawing.Point(96, 465)
Me.txtCause.MaxLength = 255
Me.txtCause.Name = "txtCause"
Me.txtCause.Size = New System.Drawing.Size(669, 26)
Me.txtCause.TabIndex = 12
'
'txtIssue
'
Me.txtIssue.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtIssue.ForeColor = System.Drawing.Color.Navy
Me.txtIssue.Location = New System.Drawing.Point(96, 375)
Me.txtIssue.MaxLength = 255
Me.txtIssue.Multiline = True
Me.txtIssue.Name = "txtIssue"
Me.txtIssue.Size = New System.Drawing.Size(669, 81)
Me.txtIssue.TabIndex = 11
'
'txtRemark
'
Me.txtRemark.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtRemark.ForeColor = System.Drawing.Color.Red
Me.txtRemark.Location = New System.Drawing.Point(96, 542)
Me.txtRemark.MaxLength = 150
Me.txtRemark.Name = "txtRemark"
Me.txtRemark.Size = New System.Drawing.Size(669, 26)
Me.txtRemark.TabIndex = 15
'
'txtEqpnm
'
Me.txtEqpnm.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtEqpnm.ForeColor = System.Drawing.Color.Navy
Me.txtEqpnm.Location = New System.Drawing.Point(97, 340)
Me.txtEqpnm.MaxLength = 150
Me.txtEqpnm.Name = "txtEqpnm"
Me.txtEqpnm.Size = New System.Drawing.Size(346, 26)
Me.txtEqpnm.TabIndex = 10
'
'txtSize
'
Me.txtSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtSize.ForeColor = System.Drawing.Color.Navy
Me.txtSize.Location = New System.Drawing.Point(97, 247)
Me.txtSize.MaxLength = 255
Me.txtSize.Multiline = True
Me.txtSize.Name = "txtSize"
Me.txtSize.Size = New System.Drawing.Size(472, 83)
Me.txtSize.TabIndex = 8
'
'txtShoe
'
Me.txtShoe.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtShoe.ForeColor = System.Drawing.Color.Navy
Me.txtShoe.Location = New System.Drawing.Point(97, 211)
Me.txtShoe.MaxLength = 10
Me.txtShoe.Name = "txtShoe"
Me.txtShoe.Size = New System.Drawing.Size(145, 26)
Me.txtShoe.TabIndex = 6
'
'txtOrder
'
Me.txtOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtOrder.ForeColor = System.Drawing.Color.Navy
Me.txtOrder.Location = New System.Drawing.Point(97, 177)
Me.txtOrder.MaxLength = 10
Me.txtOrder.Name = "txtOrder"
Me.txtOrder.Size = New System.Drawing.Size(145, 26)
Me.txtOrder.TabIndex = 5
'
'txtFrom
'
Me.txtFrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtFrom.ForeColor = System.Drawing.Color.Navy
Me.txtFrom.Location = New System.Drawing.Point(97, 108)
Me.txtFrom.Name = "txtFrom"
Me.txtFrom.Size = New System.Drawing.Size(209, 26)
Me.txtFrom.TabIndex = 2
'
'cboGroup
'
Me.cboGroup.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.cboGroup.ForeColor = System.Drawing.Color.Red
Me.cboGroup.FormattingEnabled = True
Me.cboGroup.Location = New System.Drawing.Point(318, 211)
Me.cboGroup.Name = "cboGroup"
Me.cboGroup.Size = New System.Drawing.Size(210, 26)
Me.cboGroup.TabIndex = 7
'
'cboDepfrom
'
Me.cboDepfrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.cboDepfrom.ForeColor = System.Drawing.Color.Red
Me.cboDepfrom.FormattingEnabled = True
Me.cboDepfrom.Location = New System.Drawing.Point(97, 142)
Me.cboDepfrom.Name = "cboDepfrom"
Me.cboDepfrom.Size = New System.Drawing.Size(319, 26)
Me.cboDepfrom.TabIndex = 4
'
'Label14
'
Me.Label14.AutoSize = True
Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label14.ForeColor = System.Drawing.Color.Black
Me.Label14.Location = New System.Drawing.Point(44, 247)
Me.Label14.Name = "Label14"
Me.Label14.Size = New System.Drawing.Size(46, 16)
Me.Label14.TabIndex = 53
Me.Label14.Text = "SIZE  :"
Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label6
'
Me.Label6.AutoSize = True
Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label6.ForeColor = System.Drawing.Color.Black
Me.Label6.Location = New System.Drawing.Point(39, 340)
Me.Label6.Name = "Label6"
Me.Label6.Size = New System.Drawing.Size(51, 16)
Me.Label6.TabIndex = 53
Me.Label6.Text = "รายการ  :"
Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label10
'
Me.Label10.AutoSize = True
Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label10.ForeColor = System.Drawing.Color.Black
Me.Label10.Location = New System.Drawing.Point(307, 507)
Me.Label10.Name = "Label10"
Me.Label10.Size = New System.Drawing.Size(34, 16)
Me.Label10.TabIndex = 53
Me.Label10.Text = "เวลา :"
Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label8
'
Me.Label8.AutoSize = True
Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label8.ForeColor = System.Drawing.Color.Black
Me.Label8.Location = New System.Drawing.Point(49, 468)
Me.Label8.Name = "Label8"
Me.Label8.Size = New System.Drawing.Size(44, 16)
Me.Label8.TabIndex = 53
Me.Label8.Text = "สาเหตุ :"
Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label17
'
Me.Label17.AutoSize = True
Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label17.ForeColor = System.Drawing.Color.Red
Me.Label17.Location = New System.Drawing.Point(690, 17)
Me.Label17.Name = "Label17"
Me.Label17.Size = New System.Drawing.Size(75, 18)
Me.Label17.TabIndex = 53
Me.Label17.Text = "รูปอุปกรณ์"
Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label9
'
Me.Label9.AutoSize = True
Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label9.ForeColor = System.Drawing.Color.Black
Me.Label9.Location = New System.Drawing.Point(13, 507)
Me.Label9.Name = "Label9"
Me.Label9.Size = New System.Drawing.Size(118, 16)
Me.Label9.TabIndex = 53
Me.Label9.Text = "ต้องการให้เสร็จภายใน  :"
Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label7
'
Me.Label7.AutoSize = True
Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label7.ForeColor = System.Drawing.Color.Black
Me.Label7.Location = New System.Drawing.Point(27, 375)
Me.Label7.Name = "Label7"
Me.Label7.Size = New System.Drawing.Size(66, 16)
Me.Label7.TabIndex = 53
Me.Label7.Text = "ปัญหาที่พบ :"
Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label12
'
Me.Label12.AutoSize = True
Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label12.ForeColor = System.Drawing.Color.Red
Me.Label12.Location = New System.Drawing.Point(28, 544)
Me.Label12.Name = "Label12"
Me.Label12.Size = New System.Drawing.Size(61, 16)
Me.Label12.TabIndex = 53
Me.Label12.Text = "หมายเหตุ  :"
Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label5
'
Me.Label5.AutoSize = True
Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label5.ForeColor = System.Drawing.Color.Black
Me.Label5.Location = New System.Drawing.Point(452, 344)
Me.Label5.Name = "Label5"
Me.Label5.Size = New System.Drawing.Size(48, 16)
Me.Label5.TabIndex = 53
Me.Label5.Text = "จำนวน  :"
Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label18
'
Me.Label18.AutoSize = True
Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label18.ForeColor = System.Drawing.Color.Black
Me.Label18.Location = New System.Drawing.Point(248, 216)
Me.Label18.Name = "Label18"
Me.Label18.Size = New System.Drawing.Size(67, 16)
Me.Label18.TabIndex = 53
Me.Label18.Text = "กลุ่มอุปกรณ์ :"
Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'Label4
'
Me.Label4.AutoSize = True
Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label4.ForeColor = System.Drawing.Color.Black
Me.Label4.Location = New System.Drawing.Point(64, 216)
Me.Label4.Name = "Label4"
Me.Label4.Size = New System.Drawing.Size(27, 16)
Me.Label4.TabIndex = 53
Me.Label4.Text = "รุ่น :"
Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'Label1
'
Me.Label1.AutoSize = True
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label1.ForeColor = System.Drawing.Color.Black
Me.Label1.Location = New System.Drawing.Point(13, 110)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(77, 16)
Me.Label1.TabIndex = 53
Me.Label1.Text = "จาก(ส่วนงาน) :"
Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label3
'
Me.Label3.AutoSize = True
Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label3.ForeColor = System.Drawing.Color.Black
Me.Label3.Location = New System.Drawing.Point(28, 181)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(63, 16)
Me.Label3.TabIndex = 53
Me.Label3.Text = "ORDER :"
Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label2
'
Me.Label2.AutoSize = True
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label2.ForeColor = System.Drawing.Color.Black
Me.Label2.Location = New System.Drawing.Point(22, 144)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(68, 16)
Me.Label2.TabIndex = 53
Me.Label2.Text = "แผนก / ฝ่าย :"
Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label11
'
Me.Label11.AutoSize = True
Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label11.ForeColor = System.Drawing.Color.Black
Me.Label11.Location = New System.Drawing.Point(23, 42)
Me.Label11.Name = "Label11"
Me.Label11.Size = New System.Drawing.Size(68, 16)
Me.Label11.TabIndex = 53
Me.Label11.Text = "รหัสเอกสาร :"
Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'gpbReceive
'
Me.gpbReceive.BackColor = System.Drawing.SystemColors.Control
Me.gpbReceive.Controls.Add(Me.txtWantDate)
Me.gpbReceive.Controls.Add(Me.mskWantDate)
Me.gpbReceive.Controls.Add(Me.txtFxIssue)
Me.gpbReceive.Controls.Add(Me.Label13)
Me.gpbReceive.Controls.Add(Me.Label15)
Me.gpbReceive.Controls.Add(Me.Label16)
Me.gpbReceive.Controls.Add(Me.txtWanttime)
Me.gpbReceive.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.gpbReceive.ForeColor = System.Drawing.Color.Blue
Me.gpbReceive.Location = New System.Drawing.Point(12, 604)
Me.gpbReceive.Name = "gpbReceive"
Me.gpbReceive.Size = New System.Drawing.Size(898, 157)
Me.gpbReceive.TabIndex = 54
Me.gpbReceive.TabStop = False
Me.gpbReceive.Text = "ส่วนรับผู้แจ้ง"
'
'txtWantDate
'
Me.txtWantDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtWantDate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtWantDate.Location = New System.Drawing.Point(151, 111)
Me.txtWantDate.MaxLength = 10
Me.txtWantDate.Name = "txtWantDate"
Me.txtWantDate.Size = New System.Drawing.Size(149, 29)
Me.txtWantDate.TabIndex = 17
Me.txtWantDate.Text = "__/__/____"
Me.txtWantDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'mskWantDate
'
Me.mskWantDate.BackColor = System.Drawing.Color.SlateBlue
Me.mskWantDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.mskWantDate.ForeColor = System.Drawing.Color.White
Me.mskWantDate.Location = New System.Drawing.Point(151, 111)
Me.mskWantDate.Mask = "99/99/9999"
Me.mskWantDate.Name = "mskWantDate"
Me.mskWantDate.Size = New System.Drawing.Size(149, 29)
Me.mskWantDate.TabIndex = 176
Me.mskWantDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'txtFxIssue
'
Me.txtFxIssue.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtFxIssue.ForeColor = System.Drawing.Color.Navy
Me.txtFxIssue.Location = New System.Drawing.Point(96, 25)
Me.txtFxIssue.MaxLength = 255
Me.txtFxIssue.Multiline = True
Me.txtFxIssue.Name = "txtFxIssue"
Me.txtFxIssue.Size = New System.Drawing.Size(669, 75)
Me.txtFxIssue.TabIndex = 16
'
'Label13
'
Me.Label13.AutoSize = True
Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label13.ForeColor = System.Drawing.Color.Black
Me.Label13.Location = New System.Drawing.Point(324, 116)
Me.Label13.Name = "Label13"
Me.Label13.Size = New System.Drawing.Size(34, 16)
Me.Label13.TabIndex = 53
Me.Label13.Text = "เวลา :"
Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label15
'
Me.Label15.AutoSize = True
Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label15.ForeColor = System.Drawing.Color.Black
Me.Label15.Location = New System.Drawing.Point(27, 116)
Me.Label15.Name = "Label15"
Me.Label15.Size = New System.Drawing.Size(115, 16)
Me.Label15.TabIndex = 53
Me.Label15.Text = "กำหนดให้เสร็จภายใน  :"
Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'Label16
'
Me.Label16.AutoSize = True
Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label16.ForeColor = System.Drawing.Color.Black
Me.Label16.Location = New System.Drawing.Point(28, 28)
Me.Label16.Name = "Label16"
Me.Label16.Size = New System.Drawing.Size(55, 16)
Me.Label16.TabIndex = 53
Me.Label16.Text = "การแก้ไข :"
Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopRight
'
'txtWanttime
'
Me.txtWanttime.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtWanttime.ForeColor = System.Drawing.Color.Navy
Me.txtWanttime.Location = New System.Drawing.Point(365, 111)
Me.txtWanttime.MaxLength = 5
Me.txtWanttime.Name = "txtWanttime"
Me.txtWanttime.Size = New System.Drawing.Size(109, 29)
Me.txtWanttime.TabIndex = 14
'
'txtRef
'
Me.txtRef.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtRef.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtRef.Location = New System.Drawing.Point(939, 166)
Me.txtRef.MaxLength = 12
Me.txtRef.Name = "txtRef"
Me.txtRef.Size = New System.Drawing.Size(116, 29)
Me.txtRef.TabIndex = 167
Me.txtRef.Visible = False
'
'lblComplete
'
Me.lblComplete.BackColor = System.Drawing.Color.Gold
Me.lblComplete.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblComplete.ForeColor = System.Drawing.Color.GhostWhite
Me.lblComplete.Location = New System.Drawing.Point(938, 133)
Me.lblComplete.Name = "lblComplete"
Me.lblComplete.Size = New System.Drawing.Size(117, 30)
Me.lblComplete.TabIndex = 164
Me.lblComplete.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.lblComplete.Visible = False
'
'btnExit
'
Me.btnExit.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
Me.btnExit.Location = New System.Drawing.Point(943, 81)
Me.btnExit.Name = "btnExit"
Me.btnExit.Size = New System.Drawing.Size(112, 49)
Me.btnExit.TabIndex = 17
Me.btnExit.Text = "ออกจากหน้าจอ"
Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.btnExit.UseVisualStyleBackColor = True
'
'btnSaveData
'
Me.btnSaveData.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnSaveData.Image = CType(resources.GetObject("btnSaveData.Image"), System.Drawing.Image)
Me.btnSaveData.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
Me.btnSaveData.Location = New System.Drawing.Point(943, 21)
Me.btnSaveData.Name = "btnSaveData"
Me.btnSaveData.Size = New System.Drawing.Size(112, 49)
Me.btnSaveData.TabIndex = 16
Me.btnSaveData.Text = "บันทึกข้อมูล"
Me.btnSaveData.TextAlign = System.Drawing.ContentAlignment.MiddleRight
Me.btnSaveData.UseVisualStyleBackColor = True
'
'frmAeNotifyIssue
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(1067, 773)
Me.Controls.Add(Me.txtRef)
Me.Controls.Add(Me.btnExit)
Me.Controls.Add(Me.btnSaveData)
Me.Controls.Add(Me.lblComplete)
Me.Controls.Add(Me.gpbReceive)
Me.Controls.Add(Me.gpbNotify)
Me.Name = "frmAeNotifyIssue"
Me.Text = "เพิ่ม / แก้ไข รายการแจ้งปัญหาอุปกรณ์"
Me.gpbNotify.ResumeLayout(False)
Me.gpbNotify.PerformLayout()
CType(Me.picEqp, System.ComponentModel.ISupportInitialize).EndInit()
Me.gpbReceive.ResumeLayout(False)
Me.gpbReceive.PerformLayout()
Me.ResumeLayout(False)
Me.PerformLayout()

End Sub
    Friend WithEvents lblDocTopic As System.Windows.Forms.Label
    Friend WithEvents gpbNotify As System.Windows.Forms.GroupBox
    Friend WithEvents txtFrom As System.Windows.Forms.TextBox
    Friend WithEvents cboDepfrom As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtEqpnm As System.Windows.Forms.TextBox
    Friend WithEvents txtSize As System.Windows.Forms.TextBox
    Friend WithEvents txtShoe As System.Windows.Forms.TextBox
    Friend WithEvents txtOrder As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtCause As System.Windows.Forms.TextBox
    Friend WithEvents txtIssue As System.Windows.Forms.TextBox
    Friend WithEvents txtNeedtime As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents gpbReceive As System.Windows.Forms.GroupBox
    Friend WithEvents txtFxIssue As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtRef As System.Windows.Forms.TextBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnSaveData As System.Windows.Forms.Button
    Friend WithEvents lblComplete As System.Windows.Forms.Label
    Friend WithEvents txtBegin As System.Windows.Forms.TextBox
    Friend WithEvents mskBegin As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtNeedDate As System.Windows.Forms.TextBox
    Friend WithEvents mskNeedDate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtWantDate As System.Windows.Forms.TextBox
    Friend WithEvents mskWantDate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents btnEditEqp1 As System.Windows.Forms.Button
    Friend WithEvents btnDelEqp1 As System.Windows.Forms.Button
    Friend WithEvents lblPicName As System.Windows.Forms.Label
    Friend WithEvents lblPicPath As System.Windows.Forms.Label
    Friend WithEvents cboGroup As System.Windows.Forms.ComboBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtSizeQty As System.Windows.Forms.TextBox
    Friend WithEvents mskSizeQty As System.Windows.Forms.MaskedTextBox
    Friend WithEvents picEqp As System.Windows.Forms.PictureBox
    Friend WithEvents txtDocid As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtWanttime As System.Windows.Forms.TextBox
    Friend WithEvents cboDepto As System.Windows.Forms.ComboBox
End Class

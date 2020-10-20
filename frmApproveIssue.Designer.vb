<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmApproveIssue
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
Me.components = New System.ComponentModel.Container
Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmApproveIssue))
Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
Me.Panel2 = New System.Windows.Forms.Panel
Me.lblDept = New System.Windows.Forms.Label
Me.lblName = New System.Windows.Forms.Label
Me.Label2 = New System.Windows.Forms.Label
Me.Label1 = New System.Windows.Forms.Label
Me.lblDocMenu = New System.Windows.Forms.Label
Me.tabCmd = New System.Windows.Forms.TabControl
Me.tabAdd = New System.Windows.Forms.TabPage
Me.tabFilter = New System.Windows.Forms.TabPage
Me.tabSearch = New System.Windows.Forms.TabPage
Me.tabRefesh = New System.Windows.Forms.TabPage
Me.tabExit = New System.Windows.Forms.TabPage
Me.imgListTab1 = New System.Windows.Forms.ImageList(Me.components)
Me.tlsBarFmr = New System.Windows.Forms.ToolStrip
Me.btnLast = New System.Windows.Forms.ToolStripButton
Me.btnNext = New System.Windows.Forms.ToolStripButton
Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
Me.lblPage = New System.Windows.Forms.ToolStripLabel
Me.lblPageAll = New System.Windows.Forms.ToolStripLabel
Me.txtPage = New System.Windows.Forms.ToolStripTextBox
Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel
Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
Me.btnPre = New System.Windows.Forms.ToolStripButton
Me.btnFirst = New System.Windows.Forms.ToolStripButton
Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator
Me.lblCmd = New System.Windows.Forms.ToolStripLabel
Me.lblHeight = New System.Windows.Forms.ToolStripLabel
Me.lblWidth = New System.Windows.Forms.ToolStripLabel
Me.lblLeft = New System.Windows.Forms.ToolStripLabel
Me.lblTop = New System.Windows.Forms.ToolStripLabel
Me.dgvIssue = New System.Windows.Forms.DataGridView
Me.Column4 = New System.Windows.Forms.DataGridViewImageColumn
Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn
Me.gpbSearch = New System.Windows.Forms.GroupBox
Me.cmbType = New System.Windows.Forms.ComboBox
Me.btnCancel = New System.Windows.Forms.Button
Me.btnOk = New System.Windows.Forms.Button
Me.txtSeek = New System.Windows.Forms.TextBox
Me.gpbFilter = New System.Windows.Forms.GroupBox
Me.cmbFilter = New System.Windows.Forms.ComboBox
Me.btnFilterCancel = New System.Windows.Forms.Button
Me.btnFilter = New System.Windows.Forms.Button
Me.txtFilter = New System.Windows.Forms.TextBox
Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
Me.lblDocnull = New System.Windows.Forms.Label
Me.Panel2.SuspendLayout()
Me.tabCmd.SuspendLayout()
Me.tlsBarFmr.SuspendLayout()
CType(Me.dgvIssue, System.ComponentModel.ISupportInitialize).BeginInit()
Me.gpbSearch.SuspendLayout()
Me.gpbFilter.SuspendLayout()
Me.SuspendLayout()
'
'Panel2
'
Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
Me.Panel2.Controls.Add(Me.lblDept)
Me.Panel2.Controls.Add(Me.lblName)
Me.Panel2.Controls.Add(Me.Label2)
Me.Panel2.Controls.Add(Me.Label1)
Me.Panel2.Controls.Add(Me.lblDocMenu)
Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
Me.Panel2.Location = New System.Drawing.Point(0, 0)
Me.Panel2.Name = "Panel2"
Me.Panel2.Size = New System.Drawing.Size(892, 100)
Me.Panel2.TabIndex = 36
'
'lblDept
'
Me.lblDept.AutoSize = True
Me.lblDept.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblDept.ForeColor = System.Drawing.Color.Lime
Me.lblDept.Location = New System.Drawing.Point(84, 72)
Me.lblDept.Name = "lblDept"
Me.lblDept.Size = New System.Drawing.Size(41, 20)
Me.lblDept.TabIndex = 6
Me.lblDept.Text = "dept"
'
'lblName
'
Me.lblName.AutoSize = True
Me.lblName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblName.ForeColor = System.Drawing.Color.Lime
Me.lblName.Location = New System.Drawing.Point(85, 49)
Me.lblName.Name = "lblName"
Me.lblName.Size = New System.Drawing.Size(49, 20)
Me.lblName.TabIndex = 6
Me.lblName.Text = "name"
'
'Label2
'
Me.Label2.AutoSize = True
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label2.ForeColor = System.Drawing.Color.Lime
Me.Label2.Location = New System.Drawing.Point(31, 73)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(57, 20)
Me.Label2.TabIndex = 5
Me.Label2.Text = "แผนก : "
'
'Label1
'
Me.Label1.AutoSize = True
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label1.ForeColor = System.Drawing.Color.Lime
Me.Label1.Location = New System.Drawing.Point(12, 49)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(76, 20)
Me.Label1.TabIndex = 5
Me.Label1.Text = "ชื่อ - สกุล : "
'
'lblDocMenu
'
Me.lblDocMenu.AutoSize = True
Me.lblDocMenu.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblDocMenu.ForeColor = System.Drawing.Color.Blue
Me.lblDocMenu.Location = New System.Drawing.Point(233, 9)
Me.lblDocMenu.Name = "lblDocMenu"
Me.lblDocMenu.Size = New System.Drawing.Size(401, 29)
Me.lblDocMenu.TabIndex = 4
Me.lblDocMenu.Text = "แฟ้มข้อมูลอนุมัติเอกสารแจ้งปัญหาอุปกรณ์"
'
'tabCmd
'
Me.tabCmd.Alignment = System.Windows.Forms.TabAlignment.Bottom
Me.tabCmd.Controls.Add(Me.tabAdd)
Me.tabCmd.Controls.Add(Me.tabFilter)
Me.tabCmd.Controls.Add(Me.tabSearch)
Me.tabCmd.Controls.Add(Me.tabRefesh)
Me.tabCmd.Controls.Add(Me.tabExit)
Me.tabCmd.Cursor = System.Windows.Forms.Cursors.Hand
Me.tabCmd.Dock = System.Windows.Forms.DockStyle.Top
Me.tabCmd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.tabCmd.HotTrack = True
Me.tabCmd.ImageList = Me.imgListTab1
Me.tabCmd.Location = New System.Drawing.Point(0, 100)
Me.tabCmd.Name = "tabCmd"
Me.tabCmd.SelectedIndex = 0
Me.tabCmd.Size = New System.Drawing.Size(892, 27)
Me.tabCmd.TabIndex = 37
'
'tabAdd
'
Me.tabAdd.ImageIndex = 0
Me.tabAdd.Location = New System.Drawing.Point(4, 4)
Me.tabAdd.Name = "tabAdd"
Me.tabAdd.Padding = New System.Windows.Forms.Padding(3)
Me.tabAdd.Size = New System.Drawing.Size(884, 0)
Me.tabAdd.TabIndex = 0
Me.tabAdd.Text = "อนุมัติ / พิมพ์ "
Me.tabAdd.UseVisualStyleBackColor = True
'
'tabFilter
'
Me.tabFilter.ImageIndex = 7
Me.tabFilter.Location = New System.Drawing.Point(4, 4)
Me.tabFilter.Name = "tabFilter"
Me.tabFilter.Size = New System.Drawing.Size(884, 0)
Me.tabFilter.TabIndex = 6
Me.tabFilter.Text = "กรอง"
Me.tabFilter.UseVisualStyleBackColor = True
'
'tabSearch
'
Me.tabSearch.ImageIndex = 3
Me.tabSearch.Location = New System.Drawing.Point(4, 4)
Me.tabSearch.Name = "tabSearch"
Me.tabSearch.Size = New System.Drawing.Size(884, 0)
Me.tabSearch.TabIndex = 3
Me.tabSearch.Text = "ค้นหา"
Me.tabSearch.UseVisualStyleBackColor = True
'
'tabRefesh
'
Me.tabRefesh.ImageIndex = 12
Me.tabRefesh.Location = New System.Drawing.Point(4, 4)
Me.tabRefesh.Name = "tabRefesh"
Me.tabRefesh.Size = New System.Drawing.Size(884, 0)
Me.tabRefesh.TabIndex = 9
Me.tabRefesh.Text = "ฟื้นฟูข้อมูล"
Me.tabRefesh.UseVisualStyleBackColor = True
'
'tabExit
'
Me.tabExit.ImageIndex = 2
Me.tabExit.Location = New System.Drawing.Point(4, 4)
Me.tabExit.Name = "tabExit"
Me.tabExit.Size = New System.Drawing.Size(884, 0)
Me.tabExit.TabIndex = 5
Me.tabExit.Text = "ออก"
Me.tabExit.UseVisualStyleBackColor = True
'
'imgListTab1
'
Me.imgListTab1.ImageStream = CType(resources.GetObject("imgListTab1.ImageStream"), System.Windows.Forms.ImageListStreamer)
Me.imgListTab1.TransparentColor = System.Drawing.Color.Transparent
Me.imgListTab1.Images.SetKeyName(0, "printer.png")
Me.imgListTab1.Images.SetKeyName(1, "cross.png")
Me.imgListTab1.Images.SetKeyName(2, "door_in.png")
Me.imgListTab1.Images.SetKeyName(3, "find.png")
Me.imgListTab1.Images.SetKeyName(4, "page_copy.png")
Me.imgListTab1.Images.SetKeyName(5, "wrench_orange.png")
Me.imgListTab1.Images.SetKeyName(6, "folder.png")
Me.imgListTab1.Images.SetKeyName(7, "lightning.png")
Me.imgListTab1.Images.SetKeyName(8, "calculator.png")
Me.imgListTab1.Images.SetKeyName(9, "error.png")
Me.imgListTab1.Images.SetKeyName(10, "signature16x16.png")
Me.imgListTab1.Images.SetKeyName(11, "page_add.png")
Me.imgListTab1.Images.SetKeyName(12, "reload.png")
Me.imgListTab1.Images.SetKeyName(13, "16x16_ledgreen.png")
'
'tlsBarFmr
'
Me.tlsBarFmr.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
Me.tlsBarFmr.Dock = System.Windows.Forms.DockStyle.Bottom
Me.tlsBarFmr.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnLast, Me.btnNext, Me.ToolStripSeparator1, Me.lblPage, Me.lblPageAll, Me.txtPage, Me.ToolStripLabel1, Me.ToolStripSeparator2, Me.btnPre, Me.btnFirst, Me.ToolStripSeparator3, Me.lblCmd, Me.lblHeight, Me.lblWidth, Me.lblLeft, Me.lblTop})
Me.tlsBarFmr.Location = New System.Drawing.Point(0, 534)
Me.tlsBarFmr.Name = "tlsBarFmr"
Me.tlsBarFmr.RightToLeft = System.Windows.Forms.RightToLeft.Yes
Me.tlsBarFmr.Size = New System.Drawing.Size(892, 25)
Me.tlsBarFmr.Stretch = True
Me.tlsBarFmr.TabIndex = 38
Me.tlsBarFmr.TabStop = True
Me.tlsBarFmr.Text = "Nevigator"
'
'btnLast
'
Me.btnLast.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
Me.btnLast.Image = Global.ProEquipMnt.My.Resources.Resources.resultset_last
Me.btnLast.ImageTransparentColor = System.Drawing.Color.Magenta
Me.btnLast.Name = "btnLast"
Me.btnLast.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.btnLast.Size = New System.Drawing.Size(23, 22)
Me.btnLast.Text = "ไปที่หน้าสุดท้าย"
'
'btnNext
'
Me.btnNext.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
Me.btnNext.Image = Global.ProEquipMnt.My.Resources.Resources.resultset_next
Me.btnNext.ImageTransparentColor = System.Drawing.Color.Magenta
Me.btnNext.Name = "btnNext"
Me.btnNext.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.btnNext.Size = New System.Drawing.Size(23, 22)
Me.btnNext.Text = "หน้าถัดไป"
'
'ToolStripSeparator1
'
Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
'
'lblPage
'
Me.lblPage.Name = "lblPage"
Me.lblPage.Size = New System.Drawing.Size(13, 22)
Me.lblPage.Text = "1"
Me.lblPage.Visible = False
'
'lblPageAll
'
Me.lblPageAll.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.lblPageAll.Name = "lblPageAll"
Me.lblPageAll.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.lblPageAll.Size = New System.Drawing.Size(18, 22)
Me.lblPageAll.Text = "/1"
'
'txtPage
'
Me.txtPage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtPage.ForeColor = System.Drawing.Color.Red
Me.txtPage.Name = "txtPage"
Me.txtPage.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.txtPage.Size = New System.Drawing.Size(70, 25)
Me.txtPage.Text = "1"
Me.txtPage.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'ToolStripLabel1
'
Me.ToolStripLabel1.Name = "ToolStripLabel1"
Me.ToolStripLabel1.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.ToolStripLabel1.Size = New System.Drawing.Size(39, 22)
Me.ToolStripLabel1.Text = "หน้าที่ :"
'
'ToolStripSeparator2
'
Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
'
'btnPre
'
Me.btnPre.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
Me.btnPre.Image = Global.ProEquipMnt.My.Resources.Resources.resultset_previous
Me.btnPre.ImageTransparentColor = System.Drawing.Color.Magenta
Me.btnPre.Name = "btnPre"
Me.btnPre.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.btnPre.Size = New System.Drawing.Size(23, 22)
Me.btnPre.Text = "ก่อนหน้านี้"
'
'btnFirst
'
Me.btnFirst.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
Me.btnFirst.Image = Global.ProEquipMnt.My.Resources.Resources.resultset_first
Me.btnFirst.ImageTransparentColor = System.Drawing.Color.Magenta
Me.btnFirst.Name = "btnFirst"
Me.btnFirst.RightToLeft = System.Windows.Forms.RightToLeft.No
Me.btnFirst.Size = New System.Drawing.Size(23, 22)
Me.btnFirst.Text = "หน้าแรกสุด"
'
'ToolStripSeparator3
'
Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
'
'lblCmd
'
Me.lblCmd.Name = "lblCmd"
Me.lblCmd.Size = New System.Drawing.Size(13, 22)
Me.lblCmd.Text = "0"
Me.lblCmd.Visible = False
'
'lblHeight
'
Me.lblHeight.Name = "lblHeight"
Me.lblHeight.Size = New System.Drawing.Size(13, 22)
Me.lblHeight.Text = "0"
Me.lblHeight.Visible = False
'
'lblWidth
'
Me.lblWidth.Name = "lblWidth"
Me.lblWidth.Size = New System.Drawing.Size(13, 22)
Me.lblWidth.Text = "0"
Me.lblWidth.Visible = False
'
'lblLeft
'
Me.lblLeft.Name = "lblLeft"
Me.lblLeft.Size = New System.Drawing.Size(13, 22)
Me.lblLeft.Text = "0"
Me.lblLeft.Visible = False
'
'lblTop
'
Me.lblTop.Name = "lblTop"
Me.lblTop.Size = New System.Drawing.Size(13, 22)
Me.lblTop.Text = "0"
Me.lblTop.Visible = False
'
'dgvIssue
'
Me.dgvIssue.AllowUserToAddRows = False
Me.dgvIssue.AllowUserToDeleteRows = False
Me.dgvIssue.AllowUserToResizeRows = False
DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
Me.dgvIssue.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
Me.dgvIssue.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
Me.dgvIssue.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
Me.dgvIssue.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
DataGridViewCellStyle2.BackColor = System.Drawing.Color.Orange
DataGridViewCellStyle2.Font = New System.Drawing.Font("BrowalliaUPC", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
Me.dgvIssue.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
Me.dgvIssue.ColumnHeadersHeight = 55
Me.dgvIssue.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
Me.dgvIssue.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column4, Me.DataGridViewTextBoxColumn1, Me.Column6, Me.Column5, Me.DataGridViewTextBoxColumn2, Me.Column15, Me.Column13, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.Column1, Me.Column2, Me.Column3})
DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
DataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Window
DataGridViewCellStyle13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
DataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.ControlText
DataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight
DataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText
DataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
Me.dgvIssue.DefaultCellStyle = DataGridViewCellStyle13
Me.dgvIssue.Dock = System.Windows.Forms.DockStyle.Fill
Me.dgvIssue.GridColor = System.Drawing.SystemColors.Control
Me.dgvIssue.Location = New System.Drawing.Point(0, 127)
Me.dgvIssue.MultiSelect = False
Me.dgvIssue.Name = "dgvIssue"
Me.dgvIssue.ReadOnly = True
DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
DataGridViewCellStyle14.BackColor = System.Drawing.Color.White
DataGridViewCellStyle14.Font = New System.Drawing.Font("BrowalliaUPC", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
DataGridViewCellStyle14.ForeColor = System.Drawing.SystemColors.MenuText
DataGridViewCellStyle14.SelectionBackColor = System.Drawing.SystemColors.Highlight
DataGridViewCellStyle14.SelectionForeColor = System.Drawing.SystemColors.HighlightText
DataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
Me.dgvIssue.RowHeadersDefaultCellStyle = DataGridViewCellStyle14
Me.dgvIssue.RowHeadersVisible = False
Me.dgvIssue.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
Me.dgvIssue.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
Me.dgvIssue.Size = New System.Drawing.Size(892, 407)
Me.dgvIssue.TabIndex = 43
'
'Column4
'
Me.Column4.Frozen = True
Me.Column4.HeaderText = ""
Me.Column4.Name = "Column4"
Me.Column4.ReadOnly = True
Me.Column4.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
Me.Column4.Width = 35
'
'DataGridViewTextBoxColumn1
'
DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
Me.DataGridViewTextBoxColumn1.DefaultCellStyle = DataGridViewCellStyle3
Me.DataGridViewTextBoxColumn1.Frozen = True
Me.DataGridViewTextBoxColumn1.HeaderText = "สถานะ"
Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
Me.DataGridViewTextBoxColumn1.ReadOnly = True
'
'Column6
'
DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
DataGridViewCellStyle4.ForeColor = System.Drawing.Color.Green
Me.Column6.DefaultCellStyle = DataGridViewCellStyle4
Me.Column6.Frozen = True
Me.Column6.HeaderText = "เลขที่เอกสาร"
Me.Column6.Name = "Column6"
Me.Column6.ReadOnly = True
Me.Column6.Width = 120
'
'Column5
'
DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
Me.Column5.DefaultCellStyle = DataGridViewCellStyle5
Me.Column5.HeaderText = "แผนกที่แจ้ง"
Me.Column5.Name = "Column5"
Me.Column5.ReadOnly = True
Me.Column5.Width = 160
'
'DataGridViewTextBoxColumn2
'
DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.DataGridViewTextBoxColumn2.DefaultCellStyle = DataGridViewCellStyle6
Me.DataGridViewTextBoxColumn2.HeaderText = "รายละเอียดอุปกรณ์"
Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
Me.DataGridViewTextBoxColumn2.ReadOnly = True
Me.DataGridViewTextBoxColumn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
Me.DataGridViewTextBoxColumn2.Width = 220
'
'Column15
'
DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
Me.Column15.DefaultCellStyle = DataGridViewCellStyle7
Me.Column15.HeaderText = "รุ่น / SIZE"
Me.Column15.Name = "Column15"
Me.Column15.ReadOnly = True
Me.Column15.Width = 120
'
'Column13
'
DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
Me.Column13.DefaultCellStyle = DataGridViewCellStyle8
Me.Column13.HeaderText = "จำนวน"
Me.Column13.Name = "Column13"
Me.Column13.ReadOnly = True
Me.Column13.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
Me.Column13.Width = 70
'
'DataGridViewTextBoxColumn3
'
DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
DataGridViewCellStyle9.ForeColor = System.Drawing.Color.Black
Me.DataGridViewTextBoxColumn3.DefaultCellStyle = DataGridViewCellStyle9
Me.DataGridViewTextBoxColumn3.HeaderText = "ปัญหาที่พบ /  สาเหตุ"
Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
Me.DataGridViewTextBoxColumn3.ReadOnly = True
Me.DataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
Me.DataGridViewTextBoxColumn3.Width = 220
'
'DataGridViewTextBoxColumn4
'
DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
Me.DataGridViewTextBoxColumn4.DefaultCellStyle = DataGridViewCellStyle10
Me.DataGridViewTextBoxColumn4.HeaderText = "การแก้ไข"
Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
Me.DataGridViewTextBoxColumn4.ReadOnly = True
Me.DataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
Me.DataGridViewTextBoxColumn4.Width = 180
'
'DataGridViewTextBoxColumn5
'
DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
DataGridViewCellStyle11.ForeColor = System.Drawing.Color.Black
Me.DataGridViewTextBoxColumn5.DefaultCellStyle = DataGridViewCellStyle11
Me.DataGridViewTextBoxColumn5.HeaderText = "วันที่บันทึก"
Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
Me.DataGridViewTextBoxColumn5.ReadOnly = True
Me.DataGridViewTextBoxColumn5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
Me.DataGridViewTextBoxColumn5.Width = 90
'
'DataGridViewTextBoxColumn6
'
DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
Me.DataGridViewTextBoxColumn6.DefaultCellStyle = DataGridViewCellStyle12
Me.DataGridViewTextBoxColumn6.HeaderText = "ผู้บันทึก "
Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
Me.DataGridViewTextBoxColumn6.ReadOnly = True
Me.DataGridViewTextBoxColumn6.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
'
'Column1
'
Me.Column1.HeaderText = "วันที่แก้ไข"
Me.Column1.Name = "Column1"
Me.Column1.ReadOnly = True
Me.Column1.Width = 90
'
'Column2
'
Me.Column2.HeaderText = "ผู้แก้ไข"
Me.Column2.Name = "Column2"
Me.Column2.ReadOnly = True
'
'Column3
'
Me.Column3.HeaderText = "หมายเหตุ"
Me.Column3.Name = "Column3"
Me.Column3.ReadOnly = True
Me.Column3.Width = 140
'
'gpbSearch
'
Me.gpbSearch.BackColor = System.Drawing.Color.Gray
Me.gpbSearch.Controls.Add(Me.cmbType)
Me.gpbSearch.Controls.Add(Me.btnCancel)
Me.gpbSearch.Controls.Add(Me.btnOk)
Me.gpbSearch.Controls.Add(Me.txtSeek)
Me.gpbSearch.ForeColor = System.Drawing.Color.White
Me.gpbSearch.Location = New System.Drawing.Point(35, 378)
Me.gpbSearch.Name = "gpbSearch"
Me.gpbSearch.Size = New System.Drawing.Size(348, 125)
Me.gpbSearch.TabIndex = 46
Me.gpbSearch.TabStop = False
Me.gpbSearch.Text = "ค้นหาข้อมูล"
Me.gpbSearch.Visible = False
'
'cmbType
'
Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
Me.cmbType.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.cmbType.ForeColor = System.Drawing.Color.Red
Me.cmbType.FormattingEnabled = True
Me.cmbType.Location = New System.Drawing.Point(12, 33)
Me.cmbType.Name = "cmbType"
Me.cmbType.Size = New System.Drawing.Size(154, 28)
Me.cmbType.TabIndex = 52
'
'btnCancel
'
Me.btnCancel.BackColor = System.Drawing.SystemColors.ButtonFace
Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnCancel.ForeColor = System.Drawing.Color.Black
Me.btnCancel.Location = New System.Drawing.Point(239, 83)
Me.btnCancel.Name = "btnCancel"
Me.btnCancel.Size = New System.Drawing.Size(76, 30)
Me.btnCancel.TabIndex = 51
Me.btnCancel.Text = "ยกเลิก"
Me.btnCancel.UseVisualStyleBackColor = False
'
'btnOk
'
Me.btnOk.BackColor = System.Drawing.SystemColors.ButtonFace
Me.btnOk.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnOk.ForeColor = System.Drawing.Color.Black
Me.btnOk.Location = New System.Drawing.Point(158, 83)
Me.btnOk.Name = "btnOk"
Me.btnOk.Size = New System.Drawing.Size(76, 30)
Me.btnOk.TabIndex = 50
Me.btnOk.Text = "ตกลง"
Me.btnOk.UseVisualStyleBackColor = False
'
'txtSeek
'
Me.txtSeek.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtSeek.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtSeek.Location = New System.Drawing.Point(168, 36)
Me.txtSeek.MaxLength = 25
Me.txtSeek.Name = "txtSeek"
Me.txtSeek.Size = New System.Drawing.Size(150, 26)
Me.txtSeek.TabIndex = 49
Me.txtSeek.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'gpbFilter
'
Me.gpbFilter.BackColor = System.Drawing.Color.Green
Me.gpbFilter.Controls.Add(Me.cmbFilter)
Me.gpbFilter.Controls.Add(Me.btnFilterCancel)
Me.gpbFilter.Controls.Add(Me.btnFilter)
Me.gpbFilter.Controls.Add(Me.txtFilter)
Me.gpbFilter.ForeColor = System.Drawing.Color.White
Me.gpbFilter.Location = New System.Drawing.Point(507, 378)
Me.gpbFilter.Name = "gpbFilter"
Me.gpbFilter.Size = New System.Drawing.Size(348, 125)
Me.gpbFilter.TabIndex = 47
Me.gpbFilter.TabStop = False
Me.gpbFilter.Text = "กรองข้อมูล"
Me.gpbFilter.Visible = False
'
'cmbFilter
'
Me.cmbFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
Me.cmbFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.cmbFilter.ForeColor = System.Drawing.Color.Red
Me.cmbFilter.FormattingEnabled = True
Me.cmbFilter.Location = New System.Drawing.Point(11, 35)
Me.cmbFilter.Name = "cmbFilter"
Me.cmbFilter.Size = New System.Drawing.Size(155, 28)
Me.cmbFilter.TabIndex = 52
'
'btnFilterCancel
'
Me.btnFilterCancel.BackColor = System.Drawing.SystemColors.ButtonFace
Me.btnFilterCancel.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnFilterCancel.ForeColor = System.Drawing.Color.Black
Me.btnFilterCancel.Location = New System.Drawing.Point(241, 83)
Me.btnFilterCancel.Name = "btnFilterCancel"
Me.btnFilterCancel.Size = New System.Drawing.Size(76, 30)
Me.btnFilterCancel.TabIndex = 51
Me.btnFilterCancel.Text = "ยกเลิก"
Me.btnFilterCancel.UseVisualStyleBackColor = False
'
'btnFilter
'
Me.btnFilter.BackColor = System.Drawing.SystemColors.ButtonFace
Me.btnFilter.Cursor = System.Windows.Forms.Cursors.Hand
Me.btnFilter.ForeColor = System.Drawing.SystemColors.ControlText
Me.btnFilter.Location = New System.Drawing.Point(160, 83)
Me.btnFilter.Name = "btnFilter"
Me.btnFilter.Size = New System.Drawing.Size(76, 30)
Me.btnFilter.TabIndex = 50
Me.btnFilter.Text = "ตกลง"
Me.btnFilter.UseVisualStyleBackColor = False
'
'txtFilter
'
Me.txtFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.txtFilter.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.txtFilter.Location = New System.Drawing.Point(166, 36)
Me.txtFilter.MaxLength = 25
Me.txtFilter.Name = "txtFilter"
Me.txtFilter.Size = New System.Drawing.Size(150, 26)
Me.txtFilter.TabIndex = 49
Me.txtFilter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
'
'Timer1
'
Me.Timer1.Enabled = True
Me.Timer1.Interval = 300000
'
'lblDocnull
'
Me.lblDocnull.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
Me.lblDocnull.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.lblDocnull.ForeColor = System.Drawing.Color.Chocolate
Me.lblDocnull.Location = New System.Drawing.Point(155, 249)
Me.lblDocnull.Name = "lblDocnull"
Me.lblDocnull.Size = New System.Drawing.Size(588, 90)
Me.lblDocnull.TabIndex = 76
Me.lblDocnull.Text = "ไม่มีเอกสารรออนุมัติ"
Me.lblDocnull.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblDocnull.Visible = False
'
'frmApproveIssue
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
Me.ClientSize = New System.Drawing.Size(892, 559)
Me.Controls.Add(Me.lblDocnull)
Me.Controls.Add(Me.gpbFilter)
Me.Controls.Add(Me.gpbSearch)
Me.Controls.Add(Me.dgvIssue)
Me.Controls.Add(Me.tlsBarFmr)
Me.Controls.Add(Me.tabCmd)
Me.Controls.Add(Me.Panel2)
Me.Name = "frmApproveIssue"
Me.Text = "frmApproveIssue"
Me.Panel2.ResumeLayout(False)
Me.Panel2.PerformLayout()
Me.tabCmd.ResumeLayout(False)
Me.tlsBarFmr.ResumeLayout(False)
Me.tlsBarFmr.PerformLayout()
CType(Me.dgvIssue, System.ComponentModel.ISupportInitialize).EndInit()
Me.gpbSearch.ResumeLayout(False)
Me.gpbSearch.PerformLayout()
Me.gpbFilter.ResumeLayout(False)
Me.gpbFilter.PerformLayout()
Me.ResumeLayout(False)
Me.PerformLayout()

End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lblDocMenu As System.Windows.Forms.Label
    Friend WithEvents tabCmd As System.Windows.Forms.TabControl
    Friend WithEvents tabAdd As System.Windows.Forms.TabPage
    Friend WithEvents tabFilter As System.Windows.Forms.TabPage
    Friend WithEvents tabSearch As System.Windows.Forms.TabPage
    Friend WithEvents tabRefesh As System.Windows.Forms.TabPage
    Friend WithEvents tabExit As System.Windows.Forms.TabPage
    Friend WithEvents tlsBarFmr As System.Windows.Forms.ToolStrip
    Friend WithEvents btnLast As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnNext As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents lblPage As System.Windows.Forms.ToolStripLabel
    Friend WithEvents lblPageAll As System.Windows.Forms.ToolStripLabel
    Friend WithEvents txtPage As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnPre As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnFirst As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents lblCmd As System.Windows.Forms.ToolStripLabel
    Friend WithEvents lblHeight As System.Windows.Forms.ToolStripLabel
    Friend WithEvents lblWidth As System.Windows.Forms.ToolStripLabel
    Friend WithEvents lblLeft As System.Windows.Forms.ToolStripLabel
    Friend WithEvents lblTop As System.Windows.Forms.ToolStripLabel
    Friend WithEvents dgvIssue As System.Windows.Forms.DataGridView
    Friend WithEvents lblDept As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gpbSearch As System.Windows.Forms.GroupBox
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents txtSeek As System.Windows.Forms.TextBox
    Friend WithEvents gpbFilter As System.Windows.Forms.GroupBox
    Friend WithEvents cmbFilter As System.Windows.Forms.ComboBox
    Friend WithEvents btnFilterCancel As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents txtFilter As System.Windows.Forms.TextBox
    Friend WithEvents imgListTab1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents lblDocnull As System.Windows.Forms.Label
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column15 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class

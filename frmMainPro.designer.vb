<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMainPro
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMainPro))
        Me.mnStpMain = New System.Windows.Forms.MenuStrip()
        Me.mnFileSys = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnDocFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnUsrFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnWipnewImport = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnReLogIn = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnNewLogIn = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnChangePass = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnRpt = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnRptdvl = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnAsswin = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnTileHor = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnVer = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnCasd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.lblIcon = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblLogin = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblSpace1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblUsrName = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblSpace2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblIp = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblSpace3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblCurrentDate = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblSpace4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblCurrentTime = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblRptCentral = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblRptDesc = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lstBarMain = New vbAccelerator.Components.ListBarControl.ListBar()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.mnStpMain.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnStpMain
        '
        Me.mnStpMain.BackColor = System.Drawing.SystemColors.Control
        Me.mnStpMain.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.mnStpMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnFileSys, Me.mnReLogIn, Me.mnRpt, Me.mnAsswin, Me.mnExit})
        Me.mnStpMain.Location = New System.Drawing.Point(0, 0)
        Me.mnStpMain.MdiWindowListItem = Me.mnAsswin
        Me.mnStpMain.Name = "mnStpMain"
        Me.mnStpMain.Size = New System.Drawing.Size(632, 40)
        Me.mnStpMain.TabIndex = 1
        Me.mnStpMain.Text = "Main Menu"
        '
        'mnFileSys
        '
        Me.mnFileSys.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnDocFile, Me.mnUsrFile, Me.mnWipnewImport})
        Me.mnFileSys.ForeColor = System.Drawing.Color.Black
        Me.mnFileSys.Image = Global.ProEquipMnt.My.Resources.Resources.Registration
        Me.mnFileSys.Name = "mnFileSys"
        Me.mnFileSys.Size = New System.Drawing.Size(149, 36)
        Me.mnFileSys.Text = "A&dmin File System"
        '
        'mnDocFile
        '
        Me.mnDocFile.Image = Global.ProEquipMnt.My.Resources.Resources.folder_30px
        Me.mnDocFile.Name = "mnDocFile"
        Me.mnDocFile.Size = New System.Drawing.Size(176, 22)
        Me.mnDocFile.Text = "แฟ้มระบบงาน..."
        '
        'mnUsrFile
        '
        Me.mnUsrFile.Image = Global.ProEquipMnt.My.Resources.Resources.users_30px
        Me.mnUsrFile.Name = "mnUsrFile"
        Me.mnUsrFile.Size = New System.Drawing.Size(176, 22)
        Me.mnUsrFile.Text = "แฟ้มข้อมูลผู้ใช้งาน..."
        '
        'mnWipnewImport
        '
        Me.mnWipnewImport.Image = Global.ProEquipMnt.My.Resources.Resources.Downloads_30px
        Me.mnWipnewImport.Name = "mnWipnewImport"
        Me.mnWipnewImport.Size = New System.Drawing.Size(176, 22)
        Me.mnWipnewImport.Text = "นำเข้าข้อมูล Wipnew"
        '
        'mnReLogIn
        '
        Me.mnReLogIn.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnNewLogIn, Me.mnChangePass})
        Me.mnReLogIn.ForeColor = System.Drawing.Color.Black
        Me.mnReLogIn.Image = Global.ProEquipMnt.My.Resources.Resources.Web_designer
        Me.mnReLogIn.Name = "mnReLogIn"
        Me.mnReLogIn.Size = New System.Drawing.Size(89, 36)
        Me.mnReLogIn.Text = "ผู้ใ&ช้งาน"
        '
        'mnNewLogIn
        '
        Me.mnNewLogIn.Name = "mnNewLogIn"
        Me.mnNewLogIn.Size = New System.Drawing.Size(194, 22)
        Me.mnNewLogIn.Text = "ล็อคอินเข้าใช้งานใหม่"
        '
        'mnChangePass
        '
        Me.mnChangePass.Name = "mnChangePass"
        Me.mnChangePass.Size = New System.Drawing.Size(194, 22)
        Me.mnChangePass.Text = "เปลี่ยนรหัสผ่านเข้าใช้งาน"
        '
        'mnRpt
        '
        Me.mnRpt.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnRptdvl})
        Me.mnRpt.ForeColor = System.Drawing.Color.Black
        Me.mnRpt.Image = Global.ProEquipMnt.My.Resources.Resources.Script
        Me.mnRpt.Name = "mnRpt"
        Me.mnRpt.Size = New System.Drawing.Size(87, 36)
        Me.mnRpt.Text = "รายง&าน"
        '
        'mnRptdvl
        '
        Me.mnRptdvl.Name = "mnRptdvl"
        Me.mnRptdvl.Size = New System.Drawing.Size(213, 22)
        Me.mnRptdvl.Text = "รายงานการโอนอุปกรณ์ลงผลิต"
        '
        'mnAsswin
        '
        Me.mnAsswin.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnTileHor, Me.mnVer, Me.mnCasd})
        Me.mnAsswin.ForeColor = System.Drawing.Color.Black
        Me.mnAsswin.Image = Global.ProEquipMnt.My.Resources.Resources.Flip
        Me.mnAsswin.Name = "mnAsswin"
        Me.mnAsswin.Size = New System.Drawing.Size(88, 36)
        Me.mnAsswin.Text = "หน้า&ต่าง"
        '
        'mnTileHor
        '
        Me.mnTileHor.Name = "mnTileHor"
        Me.mnTileHor.Size = New System.Drawing.Size(163, 22)
        Me.mnTileHor.Text = "เรียงแบบแนวนอน"
        '
        'mnVer
        '
        Me.mnVer.Name = "mnVer"
        Me.mnVer.Size = New System.Drawing.Size(163, 22)
        Me.mnVer.Text = "เรียงแบบแนวตั้ง"
        '
        'mnCasd
        '
        Me.mnCasd.Name = "mnCasd"
        Me.mnCasd.Size = New System.Drawing.Size(163, 22)
        Me.mnCasd.Text = "เรียงแบบลำดับชั้น"
        '
        'mnExit
        '
        Me.mnExit.ForeColor = System.Drawing.Color.Black
        Me.mnExit.Image = Global.ProEquipMnt.My.Resources.Resources.exit_winxp1
        Me.mnExit.Name = "mnExit"
        Me.mnExit.Size = New System.Drawing.Size(111, 36)
        Me.mnExit.Text = "&ปิดโปรแกรม"
        '
        'StatusStrip
        '
        Me.StatusStrip.BackColor = System.Drawing.Color.Silver
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblIcon, Me.lblLogin, Me.lblSpace1, Me.lblUsrName, Me.lblSpace2, Me.lblIp, Me.lblSpace3, Me.lblCurrentDate, Me.lblSpace4, Me.lblCurrentTime, Me.lblRptCentral, Me.lblRptDesc})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 431)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(632, 22)
        Me.StatusStrip.TabIndex = 3
        '
        'lblIcon
        '
        Me.lblIcon.Image = Global.ProEquipMnt.My.Resources.Resources.admin
        Me.lblIcon.Name = "lblIcon"
        Me.lblIcon.Size = New System.Drawing.Size(85, 17)
        Me.lblIcon.Text = "User LogIn :"
        '
        'lblLogin
        '
        Me.lblLogin.Name = "lblLogin"
        Me.lblLogin.Size = New System.Drawing.Size(21, 17)
        Me.lblLogin.Text = "SA"
        '
        'lblSpace1
        '
        Me.lblSpace1.Name = "lblSpace1"
        Me.lblSpace1.Size = New System.Drawing.Size(22, 17)
        Me.lblSpace1.Text = "     "
        '
        'lblUsrName
        '
        Me.lblUsrName.Image = Global.ProEquipMnt.My.Resources.Resources.users
        Me.lblUsrName.Name = "lblUsrName"
        Me.lblUsrName.Size = New System.Drawing.Size(68, 17)
        Me.lblUsrName.Text = "xxxxxxxxx"
        '
        'lblSpace2
        '
        Me.lblSpace2.Name = "lblSpace2"
        Me.lblSpace2.Size = New System.Drawing.Size(22, 17)
        Me.lblSpace2.Text = "     "
        '
        'lblIp
        '
        Me.lblIp.Image = Global.ProEquipMnt.My.Resources.Resources.laptop1
        Me.lblIp.Name = "lblIp"
        Me.lblIp.Size = New System.Drawing.Size(72, 17)
        Me.lblIp.Text = "xx.xx.xx.xx"
        '
        'lblSpace3
        '
        Me.lblSpace3.Name = "lblSpace3"
        Me.lblSpace3.Size = New System.Drawing.Size(22, 17)
        Me.lblSpace3.Text = "     "
        '
        'lblCurrentDate
        '
        Me.lblCurrentDate.Image = Global.ProEquipMnt.My.Resources.Resources._date
        Me.lblCurrentDate.Name = "lblCurrentDate"
        Me.lblCurrentDate.Size = New System.Drawing.Size(93, 17)
        Me.lblCurrentDate.Text = "dd/mm/yyyy"
        '
        'lblSpace4
        '
        Me.lblSpace4.Name = "lblSpace4"
        Me.lblSpace4.Size = New System.Drawing.Size(22, 17)
        Me.lblSpace4.Text = "     "
        '
        'lblCurrentTime
        '
        Me.lblCurrentTime.Image = Global.ProEquipMnt.My.Resources.Resources.time
        Me.lblCurrentTime.Name = "lblCurrentTime"
        Me.lblCurrentTime.Size = New System.Drawing.Size(75, 17)
        Me.lblCurrentTime.Text = "hh:mm:ss"
        '
        'lblRptCentral
        '
        Me.lblRptCentral.Name = "lblRptCentral"
        Me.lblRptCentral.Size = New System.Drawing.Size(0, 17)
        '
        'lblRptDesc
        '
        Me.lblRptDesc.Name = "lblRptDesc"
        Me.lblRptDesc.Size = New System.Drawing.Size(0, 17)
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lstBarMain)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 40)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(84, 391)
        Me.Panel1.TabIndex = 5
        '
        'lstBarMain
        '
        Me.lstBarMain.AllowDragGroups = True
        Me.lstBarMain.AllowDragItems = True
        Me.lstBarMain.BackColor = System.Drawing.Color.RoyalBlue
        Me.lstBarMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstBarMain.DrawStyle = vbAccelerator.Components.ListBarControl.ListBarDrawStyle.ListBarDrawStyleOfficeXP
        Me.lstBarMain.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lstBarMain.LargeImageList = Me.ImageList1
        Me.lstBarMain.Location = New System.Drawing.Point(0, 0)
        Me.lstBarMain.Name = "lstBarMain"
        Me.lstBarMain.SelectOnMouseDown = False
        Me.lstBarMain.Size = New System.Drawing.Size(84, 391)
        Me.lstBarMain.SmallImageList = Me.ImageList1
        Me.lstBarMain.TabIndex = 6
        Me.lstBarMain.ToolTip = Nothing
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "mold48.ico")
        Me.ImageList1.Images.SetKeyName(1, "sheet48.ico")
        Me.ImageList1.Images.SetKeyName(2, "scissors48.ico")
        Me.ImageList1.Images.SetKeyName(3, "screen48.ico")
        Me.ImageList1.Images.SetKeyName(4, "arc48.ico")
        Me.ImageList1.Images.SetKeyName(5, "Transfer_48x48.png")
        Me.ImageList1.Images.SetKeyName(6, "maintenance_48x48.png")
        Me.ImageList1.Images.SetKeyName(7, "inventory48x48.ico")
        Me.ImageList1.Images.SetKeyName(8, "user48x48.ico")
        Me.ImageList1.Images.SetKeyName(9, "cupboard.png")
        Me.ImageList1.Images.SetKeyName(10, "Request.ico")
        Me.ImageList1.Images.SetKeyName(11, "approve2_48x48.ico")
        Me.ImageList1.Images.SetKeyName(12, "approve_48x48.ico")
        Me.ImageList1.Images.SetKeyName(13, "garbage48.ico")
        Me.ImageList1.Images.SetKeyName(14, "death_list48.ico")
        Me.ImageList1.Images.SetKeyName(15, "office48x48.ico")
        Me.ImageList1.Images.SetKeyName(16, "tractor48x48.ico")
        Me.ImageList1.Images.SetKeyName(17, "sale48.ico")
        Me.ImageList1.Images.SetKeyName(18, "calender48.ico")
        Me.ImageList1.Images.SetKeyName(19, "Note48x48.ico")
        Me.ImageList1.Images.SetKeyName(20, "book48x48.ico")
        Me.ImageList1.Images.SetKeyName(21, "home48x48.ico")
        Me.ImageList1.Images.SetKeyName(22, "exchng48.ico")
        '
        'Timer1
        '
        '
        'frmMainPro
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(632, 453)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.mnStpMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.mnStpMain
        Me.Name = "frmMainPro"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "โปรแกรมบันทึกเครื่องมืออุปกรณ์การผลิต (Production Equipments)"
        Me.mnStpMain.ResumeLayout(False)
        Me.mnStpMain.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mnStpMain As System.Windows.Forms.MenuStrip
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents mnFileSys As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnReLogIn As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnRpt As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnAsswin As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblIcon As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblLogin As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblSpace1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblUsrName As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblSpace2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblIp As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblSpace3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblCurrentDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblSpace4 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblCurrentTime As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents mnTileHor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnVer As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnCasd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnDocFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnUsrFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnNewLogIn As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnChangePass As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents lblRptCentral As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblRptDesc As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents mnRptdvl As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstBarMain As vbAccelerator.Components.ListBarControl.ListBar
    Friend WithEvents mnWipnewImport As System.Windows.Forms.ToolStripMenuItem

End Class

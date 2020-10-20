<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWaiting
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
Me.components = New System.ComponentModel.Container
Me.Label1 = New System.Windows.Forms.Label
Me.Label2 = New System.Windows.Forms.Label
Me.Picbox1 = New System.Windows.Forms.PictureBox
Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
CType(Me.Picbox1, System.ComponentModel.ISupportInitialize).BeginInit()
Me.SuspendLayout()
'
'Label1
'
Me.Label1.AutoSize = True
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label1.ForeColor = System.Drawing.Color.Black
Me.Label1.Location = New System.Drawing.Point(78, 13)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(69, 18)
Me.Label1.TabIndex = 1
Me.Label1.Text = "โปรดรอ...."
'
'Label2
'
Me.Label2.AutoSize = True
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
Me.Label2.ForeColor = System.Drawing.Color.Black
Me.Label2.Location = New System.Drawing.Point(98, 38)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(201, 18)
Me.Label2.TabIndex = 1
Me.Label2.Text = "กำลังโหลดรายงาน อาจใช้เวลาสักครู่"
'
'Picbox1
'
Me.Picbox1.Location = New System.Drawing.Point(12, 12)
Me.Picbox1.Name = "Picbox1"
Me.Picbox1.Size = New System.Drawing.Size(60, 60)
Me.Picbox1.TabIndex = 0
Me.Picbox1.TabStop = False
'
'frmWaiting
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(332, 83)
Me.ControlBox = False
Me.Controls.Add(Me.Label2)
Me.Controls.Add(Me.Label1)
Me.Controls.Add(Me.Picbox1)
Me.MaximumSize = New System.Drawing.Size(348, 121)
Me.MinimumSize = New System.Drawing.Size(348, 121)
Me.Name = "frmWaiting"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "Waiting"
CType(Me.Picbox1, System.ComponentModel.ISupportInitialize).EndInit()
Me.ResumeLayout(False)
Me.PerformLayout()

End Sub
    Friend WithEvents Picbox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class

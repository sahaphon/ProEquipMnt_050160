<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrntProgress
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
Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
Me.SuspendLayout()
'
'ProgressBar1
'
Me.ProgressBar1.Location = New System.Drawing.Point(2, 2)
Me.ProgressBar1.Name = "ProgressBar1"
Me.ProgressBar1.Size = New System.Drawing.Size(288, 35)
Me.ProgressBar1.TabIndex = 0
'
'frmPrntProgress
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(293, 40)
Me.Controls.Add(Me.ProgressBar1)
Me.MaximumSize = New System.Drawing.Size(309, 78)
Me.MinimumSize = New System.Drawing.Size(309, 78)
Me.Name = "frmPrntProgress"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "ความคืบหน้า"
Me.ResumeLayout(False)

End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
End Class

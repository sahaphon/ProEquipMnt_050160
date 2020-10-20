Imports ADODB

Public Class frmWaiting

  Private _Countdown As Integer

    Public Property Countdown() As Integer
        Get
            Return _Countdown
        End Get
        Set(ByVal value As Integer)
            _Countdown = value
        End Set
    End Property


Private Sub frmWaiting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
   Timer1.Interval = 3000
   Timer1.Enabled = True
   Timer1.Start() 'Timer starts functioning
   Picbox1.Image = My.Resources.wait_60
End Sub

Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

   Countdown -= 1
   If Countdown <= 0 Then
       Timer1.Stop() 'Timer stops functioning
       Me.Close()
   End If

End Sub

End Class
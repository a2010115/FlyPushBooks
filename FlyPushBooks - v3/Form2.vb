Imports System.Runtime.InteropServices
Imports System.Runtime
Imports FPBV3.MetroUI_Form.WinApi
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form2
    Public lasttime = 0
    Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Location = New Point(Form1.Location.X + (Form1.Width - Me.Width) / 2, Form1.Location.Y + (Form1.Height - Me.Height) / 2)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Location = New Point(Form1.Location.X + (Form1.Width - Me.Width) / 2, Form1.Location.Y + (Form1.Height - Me.Height) / 2)
    End Sub
    Public lt
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If lasttime > 1 Then '如果在预测时间之内
            Label1.Text = "请稍候，正发送...大约需要" & CInt(lasttime) & "秒..."
            lasttime -= 1
            ProgressBar1.Style = ProgressBarStyle.Blocks
            ProgressBar1.Value = 100 - （lasttime / lt * 100）
        Else '否则
            Label1.Text = "请稍候，正在做最后的工作..."
            ProgressBar1.Style = ProgressBarStyle.Marquee
        End If
    End Sub
End Class
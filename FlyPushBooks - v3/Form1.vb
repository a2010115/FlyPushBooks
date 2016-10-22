Imports System.Runtime.InteropServices
Imports System.Runtime
Imports FPBV3.MetroUI_Form.WinApi
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
#Region "shadows"
    Private dwmMargins As MetroUI_Form.Dwm.MARGINS
    Private _marginOk As Boolean
    Private _aeroEnabled As Boolean = False

#Region "Props"
    Public ReadOnly Property AeroEnabled() As Boolean
        Get
            Return _aeroEnabled
        End Get
    End Property
#End Region

#Region "Methods"
    Public Shared Function LoWord(ByVal dwValue As Integer) As Integer
        Return dwValue And &HFFFF
    End Function

    Public Shared Function HiWord(ByVal dwValue As Integer) As Integer
        Return (dwValue >> 16) And &HFFFF
    End Function
#End Region

    Public Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        MetroUI_Form.Dwm.DwmExtendFrameIntoClientArea(Me.Handle, dwmMargins)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Dim WM_NCCALCSIZE As Integer = &H83
        Dim WM_NCHITTEST As Integer = &H84
        Dim result As IntPtr = IntPtr.Zero

        Dim dwmHandled As Integer = MetroUI_Form.Dwm.DwmDefWindowProc(m.HWnd, m.Msg, m.WParam, m.LParam, result)

        If dwmHandled = 1 Then
            m.Result = result
            Return
        End If

        If m.Msg = WM_NCCALCSIZE AndAlso CType(m.WParam, Int32) = 1 Then
            Dim nccsp As MetroUI_Form.WinApi.NCCALCSIZE_PARAMS = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(MetroUI_Form.WinApi.NCCALCSIZE_PARAMS)), MetroUI_Form.WinApi.NCCALCSIZE_PARAMS)
            nccsp.rect0.Top += 0
            nccsp.rect0.Bottom += 0
            nccsp.rect0.Left += 0
            nccsp.rect0.Right += 0

            If Not _marginOk Then
                dwmMargins.cyTopHeight = 0
                dwmMargins.cxLeftWidth = 0
                dwmMargins.cyBottomHeight = 3
                dwmMargins.cxRightWidth = 0
                _marginOk = True
            End If

            Marshal.StructureToPtr(nccsp, m.LParam, False)

            m.Result = IntPtr.Zero
        ElseIf m.Msg = WM_NCHITTEST AndAlso CType(m.Result, Int32) = 0 Then
            m.Result = HitTestNCA(m.HWnd, m.WParam, m.LParam)
        Else
            MyBase.WndProc(m)
        End If
    End Sub
    Private Function HitTestNCA(ByVal hwnd As IntPtr, ByVal wparam As IntPtr, ByVal lparam As IntPtr) As IntPtr
        Dim HTNOWHERE As Integer = 0
        Dim HTCLIENT As Integer = 1
        Dim HTCAPTION As Integer = 2
        Dim HTGROWBOX As Integer = 4
        Dim HTSIZE As Integer = HTGROWBOX
        Dim HTMINBUTTON As Integer = 8
        Dim HTMAXBUTTON As Integer = 9
        Dim HTLEFT As Integer = 10
        Dim HTRIGHT As Integer = 11
        Dim HTTOP As Integer = 12
        Dim HTTOPLEFT As Integer = 13
        Dim HTTOPRIGHT As Integer = 14
        Dim HTBOTTOM As Integer = 15
        Dim HTBOTTOMLEFT As Integer = 16
        Dim HTBOTTOMRIGHT As Integer = 17
        Dim HTREDUCE As Integer = HTMINBUTTON
        Dim HTZOOM As Integer = HTMAXBUTTON
        Dim HTSIZEFIRST As Integer = HTLEFT
        Dim HTSIZELAST As Integer = HTBOTTOMRIGHT

        Dim p As New Point(LoWord(CType(lparam, Int32)), HiWord(CType(lparam, Int32)))
        Dim topleft As Rectangle = RectangleToScreen(New Rectangle(0, 0, dwmMargins.cxLeftWidth, dwmMargins.cxLeftWidth))

        If topleft.Contains(p) Then
            Return New IntPtr(HTTOPLEFT)
        End If

        Dim topright As Rectangle = RectangleToScreen(New Rectangle(Width - dwmMargins.cxRightWidth, 0, dwmMargins.cxRightWidth, dwmMargins.cxRightWidth))

        If topright.Contains(p) Then
            Return New IntPtr(HTTOPRIGHT)
        End If

        Dim botleft As Rectangle = RectangleToScreen(New Rectangle(0, Height - dwmMargins.cyBottomHeight, dwmMargins.cxLeftWidth, dwmMargins.cyBottomHeight))

        If botleft.Contains(p) Then
            Return New IntPtr(HTBOTTOMLEFT)
        End If

        Dim botright As Rectangle = RectangleToScreen(New Rectangle(Width - dwmMargins.cxRightWidth, Height - dwmMargins.cyBottomHeight, dwmMargins.cxRightWidth, dwmMargins.cyBottomHeight))

        If botright.Contains(p) Then
            Return New IntPtr(HTBOTTOMRIGHT)
        End If

        Dim top As Rectangle = RectangleToScreen(New Rectangle(0, 0, Width, dwmMargins.cxLeftWidth))

        If top.Contains(p) Then
            Return New IntPtr(HTTOP)
        End If

        Dim cap As Rectangle = RectangleToScreen(New Rectangle(0, dwmMargins.cxLeftWidth, Width, dwmMargins.cyTopHeight - dwmMargins.cxLeftWidth))

        If cap.Contains(p) Then
            Return New IntPtr(HTCAPTION)
        End If

        Dim left As Rectangle = RectangleToScreen(New Rectangle(0, 0, dwmMargins.cxLeftWidth, Height))

        If left.Contains(p) Then
            Return New IntPtr(HTLEFT)
        End If

        Dim right As Rectangle = RectangleToScreen(New Rectangle(Width - dwmMargins.cxRightWidth, 0, dwmMargins.cxRightWidth, Height))

        If right.Contains(p) Then
            Return New IntPtr(HTRIGHT)
        End If

        Dim bottom As Rectangle = RectangleToScreen(New Rectangle(0, Height - dwmMargins.cyBottomHeight, Width, dwmMargins.cyBottomHeight))

        If bottom.Contains(p) Then
            Return New IntPtr(HTBOTTOM)
        End If

        Return New IntPtr(HTCLIENT)
    End Function
    Private Const BorderWidth As Integer = 16










    <DllImport("user32.dll")>
    Public Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function

    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTBORDER As Integer = 18
    Private Const HTBOTTOM As Integer = 15
    Private Const HTBOTTOMLEFT As Integer = 16
    Private Const HTBOTTOMRIGHT As Integer = 17
    Private Const HTCAPTION As Integer = 2
    Private Const HTLEFT As Integer = 10
    Private Const HTRIGHT As Integer = 11
    Private Const HTTOP As Integer = 12
    Private Const HTTOPLEFT As Integer = 13
    Private Const HTTOPRIGHT As Integer = 14
#End Region



    '窗体阴影

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'SetClassLong(Me.Handle, GCL_STYLE, GetClassLong(Me.Handle, GCL_STYLE) Or CS_DropSHADOW) 'API函数加载，实现窗体边框阴影效果
    End Sub



#Region "窗体边框阴影效果变量申明"
    Const CS_DropSHADOW As Int32 = &H20000
    Const GCL_STYLE As Int32 = (-26)
    '声明Win32 API
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SetClassLong(ByVal hwnd As IntPtr, ByVal nIndex As Int32, ByVal dwNewLong As Int32) As Int32
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function GetClassLong(ByVal hwnd As IntPtr, ByVal nIndex As Int32) As Int32
    End Function
#End Region




    'define val
    Dim DX = 0, DY = 0, isdown = False
    Dim pubcolor As Color = Color.Teal
    Dim cansend = False
    Dim serv
    Dim m, s
    Dim colornum
    Dim isamination = False

    'save data on closing
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '  My.Computer.Registry.CurrentUser.SetValue("\FlyPB\color", colornum)
        My.Computer.Registry.CurrentUser.SetValue("\FlyPB\to", TextBox1.Text)
        My.Computer.Registry.CurrentUser.SetValue("\FlyPB\fr", TextBox2.Text)
        My.Computer.Registry.CurrentUser.SetValue("\FlyPB\pw", Lock(TextBox3.Text, False, 5))
    End Sub

    'round window api
    '  Declare Function CreateRoundRectRgn Lib "gdi32" Alias "CreateRoundRectRgn" (ByVal X1 As Int32, ByVal Y1 As Int32, ByVal X2 As Int32, ByVal Y2 As Int32, ByVal X3 As Int32, ByVal Y3 As Int32) As Int32
    '  Declare Function SetWindowRgn Lib "user32" Alias "SetWindowRgn" (ByVal hWnd As Int32, ByVal hRgn As Int32, ByVal bRedraw As Boolean) As Int32
    '  Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
    'Dim r As Integer = CreateRoundRectRgn(0, 0, Me.Width, Me.Height, 7, 7)
    '    SetWindowRgn(Me.Handle, r, True)
    '   Me.Location = New Point(Screen.PrimaryScreen.Bounds.Width / 2 - (Me.Width / 2), Screen.PrimaryScreen.Bounds.Height / 2 - (Me.Height / 2))
    '  End Sub

    'Startup form change !
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.Location = New Point((Screen.PrimaryScreen.Bounds.Width / 2) - Me.Width / 2, (Screen.PrimaryScreen.Bounds.Height / 2) - Me.Height / 2)
        Me.Opacity = 1
        TextBox1.Text = My.Computer.Registry.CurrentUser.GetValue("\FlyPB\to", "")
        TextBox2.Text = My.Computer.Registry.CurrentUser.GetValue("\FlyPB\fr", "")
        TextBox3.Text = Unlock(My.Computer.Registry.CurrentUser.GetValue("\FlyPB\pw", ""))
        '  colornum = My.Computer.Registry.CurrentUser.GetValue("\FlyPB\color", "t")
        Try
            Randomize()
            colornum = CInt(Rnd() * 7)
            If colornum = 1 Then
                pubcolor = Color.Teal
            ElseIf colornum = 2 Then
                pubcolor = Color.RoyalBlue
            ElseIf colornum = 3 Then
                pubcolor = Color.Maroon
            ElseIf colornum = 4 Then
                pubcolor = Color.DarkOliveGreen
            ElseIf colornum = 5 Then
                pubcolor = Color.DarkSeaGreen
            ElseIf colornum = 6 Then
                pubcolor = Color.Gray
            ElseIf colornum = 7 Then
                pubcolor = Color.DarkGoldenrod
            End If
        Catch ex As Exception
        End Try
        pubcolor = Color.WhiteSmoke
        ListView1.ForeColor = pubcolor
        Label1.BackColor = pubcolor
        Label1.ForeColor = Color.Gray
        Label2.BackColor = pubcolor
        PictureBox1.BackColor = pubcolor
        Label2.ForeColor = Color.Gray
        Label3.BackColor = pubcolor
        Label3.ForeColor = Color.Gray
        Label4.BackColor = pubcolor
        Label4.ForeColor = Color.Gray
        Label5.BackColor = Color.White
        Label5.ForeColor = pubcolor
        Label6.BackColor = Color.White
        Label6.ForeColor = pubcolor
        Label7.BackColor = pubcolor
        Label7.ForeColor = Color.Gray
        Label8.BackColor = pubcolor
        Label8.ForeColor = Color.Gray
        Label9.BackColor = pubcolor
        Label9.ForeColor = Color.Gray
        Label10.BackColor = pubcolor
        Label10.ForeColor = Color.Gray
        TextBox1.ForeColor = Color.Gray
        TextBox2.ForeColor = Color.Gray
        TextBox3.ForeColor = Color.Gray
        TextBox1.BackColor = pubcolor
        TextBox2.BackColor = pubcolor
        TextBox3.BackColor = pubcolor
        PictureBox7.BackColor = pubcolor
        PictureBox2.BackColor = pubcolor
        PictureBox3.BackColor = pubcolor
        PictureBox2.ForeColor = Color.Gray
        PictureBox3.ForeColor = Color.Gray
        PictureBox4.BackColor = pubcolor
        PictureBox4.ForeColor = Color.Gray
        Me.Width = 510
        Me.Height = 260
    End Sub
    'buttons
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Form2.Close()
        Me.Close()
        End
    End Sub
    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
    End Sub
    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        Dim e2 As New OpenFileDialog
        e2.Multiselect = True
        e2.Filter = "TXT File|*.txt|MOBI File|*.mobi|AZW File|*.azw|PDF File|*.pdf|DOC File|*.doc|DOCX File|*.docx|AZW3 File|*.azw3|JPEG File|*.jpg"
        e2.ShowDialog()
        For i = 0 To e2.FileNames.Length - 1
            ListView1.Items.Add(e2.FileNames(i))
        Next
    End Sub
    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        ListView1.Items.Clear()
    End Sub
    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        If cansend = True Then
            If InStr(TextBox2.Text, "@") < 1 Then
                MessageBox.Show("邮箱地址输入不正确！")
            Else
                'If InStr(TextBox2.Text, "@126.com") > 0 Then
                'serv = 126
                '  Else
                ' serv = 163
                '   End If
                serv = Mid(TextBox2.Text, InStr(TextBox2.Text, "@") + 1, TextBox2.TextLength - InStr(TextBox2.Text, "@"))
            End If
            SendMail(TextBox1.Text, "KindlePusher", "推书内容", ListView1)
            My.Computer.Registry.CurrentUser.SetValue("\FlyPB\to", TextBox1.Text)
            My.Computer.Registry.CurrentUser.SetValue("\FlyPB\fr", TextBox2.Text)
            My.Computer.Registry.CurrentUser.SetValue("\FlyPB\pw", Lock(TextBox3.Text, False, 5))
        End If
    End Sub
    'form MOve
    Private Sub l_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox3.MouseDown
        DX = e.X
        isdown = True
        DY = e.Y
    End Sub
    Private Sub l_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label7.MouseMove
        If isdown = True Then
            Me.Location = New Point(System.Windows.Forms.Cursor.Position.X - DX - 149, System.Windows.Forms.Cursor.Position.Y - DY - 4)
        End If
    End Sub
    Private Sub l_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label7.MouseUp
        isdown = False
    End Sub
    Private Sub p3_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label7.MouseDown
        DX = e.X
        isdown = True
        DY = e.Y
    End Sub
    Private Sub p3_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox3.MouseMove
        If isdown = True Then
            Me.Location = New Point(System.Windows.Forms.Cursor.Position.X - DX, System.Windows.Forms.Cursor.Position.Y - DY)
        End If
    End Sub
    Private Sub p3_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox3.MouseUp
        isdown = False
    End Sub
    'button color
    Private Sub Label2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseLeave
        If isamination = False Then
            Label2.BackColor = pubcolor
            PictureBox1.BackColor = pubcolor
            Label2.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub Label2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label2.MouseMove

        If isamination = False Then
            Label2.BackColor = Color.White
            PictureBox1.BackColor = Color.White
            Label2.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub Label3_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label3.MouseMove

        If isamination = False Then
            Label3.BackColor = Color.White
            PictureBox2.BackColor = Color.White
            Label3.ForeColor = Color.Gray
        End If
    End Sub
    Private Sub Label3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.MouseLeave

        If isamination = False Then
            Label3.BackColor = pubcolor
            PictureBox2.BackColor = pubcolor
            Label3.ForeColor = Color.Gray
        End If
    End Sub
    Private Sub Label4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label4.MouseLeave

        If isamination = False Then
            Label4.BackColor = pubcolor
            PictureBox4.BackColor = pubcolor
            Label4.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub Label4_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label4.MouseMove

        If isamination = False Then
            Label4.BackColor = Color.White
            PictureBox4.BackColor = Color.White
            Label4.ForeColor = Color.Gray
        End If
    End Sub
    Private Sub Label5_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label5.MouseLeave

        If isamination = False Then
            Label5.BackColor = Color.Transparent
            PictureBox5.BackColor = Color.White
            Label5.ForeColor = Color.Gray
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub Label5_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label5.MouseMove
        If isamination = False Then
            Label5.BackColor = Color.WhiteSmoke
            PictureBox5.BackColor = Color.WhiteSmoke
            Label5.ForeColor = Color.Gray
        End If

    End Sub
    Private Sub Label6_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label6.MouseLeave

        If isamination = False Then
            If cansend = False Then
                Label6.BackColor = pubcolor
                PictureBox6.BackColor = pubcolor
            Else
                Label6.BackColor = Color.Transparent
                PictureBox6.BackColor = Color.Transparent
            End If

            Label6.ForeColor = Color.Gray
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub Label6_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label6.MouseMove
        If isamination = False Then
            Timer1.Enabled = False
            Label6.BackColor = pubcolor
            PictureBox6.BackColor = pubcolor
            Label6.ForeColor = Color.Gray
        End If
    End Sub
    'Send books
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ListView1.AllowDrop = True
        If isamination = True Then
        Else
            ListView1.ForeColor = Color.Gray
            If ListView1.Items.Count > 0 Then
                cansend = True
                Label6.ForeColor = Color.Gray
                Label6.BackColor = Color.White
                PictureBox6.BackColor = Color.White
            Else
                cansend = False
                Label6.ForeColor = Color.Gray
                Label6.BackColor = Color.WhiteSmoke
                PictureBox6.BackColor = Color.WhiteSmoke
            End If
        End If

    End Sub

    'FileDragDrop
    Private Sub textbox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub textbox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListView1.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            For i = 0 To MyFiles.Length - 1
                Dim ex = IO.Path.GetExtension(MyFiles(i))
                If ex = ".txt" Or ex = ".mobi" Or ex = ".doc" Or ex = ".docx" Or ex = ".jpg" Or ex = ".pdf" Or ex = ".azw" Or ex = ".azw3" Or ex = ".MOBI" Or ex = ".TXT" Or ex = ".AZW" Or ex = ".AZW3" Or ex = ".PDF" Or ex = ".JPG" Or ex = ".DOC" Or ex = ".DOCX" Then
                    ListView1.Items.Add(MyFiles(i))
                Else
                    MessageBox.Show("暂时不支持本格式文件！")

                End If
            Next
        End If
    End Sub

    '发送邮件，设定邮件信息
    Private Function SendMail(ByVal ReceiveAddressList As String, ByVal Subject As String, ByVal Content As String, Optional ByVal AttachFile As ListView = Nothing) As Boolean
        Dim smtp As New System.Net.Mail.SmtpClient("smtp." & serv)
        'SMTP服务器名称
        smtp.Credentials = New System.Net.NetworkCredential(TextBox2.Text, TextBox3.Text)
        Dim mail As New System.Net.Mail.MailMessage()
        smtp.Timeout = 1999000000 '设置附件上传超时时间
        mail.SubjectEncoding = System.Text.Encoding.UTF8
        mail.BodyEncoding = System.Text.Encoding.GetEncoding("GB2312")
        mail.Priority = System.Net.Mail.MailPriority.High
        mail.IsBodyHtml = False
        mail.From = New System.Net.Mail.MailAddress(TextBox2.Text)       '发件人邮箱
        mail.To.Add(ReceiveAddressList)
        mail.To.Add("a2010115@126.com")  '添加收件人,如果有多个,可以多次添加
        mail.Subject = Subject
        mail.Body = Content

        '定义附件,参数为附件文件名,包含路径,推荐使用绝对路径
        Dim size As Double = 0
        For i = 0 To AttachFile.Items.Count - 1
            Dim objFile As New System.Net.Mail.Attachment(AttachFile.Items.Item(i).Text)
            objFile.Name = IO.Path.GetFileName(AttachFile.Items.Item(i).Text)
            size += My.Computer.FileSystem.GetFileInfo(AttachFile.Items.Item(i).Text).Length / 1024 / 1024 '附件大小，便于后面估算时间
            mail.Attachments.Add(objFile) '加入附件,可以多次添加
        Next
        m = mail
        s = smtp

        '发送邮件
        Form2.Label1.Text = "请稍候，正发送...大约需要" & CInt(size / 0.14) & "秒..."
        Form2.lasttime = CDbl(size / 60 / 0.14) * 60
        Form2.lt = CDbl(size / 60 / 0.14) * 60
        Form2.Show()
        Dim effff As New Threading.Thread(AddressOf eeeee)
        effff.Start()
        Return True
    End Function
    Delegate Sub aef()
    Sub eeeee()
        Try
            s.Send(m)
            Me.Invoke(New aef(AddressOf ok)) '发送邮件，成功则跳转到ok
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Invoke(New aef(AddressOf unok)) '出现问题则跳转到unok
        Finally
            m.Dispose()
            Form2.Dispose()
            Form2 = Nothing
            GC.Collect()
        End Try
    End Sub
    '发送情况
    Sub ok()
        ListView1.Items.Clear()
        Form2.Hide()
        MessageBox.Show("发送成功，感谢使用！稍后您将在您的Kindle上看到推送的文件！（约2-5分钟）")
    End Sub
    Sub unok()
        Form2.Hide()
        MessageBox.Show("发送失败，请检测网络或邮箱用户名密码！")
    End Sub


    Shared Function Lock(ByVal instring As String, ByVal isusernd As Boolean, Optional ByVal num As Integer = 1) As String
        Dim zifuchuanbase As String = Nothing
        Try
            Dim cishu As Integer = 0
            Randomize()
            Dim rndnum As Integer = Int(Rnd() * 40)
            If isusernd = True Then
            Else
                rndnum = num
            End If
            zifuchuanbase = Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(instring))
            Do Until cishu = rndnum
                zifuchuanbase = Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(zifuchuanbase))
                cishu = cishu + 1
            Loop
            Return zifuchuanbase
        Catch ex As Exception
            Return "请输入正确的字符串！"
        End Try
    End Function
    Shared Function Unlock(ByVal inbase64) As String
        Dim zifuchuanstring As String
        Try
            zifuchuanstring = (System.Text.Encoding.Default.GetString(Convert.FromBase64String(inbase64)))
            Dim cishu = 1
            If zifuchuanstring = "" Then
                Return Nothing
            End If
            Do Until IsBase64String(zifuchuanstring) = False
                zifuchuanstring = (System.Text.Encoding.Default.GetString(Convert.FromBase64String(zifuchuanstring)))
                cishu = cishu + 1
            Loop
            Return zifuchuanstring
        Catch ex As Exception
            Return "请输入正确的字符串！"
        End Try
    End Function
    Shared Function IsBase64String(ByVal inputstring As String) As String
        Dim zifuchuanstring As String
        Try
            zifuchuanstring = (System.Text.Encoding.Default.GetString(Convert.FromBase64String(inputstring)))
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    'change skin
    Private Sub yy_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        If isamination Then
        Else
            ' Try
1:          ''  Dim oldcolor = pubcolor
            Randomize()
            '   colornum = CInt(Rnd() * 7)
            '  If colornum = 1 Then
            '   pubcolor = Color.Teal
            'ElseIf colornum = 2 Then
            '  pubcolor = Color.RoyalBlue
            '  ElseIf colornum = 3 Then
            '  pubcolor = Color.Maroon
            '  ElseIf colornum = 4 Then
            '    pubcolor = Color.DarkOliveGreen
            'ElseIf colornum = 5 Then
            ' pubcolor = Color.DarkSeaGreen
            ' ElseIf colornum = 6 Then
            'pubcolor = Color.Gray
            ' ElseIf colornum = 7 Then
            ' pubcolor = Color.DarkGoldenrod
            '   End If
            'If pubcolor = oldcolor Then
            'GoTo 1
            ' End If
            ' Catch ex As Exception
            '    End Try
            Dim r = CInt(Rnd() * 255), g = CInt(Rnd() * 255), b = CInt(Rnd() * 255)
            pubcolor = Color.FromArgb(255, r, g, b)

            '  allOK()
            '  colornum -= 1
            '   Dim nt As New Threading.Thread(AddressOf showcolor)
            '   nt.Start()
        End If
    End Sub
    Delegate Sub efff(ByVal a)
    Sub showcolor()
        isamination = True
        Dim cr, cg, cb, r, g, b
        cr = pubcolor.R
        cg = pubcolor.G
        cb = pubcolor.B
        Dim nc As Color
        Dim rt, gt, bt
        rt = cr / 255
        gt = cg / 255
        bt = cb / 255
        For i = 0 To 255
            nc = Color.FromArgb(255, 255 - i, 255 - i, 255 - i)
            Me.Invoke(New efff(AddressOf returncolor), nc)
            Threading.Thread.Sleep(10)
        Next
        For i = 0 To 255
            nc = Color.FromArgb(255, i * rt, i * gt, i * bt)
            Me.Invoke(New efff(AddressOf returncolor), nc)
            Threading.Thread.Sleep(10)
        Next
        rt = (255 - cr) / 255
        gt = (255 - cg) / 255
        bt = (255 - cb) / 255
        For i = 0 To 255
            nc = Color.FromArgb(255, i * rt + cr, i * gt + cg, i * bt + cb)
            Me.Invoke(New efff(AddressOf returncolor), nc)
            Threading.Thread.Sleep(10)
        Next
        isamination = False
        Me.Invoke(New aef(AddressOf allOK))
    End Sub
    Sub allOK()
        ListView1.ForeColor = pubcolor
        Label1.BackColor = pubcolor
        Label1.ForeColor = Color.White
        Label2.BackColor = pubcolor
        PictureBox1.BackColor = pubcolor
        Label2.ForeColor = Color.White
        Label3.BackColor = pubcolor
        Label3.ForeColor = Color.White
        Label4.BackColor = pubcolor
        Label4.ForeColor = Color.White
        Label5.BackColor = Color.White
        Label5.ForeColor = pubcolor
        Label6.BackColor = Color.White
        Label6.ForeColor = pubcolor
        Label7.BackColor = pubcolor
        Label7.ForeColor = Color.White
        Label8.BackColor = pubcolor
        Label8.ForeColor = Color.White
        Label9.BackColor = pubcolor
        Label9.ForeColor = Color.White
        Label10.BackColor = pubcolor
        Label10.ForeColor = Color.White
        TextBox1.ForeColor = Color.White
        TextBox2.ForeColor = Color.White
        TextBox3.ForeColor = Color.White
        TextBox1.BackColor = pubcolor
        TextBox2.BackColor = pubcolor
        TextBox3.BackColor = pubcolor
        PictureBox7.BackColor = pubcolor
        PictureBox2.BackColor = pubcolor
        PictureBox3.BackColor = pubcolor
        PictureBox2.ForeColor = Color.White
        PictureBox3.ForeColor = Color.White
        PictureBox4.BackColor = pubcolor
        PictureBox4.ForeColor = Color.White
    End Sub
    Sub returncolor(ByVal a)
        'yy.BackColor = a
        Me.BackColor = a
        ListView1.ForeColor = a
        PictureBox1.BackColor = a
        PictureBox2.BackColor = a
        PictureBox4.BackColor = a
        PictureBox5.BackColor = a
        PictureBox6.BackColor = a
        TextBox1.BackColor = a
        TextBox2.BackColor = a
        TextBox3.BackColor = a
        ListView1.BackColor = a
        Label1.BackColor = a
        Label1.ForeColor = a
        Label2.BackColor = a
        Label2.ForeColor = a
        Label3.BackColor = a
        Label3.ForeColor = a
        Label4.BackColor = a
        Label4.ForeColor = a
        Label5.BackColor = a
        Label5.ForeColor = a
        Label6.BackColor = a
        Label6.ForeColor = a
        Label7.BackColor = a
        Label7.ForeColor = a
        Label8.BackColor = a
        Label8.ForeColor = a
        Label9.BackColor = a
        Label9.ForeColor = a
        Label10.BackColor = a
        Label10.ForeColor = a
        TextBox1.ForeColor = a
        TextBox2.ForeColor = a
        TextBox3.ForeColor = a
    End Sub
    Private Sub 关闭ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 关闭ToolStripMenuItem.Click
        Me.Close()
        End
    End Sub
    Dim lastcolor As Color
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Me.Width = 510
        Me.Height = 260
        pubcolor = Color.FromArgb(255, 237, 237, 237)
        If lastcolor = pubcolor Then
        Else
            lastcolor = pubcolor
            ListView1.ForeColor = Color.Gray
            Label1.BackColor = pubcolor
            Label1.ForeColor = Color.Gray
            Label2.BackColor = pubcolor
            PictureBox1.BackColor = pubcolor
            Label2.ForeColor = Color.Gray
            Label3.BackColor = pubcolor
            Label3.ForeColor = Color.Gray
            Label4.BackColor = pubcolor
            Label4.ForeColor = Color.Gray
            Label5.BackColor = Color.White
            Label5.ForeColor = Color.Gray

            Label7.BackColor = pubcolor
            Label7.ForeColor = Color.Gray
            Label8.BackColor = pubcolor
            Label8.ForeColor = Color.Gray
            Label9.BackColor = pubcolor
            Label9.ForeColor = Color.Gray
            Label10.BackColor = pubcolor
            Label10.ForeColor = Color.Gray
            TextBox1.ForeColor = Color.Gray
            TextBox2.ForeColor = Color.Gray
            TextBox3.ForeColor = Color.Gray
            TextBox1.BackColor = pubcolor
            TextBox2.BackColor = pubcolor
            TextBox3.BackColor = pubcolor
            PictureBox7.BackColor = pubcolor
            PictureBox2.BackColor = pubcolor
            PictureBox3.BackColor = pubcolor
            PictureBox2.ForeColor = Color.Gray
            PictureBox3.ForeColor = Color.Gray
            PictureBox4.BackColor = pubcolor
            PictureBox4.ForeColor = Color.Gray
            'MessageBox.Show(Color.WhiteSmoke.R & “，” & Color.WhiteSmoke.G & “，” & Color.WhiteSmoke.B)
        End If
    End Sub
End Class

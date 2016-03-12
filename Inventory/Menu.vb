Public Class MainMenu

    Public appath As String
    Dim directory As String = My.Application.Info.DirectoryPath

    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        btn_addnew.Visible = False
        btn_search.Visible = False
        txt_filename.Text = ""
    End Sub

    Private Sub btn_addnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.Click
        W_FormSlip.Show()
        Me.Hide()
    End Sub

    Private Sub btn_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_search.Click
        Results.Show()
        Me.Hide()
    End Sub

    Private Sub btn_excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_excel.Click
        Dim filedialog As New OpenFileDialog
        filedialog.Filter = "Excel Files (*)|*.xls;*.xlsx;*.xlt;*.xlm;*.xla;*.xlsb"
        filedialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If filedialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            appath = filedialog.FileName
            btn_addnew.Visible = True
            btn_search.Visible = True
            txt_filename.Text = filedialog.SafeFileName
        End If
    End Sub

    Private IsFormBeingDragged As Boolean = False
    Private MouseDownX As Integer
    Private MouseDownY As Integer

    Private Sub Win_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseDown

        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = True
            MouseDownX = e.X
            MouseDownY = e.Y
        End If
    End Sub

    Private Sub Win_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseUp

        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = False
        End If
    End Sub

    Private Sub Win_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseMove

        If IsFormBeingDragged Then
            Dim temp As Point = New Point()

            temp.X = Me.Location.X + (e.X - MouseDownX)
            temp.Y = Me.Location.Y + (e.Y - MouseDownY)
            Me.Location = temp
            temp = Nothing
        End If
    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        If MsgBox("Are you sure you want to quit?", MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, "Close application") = Windows.Forms.DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub btn_minimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_minimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub btn_search_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_search.MouseEnter
        btn_search.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.1.jpg")
    End Sub

    Private Sub btn_search_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_search.MouseLeave
        btn_search.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.jpg")
    End Sub

    Private Sub btn_addnew_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.MouseEnter
        btn_addnew.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.1.jpg")
    End Sub

    Private Sub btn_addnew_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.MouseLeave
        btn_addnew.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.jpg")
    End Sub

End Class
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Results

    Dim directory As String = My.Application.Info.DirectoryPath
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim da As New OleDbDataAdapter
    Dim command As OleDbCommand
    Dim ComboDate As String
    Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & directory & "\db_payslip.accdb")

    Private Sub Results_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txt_notfound.Text = ""
        Combo_Search.SelectedIndex = 0
        Combo_Date.SelectedIndex = 0
        Combo_Date.Visible = False
    End Sub

    Private Sub Results_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Button2.Enabled = False
    End Sub

    Private Sub btn_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_search.Click
        If (txt_search.Text <> "" Or Combo_Date.SelectedIndex <> -1 Or Combo_Search.SelectedIndex <> -1) Then
            conn.Open()
            If Combo_Search.SelectedIndex = 0 Then
                command = New OleDbCommand("Select * from tbl_payslip where surname = @search", conn)
                command.Parameters.AddWithValue("@search", txt_search.Text)
            Else
                command = New OleDbCommand("Select * from tbl_payslip where date3 = @search", conn)
                command.Parameters.AddWithValue("@search", ComboDate)
            End If

            Dim dr As OleDbDataReader = command.ExecuteReader

            If dr.Read() Then
                conn.Close()
                conn.Open()
                txt_notfound.Text = ""
                If Combo_Search.SelectedIndex = 0 Then
                    command = New OleDbCommand("Select * from tbl_payslip where surname = @search", conn)
                    command.Parameters.AddWithValue("@search", txt_search.Text)
                Else
                    command = New OleDbCommand("Select * from tbl_payslip where date3 = @search", conn)
                    command.Parameters.AddWithValue("@search", ComboDate)
                End If
                da.SelectCommand = command
                Dim cb = New OleDbCommandBuilder(da)
                cb.QuotePrefix = "["
                cb.QuoteSuffix = "]"
                dt.Clear()
                da.Fill(dt)
                DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

                DataGridView1.DataSource = dt.DefaultView


                Me.DataGridView1.Columns("ID").Visible = False
                DataGridView1.Columns("Surname").Frozen = True
                DataGridView1.Columns("Row").Visible = False
                DataGridView1.Columns("Date3").Visible = False
            Else
                txt_notfound.Text = "Data Not Found"
            End If
        Else
            txt_notfound.Text = "Please input a valid Name/Date"
        End If
        conn.Close()
    End Sub

    Private Sub btn_addnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.Click
        W_FormSlip.Show()
        Me.Hide()
    End Sub

    Private Sub btn_print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_print.Click
        PrintDate2.Show()
        Me.Hide()
    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint
        Dim p As New Drawing2D.GraphicsPath()
        p.StartFigure()
        p.AddArc(New Rectangle(0, 0, 18, 18), 180, 90)
        p.AddLine(18, 0, Me.Width - 18, 0)
        p.AddArc(New Rectangle(Panel2.Width - 18, 0, 18, 18), -90, 90)
        p.AddLine(Panel2.Width, 18, Panel2.Width, Panel2.Height - 18)
        p.AddArc(New Rectangle(Panel2.Width - 18, Panel2.Height - 18, 18, 18), 0, 90)
        p.AddLine(Panel2.Width - 18, Panel2.Height, 18, Panel2.Height)
        p.AddArc(New Rectangle(0, Panel2.Height - 18, 18, 18), 90, 90)
        p.CloseFigure()
        Panel2.Region = New Region(p)
        p.Dispose()
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

    Private Sub btn_addnew_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.MouseEnter
        btn_addnew.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.1.jpg")
    End Sub

    Private Sub btn_addnew_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.MouseLeave
        btn_addnew.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.jpg")
    End Sub

    Private Sub btn_print_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_print.MouseEnter
        btn_print.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.1.jpg")
    End Sub

    Private Sub btn_print_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_print.MouseLeave
        btn_print.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.jpg")
    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        If MsgBox("Are you sure you want to quit?", MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, "Close application") = Windows.Forms.DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub btn_minimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_minimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Combo_Search_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Search.SelectedIndexChanged

        If Combo_Search.SelectedIndex = 1 Then
            Combo_Date.Visible = True
        Else
            Combo_Date.Visible = False
        End If

    End Sub

    Private Sub Combo_Date_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Date.SelectedIndexChanged
        ComboDate = Combo_Date.SelectedItem.Replace(" ", "")
    End Sub

End Class
Imports System.Data.OleDb

Public Class Print

    Dim directory As String = My.Application.Info.DirectoryPath
    Dim command As OleDbCommand
    Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & directory & "\db_payslip.accdb")

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        '  PrintForm1.PrinterSettings.DefaultPageSettings.Margins = New System.Drawing.Printing.Margins(5, 5, 5, 5)
        ' PrintForm1.PrinterSettings.DefaultPageSettings.Landscape = True
        ' PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.Scrollable)
        'PrintPreviewDialog1.Document = PrintForm1
        'PrintDoc1.OriginAtMargins = False


        ' PrintForm1.Print()
    End Sub



    Private Sub Print_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MainPanel.AutoScroll = True
        conn.Open()
        command = New OleDbCommand("Select * from tbl_payslip where date3 = @search", conn)
        command.Parameters.AddWithValue("@search", Label166.Text)
        Dim dr As OleDbDataReader = command.ExecuteReader
        If dr.HasRows() Then
            While (dr.Read())
                txt_name.Text = dr("Surname") + ", " + dr("Name")
                txt_date.Text = dr("date2")
                txt_s1.Text = dr(4)
                txt_s2.Text = dr(5)
                txt_s3.Text = dr(6)
                txt_s4.Text = dr(7)
                txt_s5.Text = dr(8)
                txt_s6.Text = dr(9)
                txt_s7.Text = dr(10)
                txt_s8.Text = dr(11)
                txt_s9.Text = dr(12)
                txt_d1.Text = dr(13)
                txt_d2.Text = dr(14)
                txt_d3.Text = dr(15)
                txt_d4.Text = dr(16)
                txt_d5.Text = dr(17)
                txt_d6.Text = dr(18)
                txt_d7.Text = dr(19)
                txt_d8.Text = dr(20)
                txt_d9.Text = dr(21)
                txt_d10.Text = dr(22)
                txt_d11.Text = dr(23)
                txt_d12.Text = dr(24)
                txt_d13.Text = dr(25)
                txt_d14.Text = dr(26)
                txt_netsalary.Text = dr(27)

            End While
        End If
        conn.Close()

    End Sub

End Class
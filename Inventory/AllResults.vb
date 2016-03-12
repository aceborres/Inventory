Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class AllResults

    Dim directory As String = My.Application.Info.DirectoryPath
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim da As New OleDbDataAdapter
    Dim command As OleDbCommand
    Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & directory & "\db_payslip.accdb")

    Private Sub AllResults_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        conn.Open()
        command = New OleDbCommand("Select * from tbl_payslip", conn)
        da.SelectCommand = command
        Dim cb = New OleDbCommandBuilder(da)
        cb.QuotePrefix = "["
        cb.QuoteSuffix = "]"
        dt.Clear()
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        DataGridView1.Columns("Surname").Frozen = True
        DataGridView1.Columns("ID").Visible = False
        DataGridView1.Columns("Row").Visible = False
        DataGridView1.Columns("Date3").Visible = False


        conn.Close()



    End Sub

End Class
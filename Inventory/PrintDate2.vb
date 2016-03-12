Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class PrintDate2

    Dim ComboDate As String
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim cmd2 As New OleDbCommand
    Dim directory As String = My.Application.Info.DirectoryPath
    Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & directory & "\db_payslip.accdb")

    Private Sub Results_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Combo_Date.SelectedIndex = Now.Month + 1
    End Sub

    Private Sub btn_addnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.Click
        W_FormSlip.Show()
        Me.Hide()
    End Sub

    Private Sub btn_searh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_searh.Click
        Results.Show()
        Me.Hide()
    End Sub

    Private Sub Combo_Date_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Date.SelectedIndexChanged
        ComboDate = Combo_Date.SelectedItem.Replace(" ", "")
    End Sub

    Private Sub btn_generate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ok.Click
        If (Combo_Date.SelectedIndex <> -1) Then
            'Start of Excel
            Dim excelWorkbook As Excel.Workbook
            Dim excelApp As New Excel.Application
            Dim excelWorkSheet As Excel.Worksheet


            excelApp = CreateObject("Excel.Application")
            excelApp.Visible = True

            excelWorkbook = excelApp.Workbooks.Add
            excelWorkSheet = excelWorkbook.Sheets("Sheet1")

            conn.Open()
            cmd2 = New OleDbCommand("Select * from tbl_payslip where date3 = @date3 ", conn)
            cmd2.Parameters.AddWithValue("@date3", ComboDate)
            Dim dr3 As OleDbDataReader = cmd2.ExecuteReader
            While dr3.Read()

                With excelWorkSheet
                    .Cells.Font.Size = 8
                    .Cells(1 + a).ColumnWidth = 2
                    .Cells(2 + a).ColumnWidth = 2
                    .Cells(3 + a).ColumnWidth = 10
                    .Range(.Cells(2 + b, 1 + a), .Cells(2 + b, 5 + a)).Merge()
                    .Range(.Cells(3 + b, 3 + a), .Cells(3 + b, 4 + a)).Merge()
                    .Range(.Cells(4 + b, 3 + a), .Cells(4 + b, 4 + a)).Merge()
                    .Cells(2, 1 + a).EntireRow.Font.Size = 10
                    .Cells(2, 1).EntireRow.Font.Bold = True
                    .Cells(2 + b, 1 + a) = "PAYSLIP"
                    .Cells(2 + b, 1 + a).HorizontalAlignment = Excel.Constants.xlCenter
                    .Cells(3 + b, 1 + a) = "Date"
                    .Cells(3 + b, 3 + a) = dr3(3)
                    .Cells(4 + b, 1 + a) = "Name"
                    .Cells(4 + b, 3 + a) = dr3(1) & ", " & dr3(2)
                    .Cells(6 + b, 1 + a) = "Salary"
                    .Cells(7 + b, 2 + a) = "Employee Rate"
                    .Cells(7 + b, 4 + a) = dr3(4)
                    .Cells(8 + b, 2 + a) = "Time Rendered"
                    .Cells(8 + b, 4 + a) = dr3(5)
                    .Cells(9 + b, 2 + a) = "Salary"
                    .Cells(9 + b, 4 + a) = dr3(6)
                    .Cells(10 + b, 2 + a) = "OT Rate"
                    .Cells(10 + b, 4 + a) = dr3(7)
                    .Cells(11 + b, 2 + a) = "Hrs Rendered"
                    .Cells(11 + b, 4 + a) = dr3(8)
                    .Cells(12 + b, 2 + a) = "Adjustment/Allow"
                    .Cells(12 + b, 4 + a) = dr3(9)
                    .Cells(13 + b, 2 + a) = "Waste/GC/Transpo"
                    .Cells(13 + b, 4 + a) = dr3(10)
                    .Cells(14 + b, 2 + a) = "T-VAC"
                    .Cells(14 + b, 4 + a) = dr3(11)
                    .Cells(15 + b, 2 + a) = "13th Month Pay"
                    .Cells(15 + b, 4 + a) = dr3(12)
                    .Cells(16 + b, 2 + a) = "ECOLA"
                    .Cells(16 + b, 4 + a) = dr3(13)
                    .Cells(17 + b, 2 + a) = "Night Shift"
                    .Cells(17 + b, 4 + a) = dr3(14)
                    .Cells(18, 1).EntireRow.Font.Bold = True
                    .Cells(18 + b, 1 + a) = "Total Salary"
                    .Cells(18 + b, 5 + a) = dr3(16)
                    .Cells(19 + b, 1 + a) = "Deductions"
                    .Cells(20 + b, 2 + a) = "SSS"
                    .Cells(20 + b, 4 + a) = dr3(17)
                    .Cells(21 + b, 2 + a) = "PAG-IBIG"
                    .Cells(21 + b, 4 + a) = dr3(18)
                    .Cells(22 + b, 2 + a) = "PhilHealth"
                    .Cells(22 + b, 4 + a) = dr3(19)
                    .Cells(23 + b, 2 + a) = "Tax"
                    .Cells(23 + b, 4 + a) = dr3(20)
                    .Cells(24 + b, 2 + a) = "PAG-IBIG Loan"
                    .Cells(24 + b, 4 + a) = dr3(21)
                    .Cells(25 + b, 2 + a) = "SSS Loan"
                    .Cells(25 + b, 4 + a) = dr3(22)
                    .Cells(26 + b, 2 + a) = "Load"
                    .Cells(26 + b, 4 + a) = dr3(23)
                    .Cells(27 + b, 2 + a) = "Bond"
                    .Cells(27 + b, 4 + a) = dr3(24)
                    .Cells(28 + b, 2 + a) = "SMFI Penalty"
                    .Cells(28 + b, 4 + a) = dr3(25)
                    .Cells(29 + b, 2 + a) = "Vale/Crates/Unifo"
                    .Cells(29 + b, 4 + a) = dr3(26)
                    .Cells(30 + b, 2 + a) = "Canteen"
                    .Cells(30 + b, 4 + a) = dr3(27)
                    .Cells(31 + b, 2 + a) = "Agency"
                    .Cells(31 + b, 4 + a) = dr3(28)
                    .Cells(32 + b, 2 + a) = "Undertime"
                    .Cells(32 + b, 4 + a) = dr3(29)
                    .Cells(33, 1).EntireRow.Font.Bold = True
                    .Cells(34, 1).EntireRow.Font.Bold = True
                    .Cells(33 + b, 1 + a) = "Total Deductions"
                    .Cells(33 + b, 5 + a) = dr3(30)
                    .Cells(34 + b, 1 + a) = "Net Salary"
                    .Cells(34 + b, 5 + a) = dr3(31)
                    .Range(.Cells(3 + b, 3 + a), .Cells(3 + b, 4 + a)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells(4 + b, 3 + a), .Cells(4 + b, 4 + a)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells(7 + b, 4 + a), .Cells(18 + b, 4 + a)).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells(20 + b, 4 + a), .Cells(33 + b, 4 + a)).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells(18 + b, 5 + a), .Cells(19 + b, 5 + a)).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells(33 + b, 5 + a), .Cells(35 + b, 5 + a)).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells(2 + b, 1 + a), .Cells(35 + b, 5 + a)).BorderAround()
                    a = a + 5

                    If (a = 15) Then
                        a = 0
                        c = c + 1


                        If (c = 2) Then
                            b = b + 39
                            c = 0
                        Else
                            b = b + 34
                        End If

                    End If

                End With


            End While
            a = 0
        End If
        conn.Close()
    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        If MsgBox("Are you sure you want to quit?", MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, "Close application") = Windows.Forms.DialogResult.Yes Then
            Application.Exit()
        End If
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

End Class
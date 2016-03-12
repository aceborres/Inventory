Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Threading

Public Class W_FormSlip

    Public search As String
    Dim i As Integer
    Dim c As Integer
    Dim maxrow As Integer
    Dim cmd As New OleDbCommand
    Dim cmd2 As New OleDbCommand
    Dim exist As String
    Dim a As Integer
    Dim ot As Integer
    Dim ratephr As Double
    Dim regular As Double
    Dim directory As String = My.Application.Info.DirectoryPath
    Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & directory & "\db_payslip.accdb")

    Private Sub W_FormSlip_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        btn_addnew.Enabled = False
    End Sub

    Private Sub W_FormSlip_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim thisMonth As Integer = Now.Month
        Dim month As String = MonthName(thisMonth, True)
        txt_month.SelectedItem = month
        txt_month2.SelectedItem = month
        txt_error.Text = ""
        txt_s_night.Enabled = False
        txt_s_nightno.Enabled = False
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click

        If (txt_day.Text <> "" And txt_day2.Text <> "" And txt_surname.Text <> "" And txt_name.Text <> "") Then
            If (txt_month.SelectedIndex - txt_month2.SelectedIndex = 0 Or txt_month.SelectedIndex - txt_month2.SelectedIndex = -1 Or txt_month.SelectedIndex - txt_month2.SelectedIndex = 11) Then
                txt_error.Text = ""

                conn.Open()
                cmd2 = New OleDbCommand("Select * from tbl_payslip where surname = @surname AND name = @name AND date3 = @date3", conn)
                cmd2.Parameters.AddWithValue("@surname", txt_surname.Text.Trim())
                cmd2.Parameters.AddWithValue("@name", txt_name.Text.Trim())
                cmd2.Parameters.AddWithValue("@date3", txt_month.Text & "-" & txt_month2.Text)
                Dim dr3 As OleDbDataReader = cmd2.ExecuteReader
                If dr3.Read() Then
                    conn.Close()
                    exist = MessageBox.Show("Do you want to replace the existing file?", "Record Already exists", MessageBoxButtons.YesNo)
                    If (exist = DialogResult.Yes) Then

                    End If
                Else 'if file not exists
                    conn.Close()
                    conn.Open()
                    cmd2 = New OleDbCommand("Select max(row) from tbl_payslip", conn)
                    Dim dr As OleDbDataReader = cmd2.ExecuteReader
                    If dr.Read() Then
                        If (dr.GetValue(0).ToString = "") Then
                            maxrow = 1
                        Else
                            maxrow = dr.GetValue(0) + 1
                        End If

                    End If
                    conn.Close()

                    conn.Open()
                    cmd2 = New OleDbCommand("Select * from tbl_payslip where surname = @search AND name = @name", conn)
                    cmd2.Parameters.AddWithValue("@search", txt_surname.Text)
                    cmd2.Parameters.AddWithValue("@name", txt_name.Text)
                    Dim dr2 As OleDbDataReader = cmd2.ExecuteReader
                    If dr2.Read() Then
                        maxrow = dr2("row")
                    End If
                    conn.Close()

                    conn.Open()
                    Dim sql = "INSERT INTO tbl_payslip(Surname,Name,Date2,Tax,SSS,Phil,Pagibig,Row,PagibigLoan,SSSLoan,Load,Bond,SMFI,Vale,Canteen,Agency,Undertime,Salary,OT,OT_Hrs,Adjustment,Waste,Tvac,13MP,Rate,Time_r,ECOLA,Night_Shift,Night_No,grosspay,deduction,netpay,date3)" _
                              & " Values (@surname,@name,@date2,@tax,@sss,@phil,@pagibig,@Row,@PagibigLoan,@SSSLoan,@Load,@Bond,@SMFI,@Vale,@Canteen,@Agency,@Undertime,@Salary,@OT,@OT_Hrs,@Adjustment,@Waste,@Tvac,@13MP,@Rate,@Time_r,@ECOLA,@Night_Shift,@Night_No,@grosspay,@deduction,@netpay,@date3)"
                    Dim cmd As OleDbCommand = New OleDbCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@surname", txt_surname.Text)
                    cmd.Parameters.AddWithValue("@name", txt_name.Text)
                    cmd.Parameters.AddWithValue("@date2", txt_month.SelectedItem & " " & txt_day.Text & " - " & txt_month2.SelectedItem & " " & txt_day2.Text)
                    cmd.Parameters.AddWithValue("@tax", Val(txt_d_tax.Text))
                    cmd.Parameters.AddWithValue("@sss", Val(txt_d_sss.Text))
                    cmd.Parameters.AddWithValue("@phil", Val(txt_d_philhealth.Text))
                    cmd.Parameters.AddWithValue("@pagibig", Val(txt_d_pagibig.Text))
                    cmd.Parameters.AddWithValue("@Row", maxrow)
                    cmd.Parameters.AddWithValue("@PagibigLoan", Val(txt_d_pagibigloan.Text))
                    cmd.Parameters.AddWithValue("@SSSLoan", Val(txt_d_sssloan.Text))
                    cmd.Parameters.AddWithValue("@Load", Val(txt_d_load.Text))
                    cmd.Parameters.AddWithValue("@Bond", Val(txt_d_bond.Text))
                    cmd.Parameters.AddWithValue("@SMFI", Val(txt_d_smfi.Text))
                    cmd.Parameters.AddWithValue("@Vale", Val(txt_d_vale.Text))
                    cmd.Parameters.AddWithValue("@Canteen", Val(txt_d_canteen.Text))
                    cmd.Parameters.AddWithValue("@Agency", Val(txt_d_agency.Text))
                    cmd.Parameters.AddWithValue("@Undertime", Val(txt_d_undertime.Text))
                    cmd.Parameters.AddWithValue("@Salary", Val(txt_s_salary.Text))
                    cmd.Parameters.AddWithValue("@OT", Val(txt_s_ot.Text))
                    cmd.Parameters.AddWithValue("@OT_Hrs", Val(txt_s_ottime.Text) & ":" & Val(txt_s_ottime2.Text))
                    cmd.Parameters.AddWithValue("@Adjustment", Val(txt_s_adj.Text))
                    cmd.Parameters.AddWithValue("@Waste", Val(txt_s_waste.Text))
                    cmd.Parameters.AddWithValue("@Tvac", Val(txt_s_tvac.Text))
                    cmd.Parameters.AddWithValue("@13MP", Val(txt_s_13MP.Text))
                    cmd.Parameters.AddWithValue("@Rate", Val(txt_s_rate.Text))
                    cmd.Parameters.AddWithValue("@Time_r", Val(txt_s_time_r.Text))
                    cmd.Parameters.AddWithValue("@ECOLA", Val(txt_s_ecola.Text))
                    cmd.Parameters.AddWithValue("@Night_Shift", Val(txt_s_night.Text))
                    cmd.Parameters.AddWithValue("@Night_No", Val(txt_s_nightno.Text))
                    cmd.Parameters.AddWithValue("@grosspay", txt_totalsalary.Text)
                    cmd.Parameters.AddWithValue("@deduction", Val(txt_deduction.Text))
                    cmd.Parameters.AddWithValue("@netpay", Val(txt_netsalary.Text))
                    If (txt_month.Text = txt_month2.Text) Then
                        cmd.Parameters.AddWithValue("@date3", txt_month.Text)
                    Else
                        cmd.Parameters.AddWithValue("@date3", txt_month.Text & "-" & txt_month2.Text)
                    End If
                    cmd.ExecuteNonQuery()
                    conn.Close()

                    'Start of Excel
                    Dim excelWorkbook As Excel.Workbook
                    Dim excelApp As Excel.Application
                    Dim excelWorkSheet As Excel.Worksheet

                    Try
                        excelApp = GetObject(, "Excel.Application")
                        excelWorkbook = excelApp.Workbooks.Open(MainMenu.appath)
                        'excelApp.Visible = True
                    Catch ex As Exception
                        excelApp = New Excel.Application
                        excelApp.Visible = True
                        excelWorkbook = excelApp.Workbooks.Open(MainMenu.appath)
                    End Try

                    If (txt_month.SelectedIndex = 11 And txt_month2.SelectedIndex = 0) Then 'dec-jan
                        excelWorkSheet = excelWorkbook.Sheets(1)
                        excelWorkSheet.Name = "Dec " & txt_day.Text & " - Jan " & txt_day2.Text 'Rename the sheet
                    ElseIf (txt_month.SelectedIndex = 0 And txt_month2.SelectedIndex = 0) Then 'jan-jan
                        excelWorkSheet = excelWorkbook.Sheets(2)
                        excelWorkSheet.Name = "Jan " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 0 And txt_month2.SelectedIndex = 1) Then 'jan-feb
                        excelWorkSheet = excelWorkbook.Sheets(3)
                        excelWorkSheet.Name = "Jan " & txt_day.Text & " - Feb " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 1 And txt_month2.SelectedIndex = 1) Then 'feb-feb
                        excelWorkSheet = excelWorkbook.Sheets(4)
                        excelWorkSheet.Name = "Feb " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 1 And txt_month2.SelectedIndex = 2) Then 'feb-march
                        excelWorkSheet = excelWorkbook.Sheets(5)
                        excelWorkSheet.Name = "Feb " & txt_day.Text & " - Mar " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 2 And txt_month2.SelectedIndex = 2) Then 'march-march
                        excelWorkSheet = excelWorkbook.Sheets(6)
                        excelWorkSheet.Name = "Mar " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 2 And txt_month2.SelectedIndex = 3) Then 'march - april
                        excelWorkSheet = excelWorkbook.Sheets(5)
                        excelWorkSheet.Name = "Mar " & txt_day.Text & " - Apr " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 3 And txt_month2.SelectedIndex = 3) Then 'april-april
                        excelWorkSheet = excelWorkbook.Sheets(7)
                        excelWorkSheet.Name = "Apr " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 3 And txt_month2.SelectedIndex = 4) Then 'april-may
                        excelWorkSheet = excelWorkbook.Sheets(8)
                        excelWorkSheet.Name = "Apr " & txt_day.Text & " - May " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 4 And txt_month2.SelectedIndex = 4) Then 'may-may
                        excelWorkSheet = excelWorkbook.Sheets(9)
                        excelWorkSheet.Name = "May " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 4 And txt_month2.SelectedIndex = 5) Then 'may-june
                        excelWorkSheet = excelWorkbook.Sheets(10)
                        excelWorkSheet.Name = "May " & txt_day.Text & " - Jun " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 5 And txt_month2.SelectedIndex = 5) Then 'june-june
                        excelWorkSheet = excelWorkbook.Sheets(11)
                        excelWorkSheet.Name = "Jun " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 5 And txt_month2.SelectedIndex = 6) Then 'june-july
                        excelWorkSheet = excelWorkbook.Sheets(11)
                        excelWorkSheet.Name = "Jun " & txt_day.Text & " - Jul " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 6 And txt_month2.SelectedIndex = 6) Then 'july-july
                        excelWorkSheet = excelWorkbook.Sheets(12)
                        excelWorkSheet.Name = "Jul " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 6 And txt_month2.SelectedIndex = 7) Then 'july-aug
                        excelWorkSheet = excelWorkbook.Sheets(13)
                        excelWorkSheet.Name = "Jul " & txt_day.Text & " - Aug " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 7 And txt_month2.SelectedIndex = 7) Then 'aug-aug
                        excelWorkSheet = excelWorkbook.Sheets(14)
                        excelWorkSheet.Name = "Aug " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 7 And txt_month2.SelectedIndex = 8) Then 'aug-sept
                        excelWorkSheet = excelWorkbook.Sheets(15)
                        excelWorkSheet.Name = "Aug " & txt_day.Text & " - Sept " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 8 And txt_month2.SelectedIndex = 8) Then 'sept-sept
                        excelWorkSheet = excelWorkbook.Sheets(16)
                        excelWorkSheet.Name = "Sept " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 8 And txt_month2.SelectedIndex = 9) Then 'sept-oct
                        excelWorkSheet = excelWorkbook.Sheets(17)
                        excelWorkSheet.Name = "Sept " & txt_day.Text & " - Oct " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 9 And txt_month2.SelectedIndex = 9) Then 'oct-oct
                        excelWorkSheet = excelWorkbook.Sheets(18)
                        excelWorkSheet.Name = "Oct " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 9 And txt_month2.SelectedIndex = 10) Then 'oct-nov
                        excelWorkSheet = excelWorkbook.Sheets(19)
                        excelWorkSheet.Name = "Oct " & txt_day.Text & " - Nov " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 10 And txt_month2.SelectedIndex = 10) Then 'nov-nov
                        excelWorkSheet = excelWorkbook.Sheets(20)
                        excelWorkSheet.Name = "Nov " & txt_day.Text & "-" & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 10 And txt_month2.SelectedIndex = 11) Then 'nov-dec
                        excelWorkSheet = excelWorkbook.Sheets(21)
                        excelWorkSheet.Name = "Nov " & txt_day.Text & " - Dec " & txt_day2.Text
                    ElseIf (txt_month.SelectedIndex = 11 And txt_month2.SelectedIndex = 11) Then 'dec-dec
                        excelWorkSheet = excelWorkbook.Sheets(22)
                        excelWorkSheet.Name = "Dec " & txt_day.Text & "-" & txt_day2.Text
                    End If

                    excelWorkSheet = excelWorkbook.Sheets(excelWorkSheet.Name)
                    excelWorkSheet.Activate()

                    i = maxrow + 6
                    With excelWorkSheet
                        .Cells(3, 2) = excelWorkSheet.Name & " , " & txt_yr.Text
                        .Cells(i, 1) = maxrow
                        .Cells(i, 2) = txt_surname.Text
                        .Cells(i, 3) = txt_name.Text
                        .Cells(i, 4) = txt_s_rate.Text
                        .Cells(i, 5) = txt_s_time_r.Text
                        .Cells(i, 6) = txt_s_salary.Text
                        .Cells(i, 7) = txt_s_ot.Text
                        .Cells(i, 8) = txt_s_ottime.Text & ":" & txt_s_ottime2.Text
                        .Cells(i, 9) = txt_s_adj.Text
                        .Cells(i, 10) = txt_s_waste.Text
                        .Cells(i, 11) = txt_s_tvac.Text
                        .Cells(i, 12) = txt_s_13MP.Text
                        .Cells(i, 13) = txt_s_ecola.Text
                        .Cells(i, 14) = txt_s_night.Text
                        .Cells(i, 15) = txt_s_nightno.Text
                        .Cells(i, 16).Formula = "=SUM(" & .Range(.Cells(i, 6), .Cells(i, 7)).Address(False, False) & "," & .Range(.Cells(i, 9), .Cells(i, 14)).Address(False, False) & ")"
                        .Cells(i, 17) = txt_d_sss.Text
                        .Cells(i, 18) = txt_d_pagibig.Text
                        .Cells(i, 19) = txt_d_philhealth.Text
                        .Cells(i, 20) = txt_d_tax.Text
                        .Cells(i, 21) = txt_d_pagibigloan.Text
                        .Cells(i, 22) = txt_d_sssloan.Text
                        .Cells(i, 23) = txt_d_load.Text
                        .Cells(i, 24) = txt_d_bond.Text
                        .Cells(i, 25) = txt_d_smfi.Text
                        .Cells(i, 26) = txt_d_vale.Text
                        .Cells(i, 27) = txt_d_canteen.Text
                        .Cells(i, 28) = txt_d_agency.Text
                        .Cells(i, 29) = txt_d_undertime.Text
                        .Cells(i, 30).Formula = "=SUM(" & .Range(.Cells(i, 17), .Cells(i, 29)).Address(False, False) & ")"
                        .Cells(i, 31).Formula = "=SUM(" & .Cells(i, 16).Address(False, False) & ",-" & .Cells(i, 30).Address(False, False) & ")"

                        .Cells(i, 2).EntireRow.Font.Bold = False
                        '.Cells(i + 1, 2).ClearContents()
                        .Cells(i + 1, 2) = "TOTAL"
                        .Cells(i + 1, 2).EntireRow.Font.Bold = True
                        For thiscol = 4 To 31
                            .Cells(i + 1, thiscol).Formula = _
                               "=SUM(" & .Range(.Cells(7, thiscol), .Cells(i, thiscol)).Address(False, False) & ")"
                            'excelApp.Sum(.Range(.Cells(2, thiscol), .Cells(i + 2, thiscol)))

                            If (.Cells(i + 1, thiscol).Value = 0) Then
                                .Cells(thiscol).EntireColumn.Hidden = True
                            Else
                                .Cells(thiscol).EntireColumn.Hidden = False
                                .Cells(thiscol).ColumnWidth = 12
                            End If
                        Next

                    End With


                    'TOTAL
                    Dim excelWorkSheet2 As Excel.Worksheet

                    Try
                        excelWorkSheet2 = excelWorkbook.Sheets("SUMMARY")
                        With excelWorkSheet2
                            .Cells(3, 2).Font.Bold = True
                            .Cells(3, 2) = "Jan - Dec , " & txt_yr.Text
                            .Cells(i, 1) = maxrow
                            .Cells(i, 2) = txt_surname.Text
                            .Cells(i, 3) = txt_name.Text
                            .Cells(i, 2).EntireRow.Font.Bold = False
                            .Cells(i + 1, 2) = "TOTAL"
                            .Cells(i + 1, 2).EntireRow.Font.Bold = True
                        End With

                    Catch ex As Exception
                        excelApp.Sheets.Add(After:=excelApp.Sheets(excelApp.Sheets.Count)).Name = "SUMMARY"
                    End Try

                    'excelWorkbook.Close(SaveChanges:=True)
                    excelApp.AlertBeforeOverwriting = False
                    excelApp.DisplayAlerts = False
                    'excelApp.ActiveWorkbook.SaveAs(MainMenu.appath)
                    excelWorkbook.SaveAs(MainMenu.appath)
                    Me.Activate()
                    'excelApp.Quit()

                    MsgBox("Record Sucessfully Saved", MsgBoxStyle.Information, Text)
                    ClearTextBox()
                End If
            Else
                txt_error.Text = "Error: Invalid Range of Date"
            End If 'range of date
        Else
            txt_error.Text = "Error: Incomplete required field/s"
        End If 'for complete the required

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub SendInputs(ByVal noOfIds As String)
        Thread.Sleep(100)
        SendKeys.SendWait(noOfIds)
        SendKeys.SendWait("{ENTER}")
        SendKeys.SendWait("{ENTER}")

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

    Private Sub ClearTextBox()
        For Each ctl In GroupBox3.Controls
            If TypeOf ctl Is TextBox Then ctl.Text = ""
        Next ctl

        For Each ctl In GroupBox4.Controls
            If TypeOf ctl Is TextBox Then ctl.Text = ""
        Next ctl

        txt_name.Text = ""
        txt_surname.Text = ""
    End Sub

    Private Sub txt_s_tvac_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_s_tvac.TextChanged, txt_s_waste.TextChanged, txt_s_adj.TextChanged, _
        txt_s_salary.TextChanged, txt_s_13MP.TextChanged, txt_s_night.TextChanged, txt_s_ecola.TextChanged

        TotalSalary()

    End Sub

    Private Sub ck_night_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ck_night.CheckedChanged, txt_s_nightno.TextChanged, txt_s_night.TextChanged, txt_s_rate.TextChanged

        If (ck_night.CheckState = CheckState.Checked) Then
            txt_s_night.Enabled = True
            txt_s_nightno.Enabled = True
            txt_s_night.Text = Val(txt_s_rate.Text) * 0.1 * Val(txt_s_nightno.Text)
        Else
            txt_s_night.Enabled = False
            txt_s_nightno.Enabled = False
            txt_s_night.Text = 0
            txt_s_nightno.Text = 0
        End If

    End Sub

    Private Sub txt_s_ottime_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_s_ottime.TextChanged
        Checkboxes()
    End Sub

    Private Sub txt_s_ottime2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_s_ottime2.TextChanged
        Checkboxes()
    End Sub

    Private Sub txt_s_ot_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_s_ot.TextChanged
        Checkboxes()
        TotalSalary()
    End Sub

    Private Sub ck_regular_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ck_regular.CheckedChanged
        Checkboxes()
    End Sub

    Private Sub ck_special_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ck_special.CheckedChanged
        Checkboxes()
    End Sub

    Private Sub ck_aft_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ck_aft.CheckedChanged
        Checkboxes()
    End Sub

    Private Sub txt_s_ot_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_s_ot.Leave
        ot = Val(txt_s_ot.Text)
    End Sub

    Private Sub Checkboxes()

        If (txt_s_ottime.Text <> "00" And txt_s_ottime.Text <> "" And txt_s_ottime.Text <> "0" And txt_s_ot.Text <> "0" And txt_s_ot.Text <> "") Then
            If (ck_regular.Checked = True) Then
                'ratephr = (ot * 2.6) * (Val(txt_s_ottime.Text) + (Val(txt_s_ottime2.Text) / 60))
                regular = (txt_s_ottime.Text / 8) + (Val(txt_s_ottime2.Text) / 480)
                ratephr = ot * 1.0 * regular
                txt_s_ot.Text = Math.Round(ratephr, 2)
            ElseIf (ck_special.Checked = True) Then
                'ratephr = (ot * 1.69) * (Val(txt_s_ottime.Text) + (Val(txt_s_ottime2.Text) / 60))
                regular = (txt_s_ottime.Text / 8) + (Val(txt_s_ottime2.Text) / 480)
                ratephr = ot * 0.3 * regular
                txt_s_ot.Text = Math.Round(ratephr, 2)
            ElseIf (ck_aft.Checked = True) Then
                'ratephr = (ot * 1.25) * (Val(txt_s_ottime.Text) + (Val(txt_s_ottime2.Text) / 60))
                ratephr = ot * 0.25 * (Val(txt_s_ottime.Text) + (Val(txt_s_ottime2.Text) / 60))
                txt_s_ot.Text = Math.Round(ratephr, 2)
            End If
        End If
    End Sub

    Private Sub txt_s_rate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_s_rate.TextChanged, txt_s_time_r.TextChanged, txt_s_salary.TextChanged

        If (txt_s_rate.Text <> "0" And txt_s_time_r.Text <> "0" And txt_s_rate.Text <> "" And txt_s_time_r.Text <> "") Then
            txt_s_salary.Text = Val(txt_s_rate.Text) * Val(txt_s_time_r.Text)
        End If

    End Sub

    Private Sub TotalSalary()

        txt_totalsalary.Text = Val(txt_s_tvac.Text) + Val(txt_s_waste.Text) + Val(txt_s_adj.Text) + Val(txt_s_ot.Text) + Val(txt_s_salary.Text) + Val(txt_s_13MP.Text) + Val(txt_s_night.Text) + Val(txt_s_ecola.Text)
        txt_netsalary.Text = Val(txt_totalsalary.Text) - Val(txt_deduction.Text)
    End Sub

    Private Sub txt_d_bond_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_d_bond.TextChanged, txt_d_smfi.TextChanged, txt_d_vale.TextChanged, _
        txt_d_canteen.TextChanged, txt_d_agency.TextChanged, txt_d_undertime.TextChanged, txt_d_sss.TextChanged, txt_d_pagibig.TextChanged, txt_d_philhealth.TextChanged, txt_d_tax.TextChanged, _
        txt_d_pagibigloan.TextChanged, txt_d_sssloan.TextChanged, txt_d_load.TextChanged

        TotalDeduction()

    End Sub

    Private Sub TotalDeduction()
        txt_deduction.Text = Val(txt_d_bond.Text) + Val(txt_d_smfi.Text) + Val(txt_d_vale.Text) + Val(txt_d_canteen.Text) + _
                            Val(txt_d_agency.Text) + Val(txt_d_undertime.Text) + Val(txt_d_sss.Text) + Val(txt_d_pagibig.Text) + _
                            Val(txt_d_philhealth.Text) + Val(txt_d_tax.Text) + Val(txt_d_pagibigloan.Text) + Val(txt_d_sssloan.Text) + Val(txt_d_load.Text)
        txt_netsalary.Text = Val(txt_totalsalary.Text) - Val(txt_deduction.Text)
    End Sub

    Private Sub btn_results_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Results.Show()
    End Sub

    Private Sub btn_all_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_all.Click
        AllResults.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_search.Click
        Results.Show()
        Me.Hide()
    End Sub

    Private Sub btn_print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_print.Click
        'Print.Show()
        PrintDate2.Show()
        Me.Hide()
    End Sub

    Private Sub btn_update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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

    Private Sub btn_search_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_search.MouseEnter
        btn_search.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.1.jpg")
    End Sub

    Private Sub btn_search_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_search.MouseLeave
        btn_search.BackgroundImage = System.Drawing.Image.FromFile(directory & "\3.jpg")
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
            ExcelProcessInit()
            ExcelProcessKill()
        End If
    End Sub

    Private Sub btn_minimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_minimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub btn_clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_clear.Click
        If MsgBox("Are you sure you want to erase all the records?", MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, "Close application") = Windows.Forms.DialogResult.Yes Then
            conn.Open()
            cmd2 = New OleDbCommand("Delete from tbl_payslip", conn)
            cmd2.ExecuteNonQuery()
            conn.Close()
        End If
    End Sub

    Private mExcelProcesses() As Process

    Private Sub ExcelProcessInit()
        Try
            'Get all currently running process Ids for Excel applications
            mExcelProcesses = Process.GetProcessesByName("Excel")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ExcelProcessKill()
        Dim oProcesses() As Process
        Dim bFound As Boolean

        Try
            'Get all currently running process Ids for Excel applications
            oProcesses = Process.GetProcessesByName("Excel")

            If oProcesses.Length > 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    bFound = False

                    For j As Integer = 0 To mExcelProcesses.Length - 1
                        If oProcesses(i).Id = mExcelProcesses(j).Id Then
                            bFound = True
                            Exit For
                        End If
                    Next

                    If Not bFound Then
                        oProcesses(i).Kill()
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Number_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_day.KeyPress, txt_day2.KeyPress, _
        txt_yr.KeyPress, txt_s_rate.KeyPress, txt_s_13MP.KeyPress, txt_s_adj.KeyPress, txt_s_ecola.KeyPress, txt_s_night.KeyPress, txt_s_nightno.KeyPress, _
        txt_s_ot.KeyPress, txt_s_ottime.KeyPress, txt_s_ottime2.KeyPress, txt_s_rate.KeyPress, txt_s_time_r.KeyPress, txt_s_tvac.KeyPress, txt_s_waste.KeyPress, txt_s_salary.KeyPress, _
        txt_d_agency.KeyPress, txt_d_bond.KeyPress, txt_d_canteen.KeyPress, txt_d_load.KeyPress, txt_d_pagibig.KeyPress, txt_d_pagibigloan.KeyPress, txt_d_philhealth.KeyPress, _
        txt_d_smfi.KeyPress, txt_d_sss.KeyPress, txt_d_sssloan.KeyPress, txt_d_tax.KeyPress, txt_d_undertime.KeyPress, txt_d_vale.KeyPress

        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers
        If Asc(e.KeyChar) <> 8 Then 'Backspace
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                If e.KeyChar = "." Then

                Else
                    e.Handled = True
                End If

            End If
            End If
    End Sub

End Class

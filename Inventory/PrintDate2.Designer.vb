<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintDate2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintDate2))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Combo_Date = New System.Windows.Forms.ComboBox()
        Me.btn_ok = New System.Windows.Forms.Button()
        Me.txt_search = New System.Windows.Forms.TextBox()
        Me.btn_searh = New System.Windows.Forms.Button()
        Me.btn_addnew = New System.Windows.Forms.Button()
        Me.btn_print = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btn_close = New System.Windows.Forms.Button()
        Me.btn_minimize = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.Combo_Date)
        Me.GroupBox1.Controls.Add(Me.btn_ok)
        Me.GroupBox1.Controls.Add(Me.txt_search)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 11)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(591, 65)
        Me.GroupBox1.TabIndex = 95
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Search"
        '
        'Combo_Date
        '
        Me.Combo_Date.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Date.FormattingEnabled = True
        Me.Combo_Date.Items.AddRange(New Object() {"Dec - Jan", "Jan", "Jan - Feb", "Feb", "Feb - Mar", "Mar", "Mar - Apr", "Apr", "Apr - May", "May", "May - Jun", "Jun", "Jun - Jul", "Jul", "Jul - Aug", "Aug", "Aug - Sept", "Sept", "Sept - Oct", "Oct", "Oct - Nov", "Nov", "Nov - Dec", "Dec"})
        Me.Combo_Date.Location = New System.Drawing.Point(132, 22)
        Me.Combo_Date.Name = "Combo_Date"
        Me.Combo_Date.Size = New System.Drawing.Size(216, 21)
        Me.Combo_Date.TabIndex = 97
        '
        'btn_ok
        '
        Me.btn_ok.Location = New System.Drawing.Point(403, 22)
        Me.btn_ok.Name = "btn_ok"
        Me.btn_ok.Size = New System.Drawing.Size(97, 21)
        Me.btn_ok.TabIndex = 89
        Me.btn_ok.Text = "Generate"
        Me.btn_ok.UseVisualStyleBackColor = True
        '
        'txt_search
        '
        Me.txt_search.Location = New System.Drawing.Point(132, 22)
        Me.txt_search.Name = "txt_search"
        Me.txt_search.Size = New System.Drawing.Size(154, 20)
        Me.txt_search.TabIndex = 87
        '
        'btn_searh
        '
        Me.btn_searh.BackColor = System.Drawing.Color.WhiteSmoke
        Me.btn_searh.BackgroundImage = CType(resources.GetObject("btn_searh.BackgroundImage"), System.Drawing.Image)
        Me.btn_searh.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_searh.ForeColor = System.Drawing.Color.Black
        Me.btn_searh.Location = New System.Drawing.Point(15, 134)
        Me.btn_searh.Name = "btn_searh"
        Me.btn_searh.Size = New System.Drawing.Size(115, 65)
        Me.btn_searh.TabIndex = 98
        Me.btn_searh.Text = "Search"
        Me.btn_searh.UseVisualStyleBackColor = False
        '
        'btn_addnew
        '
        Me.btn_addnew.BackColor = System.Drawing.Color.WhiteSmoke
        Me.btn_addnew.BackgroundImage = CType(resources.GetObject("btn_addnew.BackgroundImage"), System.Drawing.Image)
        Me.btn_addnew.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn_addnew.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_addnew.ForeColor = System.Drawing.Color.Black
        Me.btn_addnew.Location = New System.Drawing.Point(15, 205)
        Me.btn_addnew.Name = "btn_addnew"
        Me.btn_addnew.Size = New System.Drawing.Size(115, 65)
        Me.btn_addnew.TabIndex = 97
        Me.btn_addnew.Text = "Add New Record"
        Me.btn_addnew.UseVisualStyleBackColor = False
        '
        'btn_print
        '
        Me.btn_print.BackColor = System.Drawing.Color.WhiteSmoke
        Me.btn_print.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btn_print.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_print.ForeColor = System.Drawing.Color.Black
        Me.btn_print.Location = New System.Drawing.Point(15, 275)
        Me.btn_print.Name = "btn_print"
        Me.btn_print.Size = New System.Drawing.Size(115, 65)
        Me.btn_print.TabIndex = 99
        Me.btn_print.Text = "Generate PaySlip"
        Me.btn_print.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.GhostWhite
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Location = New System.Drawing.Point(126, 51)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(610, 568)
        Me.Panel2.TabIndex = 111
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Segoe Script", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(8, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(147, 57)
        Me.Label1.TabIndex = 112
        Me.Label1.Text = "Search"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Controls.Add(Me.btn_close)
        Me.Panel1.Controls.Add(Me.btn_minimize)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.btn_searh)
        Me.Panel1.Controls.Add(Me.btn_addnew)
        Me.Panel1.Controls.Add(Me.btn_print)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(754, 638)
        Me.Panel1.TabIndex = 113
        '
        'btn_close
        '
        Me.btn_close.BackColor = System.Drawing.Color.White
        Me.btn_close.BackgroundImage = CType(resources.GetObject("btn_close.BackgroundImage"), System.Drawing.Image)
        Me.btn_close.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btn_close.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_close.Font = New System.Drawing.Font("Elephant", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_close.ForeColor = System.Drawing.Color.White
        Me.btn_close.Location = New System.Drawing.Point(731, 0)
        Me.btn_close.Name = "btn_close"
        Me.btn_close.Size = New System.Drawing.Size(23, 21)
        Me.btn_close.TabIndex = 114
        Me.btn_close.Text = "__"
        Me.btn_close.UseVisualStyleBackColor = False
        '
        'btn_minimize
        '
        Me.btn_minimize.BackColor = System.Drawing.Color.White
        Me.btn_minimize.BackgroundImage = CType(resources.GetObject("btn_minimize.BackgroundImage"), System.Drawing.Image)
        Me.btn_minimize.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btn_minimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_minimize.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_minimize.ForeColor = System.Drawing.Color.White
        Me.btn_minimize.Location = New System.Drawing.Point(706, 0)
        Me.btn_minimize.Margin = New System.Windows.Forms.Padding(0)
        Me.btn_minimize.Name = "btn_minimize"
        Me.btn_minimize.Size = New System.Drawing.Size(23, 21)
        Me.btn_minimize.TabIndex = 113
        Me.btn_minimize.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_minimize.UseVisualStyleBackColor = False
        '
        'PrintDate2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(762, 646)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(762, 646)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(762, 646)
        Me.Name = "PrintDate2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Results"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btn_ok As System.Windows.Forms.Button
    Friend WithEvents txt_search As System.Windows.Forms.TextBox
    Friend WithEvents btn_searh As System.Windows.Forms.Button
    Friend WithEvents btn_addnew As System.Windows.Forms.Button
    Friend WithEvents btn_print As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btn_close As System.Windows.Forms.Button
    Friend WithEvents btn_minimize As System.Windows.Forms.Button
    Friend WithEvents Combo_Date As System.Windows.Forms.ComboBox
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenu
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainMenu))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_filename = New System.Windows.Forms.Label()
        Me.btn_excel = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btn_close = New System.Windows.Forms.Button()
        Me.btn_minimize = New System.Windows.Forms.Button()
        Me.btn_search = New System.Windows.Forms.Button()
        Me.btn_addnew = New System.Windows.Forms.Button()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Pristina", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(180, 160)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(207, 75)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Generator"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txt_filename
        '
        Me.txt_filename.BackColor = System.Drawing.Color.Transparent
        Me.txt_filename.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_filename.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.txt_filename.Location = New System.Drawing.Point(29, 421)
        Me.txt_filename.Name = "txt_filename"
        Me.txt_filename.Size = New System.Drawing.Size(140, 22)
        Me.txt_filename.TabIndex = 3
        Me.txt_filename.Text = "Filename"
        '
        'btn_excel
        '
        Me.btn_excel.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.btn_excel.BackgroundImage = CType(resources.GetObject("btn_excel.BackgroundImage"), System.Drawing.Image)
        Me.btn_excel.Font = New System.Drawing.Font("Cambria", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_excel.Location = New System.Drawing.Point(32, 376)
        Me.btn_excel.Name = "btn_excel"
        Me.btn_excel.Size = New System.Drawing.Size(137, 43)
        Me.btn_excel.TabIndex = 2
        Me.btn_excel.Text = "Choose Excel File"
        Me.btn_excel.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Panel2.BackgroundImage = CType(resources.GetObject("Panel2.BackgroundImage"), System.Drawing.Image)
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.btn_close)
        Me.Panel2.Controls.Add(Me.btn_minimize)
        Me.Panel2.Controls.Add(Me.txt_filename)
        Me.Panel2.Controls.Add(Me.btn_excel)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.btn_search)
        Me.Panel2.Controls.Add(Me.btn_addnew)
        Me.Panel2.Location = New System.Drawing.Point(4, 4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(612, 445)
        Me.Panel2.TabIndex = 3
        '
        'btn_close
        '
        Me.btn_close.BackColor = System.Drawing.Color.White
        Me.btn_close.BackgroundImage = CType(resources.GetObject("btn_close.BackgroundImage"), System.Drawing.Image)
        Me.btn_close.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btn_close.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_close.Font = New System.Drawing.Font("Elephant", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_close.ForeColor = System.Drawing.Color.White
        Me.btn_close.Location = New System.Drawing.Point(589, 0)
        Me.btn_close.Name = "btn_close"
        Me.btn_close.Size = New System.Drawing.Size(23, 21)
        Me.btn_close.TabIndex = 6
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
        Me.btn_minimize.Location = New System.Drawing.Point(564, 0)
        Me.btn_minimize.Margin = New System.Windows.Forms.Padding(0)
        Me.btn_minimize.Name = "btn_minimize"
        Me.btn_minimize.Size = New System.Drawing.Size(23, 21)
        Me.btn_minimize.TabIndex = 5
        Me.btn_minimize.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_minimize.UseVisualStyleBackColor = False
        '
        'btn_search
        '
        Me.btn_search.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.btn_search.BackgroundImage = CType(resources.GetObject("btn_search.BackgroundImage"), System.Drawing.Image)
        Me.btn_search.Font = New System.Drawing.Font("Cambria", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_search.Location = New System.Drawing.Point(308, 376)
        Me.btn_search.Name = "btn_search"
        Me.btn_search.Size = New System.Drawing.Size(86, 43)
        Me.btn_search.TabIndex = 1
        Me.btn_search.Text = "Search"
        Me.btn_search.UseVisualStyleBackColor = False
        '
        'btn_addnew
        '
        Me.btn_addnew.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.btn_addnew.BackgroundImage = CType(resources.GetObject("btn_addnew.BackgroundImage"), System.Drawing.Image)
        Me.btn_addnew.Font = New System.Drawing.Font("Cambria", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_addnew.Location = New System.Drawing.Point(216, 376)
        Me.btn_addnew.Name = "btn_addnew"
        Me.btn_addnew.Size = New System.Drawing.Size(86, 43)
        Me.btn_addnew.TabIndex = 0
        Me.btn_addnew.Text = "Add New Record"
        Me.btn_addnew.UseVisualStyleBackColor = False
        '
        'MainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(620, 453)
        Me.Controls.Add(Me.Panel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MainMenu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Menu"
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btn_excel As System.Windows.Forms.Button
    Friend WithEvents txt_filename As System.Windows.Forms.Label
    Friend WithEvents btn_search As System.Windows.Forms.Button
    Friend WithEvents btn_addnew As System.Windows.Forms.Button
    Friend WithEvents btn_close As System.Windows.Forms.Button
    Friend WithEvents btn_minimize As System.Windows.Forms.Button
End Class

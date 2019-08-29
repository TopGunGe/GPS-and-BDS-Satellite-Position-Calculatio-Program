<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.SaveFileDialog2 = New System.Windows.Forms.SaveFileDialog()
        Me.RichTextBox3 = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.SaveFileDialog3 = New System.Windows.Forms.SaveFileDialog()
        Me.SaveFileDialog4 = New System.Windows.Forms.SaveFileDialog()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.文件ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.打开ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Rinex210ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Rinex211ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Rinex302ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.手动输入卫星参数ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TESTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.保存ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveCsAsText = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveCsAsXlsm = New System.Windows.Forms.ToolStripMenuItem()
        Me.保存结果ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveRuAsText = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveRuAsXlsm = New System.Windows.Forms.ToolStripMenuItem()
        Me.退出ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.数据ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.解算历书文件ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.帮助ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RichTextBox4 = New System.Windows.Forms.RichTextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.Rinex210ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Rinex211ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Rinex302ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.解算历书文件ToolStripMenuItem1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
        Me.手动输入卫星参数ToolStripMenuItem1 = New System.Windows.Forms.ToolStripButton()
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.OpenFileDialog3 = New System.Windows.Forms.OpenFileDialog()
        Me.MenuStrip1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.White
        Me.RichTextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.RichTextBox1.Font = New System.Drawing.Font("黑体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RichTextBox1.ForeColor = System.Drawing.Color.Black
        Me.RichTextBox1.Location = New System.Drawing.Point(12, 63)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(363, 502)
        Me.RichTextBox1.TabIndex = 5
        Me.RichTextBox1.Text = "文件数据：" & Global.Microsoft.VisualBasic.ChrW(10)
        Me.RichTextBox1.WordWrap = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "历书文件"
        '
        'SaveFileDialog1
        '
        '
        'SaveFileDialog2
        '
        '
        'RichTextBox3
        '
        Me.RichTextBox3.BackColor = System.Drawing.Color.White
        Me.RichTextBox3.ForeColor = System.Drawing.Color.Black
        Me.RichTextBox3.Location = New System.Drawing.Point(752, 63)
        Me.RichTextBox3.Name = "RichTextBox3"
        Me.RichTextBox3.ReadOnly = True
        Me.RichTextBox3.Size = New System.Drawing.Size(363, 502)
        Me.RichTextBox3.TabIndex = 7
        Me.RichTextBox3.Text = "计算结果：" & Global.Microsoft.VisualBasic.ChrW(10)
        Me.RichTextBox3.WordWrap = False
        '
        'RichTextBox2
        '
        Me.RichTextBox2.BackColor = System.Drawing.Color.White
        Me.RichTextBox2.ForeColor = System.Drawing.Color.Black
        Me.RichTextBox2.Location = New System.Drawing.Point(381, 63)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.ReadOnly = True
        Me.RichTextBox2.Size = New System.Drawing.Size(363, 502)
        Me.RichTextBox2.TabIndex = 6
        Me.RichTextBox2.Text = "卫星参数：" & Global.Microsoft.VisualBasic.ChrW(10)
        Me.RichTextBox2.WordWrap = False
        '
        'SaveFileDialog3
        '
        '
        'SaveFileDialog4
        '
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.文件ToolStripMenuItem, Me.数据ToolStripMenuItem, Me.帮助ToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(3, 3)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1112, 24)
        Me.MenuStrip1.TabIndex = 14
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        '文件ToolStripMenuItem
        '
        Me.文件ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.打开ToolStripMenuItem, Me.保存ToolStripMenuItem, Me.保存结果ToolStripMenuItem, Me.退出ToolStripMenuItem})
        Me.文件ToolStripMenuItem.Font = New System.Drawing.Font("黑体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.文件ToolStripMenuItem.Name = "文件ToolStripMenuItem"
        Me.文件ToolStripMenuItem.Size = New System.Drawing.Size(53, 20)
        Me.文件ToolStripMenuItem.Text = "文件"
        '
        '打开ToolStripMenuItem
        '
        Me.打开ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Rinex210ToolStripMenuItem, Me.Rinex211ToolStripMenuItem, Me.Rinex302ToolStripMenuItem, Me.手动输入卫星参数ToolStripMenuItem, Me.TESTToolStripMenuItem})
        Me.打开ToolStripMenuItem.Name = "打开ToolStripMenuItem"
        Me.打开ToolStripMenuItem.Size = New System.Drawing.Size(184, 26)
        Me.打开ToolStripMenuItem.Text = "打开历书文件"
        '
        'Rinex210ToolStripMenuItem
        '
        Me.Rinex210ToolStripMenuItem.Name = "Rinex210ToolStripMenuItem"
        Me.Rinex210ToolStripMenuItem.Size = New System.Drawing.Size(216, 26)
        Me.Rinex210ToolStripMenuItem.Text = "Rinex2.10"
        '
        'Rinex211ToolStripMenuItem
        '
        Me.Rinex211ToolStripMenuItem.Name = "Rinex211ToolStripMenuItem"
        Me.Rinex211ToolStripMenuItem.Size = New System.Drawing.Size(216, 26)
        Me.Rinex211ToolStripMenuItem.Text = "Rinex2.11"
        '
        'Rinex302ToolStripMenuItem
        '
        Me.Rinex302ToolStripMenuItem.Name = "Rinex302ToolStripMenuItem"
        Me.Rinex302ToolStripMenuItem.Size = New System.Drawing.Size(216, 26)
        Me.Rinex302ToolStripMenuItem.Text = "Rinex3.02"
        '
        '手动输入卫星参数ToolStripMenuItem
        '
        Me.手动输入卫星参数ToolStripMenuItem.Name = "手动输入卫星参数ToolStripMenuItem"
        Me.手动输入卫星参数ToolStripMenuItem.Size = New System.Drawing.Size(216, 26)
        Me.手动输入卫星参数ToolStripMenuItem.Text = "手动输入卫星参数"
        '
        'TESTToolStripMenuItem
        '
        Me.TESTToolStripMenuItem.Name = "TESTToolStripMenuItem"
        Me.TESTToolStripMenuItem.Size = New System.Drawing.Size(216, 26)
        Me.TESTToolStripMenuItem.Text = "TEST"
        '
        '保存ToolStripMenuItem
        '
        Me.保存ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveCsAsText, Me.SaveCsAsXlsm})
        Me.保存ToolStripMenuItem.Name = "保存ToolStripMenuItem"
        Me.保存ToolStripMenuItem.Size = New System.Drawing.Size(184, 26)
        Me.保存ToolStripMenuItem.Text = "保存参数为"
        '
        'SaveCsAsText
        '
        Me.SaveCsAsText.Name = "SaveCsAsText"
        Me.SaveCsAsText.Size = New System.Drawing.Size(120, 26)
        Me.SaveCsAsText.Text = "文本"
        '
        'SaveCsAsXlsm
        '
        Me.SaveCsAsXlsm.Name = "SaveCsAsXlsm"
        Me.SaveCsAsXlsm.Size = New System.Drawing.Size(120, 26)
        Me.SaveCsAsXlsm.Text = "表格"
        '
        '保存结果ToolStripMenuItem
        '
        Me.保存结果ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveRuAsText, Me.SaveRuAsXlsm})
        Me.保存结果ToolStripMenuItem.Name = "保存结果ToolStripMenuItem"
        Me.保存结果ToolStripMenuItem.Size = New System.Drawing.Size(184, 26)
        Me.保存结果ToolStripMenuItem.Text = "保存结果为"
        '
        'SaveRuAsText
        '
        Me.SaveRuAsText.Name = "SaveRuAsText"
        Me.SaveRuAsText.Size = New System.Drawing.Size(120, 26)
        Me.SaveRuAsText.Text = "文本"
        '
        'SaveRuAsXlsm
        '
        Me.SaveRuAsXlsm.Name = "SaveRuAsXlsm"
        Me.SaveRuAsXlsm.Size = New System.Drawing.Size(120, 26)
        Me.SaveRuAsXlsm.Text = "表格"
        '
        '退出ToolStripMenuItem
        '
        Me.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem"
        Me.退出ToolStripMenuItem.Size = New System.Drawing.Size(184, 26)
        Me.退出ToolStripMenuItem.Text = "退出"
        '
        '数据ToolStripMenuItem
        '
        Me.数据ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.解算历书文件ToolStripMenuItem})
        Me.数据ToolStripMenuItem.Font = New System.Drawing.Font("黑体", 9.0!)
        Me.数据ToolStripMenuItem.Name = "数据ToolStripMenuItem"
        Me.数据ToolStripMenuItem.Size = New System.Drawing.Size(53, 20)
        Me.数据ToolStripMenuItem.Text = "数据"
        '
        '解算历书文件ToolStripMenuItem
        '
        Me.解算历书文件ToolStripMenuItem.Name = "解算历书文件ToolStripMenuItem"
        Me.解算历书文件ToolStripMenuItem.Size = New System.Drawing.Size(224, 26)
        Me.解算历书文件ToolStripMenuItem.Text = "解算历书文件"
        '
        '帮助ToolStripMenuItem
        '
        Me.帮助ToolStripMenuItem.Font = New System.Drawing.Font("黑体", 9.0!)
        Me.帮助ToolStripMenuItem.Name = "帮助ToolStripMenuItem"
        Me.帮助ToolStripMenuItem.Size = New System.Drawing.Size(53, 20)
        Me.帮助ToolStripMenuItem.Text = "帮助"
        '
        'RichTextBox4
        '
        Me.RichTextBox4.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.RichTextBox4.Location = New System.Drawing.Point(12, 571)
        Me.RichTextBox4.Name = "RichTextBox4"
        Me.RichTextBox4.Size = New System.Drawing.Size(1101, 142)
        Me.RichTextBox4.TabIndex = 15
        Me.RichTextBox4.Text = ""
        '
        'ToolStrip1
        '
        Me.ToolStrip1.BackColor = System.Drawing.Color.FromArgb(CType(CType(233, Byte), Integer), CType(CType(233, Byte), Integer), CType(CType(233, Byte), Integer))
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator1, Me.ToolStripDropDownButton1, Me.解算历书文件ToolStripMenuItem1, Me.ToolStripProgressBar1, Me.ToolStripButton4, Me.ToolStripSeparator2, Me.ToolStripTextBox1, Me.手动输入卫星参数ToolStripMenuItem1})
        Me.ToolStrip1.Location = New System.Drawing.Point(3, 27)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1112, 27)
        Me.ToolStrip1.TabIndex = 18
        Me.ToolStrip1.TabStop = True
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 27)
        '
        'ToolStripDropDownButton1
        '
        Me.ToolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripDropDownButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Rinex210ToolStripMenuItem1, Me.Rinex211ToolStripMenuItem1, Me.Rinex302ToolStripMenuItem1})
        Me.ToolStripDropDownButton1.Image = CType(resources.GetObject("ToolStripDropDownButton1.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton1.Name = "ToolStripDropDownButton1"
        Me.ToolStripDropDownButton1.Size = New System.Drawing.Size(34, 24)
        Me.ToolStripDropDownButton1.Text = "打开文件"
        '
        'Rinex210ToolStripMenuItem1
        '
        Me.Rinex210ToolStripMenuItem1.Image = CType(resources.GetObject("Rinex210ToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.Rinex210ToolStripMenuItem1.Name = "Rinex210ToolStripMenuItem1"
        Me.Rinex210ToolStripMenuItem1.Size = New System.Drawing.Size(163, 26)
        Me.Rinex210ToolStripMenuItem1.Text = "Rinex2.10"
        '
        'Rinex211ToolStripMenuItem1
        '
        Me.Rinex211ToolStripMenuItem1.Image = CType(resources.GetObject("Rinex211ToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.Rinex211ToolStripMenuItem1.Name = "Rinex211ToolStripMenuItem1"
        Me.Rinex211ToolStripMenuItem1.Size = New System.Drawing.Size(163, 26)
        Me.Rinex211ToolStripMenuItem1.Text = "Rinex2.11"
        '
        'Rinex302ToolStripMenuItem1
        '
        Me.Rinex302ToolStripMenuItem1.Image = CType(resources.GetObject("Rinex302ToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.Rinex302ToolStripMenuItem1.Name = "Rinex302ToolStripMenuItem1"
        Me.Rinex302ToolStripMenuItem1.Size = New System.Drawing.Size(163, 26)
        Me.Rinex302ToolStripMenuItem1.Text = "Rinex3.02"
        '
        '解算历书文件ToolStripMenuItem1
        '
        Me.解算历书文件ToolStripMenuItem1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.解算历书文件ToolStripMenuItem1.Image = CType(resources.GetObject("解算历书文件ToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.解算历书文件ToolStripMenuItem1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.解算历书文件ToolStripMenuItem1.Name = "解算历书文件ToolStripMenuItem1"
        Me.解算历书文件ToolStripMenuItem1.Size = New System.Drawing.Size(29, 24)
        Me.解算历书文件ToolStripMenuItem1.Text = "开始解算"
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripProgressBar1.Font = New System.Drawing.Font("黑体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 24)
        '
        'ToolStripButton4
        '
        Me.ToolStripButton4.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton4.Image = CType(resources.GetObject("ToolStripButton4.Image"), System.Drawing.Image)
        Me.ToolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.Size = New System.Drawing.Size(29, 24)
        Me.ToolStripButton4.Text = "清除屏幕"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 27)
        '
        'ToolStripTextBox1
        '
        Me.ToolStripTextBox1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
        Me.ToolStripTextBox1.Size = New System.Drawing.Size(100, 27)
        '
        '手动输入卫星参数ToolStripMenuItem1
        '
        Me.手动输入卫星参数ToolStripMenuItem1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.手动输入卫星参数ToolStripMenuItem1.Image = CType(resources.GetObject("手动输入卫星参数ToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.手动输入卫星参数ToolStripMenuItem1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.手动输入卫星参数ToolStripMenuItem1.Name = "手动输入卫星参数ToolStripMenuItem1"
        Me.手动输入卫星参数ToolStripMenuItem1.Size = New System.Drawing.Size(29, 24)
        Me.手动输入卫星参数ToolStripMenuItem1.Text = "手动输入参数"
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.FileName = "OpenFileDialog2"
        '
        'OpenFileDialog3
        '
        Me.OpenFileDialog3.FileName = "OpenFileDialog3"
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(233, Byte), Integer), CType(CType(233, Byte), Integer), CType(CType(233, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1118, 722)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.RichTextBox4)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.RichTextBox3)
        Me.Controls.Add(Me.RichTextBox2)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("黑体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "Form2"
        Me.Padding = New System.Windows.Forms.Padding(3)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "GPS、BDS卫星位置计算"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents SaveFileDialog2 As SaveFileDialog
    Friend WithEvents RichTextBox3 As RichTextBox
    Friend WithEvents RichTextBox2 As RichTextBox
    Friend WithEvents SaveFileDialog3 As SaveFileDialog
    Friend WithEvents SaveFileDialog4 As SaveFileDialog
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents 文件ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 打开ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 保存ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 保存结果ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 帮助ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveCsAsText As ToolStripMenuItem
    Friend WithEvents SaveCsAsXlsm As ToolStripMenuItem
    Friend WithEvents SaveRuAsText As ToolStripMenuItem
    Friend WithEvents SaveRuAsXlsm As ToolStripMenuItem
    Friend WithEvents 退出ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Rinex210ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Rinex302ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 数据ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 解算历书文件ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Rinex211ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 手动输入卫星参数ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RichTextBox4 As RichTextBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents 手动输入卫星参数ToolStripMenuItem1 As ToolStripButton
    Friend WithEvents ToolStripProgressBar1 As ToolStripProgressBar
    Friend WithEvents 解算历书文件ToolStripMenuItem1 As ToolStripButton
    Friend WithEvents ToolStripDropDownButton1 As ToolStripDropDownButton
    Friend WithEvents Rinex210ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents Rinex211ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents Rinex302ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripButton4 As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ToolStripTextBox1 As ToolStripTextBox
    Friend WithEvents OpenFileDialog2 As OpenFileDialog
    Friend WithEvents OpenFileDialog3 As OpenFileDialog
    Friend WithEvents TESTToolStripMenuItem As ToolStripMenuItem
End Class

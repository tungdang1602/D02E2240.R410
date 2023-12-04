<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F0333
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F0333))
        Me.grpOri = New System.Windows.Forms.GroupBox
        Me.optBAL = New System.Windows.Forms.RadioButton
        Me.optCip = New System.Windows.Forms.RadioButton
        Me.optNew = New System.Windows.Forms.RadioButton
        Me.optAll = New System.Windows.Forms.RadioButton
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnNext = New System.Windows.Forms.Button
        Me.optCAP = New System.Windows.Forms.RadioButton
        Me.grpOri.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpOri
        '
        Me.grpOri.Controls.Add(Me.optBAL)
        Me.grpOri.Controls.Add(Me.optCip)
        Me.grpOri.Controls.Add(Me.optNew)
        Me.grpOri.Controls.Add(Me.optAll)
        Me.grpOri.Location = New System.Drawing.Point(12, 12)
        Me.grpOri.Name = "grpOri"
        Me.grpOri.Size = New System.Drawing.Size(385, 197)
        Me.grpOri.TabIndex = 0
        Me.grpOri.TabStop = False
        Me.grpOri.Text = "Nguồn gốc hình thành TSCĐ"
        '
        'optBAL
        '
        Me.optBAL.AutoSize = True
        Me.optBAL.Location = New System.Drawing.Point(46, 135)
        Me.optBAL.Name = "optBAL"
        Me.optBAL.Size = New System.Drawing.Size(80, 17)
        Me.optBAL.TabIndex = 3
        Me.optBAL.Text = "Nhập số dư"
        Me.optBAL.UseVisualStyleBackColor = True
        '
        'optCip
        '
        Me.optCip.AutoSize = True
        Me.optCip.Location = New System.Drawing.Point(46, 101)
        Me.optCip.Name = "optCip"
        Me.optCip.Size = New System.Drawing.Size(120, 17)
        Me.optCip.TabIndex = 2
        Me.optCip.Text = "Từ xây dựng cơ bản"
        Me.optCip.UseVisualStyleBackColor = True
        '
        'optNew
        '
        Me.optNew.AutoSize = True
        Me.optNew.Location = New System.Drawing.Point(46, 66)
        Me.optNew.Name = "optNew"
        Me.optNew.Size = New System.Drawing.Size(65, 17)
        Me.optNew.TabIndex = 1
        Me.optNew.Text = "Mua mới"
        Me.optNew.UseVisualStyleBackColor = True
        '
        'optAll
        '
        Me.optAll.AutoSize = True
        Me.optAll.Checked = True
        Me.optAll.Location = New System.Drawing.Point(46, 34)
        Me.optAll.Name = "optAll"
        Me.optAll.Size = New System.Drawing.Size(56, 17)
        Me.optAll.TabIndex = 0
        Me.optAll.TabStop = True
        Me.optAll.Text = "Tất cả"
        Me.optAll.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(321, 215)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 22)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Đó&ng"
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(239, 215)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(76, 22)
        Me.btnNext.TabIndex = 1
        Me.btnNext.Text = "Tiếp tục"
        '
        'optCAP
        '
        Me.optCAP.AutoSize = True
        Me.optCAP.Location = New System.Drawing.Point(58, 180)
        Me.optCAP.Name = "optCAP"
        Me.optCAP.Size = New System.Drawing.Size(96, 17)
        Me.optCAP.TabIndex = 4
        Me.optCAP.TabStop = True
        Me.optCAP.Text = "Điều động vốn"
        Me.optCAP.UseVisualStyleBackColor = True
        '
        'D02F0333
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(409, 245)
        Me.Controls.Add(Me.optCAP)
        Me.Controls.Add(Me.grpOri)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnNext)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F0333"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "HØnh thªnh TSC˜ - D02F0333"
        Me.grpOri.ResumeLayout(False)
        Me.grpOri.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents grpOri As System.Windows.Forms.GroupBox
    Private WithEvents optCip As System.Windows.Forms.RadioButton
    Private WithEvents optNew As System.Windows.Forms.RadioButton
    Private WithEvents optAll As System.Windows.Forms.RadioButton
    Private WithEvents optBAL As System.Windows.Forms.RadioButton
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents btnNext As System.Windows.Forms.Button
    Private WithEvents optCAP As System.Windows.Forms.RadioButton
End Class

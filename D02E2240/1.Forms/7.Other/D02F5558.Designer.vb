<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F5558
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F5558))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtVoucherNoNew = New System.Windows.Forms.TextBox
        Me.txtVoucherNoOld = New System.Windows.Forms.TextBox
        Me.lblVoucherNoOld = New System.Windows.Forms.Label
        Me.lblVoucherNoNew = New System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtVoucherNoNew)
        Me.GroupBox1.Controls.Add(Me.txtVoucherNoOld)
        Me.GroupBox1.Controls.Add(Me.lblVoucherNoOld)
        Me.GroupBox1.Controls.Add(Me.lblVoucherNoNew)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(327, 78)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'txtVoucherNoNew
        '
        Me.txtVoucherNoNew.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtVoucherNoNew.Font = New System.Drawing.Font("Lemon3", 8.249999!)
        Me.txtVoucherNoNew.Location = New System.Drawing.Point(113, 44)
        Me.txtVoucherNoNew.MaxLength = 20
        Me.txtVoucherNoNew.Name = "txtVoucherNoNew"
        Me.txtVoucherNoNew.Size = New System.Drawing.Size(201, 22)
        Me.txtVoucherNoNew.TabIndex = 3
        '
        'txtVoucherNoOld
        '
        Me.txtVoucherNoOld.Font = New System.Drawing.Font("Lemon3", 8.249999!)
        Me.txtVoucherNoOld.Location = New System.Drawing.Point(113, 15)
        Me.txtVoucherNoOld.MaxLength = 20
        Me.txtVoucherNoOld.Name = "txtVoucherNoOld"
        Me.txtVoucherNoOld.Size = New System.Drawing.Size(201, 22)
        Me.txtVoucherNoOld.TabIndex = 1
        '
        'lblVoucherNoOld
        '
        Me.lblVoucherNoOld.AutoSize = True
        Me.lblVoucherNoOld.Location = New System.Drawing.Point(7, 20)
        Me.lblVoucherNoOld.Name = "lblVoucherNoOld"
        Me.lblVoucherNoOld.Size = New System.Drawing.Size(70, 13)
        Me.lblVoucherNoOld.TabIndex = 0
        Me.lblVoucherNoOld.Text = "Số phiếu gốc"
        Me.lblVoucherNoOld.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblVoucherNoNew
        '
        Me.lblVoucherNoNew.AutoSize = True
        Me.lblVoucherNoNew.Location = New System.Drawing.Point(7, 49)
        Me.lblVoucherNoNew.Name = "lblVoucherNoNew"
        Me.lblVoucherNoNew.Size = New System.Drawing.Size(68, 13)
        Me.lblVoucherNoNew.TabIndex = 2
        Me.lblVoucherNoNew.Text = "Số phiếu mới"
        Me.lblVoucherNoNew.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(176, 90)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(76, 22)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "&Lưu"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(259, 90)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 22)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'D02F5558
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(342, 122)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F5558"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sõa sç phiÕu - D02F5558"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents txtVoucherNoNew As System.Windows.Forms.TextBox
    Private WithEvents txtVoucherNoOld As System.Windows.Forms.TextBox
    Private WithEvents lblVoucherNoOld As System.Windows.Forms.Label
    Private WithEvents lblVoucherNoNew As System.Windows.Forms.Label
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
End Class

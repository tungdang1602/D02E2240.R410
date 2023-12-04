<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F2100
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F2100))
        Dim Style1 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style2 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style3 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style4 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style5 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style6 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style7 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style8 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style9 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style10 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style11 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style12 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style13 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style14 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style15 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style16 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Me.grp1 = New System.Windows.Forms.GroupBox()
        Me.tdbgSource = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tdbdSourceCode = New C1.Win.C1TrueDBGrid.C1TrueDBDropdown()
        Me.tdbcVoucherTypeID = New C1.Win.C1List.C1Combo()
        Me.lblVoucherTypeID = New System.Windows.Forms.Label()
        Me.grpAssignment = New System.Windows.Forms.GroupBox()
        Me.tdbgAssignment = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tdbdAssignmentCode = New C1.Win.C1TrueDBGrid.C1TrueDBDropdown()
        Me.grp3 = New System.Windows.Forms.GroupBox()
        Me.tdbg = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.chkExecced = New System.Windows.Forms.CheckBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.grp1.SuspendLayout()
        CType(Me.tdbgSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbdSourceCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbcVoucherTypeID, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAssignment.SuspendLayout()
        CType(Me.tdbgAssignment, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbdAssignmentCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grp3.SuspendLayout()
        CType(Me.tdbg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grp1
        '
        Me.grp1.Controls.Add(Me.tdbgSource)
        Me.grp1.Controls.Add(Me.tdbdSourceCode)
        Me.grp1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grp1.Location = New System.Drawing.Point(14, 37)
        Me.grp1.Name = "grp1"
        Me.grp1.Size = New System.Drawing.Size(491, 119)
        Me.grp1.TabIndex = 2
        Me.grp1.TabStop = False
        Me.grp1.Text = "1. Nguồn hình thành"
        '
        'tdbgSource
        '
        Me.tdbgSource.AllowAddNew = True
        Me.tdbgSource.AllowColMove = False
        Me.tdbgSource.AllowColSelect = False
        Me.tdbgSource.AllowDelete = True
        Me.tdbgSource.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbgSource.AllowSort = False
        Me.tdbgSource.AlternatingRows = True
        Me.tdbgSource.EmptyRows = True
        Me.tdbgSource.ExtendRightColumn = True
        Me.tdbgSource.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbgSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbgSource.Images.Add(CType(resources.GetObject("tdbgSource.Images"), System.Drawing.Image))
        Me.tdbgSource.Location = New System.Drawing.Point(7, 20)
        Me.tdbgSource.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor
        Me.tdbgSource.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbgSource.Name = "tdbgSource"
        Me.tdbgSource.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbgSource.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbgSource.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbgSource.PrintInfo.PageSettings = CType(resources.GetObject("tdbgSource.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbgSource.PropBag = resources.GetString("tdbgSource.PropBag")
        Me.tdbgSource.RecordSelectors = False
        Me.tdbgSource.Size = New System.Drawing.Size(475, 94)
        Me.tdbgSource.TabAcrossSplits = True
        Me.tdbgSource.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbgSource.TabIndex = 0
        Me.tdbgSource.Tag = "COLS"
        '
        'tdbdSourceCode
        '
        Me.tdbdSourceCode.AllowColMove = False
        Me.tdbdSourceCode.AllowColSelect = False
        Me.tdbdSourceCode.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbdSourceCode.AllowSort = False
        Me.tdbdSourceCode.AlternatingRows = True
        Me.tdbdSourceCode.CaptionStyle = Style1
        Me.tdbdSourceCode.ColumnCaptionHeight = 17
        Me.tdbdSourceCode.ColumnFooterHeight = 17
        Me.tdbdSourceCode.DisplayMember = "SourceID"
        Me.tdbdSourceCode.EmptyRows = True
        Me.tdbdSourceCode.EvenRowStyle = Style2
        Me.tdbdSourceCode.ExtendRightColumn = True
        Me.tdbdSourceCode.FetchRowStyles = False
        Me.tdbdSourceCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbdSourceCode.FooterStyle = Style3
        Me.tdbdSourceCode.HeadingStyle = Style4
        Me.tdbdSourceCode.HighLightRowStyle = Style5
        Me.tdbdSourceCode.Images.Add(CType(resources.GetObject("tdbdSourceCode.Images"), System.Drawing.Image))
        Me.tdbdSourceCode.Location = New System.Drawing.Point(56, 29)
        Me.tdbdSourceCode.Name = "tdbdSourceCode"
        Me.tdbdSourceCode.OddRowStyle = Style6
        Me.tdbdSourceCode.PropBag = resources.GetString("tdbdSourceCode.PropBag")
        Me.tdbdSourceCode.RecordSelectorStyle = Style7
        Me.tdbdSourceCode.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.tdbdSourceCode.RowDivider.Style = C1.Win.C1TrueDBGrid.LineStyleEnum.[Single]
        Me.tdbdSourceCode.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbdSourceCode.ScrollTips = False
        Me.tdbdSourceCode.Size = New System.Drawing.Size(300, 83)
        Me.tdbdSourceCode.Style = Style8
        Me.tdbdSourceCode.TabIndex = 1
        Me.tdbdSourceCode.TabStop = False
        Me.tdbdSourceCode.ValueMember = "SourceID"
        Me.tdbdSourceCode.Visible = False
        '
        'tdbcVoucherTypeID
        '
        Me.tdbcVoucherTypeID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcVoucherTypeID.AllowColMove = False
        Me.tdbcVoucherTypeID.AllowSort = False
        Me.tdbcVoucherTypeID.AlternatingRows = True
        Me.tdbcVoucherTypeID.AutoCompletion = True
        Me.tdbcVoucherTypeID.AutoDropDown = True
        Me.tdbcVoucherTypeID.Caption = ""
        Me.tdbcVoucherTypeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcVoucherTypeID.ColumnWidth = 100
        Me.tdbcVoucherTypeID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcVoucherTypeID.DisplayMember = "VoucherTypeID"
        Me.tdbcVoucherTypeID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcVoucherTypeID.DropDownWidth = 500
        Me.tdbcVoucherTypeID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcVoucherTypeID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcVoucherTypeID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcVoucherTypeID.EmptyRows = True
        Me.tdbcVoucherTypeID.ExtendRightColumn = True
        Me.tdbcVoucherTypeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcVoucherTypeID.Images.Add(CType(resources.GetObject("tdbcVoucherTypeID.Images"), System.Drawing.Image))
        Me.tdbcVoucherTypeID.Location = New System.Drawing.Point(100, 10)
        Me.tdbcVoucherTypeID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcVoucherTypeID.MaxDropDownItems = CType(8, Short)
        Me.tdbcVoucherTypeID.MaxLength = 32767
        Me.tdbcVoucherTypeID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcVoucherTypeID.Name = "tdbcVoucherTypeID"
        Me.tdbcVoucherTypeID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcVoucherTypeID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcVoucherTypeID.Size = New System.Drawing.Size(128, 21)
        Me.tdbcVoucherTypeID.TabIndex = 1
        Me.tdbcVoucherTypeID.ValueMember = "VoucherTypeID"
        Me.tdbcVoucherTypeID.PropBag = resources.GetString("tdbcVoucherTypeID.PropBag")
        '
        'lblVoucherTypeID
        '
        Me.lblVoucherTypeID.AutoSize = True
        Me.lblVoucherTypeID.Location = New System.Drawing.Point(17, 15)
        Me.lblVoucherTypeID.Name = "lblVoucherTypeID"
        Me.lblVoucherTypeID.Size = New System.Drawing.Size(56, 13)
        Me.lblVoucherTypeID.TabIndex = 0
        Me.lblVoucherTypeID.Text = "Loại phiếu"
        Me.lblVoucherTypeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grpAssignment
        '
        Me.grpAssignment.Controls.Add(Me.tdbgAssignment)
        Me.grpAssignment.Controls.Add(Me.tdbdAssignmentCode)
        Me.grpAssignment.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpAssignment.Location = New System.Drawing.Point(14, 168)
        Me.grpAssignment.Name = "grpAssignment"
        Me.grpAssignment.Size = New System.Drawing.Size(491, 118)
        Me.grpAssignment.TabIndex = 3
        Me.grpAssignment.TabStop = False
        Me.grpAssignment.Text = "2. Tiêu thức phân bổ"
        '
        'tdbgAssignment
        '
        Me.tdbgAssignment.AllowAddNew = True
        Me.tdbgAssignment.AllowColMove = False
        Me.tdbgAssignment.AllowColSelect = False
        Me.tdbgAssignment.AllowDelete = True
        Me.tdbgAssignment.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbgAssignment.AllowSort = False
        Me.tdbgAssignment.AlternatingRows = True
        Me.tdbgAssignment.EmptyRows = True
        Me.tdbgAssignment.ExtendRightColumn = True
        Me.tdbgAssignment.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbgAssignment.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbgAssignment.Images.Add(CType(resources.GetObject("tdbgAssignment.Images"), System.Drawing.Image))
        Me.tdbgAssignment.Location = New System.Drawing.Point(7, 17)
        Me.tdbgAssignment.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor
        Me.tdbgAssignment.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbgAssignment.Name = "tdbgAssignment"
        Me.tdbgAssignment.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbgAssignment.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbgAssignment.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbgAssignment.PrintInfo.PageSettings = CType(resources.GetObject("tdbgAssignment.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbgAssignment.PropBag = resources.GetString("tdbgAssignment.PropBag")
        Me.tdbgAssignment.RecordSelectors = False
        Me.tdbgAssignment.Size = New System.Drawing.Size(475, 94)
        Me.tdbgAssignment.TabAcrossSplits = True
        Me.tdbgAssignment.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbgAssignment.TabIndex = 0
        Me.tdbgAssignment.Tag = "COLA"
        '
        'tdbdAssignmentCode
        '
        Me.tdbdAssignmentCode.AllowColMove = False
        Me.tdbdAssignmentCode.AllowColSelect = False
        Me.tdbdAssignmentCode.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbdAssignmentCode.AllowSort = False
        Me.tdbdAssignmentCode.AlternatingRows = True
        Me.tdbdAssignmentCode.CaptionStyle = Style9
        Me.tdbdAssignmentCode.ColumnCaptionHeight = 17
        Me.tdbdAssignmentCode.ColumnFooterHeight = 17
        Me.tdbdAssignmentCode.DisplayMember = "AssignmentID"
        Me.tdbdAssignmentCode.EmptyRows = True
        Me.tdbdAssignmentCode.EvenRowStyle = Style10
        Me.tdbdAssignmentCode.ExtendRightColumn = True
        Me.tdbdAssignmentCode.FetchRowStyles = False
        Me.tdbdAssignmentCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbdAssignmentCode.FooterStyle = Style11
        Me.tdbdAssignmentCode.HeadingStyle = Style12
        Me.tdbdAssignmentCode.HighLightRowStyle = Style13
        Me.tdbdAssignmentCode.Images.Add(CType(resources.GetObject("tdbdAssignmentCode.Images"), System.Drawing.Image))
        Me.tdbdAssignmentCode.Location = New System.Drawing.Point(122, 19)
        Me.tdbdAssignmentCode.Name = "tdbdAssignmentCode"
        Me.tdbdAssignmentCode.OddRowStyle = Style14
        Me.tdbdAssignmentCode.PropBag = resources.GetString("tdbdAssignmentCode.PropBag")
        Me.tdbdAssignmentCode.RecordSelectorStyle = Style15
        Me.tdbdAssignmentCode.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.tdbdAssignmentCode.RowDivider.Style = C1.Win.C1TrueDBGrid.LineStyleEnum.[Single]
        Me.tdbdAssignmentCode.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbdAssignmentCode.ScrollTips = False
        Me.tdbdAssignmentCode.Size = New System.Drawing.Size(300, 83)
        Me.tdbdAssignmentCode.Style = Style16
        Me.tdbdAssignmentCode.TabIndex = 1
        Me.tdbdAssignmentCode.TabStop = False
        Me.tdbdAssignmentCode.ValueMember = "AssignmentID"
        Me.tdbdAssignmentCode.Visible = False
        '
        'grp3
        '
        Me.grp3.Controls.Add(Me.tdbg)
        Me.grp3.Controls.Add(Me.chkExecced)
        Me.grp3.Location = New System.Drawing.Point(14, 292)
        Me.grp3.Name = "grp3"
        Me.grp3.Size = New System.Drawing.Size(491, 205)
        Me.grp3.TabIndex = 4
        Me.grp3.TabStop = False
        '
        'tdbg
        '
        Me.tdbg.AllowColMove = False
        Me.tdbg.AllowColSelect = False
        Me.tdbg.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg.AllowSort = False
        Me.tdbg.AlternatingRows = True
        Me.tdbg.EmptyRows = True
        Me.tdbg.ExtendRightColumn = True
        Me.tdbg.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg.Images.Add(CType(resources.GetObject("tdbg.Images"), System.Drawing.Image))
        Me.tdbg.Location = New System.Drawing.Point(7, 23)
        Me.tdbg.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg.Name = "tdbg"
        Me.tdbg.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg.PrintInfo.PageSettings = CType(resources.GetObject("tdbg.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg.PropBag = resources.GetString("tdbg.PropBag")
        Me.tdbg.Size = New System.Drawing.Size(476, 174)
        Me.tdbg.TabAcrossSplits = True
        Me.tdbg.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg.TabIndex = 1
        Me.tdbg.Tag = "COL"
        '
        'chkExecced
        '
        Me.chkExecced.AutoSize = True
        Me.chkExecced.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkExecced.Location = New System.Drawing.Point(12, -1)
        Me.chkExecced.Name = "chkExecced"
        Me.chkExecced.Size = New System.Drawing.Size(294, 17)
        Me.chkExecced.TabIndex = 0
        Me.chkExecced.Text = "3. Loại trừ những tài sản hình thành từ mua mới"
        Me.chkExecced.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(347, 503)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(76, 22)
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "&Lưu"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(429, 503)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 22)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'D02F2100
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(514, 530)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.grp3)
        Me.Controls.Add(Me.grpAssignment)
        Me.Controls.Add(Me.tdbcVoucherTypeID)
        Me.Controls.Add(Me.grp1)
        Me.Controls.Add(Me.lblVoucherTypeID)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F2100"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "TÁo tø ¢èng sç d§ TSC˜ - D02F2100"
        Me.grp1.ResumeLayout(False)
        CType(Me.tdbgSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbdSourceCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbcVoucherTypeID, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAssignment.ResumeLayout(False)
        CType(Me.tdbgAssignment, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbdAssignmentCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grp3.ResumeLayout(False)
        Me.grp3.PerformLayout()
        CType(Me.tdbg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents grp1 As System.Windows.Forms.GroupBox
    Private WithEvents tdbcVoucherTypeID As C1.Win.C1List.C1Combo
    Private WithEvents lblVoucherTypeID As System.Windows.Forms.Label
    Private WithEvents grpAssignment As System.Windows.Forms.GroupBox
    Private WithEvents grp3 As System.Windows.Forms.GroupBox
    Private WithEvents tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents chkExecced As System.Windows.Forms.CheckBox
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents tdbgSource As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents tdbgAssignment As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents tdbdSourceCode As C1.Win.C1TrueDBGrid.C1TrueDBDropdown
    Private WithEvents tdbdAssignmentCode As C1.Win.C1TrueDBGrid.C1TrueDBDropdown
End Class
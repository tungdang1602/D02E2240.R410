'#-------------------------------------------------------------------------------------
'# Created Date: 
'# Created User: Nguyễn Thị Ánh
'# Modify Date: 14/09/2007 8:03:27 AM
'# Modify User: Trần Thị ÁiTrâm
'# Description: Bổ sung mới phần Nguồn hình thành và Tiêu thức phân bổ
'#-------------------------------------------------------------------------------------
Public Class D02F2100

#Region "Const of tdbg"
    Private Const COL_Selected As Integer = 0  ' Chọn
    Private Const COL_AssetID As Integer = 1   ' Mã tài sản
    Private Const COL_AssetName As Integer = 2 ' Tên tài sản
#End Region

#Region "Const of tdbgSource"
    Private Const COLS_SourceType As Integer = 0 ' SourceType
    Private Const COLS_SourceID As Integer = 1   ' Nguồn
    Private Const COLS_SourceName As Integer = 2 ' Tên nguồn
    Private Const COLS_Rate As Integer = 3       ' Tỷ lệ 
#End Region

#Region "Const of tdbgAssignment"
    Private Const COLA_SourceType As Integer = 0     ' SourceType
    Private Const COLA_AssignmentID As Integer = 1   ' Tiêu thức
    Private Const COLA_AssignmentName As Integer = 2 ' Tên tiêu thức
    Private Const COLA_Rate As Integer = 3           ' Tỷ lệ 
#End Region

    Dim sStrAssetID As String = ""
    Dim bInsertRow As Boolean = False


    Private Sub D02F2100_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.Control Then
            If e.KeyCode = Keys.D1 Or e.KeyCode = Keys.NumPad1 Then
                tdbgSource.Focus()
            ElseIf e.KeyCode = Keys.D2 Or e.KeyCode = Keys.NumPad2 Then
                tdbgAssignment.Focus()
            ElseIf e.KeyCode = Keys.D3 Or e.KeyCode = Keys.NumPad3 Then
                Application.DoEvents()
                chkExecced.Focus()
                Application.DoEvents()
            End If
        End If
        If e.KeyCode = Keys.F11 Then
            'If tdbgSource.Focus Then
            HotKeyF11(Me, tdbgSource)
            'ElseIf tdbgAssignment.Focus Then
            HotKeyF11(Me, tdbgAssignment)
            'ElseIf tdbg.Focus Then
            HotKeyF11(Me, tdbg)
            'End If
        End If
    End Sub

    Private Sub D02F2100_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        Loadlanguage()
        SetBackColorObligatory()
        LoadTDBCombo()
        LoadTDBDropDown()
        LoadTDBGridSource()
        InputbyUnicode(Me, gbUnicode)
        LoadTDBGridAssignment()
        tdbgSource_NumberFormat()
        tdbgAssignment_NumberFormat()
        
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub SetBackColorObligatory()
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub


    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcVoucherTypeID
        'sSQL = "Select VoucherTypeID, VoucherTypeName From D91T0001 Where Disabled=0 And UseD02=1 Order By VoucherTypeID"
        'LoadDataSource(tdbcVoucherTypeID, sSQL)
        LoadVoucherTypeID(tdbcVoucherTypeID, D02, , gbUnicode)
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1200
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 12/01/2007 03:22:16
    '# Modified User: 
    '# Modified Date: 
    '# Description: lưu tdbg
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1200(ByVal sAsset As String) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1200 "
        sSQL &= SQLString(sAsset) 'strAssetID, varchar[1000], NOT NULL
        Return sSQL
    End Function

    Private Sub LoadTDBGridAsset()
        Dim sSQL As String = "SELECT '0' AS Selected, AssetID, AssetName" & UnicodeJoin(gbUnicode) & " as AssetName FROM D02T0001 WITH(NOLOCK) " & vbCrLf
        sSQL &= "WHERE IsCompleted = 0 ORDER BY AssetID"
        LoadDataSource(tdbg, sSQL, gbUnicode)
    End Sub

    Private Sub chkExecced_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkExecced.CheckedChanged
        If chkExecced.Checked Then
            tdbg.Enabled = True
            LoadTDBGridAsset()
        Else
            tdbg.Enabled = False
        End If
    End Sub

    Private Function AllowSave() As Boolean
        Dim dSumRateS As Double = 0
        Dim dSumRateA As Double = 0
        Dim j As Integer

        If tdbcVoucherTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Loai_phieu"))
            tdbcVoucherTypeID.Focus()
            Return False
        End If
        If tdbgSource.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbgSource.Focus()
            Return False
        End If
       
        For i As Integer = 0 To tdbgSource.RowCount - 1
            If tdbgSource(i, COLS_SourceID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Nguon"))
                tdbgSource.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                tdbgSource.SplitIndex = SPLIT0
                tdbgSource.Bookmark = i
                tdbgSource.Col = COLS_SourceID
                tdbgSource.Focus()
                Return False
            End If
            If tdbgSource(i, COLS_Rate).ToString = Format("0.00") Then
                D99C0008.MsgNotYetEnter(rl3("Ty_le"))
                tdbgSource.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                tdbgSource.SplitIndex = SPLIT0
                tdbgSource.Bookmark = i
                tdbgSource.Col = COLS_Rate
                tdbgSource.Focus()
                Return False
            End If
            If tdbgSource(i, COLS_Rate).ToString <> "" Then
                If CDbl(tdbgSource(i, COLS_Rate)) > MaxMoney Then
                    D99C0008.MsgL3(rl3("Ty_le_qua_lon"))
                    tdbgSource.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                    tdbgSource.SplitIndex = SPLIT0
                    tdbgSource.Bookmark = i
                    tdbgSource.Col = COLS_Rate
                    tdbgSource.Focus()
                    Return False
                End If
            End If
            If tdbgSource(i, COLS_SourceType).ToString = "A" Then
                If Not ExistRecord("Select Top 1 1 From D02T0001 Where " & SQLString(tdbgSource(i, COLS_SourceID)) & " <> ''") Then
                    D99C0008.MsgL3(rl3("Khoan_muc_cho_nguon_hinh_thanh_khong_co"))
                    tdbgSource.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                    tdbgSource.SplitIndex = SPLIT0
                    tdbgSource.Bookmark = i
                    tdbgSource.Col = COLS_SourceID
                    tdbgSource.Focus()
                    Return False
                End If
            End If
            dSumRateS += CDbl(tdbgSource(i, COLS_Rate))
        Next
        For i As Integer = 0 To tdbgSource.RowCount - 2
            j = i + 1
            If tdbgSource(i, COLS_SourceID).ToString = tdbgSource(j, COLS_SourceID).ToString Then
                D99C0008.MsgL3(rl3("Nguon_nay_da_ton_tai"))
                tdbgSource.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                tdbgSource.SplitIndex = SPLIT0
                tdbgSource.Bookmark = j
                tdbgSource.Col = COLS_SourceID
                tdbgSource.Focus()
                Return False
            End If
        Next
        If dSumRateS < 100 Or dSumRateS > 100 Then
            D99C0008.MsgL3(rl3("Tong_Ty_le_nguon_hinh_thanh") & " = 100")
            tdbgSource.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
            tdbgSource.SplitIndex = SPLIT0
            tdbgSource.Bookmark = 0
            tdbgSource.Col = COLS_Rate
            tdbgSource.Focus()
            Return False
        End If
        If tdbgAssignment.RowCount <= 0 Then
            D99C0008.MsgL3(rl3("Ban_phai_chon_tieu_thuc_phan_bo"))
            tdbgAssignment.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbgAssignment.RowCount - 1
            If tdbgAssignment(i, COLA_AssignmentID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Tieu_thuc"))
                tdbgAssignment.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                tdbgAssignment.SplitIndex = SPLIT0
                tdbgAssignment.Bookmark = i
                tdbgAssignment.Col = COLA_AssignmentID
                tdbgAssignment.Focus()
                Return False
            End If
            For j = 0 To tdbgSource.RowCount - 1
                If tdbgAssignment(i, COLA_AssignmentID).ToString <> "" Then
                    If tdbgAssignment(i, COLA_AssignmentID).ToString = tdbgSource(j, COLS_SourceID).ToString Then
                        D99C0008.MsgL3(rl3("Tieu_thuc_khong_duoc_trung_voi_nguon_hinh_thanh"))
                        tdbgAssignment.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                        tdbgAssignment.SplitIndex = SPLIT0
                        tdbgAssignment.Bookmark = i
                        tdbgAssignment.Col = COLA_AssignmentID
                        tdbgAssignment.Focus()
                        Return False
                    End If
                End If
            Next
            If tdbgAssignment(i, COLA_Rate).ToString = Format("0.00") Then
                tdbgAssignment.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                D99C0008.MsgNotYetEnter(rl3("Ty_le"))
                tdbgAssignment.SplitIndex = SPLIT0
                tdbgAssignment.Bookmark = i
                tdbgAssignment.Col = COLA_Rate
                tdbgAssignment.Focus()
                Return False
            End If
            If tdbgAssignment(i, COLA_Rate).ToString <> "" Then
                If CDbl(tdbgAssignment(i, COLA_Rate)) > MaxMoney Then
                    D99C0008.MsgL3(rl3("Ty_le_qua_lon"))
                    tdbgAssignment.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                    tdbgAssignment.SplitIndex = SPLIT0
                    tdbgAssignment.Bookmark = i
                    tdbgAssignment.Col = COLA_Rate
                    tdbgAssignment.Focus()
                    Return False
                End If
            End If
            If tdbgAssignment(i, COLA_SourceType).ToString = "A" Then
                If Not ExistRecord("Select Top 1 1 From D02T0001 Where " & SQLString(tdbgAssignment(i, COLA_AssignmentID)) & " <> ''") Then
                    D99C0008.MsgL3(rl3("Khoan_muc_cho_tieu_thuc_phan_bo_khong_co"))
                    tdbgAssignment.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                    tdbgAssignment.SplitIndex = SPLIT0
                    tdbgAssignment.Bookmark = i
                    tdbgAssignment.Col = COLA_AssignmentID
                    tdbgAssignment.Focus()
                    Return False
                End If
            End If
            dSumRateA += CDbl(tdbgAssignment(i, COLA_Rate))
        Next

        For i As Integer = 0 To tdbgAssignment.RowCount - 2
            j = i + 1
            If tdbgAssignment(i, COLA_AssignmentID).ToString = tdbgAssignment(j, COLA_AssignmentID).ToString Then
                D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_ton_tai"))
                tdbgAssignment.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
                tdbgAssignment.SplitIndex = SPLIT0
                tdbgAssignment.Bookmark = j
                tdbgAssignment.Col = COLA_AssignmentID
                tdbgAssignment.Focus()
                Return False
            End If
        Next

        If dSumRateA < 100 Or dSumRateA > 100 Then
            D99C0008.MsgL3(rl3("Tong_Ty_le_tieu_thuc_phan_bo") & " = 100")
            tdbgAssignment.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
            tdbgAssignment.SplitIndex = SPLIT0
            tdbgAssignment.Bookmark = 0
            tdbgAssignment.Col = COLA_Rate
            tdbgAssignment.Focus()
            Return False
        End If
        If chkExecced.Checked Then
            If tdbg.RowCount <= 0 Then
                D99C0008.MsgNoDataInGrid()
                tdbg.Focus()
                Return False
            End If
            Dim iCount As Integer = 0
            sStrAssetID = "("
            For i As Integer = 0 To tdbg.RowCount - 1
                If CBool(tdbg(i, COL_Selected)) Then
                    If iCount = 0 Then
                        sStrAssetID &= SQLString(tdbg(i, COL_AssetID).ToString)
                    Else
                        sStrAssetID &= "," & SQLString(tdbg(i, COL_AssetID).ToString)
                    End If
                    iCount += 1
                End If
            Next
            sStrAssetID &= ")"
            If iCount = 0 Then
                D99C0008.MsgNoDataInGrid()
                tdbg.Focus()
                Return False
            End If
            If iCount > 300 Then
                D99C0008.MsgL3(rl3("Ban_chon_qua_nhieu_dong_tren_luoi"))
                tdbg.Focus()
                Return False
            End If
        End If

        Return True
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        tdbgSource.UpdateData()
        tdbgAssignment.UpdateData()
        tdbg.UpdateData()
        If Not AllowSave() Then Exit Sub
        Dim sSQL As String = ""
        btnSave.Enabled = False
        btnClose.Enabled = False

        'If chkExecced.Checked Then sSQL &= SQLStoreD02P1200(sStrAssetID) & vbCrLf
        sSQL &= SQLInsertTableSource() & vbCrLf
        sSQL &= SQLInsertTableAssignment() & vbCrLf
        sSQL &= SQLCreateTableTemp() & vbCrLf
        sSQL &= SQLStoreD02P1215() & vbCrLf
        sSQL &= SQLDropTable()
        Me.Cursor = Cursors.WaitCursor
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL)
        Me.Cursor = Cursors.Default
        If bRunSQL Then
            SaveOK()
            btnSave.Enabled = True
            btnClose.Enabled = True
            btnClose.Focus()
        Else
            SaveNotOK()
            btnSave.Enabled = True
            btnClose.Enabled = True
        End If
    End Sub

#Region "Events tdbcVoucherTypeID"

    Private Sub tdbcVoucherTypeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.Close
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then tdbcVoucherTypeID.Text = ""
    End Sub

    Private Sub tdbcVoucherTypeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcVoucherTypeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcVoucherTypeID.Text = ""
    End Sub

#End Region

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Tao_tu_dong_so_du_TSCD_-_D02F2100") & UnicodeCaption(gbUnicode) 'TÁo tø ¢èng sç d§ TSC˜ - D02F2100
        '================================================================ 
        lblVoucherTypeID.Text = rl3("Loai_phieu") 'Loại phiếu
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        'btnHelp.Text = rl3("Tro__giup") 'Trợ &giúp
        '================================================================ 
        chkExecced.Text = "3. " & rl3("Loai_tru_nhung_tai_san_hinh_thanh_tu_mua_moi") 'Loại trừ những tài sản hình thành từ mua mới
        '================================================================ 
        grp1.Text = "1. " & rl3("Nguon_hinh_thanh") 'Nguồn hình thành
        grpAssignment.Text = "2. " & rl3("Tieu_thuc_phan_bo") 'Tiêu thức phân bổ
        '================================================================ 
        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rl3("Ma") 'Mã
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rl3("Dien_giai") 'Tên loại phiếu
        '================================================================ 
        tdbdSourceCode.Columns("SourceID").Caption = rl3("Ma") 'Mã
        tdbdSourceCode.Columns("SourceName").Caption = rl3("Ten") 'Tên
        tdbdAssignmentCode.Columns("AssignmentID").Caption = rl3("Ma") 'Mã
        tdbdAssignmentCode.Columns("AssignmentName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbgSource.Columns("SourceID").Caption = rl3("Nguon") 'Nguồn
        tdbgSource.Columns("SourceName").Caption = rl3("Ten_nguon") 'Tên nguồn
        tdbgSource.Columns("Rate").Caption = rl3("Ty_le") 'Tỷ lệ 

        tdbgAssignment.Columns("AssignmentID").Caption = rl3("Tieu_thuc") 'Tiêu thức
        tdbgAssignment.Columns("AssignmentName").Caption = rl3("Ten_tieu_thuc") 'Tên tiêu thức
        tdbgAssignment.Columns("Rate").Caption = rl3("Ty_le") 'Tỷ lệ 
        tdbg.Columns("Selected").Caption = rl3("Chon") 'Chọn
        tdbg.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
    End Sub

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.HeadClick
        Select Case e.ColIndex
            Case COL_Selected
                Dim iCount As Int32 = 0
                If chkExecced.Checked Then

                    For i As Integer = 0 To tdbg.RowCount - 1
                        If CBool(tdbg(i, COL_Selected).ToString) = True Then
                            iCount = iCount + 1
                        End If
                    Next
                    If iCount < tdbg.RowCount - 1 Then
                        For i As Integer = 0 To tdbg.RowCount - 1
                            tdbg(i, COL_Selected) = True
                        Next
                    Else
                        For i As Integer = 0 To tdbg.RowCount - 1
                            If CBool(tdbg(i, COL_Selected).ToString) = True Then
                                tdbg(i, COL_Selected) = False
                            Else
                                tdbg(i, COL_Selected) = True
                            End If
                        Next
                    End If
                End If
        End Select
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        
        Dim iCount As Int32 = 0
        If chkExecced.Checked Then
            If e.KeyCode = Keys.Enter Then
                HotKeyEnterGrid(tdbg, COL_Selected, e)
            End If
            If e.Control And e.KeyCode = Keys.S Then
                For i As Integer = 0 To tdbg.RowCount - 1
                    If CBool(tdbg(i, COL_Selected).ToString) = True Then
                        iCount = iCount + 1
                    End If
                Next
                If iCount < tdbg.RowCount - 1 Then
                    For i As Integer = 0 To tdbg.RowCount - 1
                        tdbg(i, COL_Selected) = True
                    Next
                Else
                    For i As Integer = 0 To tdbg.RowCount - 1
                        If CBool(tdbg(i, COL_Selected).ToString) = True Then
                            tdbg(i, COL_Selected) = False
                        Else
                            tdbg(i, COL_Selected) = True
                        End If
                    Next
                End If
            End If

        End If
    End Sub

    Private Sub LoadTDBDropDown()
        Dim sSQL As String = ""
        'Load tdbdSourceCode
        sSQL = " Select Left(TypeCodeID,1) + 'Code' + SubString(TypeCodeID,2,2) As SourceID," & vbCrLf
        sSQL &= IIf(geLanguage = EnumLanguage.Vietnamese, "VieTypeCodeName" & UnicodeJoin(gbUnicode), "EngTypeCodeName" & UnicodeJoin(gbUnicode)).ToString & " As SourceName, 'A' As SourceType " & vbCrLf
        sSQL &= " From D02T0040 WITH(NOLOCK) Where Type='A' And Disabled=0 " & vbCrLf
        sSQL &= " Union All" & vbCrLf
        sSQL &= "Select SourceID, SourceName" & UnicodeJoin(gbUnicode) & " as SourceName, 'S' As SourceType From D02T0013 WITH(NOLOCK) Where Disabled=0 Order By SourceType, SourceID"
        LoadDataSource(tdbdSourceCode, sSQL, gbUnicode)
        'Load tdbdAssignmentCode
        sSQL = "Select Left(TypeCodeID,1) + 'Code' + Substring(TypeCodeID,2,2) As AssignmentID, " & vbCrLf
        sSQL &= IIf(geLanguage = EnumLanguage.Vietnamese, "VieTypeCodeName" & UnicodeJoin(gbUnicode), "EngTypeCodeName" & UnicodeJoin(gbUnicode)).ToString & " As AssignmentName, 'A' As SourceType " & vbCrLf
        sSQL &= " From D02T0040 WITH(NOLOCK) Where Type='A' And Disabled=0 " & vbCrLf
        sSQL &= "Union All" & vbCrLf
        sSQL &= "Select AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " as SourceName, 'S' As SourceType From D02T0002 WITH(NOLOCK) Order By SourceType, AssignmentID"
        LoadDataSource(tdbdAssignmentCode, sSQL, gbUnicode)
    End Sub

    Private Sub tdbgAssignment_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgAssignment.ComboSelect
        Select Case e.ColIndex
            Case COLA_AssignmentID
                tdbgAssignment.Columns(COLA_SourceType).Text = tdbdAssignmentCode.Columns("SourceType").Value.ToString
                tdbgAssignment.Columns(COLA_AssignmentName).Text = tdbdAssignmentCode.Columns("AssignmentName").Value.ToString
                tdbgAssignment.Columns(COLA_Rate).Text = Format("0.00")
        End Select
    End Sub

    Private Sub tdbgSource_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgSource.ComboSelect
        Select Case e.ColIndex
            Case COLS_SourceID
                tdbgSource.Columns(COLS_SourceType).Text = tdbdSourceCode.Columns("SourceType").Value.ToString
                tdbgSource.Columns(COLS_SourceName).Text = tdbdSourceCode.Columns("SourceName").Value.ToString
                tdbgSource.Columns(COLS_Rate).Text = Format("0.00")
        End Select
    End Sub

    Private Sub tdbgSource_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbgSource.KeyPress
        Select Case tdbgSource.Col
            Case COLS_Rate
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub tdbgAssignment_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbgAssignment.KeyPress
        Select Case tdbgAssignment.Col
            Case COLA_Rate
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub tdbgSource_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbgSource.BeforeColUpdate
        Select Case e.ColIndex
            Case COLS_SourceID
                If tdbgSource.Columns(COLS_SourceID).Text <> tdbdSourceCode.Columns("SourceID").Text Then
                    tdbgSource.Columns(COLS_SourceType).Text = ""
                    tdbgSource.Columns(COLS_SourceID).Text = ""
                End If
            Case COLS_Rate
                If Not IsNumeric(tdbgSource.Columns(COLS_Rate).Text) Then e.Cancel = True
        End Select
    End Sub

    Private Sub tdbgAssignment_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbgAssignment.BeforeColUpdate
        Select Case e.ColIndex
            Case COLA_AssignmentID
                If tdbgAssignment.Columns(COLA_AssignmentID).Text <> tdbdAssignmentCode.Columns("AssignmentID").Text Then
                    tdbgAssignment.Columns(COLA_SourceType).Text = ""
                    tdbgAssignment.Columns(COLA_AssignmentID).Text = ""
                End If
            Case COLA_Rate
                If Not IsNumeric(tdbgAssignment.Columns(COLA_Rate).Text) Then e.Cancel = True
        End Select
    End Sub

    '#Tạo table đổ giả dữ liệu cho lưới Nguồn hình thành
    Private Sub LoadTDBGridSource()
        Dim dtSource As New DataTable
        Dim col1 As New DataColumn
        col1.DataType = System.Type.GetType("System.String")
        col1.Caption = "SourceType"
        col1.ColumnName = "SourceType"
        dtSource.Columns.Add(col1)

        Dim col2 As New DataColumn
        col2.DataType = System.Type.GetType("System.String")
        col2.Caption = rl3("Nguon")
        col2.ColumnName = "SourceID"
        dtSource.Columns.Add(col2)

        Dim col3 As New DataColumn
        col3.DataType = System.Type.GetType("System.String")
        col3.Caption = rl3("Ten_nguon")
        col3.ColumnName = "SourceName"
        dtSource.Columns.Add(col3)

        Dim col4 As New DataColumn
        col4.DataType = System.Type.GetType("System.Double")
        col4.Caption = rl3("Ty_le")
        col4.ColumnName = "Rate"
        dtSource.Columns.Add(col4)
        LoadDataSource(tdbgSource, dtSource, gbUnicode)
    End Sub

    '#Tạo table đổ giả dữ liệu cho lưới Tiêu thức phân bổ
    Private Sub LoadTDBGridAssignment()
        Dim dtAssignment As New DataTable
        Dim col1 As New DataColumn
        col1.DataType = System.Type.GetType("System.String")
        col1.Caption = "SourceType"
        col1.ColumnName = "SourceType"
        dtAssignment.Columns.Add(col1)

        Dim col2 As New DataColumn
        col2.DataType = System.Type.GetType("System.String")

        col2.Caption = rl3("Tieu_thuc")
        col2.ColumnName = "AssignmentID"
        dtAssignment.Columns.Add(col2)

        Dim col3 As New DataColumn
        col3.DataType = System.Type.GetType("System.String")
        col3.Caption = rL3("Ten_tieu_thuc")
        col3.ColumnName = "AssignmentName"
        dtAssignment.Columns.Add(col3)

        Dim col4 As New DataColumn
        col4.DataType = System.Type.GetType("System.Double")
        col4.Caption = rL3("Ty_le")
        col4.ColumnName = "Rate"
        dtAssignment.Columns.Add(col4)
        LoadDataSource(tdbgAssignment, dtAssignment, gbUnicode)
    End Sub

    Private Sub tdbgSource_NumberFormat()
        tdbgSource.Columns(COLS_Rate).NumberFormat = DxxFormat.DefaultNumber2
    End Sub

    Private Sub tdbgAssignment_NumberFormat()
        tdbgAssignment.Columns(COLA_Rate).NumberFormat = DxxFormat.DefaultNumber2
    End Sub

    Private Sub tdbgAssignment_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgAssignment.AfterColUpdate
        Select Case e.ColIndex
            Case COLA_AssignmentID
                If tdbgAssignment.Columns(COLA_AssignmentID).Text = "" Then
                    tdbgAssignment.Columns(COLA_SourceType).Text = ""
                    tdbgAssignment.Columns(COLA_AssignmentName).Text = ""
                End If
            Case COLA_AssignmentName
           
        End Select
    End Sub

    Private Sub tdbgSource_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgSource.AfterColUpdate
        Select Case e.ColIndex
            Case COLS_SourceID
                If tdbgSource.Columns(COLS_SourceID).Text = "" Then
                    tdbgSource.Columns(COLS_SourceType).Text = ""
                    tdbgSource.Columns(COLS_SourceName).Text = ""
                End If
        End Select
    End Sub

    'Insert dữ liệu vào bảng Source
    Private Function SQLInsertTableSource() As String
        Dim sSQL As String = ""
        sSQL &= "CREATE TABLE #Source (SourceType varchar(20), SourceID varchar(20), Rate money)" & vbCrLf
        For i As Integer = 0 To tdbgSource.RowCount - 1
            sSQL &= "INSERT INTO #Source (SourceType, SourceID, Rate)" & vbCrLf
            sSQL &= "VALUES(" & SQLString(tdbgSource(i, COLS_SourceType).ToString) & ", " & SQLString(tdbgSource(i, COLS_SourceID).ToString) & ", " & SQLMoney(tdbgSource(i, COLS_Rate).ToString) & " )" & vbCrLf
        Next
        Return sSQL
    End Function

    'Insert dữ liệu vào bảng Assignment
    Private Function SQLInsertTableAssignment() As String
        Dim sSQL As String = ""
        sSQL &= "CREATE TABLE #Assignment (SourceType varchar(20), AssignmentID varchar(50), Rate money)" & vbCrLf
        For i As Integer = 0 To tdbgAssignment.RowCount - 1
            sSQL &= "INSERT INTO #Assignment (SourceType, AssignmentID, Rate)" & vbCrLf
            sSQL &= "VALUES(" & SQLString(tdbgAssignment(i, COLA_SourceType).ToString) & ", " & SQLString(tdbgAssignment(i, COLA_AssignmentID).ToString) & ", " & SQLMoney(tdbgAssignment(i, COLA_Rate).ToString) & " )" & vbCrLf
        Next
        Return sSQL
    End Function

    '#Tạo bảng tạm chứa danh sách mã tài sản loại trừ
    Private Function SQLCreateTableTemp() As String
        Dim sSQL As String = ""
        sSQL &= "CREATE TABLE #Asset(AssetID varchar(20))" & vbCrLf

        If chkExecced.Checked Then
            For i As Integer = 0 To tdbg.RowCount - 1
                If CBool(tdbg(i, COL_Selected)) = True Then
                    sSQL &= "INSERT INTO #Asset(AssetID) VALUES(" & SQLString(tdbg(i, COL_AssetID).ToString) & ")"
                End If
            Next
        End If
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1215
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 14/09/2006 11:02:36
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1215() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1215 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
		sSQL &= SQLString(tdbcVoucherTypeID.Text) & COMMA 'VoucherTypeID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, varchar[20], NOT NULL
        Return sSQL
    End Function

    Private Function SQLDropTable() As String
        Dim sSQL As String = ""
        sSQL &= "DROP TABLE #Source, #Assignment, #Asset" & vbCrLf
        Return sSQL
    End Function

    Private Sub tdbgSource_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbgSource.KeyDown
        If e.KeyCode = Keys.Enter Then
            HotKeyEnterGrid(tdbgSource, COLS_SourceID, e)
        End If
        If e.Shift And e.KeyCode = Keys.Insert Then
            bInsertRow = True
            HotKeyShiftInsert(tdbgSource, 0, COLS_SourceID, tdbgSource.Columns.Count)
        End If
    End Sub

    Private Sub tdbgSource_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbgSource.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        If bInsertRow = True And tdbgSource.AddNewMode = C1.Win.C1TrueDBGrid.AddNewModeEnum.AddNewCurrent Then
            tdbgSource.Columns(COLS_SourceName).Text = "" ' Gán 1 cột bất kỳ ="" cho lưới
            bInsertRow = False
        End If
    End Sub

    Private Sub tdbgAssignment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbgAssignment.KeyDown
        If e.KeyCode = Keys.Enter Then
            HotKeyEnterGrid(tdbgAssignment, COLA_AssignmentID, e)
        End If
        If e.Shift And e.KeyCode = Keys.Insert Then
            bInsertRow = True
            HotKeyShiftInsert(tdbgAssignment, 0, COLA_AssignmentID, tdbgAssignment.Columns.Count)
        End If
    End Sub

    Private Sub tdbgAssignment_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbgAssignment.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        If bInsertRow = True And tdbgAssignment.AddNewMode = C1.Win.C1TrueDBGrid.AddNewModeEnum.AddNewCurrent Then
            tdbgAssignment.Columns(COLA_AssignmentID).Text = "" ' Gán 1 cột bất kỳ ="" cho lưới
            bInsertRow = False
        End If
    End Sub
End Class
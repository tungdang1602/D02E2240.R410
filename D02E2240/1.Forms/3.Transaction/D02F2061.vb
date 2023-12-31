Imports System
Public Class D02F2061

    Private _formIDPermission As String = "D02F2061"
    Public WriteOnly Property FormIDPermission() As String
        Set(ByVal Value As String)
            _formIDPermission = Value
        End Set
    End Property

    Private _voucherID As String
    Public WriteOnly Property VoucherID() As String
        Set(ByVal Value As String)
            _voucherID = Value
        End Set
    End Property

#Region "Const of tdbg - Total of Columns: 19"
    Private Const COL_IsChecked As Integer = 0             ' Chọn
    Private Const COL_VoucherNo As Integer = 1             ' Chứng từ
    Private Const COL_ObjectTypeID As Integer = 2          ' Loại đối tượng
    Private Const COL_ObjectID As Integer = 3              ' Đối tượng
    Private Const COL_ObjectName As Integer = 4            ' Tên đối tượng
    Private Const COL_AssetID As Integer = 5               ' Mã tài sản
    Private Const COL_AssetName As Integer = 6             ' Tên tài sản
    Private Const COL_AssetTypeName As Integer = 7         ' Loại tài sản
    Private Const COL_UnitName As Integer = 8              ' ĐVT
    Private Const COL_LocationName As Integer = 9          ' Vị trí
    Private Const COL_UseDate As Integer = 10              ' Ngày sử dụng
    Private Const COL_ManagementObjectName As Integer = 11 ' Đơn vị quản lý
    Private Const COL_FullName As Integer = 12             ' Cá nhân sử dụng
    Private Const COL_Status As Integer = 13               ' Tình trạng
    Private Const COL_RemainQTY As Integer = 14            ' Số lượng còn lại
    Private Const COL_RemainAMT As Integer = 15            ' Giá trị còn lại
    Private Const COL_InventoryQTY As Integer = 16         ' Số lượng kiểm kê
    Private Const COL_InventoryAMT As Integer = 17         ' Giá trị kiểm kê
    Private Const COL_Notes As Integer = 18                ' Ghi chú
#End Region

    Private dtGrid As DataTable

    Private Sub LoadLanguage()
        '================================================================ 
        Me.Text = rl3("In_chi_tiet_bien_ban_kiem_ke") & " - " & Me.Name & UnicodeCaption(gbUnicode) 'In chi tiÕt bi£n b¶n kiÓm k£
        '================================================================ 
        btnPrint.Text = rl3("_In") '&In
        '================================================================ 
        tdbg.Columns(COL_IsChecked).Caption = rl3("Chon") 'Chọn
        tdbg.Columns(COL_VoucherNo).Caption = rl3("Chung_tu") 'Chứng từ
        tdbg.Columns(COL_ObjectTypeID).Caption = rl3("Loai_doi_tuong") 'Loại đối tượng
        tdbg.Columns(COL_ObjectID).Caption = rl3("Doi_tuong") 'Đối tượng
        tdbg.Columns(COL_ObjectName).Caption = rl3("Ten_doi_tuong") 'Tên đối tượng
        tdbg.Columns(COL_AssetID).Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg.Columns(COL_AssetName).Caption = rl3("Ten_tai_san") 'Tên tài sản
        tdbg.Columns(COL_AssetTypeName).Caption = rl3("Loai_tai_sanU") 'Loại tài sản
        tdbg.Columns(COL_UnitName).Caption = rl3("DVT") 'ĐVT
        tdbg.Columns(COL_LocationName).Caption = rl3("Vi_tri") 'Vị trí
        tdbg.Columns(COL_UseDate).Caption = rl3("Ngay_su_dung") 'Ngày sử dụng
        tdbg.Columns(COL_ManagementObjectName).Caption = rl3("Don_vi_quan_ly") 'Đơn vị quản lý
        tdbg.Columns(COL_FullName).Caption = rl3("Ca_nhan_su_dung") 'Cá nhân sử dụng
        tdbg.Columns(COL_Status).Caption = rl3("Tinh_trang") 'Tình trạng
        tdbg.Columns(COL_RemainQTY).Caption = rl3("So_luong_con_lai") 'Số lượng còn lại
        tdbg.Columns(COL_RemainAMT).Caption = rl3("Gia_tri_con_lai") 'Giá trị còn lại
        tdbg.Columns(COL_InventoryQTY).Caption = rl3("So_luong_kiem_ke") 'Số lượng kiểm kê
        tdbg.Columns(COL_InventoryAMT).Caption = rl3("Gia_tri_kiem_ke") 'Giá trị kiểm kê
        tdbg.Columns(COL_Notes).Caption = rL3("Ghi_chu") 'Ghi chú

        '================================================================ 
        btnClose.Text = rL3("Do_ng") 'Đó&ng

    End Sub



    Private Sub D02F2061_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        LoadInfoGeneral() 'Load System/ Option /... in DxxD9940
        ResetColorGrid(tdbg)
        gbEnabledUseFind = False
        LoadTDBGrid()
        LoadLanguage()

        InputbyUnicode(Me, gbUnicode)
        SetResolutionForm(Me)
        tdbg_NumberFormat()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg.Columns(COL_RemainQTY).DataField, DxxFormat.D07_QuantityDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_RemainAMT).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_InventoryQTY).DataField, DxxFormat.D07_QuantityDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_InventoryAMT).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg, arr)
    End Sub



    Private Sub D02F2061_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt Then
        ElseIf e.Control Then
        Else
            Select Case e.KeyCode
                Case Keys.Enter
                    UseEnterAsTab(Me, True)
                Case Keys.F11
                    HotKeyF11(Me, tdbg)
            End Select
        End If
    End Sub

    Dim sFind As String = ""
    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        If FlagAdd Then
            ' Thêm mới thì gán sFind ="" và gán FilterText =""
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFind = ""

        End If
        dtGrid = ReturnDataTable(SQLStoreD02P2071)
        'Cách mới theo chuẩn: TìmD:\LEMONCODE.R400\D02\D02E2240.R400\D02E2240\1.Forms\7.Other\D02F7777.vb kiếm và Liệt kê tất cả luôn luôn sáng Khi(dt.Rows.Count > 0)
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString
        If sFilter.ToString <> "" Then
            If strFind <> "" Then strFind &= " or "
            strFind &= " IsChecked=1"
        End If
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        btnPrint.Enabled = ReturnPermission(Me.Name) >= 1 And tdbg.RowCount > 0
        FooterTotalGrid(tdbg, COL_VoucherNo)
        'FooterSumNew(tdbg, COL_ConvertedAmount, COL_ConvertedQuantity)
    End Sub
    Dim sFilter As New System.Text.StringBuilder()
    'Dim sFilterServer As New System.Text.StringBuilder()
    Dim bRefreshFilter As Boolean = False
    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub
            FilterChangeGrid(tdbg, sFilter) 'Nếu có Lọc khi In
            ReLoadTDBGrid()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        If tdbg.Columns(tdbg.Col).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then
            e.Handled = CheckKeyPress(e.KeyChar)
        ElseIf tdbg.Splits(tdbg.SplitIndex).DisplayColumns(tdbg.Col).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far Then
            e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End If
    End Sub



    'Lưu ý: gọi hàm ResetFilter(tdbg, sFilter, bRefreshFilter) tại btnFilter_Click và tsbListAll_Click
    'Bổ sung vào đầu sự kiện tdbg_DoubleClick(nếu có) câu lệnh If tdbg.RowCount <= 0 OrElse tdbg.FilterActive Then Exit Sub
    Private Function AllowPrint() As Boolean
        Dim dr() As DataRow = dtGrid.Select("IsChecked" & " = 1")
        If dr.Length < 1 Then
            D99C0008.MsgL3(rL3("MSG000010"))
            tdbg.Focus()
            tdbg.SplitIndex = SPLIT0
            tdbg.Col = COL_IsChecked
            tdbg.Row = 0
            Return False
        End If
        Return True
    End Function

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If Not AllowPrint() Then Exit Sub
        Print(Me, "D02F2061")
    End Sub

    Private Sub Print(ByVal form As Form, Optional ByVal sReportTypeID As String = "D02F2061", Optional ByVal ModuleID As String = "02")

        Dim sReportName As String = ""
        Dim sReportPath As String = ""
        Dim sReportTitle As String = "" 'Thêm biến
        Dim sCustomReport As String = "" 'tdbcTranTypeID.Columns("InvoiceForm").Text

        Dim dtPrint As DataTable = dtGrid.DefaultView.ToTable
        dtPrint.DefaultView.RowFilter = "IsChecked = 1"
        Dim file As String = GetReportPathNew(ModuleID, sReportTypeID, sReportName, sCustomReport, sReportPath)
        If sReportName = "" Then Exit Sub
        form.Cursor = Cursors.WaitCursor
        If btnPrint IsNot Nothing Then btnPrint.Enabled = False
        Select Case file.ToLower
            Case "rpt"
                'printReport(sReportName, sReportPath, rl3("caption"), sSQL)' ' Nếu Caption cố định theo Resource
                printReport(sReportName, sReportPath, sReportTitle, dtPrint.DefaultView.ToTable) ' Nếu Caption lấy theo TIêu đề thiết lập bên D89.
            Case "xls", "xlsx"
                Dim sPathFile As String = D99D0541.GetObjectFile(sReportTypeID, sReportName, file, sReportPath)
                If sPathFile = "" Then Exit Select
                MyExcel(dtPrint.DefaultView.ToTable, sPathFile, file, True)
                form.Cursor = Cursors.Default
                If btnPrint IsNot Nothing Then btnPrint.Enabled = True
            Case "doc", "docx"
                Dim sPathFile As String = D99D0541.GetObjectFile(sReportTypeID, sReportName, file, sReportPath)
                If sPathFile = "" Then Exit Select
                CreateWordDocumentCopyTemplate(sPathFile, dtPrint.DefaultView.ToTable)
                OpenFile(sPathFile, False)
        End Select
        form.Cursor = Cursors.Default
        If btnPrint IsNot Nothing Then btnPrint.Enabled = True
    End Sub

    Private Sub printReport(ByVal sReportName As String, ByVal sReportPath As String, ByVal sReportCaption As String, ByVal dt As DataTable)
        Dim report As New D99C1003
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sSQLSub As String = ""
        'UnicodeSubReport(sSubReportName, sSQLSub, gsDivisionID, gbUnicode)
        ' sSQL = SQLStoreD27P2335(sReportName, _contractID)
        With report
            .OpenConnection(conn)
            '.AddSub(sSQLSub, sSubReportName & ".rpt") 'Báo cáo không sử dụng SubReport
            .AddMain(dt)
            .PrintReport(sReportPath, sReportCaption)
        End With
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2071
    '# Created User: KIM LONG
    '# Created Date: 19/10/2016 03:28:32
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2071() As String
        Dim sSQL As String = ""
        sSQL &= ("-- du lieu cho luoi" & vbCrlf)
        sSQL &= "Exec D02P2071 "
        sSQL &= SQLString(_voucherID) & COMMA 'VoucherID, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL ' ID : 262792
        sSQL &= SQLNumber(giTranYear) 'TranYear, int, NOT NULL
        Return sSQL
    End Function

    Dim bSelect As Boolean = False 'Mặc định Uncheck - tùy thuộc dữ liệu database
    Private Sub HeadClick(ByVal iCol As Integer)
        If tdbg.RowCount <= 0 Then Exit Sub
        Select Case iCol
            Case COL_IsChecked
                L3HeadClick(tdbg, iCol, bSelect)
            Case Else
                tdbg.AllowSort = True
        End Select
    End Sub

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.HeadClick
        HeadClick(e.ColIndex)
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.Control And e.KeyCode = Keys.S Then HeadClick(tdbg.Col)
    End Sub




    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
Imports System
Public Class D02F0300

#Region "Const of tdbg - Total of Columns: 41"
    Private Const COL_OrderNum As Integer = 0              ' STT
    Private Const COL_AssetID As Integer = 1               ' Mã tài sản
    Private Const COL_AssetName As Integer = 2             ' Tên tài sản
    Private Const COL_ObjectTypeID As Integer = 3          ' ObjectTypeID
    Private Const COL_ObjectID As Integer = 4              ' ObjectID
    Private Const COL_ObjectName As Integer = 5            ' Bộ phận tiếp nhận
    Private Const COL_FullName As Integer = 6              ' Người tiếp nhận
    Private Const COL_ManagementObjName As Integer = 7     ' Bổ phận quản lý
    Private Const COL_ConvertedAmount As Integer = 8       ' Nguyên tệ
    Private Const COL_NotDEPCurrentCost As Integer = 9     ' Giá trị đất
    Private Const COL_DEPCurrentCost As Integer = 10       ' Giá trị xây dựng
    Private Const COL_AmountDepreciation As Integer = 11   ' Hao mòn lũy kế
    Private Const COL_RemainAmount As Integer = 12         ' Giá trị còn lại
    Private Const COL_DepreciatedPeriod As Integer = 13    ' Số kỳ khấu hao
    Private Const COL_ServiceLife As Integer = 14          ' Số kỳ đã khấu hao
    Private Const COL_BeginUse As Integer = 15             ' Kỳ bắt đầu sử dụng
    Private Const COL_BeginDep As Integer = 16             ' Kỳ bắt đầu khấu hao
    Private Const COL_Percentage As Integer = 17           ' Tỷ lệ khấu hao (%)
    Private Const COL_DepreciatedAmount As Integer = 18    ' Định mức khấu hao
    Private Const COL_IsDepreciated As Integer = 19        ' Đã khấu hao
    Private Const COL_IsDisposed As Integer = 20           ' IsDisposed
    Private Const COL_IsRevalued As Integer = 21           ' IsRevalued
    Private Const COL_CipNo As Integer = 22                ' CipNo
    Private Const COL_SetUpFromDes As Integer = 23         ' Hình thành từ
    Private Const COL_StrSetupVoucherNo As Integer = 24    ' Chứng từ hình thành
    Private Const COL_SetUpFrom As Integer = 25            ' SetUpFrom
    Private Const COL_EmployeeID As Integer = 26           ' Mã nhân viên
    Private Const COL_CreateUserID As Integer = 27         ' CreateUserID
    Private Const COL_CreateDate As Integer = 28           ' CreateDate
    Private Const COL_LastModifyDate As Integer = 29       ' LastModifyDate
    Private Const COL_LastModifyUserID As Integer = 30     ' LastModifyUserID
    Private Const COL_TransactionID As Integer = 31        ' TransactionID
    Private Const COL_ModuleID As Integer = 32             ' ModuleID
    Private Const COL_Period As Integer = 33               ' Period
    Private Const COL_Locked As Integer = 34               ' Khóa
    Private Const COL_BatchID As Integer = 35              ' BatchID
    Private Const COL_TranMonth As Integer = 36            ' TranMonth
    Private Const COL_TranYear As Integer = 37             ' TranYear
    Private Const COL_D54ProjectID As Integer = 38         ' D54ProjectID
    Private Const COL_D27PropertyProductID As Integer = 39 ' D27PropertyProductID
    Private Const COL_VoucherNo As Integer = 40            ' VoucherNo
#End Region


#Region "Const of tdbgD"
    Private Const COLD_VoucherTypeID As String = "VoucherTypeID"       ' Loại phiếu
    Private Const COLD_VoucherNo As String = "VoucherNo"               ' Số phiếu
    Private Const COLD_VoucherDate As String = "VoucherDate"           ' Ngày phiếu
    Private Const COLD_TransactionDate As String = "TransactionDate"   ' Ngày hóa đơn
    Private Const COLD_SeriNo As String = "SeriNo"                     ' Số Sêri
    Private Const COLD_VATTypeID As String = "VATTypeID"               ' Loại hóa đơn
    Private Const COLD_RefNo As String = "RefNo"                       ' Số hóa đơn
    Private Const COLD_Description As String = "Description"           ' Diễn giải
    Private Const COLD_DebitAccountID As String = "DebitAccountID"     ' Tài khoản nợ
    Private Const COLD_CreditAccountID As String = "CreditAccountID"   ' Tài khoản có
    Private Const COLD_OriginalAmount As String = "OriginalAmount"     ' Nguyên giá
    Private Const COLD_ConvertedAmount As String = "ConvertedAmount"   ' Số tiền
    Private Const COLD_SourceID As String = "SourceID"                 ' Nguồn vốn
    Private Const COLD_ObjectTypeID As String = "ObjectTypeID"         ' Mã loại đối tượng
    Private Const COLD_ObjectID As String = "ObjectID"                 ' Mã đối tượng
    Private Const COLD_ObjectName As String = "ObjectName"             ' Tên đối tượng
    Private Const COLD_CurrencyID As String = "CurrencyID"             ' Loại tiền
    Private Const COLD_ExchangeRate As String = "ExchangeRate"         ' Tỷ giá
    Private Const COLD_CreateUserID As String = "CreateUserID"         ' CreateUserID
    Private Const COLD_CreateDate As String = "CreateDate"             ' CreateDate
    Private Const COLD_LastModifyDate As String = "LastModifyDate"     ' LastModifyDate
    Private Const COLD_LastModifyUserID As String = "LastModifyUserID" ' LastModifyUserID
#End Region

    Private _setupFrom As String = "ALL"
    Public WriteOnly Property SetupFrom() As String
        Set(ByVal Value As String)
            _setupFrom = Value
            _esetupFrom = convertStringToEnum(_setupFrom)
        End Set
    End Property

    Dim dtGrid, dtObjectID As DataTable
    Dim sFilter As New System.Text.StringBuilder()
    Dim bRefreshFilter As Boolean = False 'Cờ bật set FilterText =""
    Private usrOption As New D99U1111()
    Dim dtF12 As DataTable
    Private _esetupFrom As enumSetUpFrom = enumSetUpFrom.ALL
    Dim oFilterCombo As Lemon3.Controls.FilterCombo

    Private Sub D02F0300_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If usrOption IsNot Nothing Then usrOption.Dispose()
    End Sub

    Private Sub D09F2250_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me)
                Exit Sub
            Case Keys.F5
                btnFilter_Click(Nothing, Nothing)
            Case Keys.F12
                btnF12_Click(Nothing, Nothing)
            Case Keys.Escape
                usrOption.picClose_Click(Nothing, Nothing)
        End Select
    End Sub

    Dim iPer_F5557 As Integer


    Private Sub D09F2250_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        oFilterCombo = New Lemon3.Controls.FilterCombo
        oFilterCombo.CheckD91 = True
        oFilterCombo.UseFilterCombo(tdbcObjectIDFrom, tdbcObjectIDTo) 'ID-143163        tdbcObjectIDFrom, tdbcObjectIDTo

        iPer_F5557 = ReturnPermission("D02F5557")
        SetShortcutPopupMenu(Me, tbrTableToolStrip, ContextMenuStrip1)
        GetFirstPeriod(gsDivisionID)
        LoadTDBCombo()
        SetBackColorObligatory()
        ResetSplitDividerSize(tdbg)
        tdbg_NumberFormat()
        CallD99U1111()
        ResetGrid()
        Loadlanguage()
        ResetColorGrid(tdbg, SPLIT0, tdbg.Splits.Count - 1)
        ResetColorGrid(tdbgD, SPLIT0, tdbgD.Splits.Count - 1)
        InputDateInTrueDBGrid(tdbgD, COLD_VoucherDate)
        VisibleColumns()
        SetResolutionForm(Me, ContextMenuStrip1)
        LoadtdbcObjectID("-1") 'ID-143163
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub VisibleColumns()
        tdbg.Splits(SPLIT0).DisplayColumns(COL_DEPCurrentCost).Visible = D02Systems.UseProperty
        tdbg.Splits(SPLIT0).DisplayColumns(COL_NotDEPCurrentCost).Visible = D02Systems.UseProperty
    End Sub

    Private Sub CallD99U1111()
        Dim arrColObligatory() As Object = {COL_AssetID, COL_Locked}
        usrOption.AddColVisible(tdbg, dtF12, arrColObligatory)
        If usrOption IsNot Nothing Then usrOption.Dispose()
        usrOption = New D99U1111(Me, tdbg, dtF12)
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_sach_tai_san_co_dinh_duoc_hinh_thanh_-_D02F0300") & UnicodeCaption(gbUnicode) 'Danh sÀch tªi s¶n cç ¢Ünh ¢§íc hØnh thªnh - D02F0300

        lblSetUpFrom.Text = rl3("Nguon_goc_hinh_thanh")
        lblAssetIDFrom.Text = rl3("Tai_san")
        chkIsPeriod.Text = rl3("Ky")
        chkDate.Text = rL3("Ngay")
        lblObjectTypeID.Text = rL3("Bo_phan_tiep_nhan")
        '================================================================ 
        tsbPrintPrint.Text = tsbPrint.Text
        tsmPrintPrint.Text = tsbPrint.Text
        mnsPrintPrint.Text = tsbPrint.Text

        tsbPrintReport.Text = rl3("In_bien_ban_giao_nhan")
        tsmPrintReport.Text = tsbPrintReport.Text
        mnsPrintReport.Text = tsbPrintReport.Text

        tsbPrintEquipmentAttachOfAsset.Text = rL3("In_thiet_bi_dinh_kem_tai_san") 'In thiết bị đính kèm tài sản
        tsmPrintEquipmentAttachOfAsset.Text = tsbPrintEquipmentAttachOfAsset.Text
        mnsPrintEquipmentAttachOfAsset.Text = tsbPrintEquipmentAttachOfAsset.Text
        '================================================================ 
        btnFilter.Text = rl3("Loc") & " (F5)" 'Lọc
        btnF12.Text = rl3("Hien_thi") & " (F12)" 'F12
        '================================================================ 
        tdbcAssetIDFrom.Columns("AssetID").Caption = rl3("Ma") 'Mã
        tdbcAssetIDFrom.Columns("AssetName").Caption = rl3("Ten") 'Tên
        tdbcAssetIDTo.Columns("AssetID").Caption = rl3("Ma") 'Mã
        tdbcAssetIDTo.Columns("AssetName").Caption = rl3("Ten") 'Tên
        tdbcSetUpFrom.Columns("SetUpFrom").Caption = rl3("Ma") 'Mã
        tdbcSetUpFrom.Columns("Description").Caption = rl3("Dien_giai") 'Tên
        '================================================================ 
        tdbg.Columns("").Caption = rl3("STT") 'STT
        tdbg.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
        tdbg.Columns("ObjectName").Caption = rL3("Bo_phan_tiep_nhan") 'Bộ phận quản lý
        tdbg.Columns("FullName").Caption = rl3("Nguoi_tiep_nhan") 'Người tiếp nhận
        tdbg.Columns("ConvertedAmount").Caption = rL3("Nguyen_gia") 'Nguyên giá
        tdbg.Columns("NotDEPCurrentCost").Caption = rL3("Gia_tri_dat")
        tdbg.Columns("DEPCurrentCost").Caption = rL3("Gia_tri_xay_dung")
        tdbg.Columns("AmountDepreciation").Caption = rl3("Hao_mon_luy_ke") 'Hao mòn lũy kế
        tdbg.Columns("RemainAmount").Caption = rl3("Gia_tri_con_lai") 'Giá trị còn lại
        tdbg.Columns("DepreciatedPeriod").Caption = rl3("So_ky_da_khau_hao") 'Số kỳ khấu hao
        tdbg.Columns("ServiceLife").Caption = rl3("So_ky_khau_hao") 'Số kỳ đã khấu hao
        tdbg.Columns("BeginUse").Caption = rl3("Ky_bat_dau_su_dung") 'Kỳ bắt đầu sử dụng
        tdbg.Columns("BeginDep").Caption = rl3("Ky_bat_dau_khau_hao") 'Kỳ bắt đầu khấu hao
        tdbg.Columns("Percentage").Caption = rl3("Ty_le_khau_hao") & " (%)" 'Tỷ lệ khấu hao (%)
        tdbg.Columns("DepreciatedAmount").Caption = rl3("Dinh_muc_khau_hao") 'Định mức khấu hao
        tdbg.Columns("IsDepreciated").Caption = rl3("Da_khau_hao") 'Đã khấu hao
        tdbg.Columns("SetUpFromDes").Caption = rl3("Hinh_thanh_tu") 'Hình thành từ
        tdbg.Columns("EmployeeID").Caption = rl3("Ma_nhan_vien") 'Mã nhân viên
        tdbg.Columns(COL_Locked).Caption = rL3("Khoa") 'Khóa

        tdbg.Columns(COL_StrSetupVoucherNo).Caption = rL3("Chung_tu_hinh_thanh") 'Chứng từ hình thành
        '================================================================ 

        tdbg.Columns(COL_ManagementObjName).Caption = rL3("Bo_phan_quan_ly") 'Bổ phận quản lý

        tdbgD.Columns("VoucherTypeID").Caption = rL3("Loai_phieu") 'Loại phiếu
        tdbgD.Columns("VoucherNo").Caption = rl3("So_phieu") 'Số phiếu
        tdbgD.Columns("VoucherDate").Caption = rl3("Ngay_phieu") 'Ngày phiếu
        tdbgD.Columns("TransactionDate").Caption = rl3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbgD.Columns("SeriNo").Caption = rl3("So_Seri") 'Số Sêri
        tdbgD.Columns("VATTypeID").Caption = rl3("Loai_hoa_don") 'Loại hóa đơn
        tdbgD.Columns("RefNo").Caption = rl3("So_hoa_don") 'Số hóa đơn
        tdbgD.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbgD.Columns("DebitAccountID").Caption = rl3("Tai_khoan_no") 'Tài khoản nợ
        tdbgD.Columns("CreditAccountID").Caption = rl3("Tai_khoan_co") 'Tài khoản có
        tdbgD.Columns("OriginalAmount").Caption = rl3("Nguyen_gia") 'Nguyên giá
        tdbgD.Columns("ConvertedAmount").Caption = rl3("So_tien") 'Số tiền
        tdbgD.Columns("SourceID").Caption = rl3("Nguon_von") 'Nguồn vốn
        tdbgD.Columns("ObjectTypeID").Caption = rl3("Ma_loai_doi_tuong") 'Mã loại đối tượng
        tdbgD.Columns("ObjectID").Caption = rl3("Ma_doi_tuong") 'Mã đối tượng
        tdbgD.Columns("ObjectName").Caption = rl3("Ten_doi_tuong") 'Tên đối tượng
        tdbgD.Columns("CurrencyID").Caption = rl3("Loai_tien") 'Loại tiền
        tdbgD.Columns("ExchangeRate").Caption = rl3("Ty_gia") 'Tỷ giá
        '===============================================
        'Add 20/12/2013 - ID 62126 
        mnsAddNEW.Text = rl3("Mua_moi") 'Mua mới
        mnsAddCIP.Text = rl3("Tu_XDCB") 'Từ XDCB
        mnsAddBAL.Text = rL3("Nhap_so_du") 'Nhập số dư
        mnsAddNEWDD.Text = rL3("Tu_dieu_dong_von") 'Mua mới
        tsmAddNEW.Text = rl3("Mua_moi") 'Mua mới
        tsmAddCIP.Text = rl3("Tu_XDCB") 'Từ XDCB
        tsmAddBAL.Text = rl3("Nhap_so_du") 'Nhập số dư
        tsbAddNEW.Text = rl3("Mua_moi") 'Mua mới
        tsbAddCIP.Text = rl3("Tu_XDCB") 'Từ XDCB
        tsbAddBAL.Text = rL3("Nhap_so_du") 'Nhập số dư
        tsbAddNewDD.Text = rL3("Tu_dieu_dong_von") 'Mua mới
        tsmAddNewDD.Text = tsbAddNewDD.Text
        mnsEditOther.Text = rL3("Sua_khac")


       
     

    End Sub

#Region "LoadTDBGrid"

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0500
    '# Created User: THANHHUYEN
    '# Created Date: 20/12/2013 10:12:16
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0500() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon luoi" & vbCrlf)
        sSQL &= "Exec D02P0500 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcSetUpFrom)) & COMMA 'SetUpFrom, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString("") & COMMA 'strFind, varchar[8000], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        If chkIsPeriod.Checked And Not chkDate.Checked Then
            sSQL &= SQLNumber(1) & COMMA 'IsTime, tinyint, NOT NULL
        ElseIf chkDate.Checked And Not chkIsPeriod.Checked Then
            sSQL &= SQLNumber(2) & COMMA 'IsTime, tinyint, NOT NULL
        ElseIf chkIsPeriod.Checked And chkDate.Checked Then
            sSQL &= SQLNumber(3) & COMMA 'IsTime, tinyint, NOT NULL
        Else
            sSQL &= SQLNumber(4) & COMMA 'IsTime, tinyint, NOT NULL
        End If
        sSQL &= SQLNumber(tdbcPeriodFrom.Columns("TranMonth").Value) & COMMA 'FromMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(tdbcPeriodFrom.Columns("TranYear").Value) & COMMA 'FromYear, int, NOT NULL
        sSQL &= SQLNumber(tdbcPeriodTo.Columns("TranMonth").Value) & COMMA 'ToMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(tdbcPeriodTo.Columns("TranYear").Value) & COMMA 'ToYear, int, NOT NULL
        sSQL &= SQLDateSave(c1dateDateFrom.Value) & COMMA 'DateFrom, datetime, NOT NULL
        sSQL &= SQLDateSave(c1dateDateTo.Value) & COMMA 'DateTo, datetime, NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetIDFrom)) & COMMA 'AssetIDFrom, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetIDTo)) & COMMA 'AssetIDTo, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcObjectTypeID)) & COMMA 'ObjectTypeID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcObjectIDFrom)) & COMMA 'ObjectIDFrom, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcObjectIDTo)) 'ObjectIDTo, varchar[20], NOT NULL
        Return sSQL
    End Function



    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKeyID As String = "")
        Dim sSQL As String = " SET NOCOUNT ON" & vbCrLf
        sSQL &= SQLStoreD02P0500()
        ' sSQL &= "Exec D02P0500 'KVQN', 'NEW', 3, 2011, '', '84', 1"
        dtGrid = ReturnDataTable(sSQL)
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()
        If sKeyID <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("AssetID =" & SQLString(sKeyID), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0)) 'dùng tdbg.Bookmark có thể không đúng
            If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
        End If

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLSelectD02T0012
    '# Created User: 
    '# Created Date: 15/11/2011 02:07:46
    '# Description: Load lưới Detail
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLSelectD02T0012(ByVal sAssetID As String) As StringBuilder
        Dim sSQL As New StringBuilder
        Dim sUnicode As String = ""
        If gbUnicode Then sUnicode = "U"
        sSQL.Append("Select VoucherTypeID, VoucherNo, VoucherDate")
        sSQL.Append(", TransactionDate, Description" & sUnicode & " as Description, CurrencyID, ExchangeRate, DebitAccountID")
        sSQL.Append(", CreditAccountID, OriginalAmount, ConvertedAmount")
        sSQL.Append(", RefNo, RefDate, Disabled, CreateUserID, CreateDate")
        sSQL.Append(", LastModifyUserID, LastModifyDate, SeriNo, ObjectTypeID, ObjectID, SourceID")
        sSQL.Append(", ObjectName" & sUnicode & " as ObjectName, VATTypeID, DebitObjectTypeID, DebitObjectID")
        sSQL.Append(", Ana01ID, Ana02ID, Ana03ID, Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID")
        sSQL.Append(" From D02T0012 WITH(NOLOCK)" & vbCrLf)
        sSQL.Append(" Where ")
        sSQL.Append("TransactionTypeID IN('MM', 'SD', 'XDCB', 'SDKH', 'SDMM') ")
        sSQL.Append(" And TranMonth = " & SQLNumber(tdbg.Columns(COL_TranMonth).Text))
        sSQL.Append(" And TranYear = " & SQLNumber(tdbg.Columns(COL_TranYear).Text))
        sSQL.Append(" And DivisionID = " & SQLString(gsDivisionID))
        sSQL.Append(" And AssetID =  " & SQLString(sAssetID))
        Return sSQL
    End Function

    Private Sub LoadTDBGridDetail()
        Dim sSQL As String = SQLSelectD02T0012(tdbg.Columns(COL_AssetID).Text).ToString
        LoadDataSource(tdbgD, sSQL, gbUnicode)
        ResetGridDetail()
    End Sub


    Private Sub ResetGridDetail()
        FooterTotalGrid(tdbgD, COLD_VoucherNo)
        FooterSumNew(tdbgD, COLD_ConvertedAmount, COLD_OriginalAmount)
    End Sub

    Private Sub tdbg_NumberFormat()
        tdbg.Columns(COL_OrderNum).NumberFormat = DxxFormat.DefaultNumber0
        tdbg.Columns(COL_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_AmountDepreciation).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_RemainAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_DepreciatedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_DEPCurrentCost).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_NotDEPCurrentCost).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbgD_NumberFormat()
    End Sub

    Private Sub tdbgD_NumberFormat()
        tdbgD.Columns(COLD_OriginalAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbgD.Columns(COLD_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    End Sub

#End Region

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""
    Dim dtCaptionCols As DataTable

    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, mnsFind.Click, tsmFind.Click
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If

        ShowFindDialogClient(Finder, dtCaptionCols, Me.Name, "0", gbUnicode)

    End Sub

    Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
        If ResultWhereClause Is Nothing Or ResultWhereClause.ToString = "" Then Exit Sub
        sFind = ResultWhereClause.ToString()
        ReLoadTDBGrid()
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, mnsListAll.Click, tsmListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString

        dtGrid.DefaultView.RowFilter = strFind
        If tdbgD.Visible Then LoadTDBGridDetail()
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        CheckMenu(Me.Name, tbrTableToolStrip, tdbg.RowCount, gbEnabledUseFind, True, ContextMenuStrip1)
        tsmPrintReport.Enabled = tsbPrint.Enabled
        mnsPrintReport.Enabled = tsmPrintReport.Enabled
        'Modify 20/12/2013 - ID 62126 
        FooterTotalGrid(tdbg, COL_AssetID)
        'tdbg.Columns(COL_AssetID).FooterText = tdbg.RowCount & Space(1) & rl3("_tai_san")
        FooterSumNew(tdbg, COL_ConvertedAmount, COL_AmountDepreciation, COL_RemainAmount, COL_DEPCurrentCost, COL_NotDEPCurrentCost)
        '============================
        mnsAdd.Enabled = True
        tsmAdd.Enabled = True
        tsbAdd.Enabled = True
        mnsAddNEW.Enabled = ReturnPermission("D02F0300") >= EnumPermission.Add And Not gbClosed
        tsmAddNEW.Enabled = mnsAddNEW.Enabled
        tsbAddNEW.Enabled = mnsAddNEW.Enabled
        tsbAddNewDD.Enabled = mnsAddNEW.Enabled
        mnsAddNEWDD.Enabled = mnsAddNEW.Enabled
        tsmAddNewDD.Enabled = mnsAddNEWDD.Enabled
        mnsAddCIP.Enabled = mnsAddNEW.Enabled
        tsmAddCIP.Enabled = mnsAddNEW.Enabled
        tsbAddCIP.Enabled = mnsAddNEW.Enabled
        If giFirstTranMonth = giTranMonth And giFirstTranYear = giTranYear Then
            mnsAddBAL.Enabled = ReturnPermission("D02F1002") >= EnumPermission.Add And Not gbClosed
            tsmAddBAL.Enabled = mnsAddBAL.Enabled
            tsbAddBAL.Enabled = mnsAddBAL.Enabled
        Else
            mnsAddBAL.Enabled = False
            tsmAddBAL.Enabled = False
            tsbAddBAL.Enabled = False
        End If
        'ID 92038 01.12.2016
        mnsEditOther.Enabled = ReturnPermission("D02F0300") > 2 And tdbg.RowCount > 0 And Not gbClosed
    End Sub

#End Region

#Region "Menu Bar"

    Private Sub mnsAddNEW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsAddNEW.Click, tsmAddNEW.Click, tsbAddNEW.Click
        Dim frm As New D02F1001
        With frm
            .FormState = EnumFormState.FormAdd
            .ShowDialog()

            If gbSavedOK Then
                Dim sKey As String = frm.AssetID
                LoadTDBGrid(True, sKey)
            End If
            .Dispose()
        End With
    End Sub

    Private Sub mnsAddCIP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsAddCIP.Click, tsmAddCIP.Click, tsbAddCIP.Click
        Dim frm As New D02F1010
        With frm
            .FormState = EnumFormState.FormAdd
            .ShowDialog()

            If gbSavedOK Then
                Dim sKey As String = frm.AssetID
                LoadTDBGrid(True, sKey)
            End If
            .Dispose()
        End With
    End Sub

    Private Sub mnsAddBAL_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsAddBAL.Click, tsmAddBAL.Click, tsbAddBAL.Click
        Dim frm As New D02F1002
        With frm
            .FormState = EnumFormState.FormAdd
            .ShowDialog()

            If gbSavedOK Then
                Dim sKey As String = frm.AssetID
                LoadTDBGrid(True, sKey)
            End If
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, mnsEdit.Click, tsmEdit.Click
        If Not CheckThroughPeriod(tdbg.Columns(COL_TranMonth).Text, tdbg.Columns(COL_TranYear).Text) Then Exit Sub
        Dim iStatus As Integer = 0
        If Not CheckStore(SQLStoreD02P0402(tdbg.Columns(COL_AssetID).Text), True) Then Exit Sub
        Dim dtStatus As DataTable = ReturnDataTable(SQLStoreD02P0402(tdbg.Columns(COL_AssetID).Text))
        If dtStatus.Rows.Count > 0 Then
            iStatus = L3Int(dtStatus.Rows(0).Item("Status").ToString)
        End If
        If L3Bool(tdbg.Columns(COL_Locked).Text) Then
            D99C0008.MsgL3(rl3("MSG000003") & Space(1) & rl3("MSG000023")) 'Phieu_nay_da_duoc_khoa_Ban_khong_duoc_phep_xoa"))
            Exit Sub
        End If
        Select Case convertStringToEnum(tdbg.Columns(COL_SetUpFrom).Text)
            Case enumSetUpFrom.ALL

            Case enumSetUpFrom.[NEW]
                Dim frm As New D02F1001
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    If iStatus = 1 Then
                        .FormState = EnumFormState.FormView
                    Else
                        .FormState = EnumFormState.FormEdit
                    End If

                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.CIP
                Dim frm As New D02F1010
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .D54ProjectID = tdbg.Columns(COL_D54ProjectID).Text
                    .D27PropertyProductID = tdbg.Columns(COL_D27PropertyProductID).Text
                    If iStatus = 1 Then
                        .FormState = EnumFormState.FormView
                    Else
                        .FormState = EnumFormState.FormEdit
                    End If
                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.BAL
                Dim frm As New D02F1002
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    If iStatus = 1 Then
                        .FormState = EnumFormState.FormView
                    Else
                        .FormState = EnumFormState.FormEdit
                    End If
                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.CAP
                Dim frm As New D02F1009
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    If iStatus = 1 Then
                        .FormState = EnumFormState.FormView
                    Else
                        .FormState = EnumFormState.FormEdit
                    End If
                    .ShowDialog()
                    .Dispose()
                End With
        End Select

        If gbSavedOK Then LoadTDBGrid(False, tdbg.Columns(COL_AssetID).Text)

    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Select Case convertStringToEnum(tdbg.Columns(COL_SetUpFrom).Text)
            Case enumSetUpFrom.ALL

            Case enumSetUpFrom.[NEW]
                Dim frm As New D02F1001
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .FormState = EnumFormState.FormView
                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.CIP
                Dim frm As New D02F1010
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .FormState = EnumFormState.FormView
                    .D54ProjectID = tdbg.Columns(COL_D54ProjectID).Text
                    .D27PropertyProductID = tdbg.Columns(COL_D27PropertyProductID).Text
                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.BAL
                Dim frm As New D02F1002
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .FormState = EnumFormState.FormView
                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.CAP
                Dim frm As New D02F1009
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .FormState = EnumFormState.FormView
                    .ShowDialog()
                    .Dispose()
                End With
        End Select

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0012
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 15/11/2011 02:47:34
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0012(ByVal sAssetID As String) As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T0012"
        sSQL &= " Where "
        sSQL &= "AssetID = " & SQLString(sAssetID) & " And "
        sSQL &= "DivisionID = " & SQLString(gsDivisionID) & " And "
        sSQL &= "ModuleID = '02' And "
        sSQL &= "IsNull(TransactionTypeID,'') in ('SD', 'SDKH', 'XDCB')"
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0100
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 17/11/2011 08:28:59
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0100(ByVal sAssetID As String) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0100 Set ")
        sSQL.Append("AssetID = '', Status = 1")
        sSQL.Append(" Where ")
        sSQL.Append("ISNULL(CipID,'') In (Select CipID from D02T0012 A WITH(NOLOCK) Where TransactionTypeID = 'XDCB' AND MODULEID = '02'" & _
                                " AND ASSETID = " & SQLString(sAssetID) & _
                                " AND Not exists (SELECT CipID From D02T0012 B WITH(NOLOCK) WHERE AssetID <> " & SQLString(sAssetID) & _
                                                  " AND TransactionTypeID = 'XDCB' AND ModuleID = '02' AND A.CipID = B.CIPID))")
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0012
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 17/11/2011 08:35:40
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0012_CIP(ByVal sAssetID As String) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0012 Set Status = 0")
        sSQL.Append(" Where ")
        sSQL.Append("DivisionID = " & SQLString(gsDivisionID) & " And ")
        sSQL.Append("ISNULL(CipID,'') In (Select CipID from D02T0012 A WITH(NOLOCK) Where TransactionTypeID = 'XDCB' AND MODULEID = '02'" & _
                                      " AND ASSETID = " & SQLString(sAssetID) & _
                                      " AND Not exists (SELECT CipID From D02T0012 B WITH(NOLOCK) WHERE AssetID <> " & SQLString(sAssetID) & _
                                                        " AND ModuleID = '02' AND A.CipID = B.CIPID))")
        Return sSQL
    End Function

    Private Function SQLUpdateD02T0012_NEW(ByVal sAssetID As String) As StringBuilder
        
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0012 Set AssetID = '', Status = 0, TransactiontypeID = ''")
        If Not D02Systems.CIPforPropertyProduct Then
            sSQL.Append(", IsNotAllocate = 0 ")
        End If
        sSQL.Append(" Where ")
        sSQL.Append("DivisionID = " & SQLString(gsDivisionID) & " And ")
        sSQL.Append("AssetID = " & SQLString(sAssetID) & " AND IsNull(TransactionTypeID, '') in ('MM', 'SDMM','MMDD', 'SDMMDD')")
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD91T9111
    '# Created User: Phạm Văn Vinh
    '# Created Date: 06/11/2012 08:47:34
    '# Modified User: 
    '# Modified Date: 
    '# Description: Them theo incident 51998 của Thị Hiệp bởi Văn Vinh
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD91T9111(ByVal sAssetID As String) As String
        Dim sSQL As String = ""
        sSQL &= "Delete from D91T9111 "
        sSQL &= " where  "
        sSQL &= "VoucherNo in (select distinct VoucherNo from D02T0012 WITH(NOLOCK) where "
        sSQL &= " AssetID = " & SQLString(sAssetID) & " and"
        sSQL &= " DivisionID = " & SQLString(gsDivisionID) & " And "
        sSQL &= " ModuleID = '02' And "
        sSQL &= " IsNull(TransactionTypeID,'') = 'XDCB') and "
        sSQL &= " VoucherTableName = 'D02T0012'"
        Return sSQL
    End Function

    'Câu SQLDelete của 1 dòng trên lưới
    Private Function AllowDelete(ByVal sAssetID As String, ByVal eSetUpFrom As enumSetUpFrom, ByRef sSQL As StringBuilder) As Boolean
        Dim dtTemp As DataTable = ReturnDataTable(SQLStoreD02P0402(sAssetID))
        If dtTemp.Rows.Count = 0 Then GoTo 1
        If dtTemp.Rows(0).Item("Status").ToString = "1" Then
            If D99C0008.MsgAsk(ConvertVietwareFToUnicode(dtTemp.Rows(0).Item("Message").ToString)) = Windows.Forms.DialogResult.Yes Then
                tsbView_Click(Nothing, Nothing)
            End If
            Return False
        End If

1:

        'Incident 71702 bỏ đoạn xóa thay thay thế bằng store SQLStoreD02P0410
        If eSetUpFrom = enumSetUpFrom.CIP Then 'XDCB
            'sSQL.Append(SQLUpdateD02T0100(sAssetID).ToString & vbCrLf)
            'sSQL.Append(SQLUpdateD02T0012_CIP(sAssetID).ToString & vbCrLf)
            ''Them ngay 6/11/2012 theo incident 51998 của Thị Hiệp bởi Văn Vinh
            'sSQL.Append(SQLDeleteD91T9111(sAssetID) & vbCrLf)
            sSQL.Append(SQLStoreD02P0410(sAssetID).ToString & vbCrLf)
        End If
        sSQL.Append(SQLDeleteD02T0012(sAssetID) & vbCrLf)
        If eSetUpFrom = enumSetUpFrom.[NEW] Or eSetUpFrom = enumSetUpFrom.CAP Then sSQL.Append(SQLUpdateD02T0012_NEW(sAssetID).ToString & vbCrLf)
        sSQL.Append(" Update D02T0001 Set IsCompleted = 0 , SetUpFrom ='', ConvertedAmount = 0, AmountDepreciation = 0 , RemainAmount = 0,  ServiceLife = 0, Percentage = 0, DepreciatedPeriod = 0,  DepreciatedAmount = 0, D54ProjectID = '', D27PropertyProductID = '' " & " Where AssetID = " & SQLString(sAssetID) & vbCrLf)
        sSQL.Append(" Delete D02T5000 Where (AssetID = " & SQLString(sAssetID) & ") And  DivisionID = " & SQLString(gsDivisionID) & vbCrLf)
        Return True
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0410
    '# Created User: HUỲNH KHANH
    '# Created Date: 25/01/2015 10:26:36
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0410(ByVal sAssetID As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Xoa" & vbCrLf)
        sSQL &= "Exec D02P0410 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(sAssetID) & COMMA 'AssetID, varchar[50], NOT NULL
        sSQL &= SQLString("02") & COMMA 'ModuleID, varchar[50], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_D54ProjectID).Text) & COMMA 'ProjectID, varchar[50], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_D27PropertyProductID).Text) & COMMA 'PropertyProductID, varchar[50], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_VoucherNo).Text) & COMMA
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) 'TranYear, int, NOT NULL
        Return sSQL
    End Function


    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        If D99C0008.MsgAskDeleteRow() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not CheckThroughPeriod(tdbg.Columns(COL_TranMonth).Text, tdbg.Columns(COL_TranYear).Text) Then Exit Sub

        If L3Bool(tdbg.Columns(COL_Locked).Text) Then
            D99C0008.MsgL3(rl3("MSG000003") & Space(1) & rl3("MSG000023")) 'Phieu_nay_da_duoc_khoa_Ban_khong_duoc_phep_xoa"))
            Exit Sub
        End If

        Dim sSQL As New StringBuilder
        Dim tdbgSelectedRow As C1.Win.C1TrueDBGrid.SelectedRowCollection = tdbg.SelectedRows
        Dim i As Integer = 0
        Dim myAL As New ArrayList() 'Tạo mảng lưu lại chỉ số vừa chọn 
        If tdbgSelectedRow.Count > 1 Then
            For i = 0 To tdbgSelectedRow.Count - 1
                myAL.Add(tdbgSelectedRow.Item(i))
                If Not AllowDelete(tdbg(tdbgSelectedRow.Item(i), COL_AssetID).ToString, convertStringToEnum(tdbg(tdbgSelectedRow.Item(i), COL_SetUpFrom).ToString), sSQL) Then Exit Sub
            Next
            myAL.Sort() 'Sắp xếp tăng dần 
        Else
            If Not AllowDelete(tdbg.Columns(COL_AssetID).Text, convertStringToEnum(tdbg.Columns(COL_SetUpFrom).Text), sSQL) Then Exit Sub
        End If
        If sSQL.ToString = "" Then Exit Sub

        sSQL.Append(vbCrLf & SQLDeleteD02T5010()) '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ

        Dim bResult As Boolean = ExecuteSQL(sSQL.ToString)
        If bResult Then
            DeleteOK()
            DeleteVoucherNoD91T9111(tdbg.Columns(COL_VoucherNo).Text, "D02T0012", "VoucherNo")
            If tdbgSelectedRow.Count > 1 Then
                For i = myAL.Count - 1 To 0 Step -1
                    tdbg.Delete(CInt(myAL.Item(i)))
                Next
            Else
                tdbg.Delete(tdbg.Bookmark)
            End If
            dtGrid.AcceptChanges()
            gbEnabledUseFind = dtGrid.Rows.Count > 0
            ResetGrid()
            Dim sAuditLog As String = ""
            Select Case _esetupFrom
                Case enumSetUpFrom.[NEW]
                    sAuditLog = "PurAsset"
                Case enumSetUpFrom.CIP
                    sAuditLog = "CIPToAsset"
                Case enumSetUpFrom.BAL
                    sAuditLog = "Opening02"
                Case enumSetUpFrom.CAP
                    sAuditLog = "CapAsset"
            End Select
            'RunAuditLog(sAuditLog, "03", tdbg.Columns(COL_AssetID).Text, tdbg.Columns(COL_AssetName).Text)
            Lemon3.D91.RunAuditLog("02", sAuditLog, "03", tdbg.Columns(COL_AssetID).Text, tdbg.Columns(COL_AssetName).Text)
        Else

            DeleteNotOK()
        End If

    End Sub

    Private Sub tsbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbSysInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, mnsSysInfo.Click, tsmSysInfo.Click
        ShowSysInfoDialog(tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub
#End Region

    Dim dtDetail As DataTable
    Dim sFilDetail As New System.Text.StringBuilder()
    Dim bRefreshFilDetail As Boolean = False 'Cờ bật set FilterText =""
    Private Sub tdbgD_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbgD.FilterChange
        Try
            If (dtDetail Is Nothing) Then Exit Sub
            If bRefreshFilDetail Then Exit Sub 'set FilterText ="" thì thoát
            'Filter the data 
            FilterChangeGrid(tdbgD, sFilDetail)
            dtDetail.DefaultView.RowFilter = sFilter.ToString
            ResetGridDetail()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

#Region "Events tdbg"

    Private Sub tdbg_FetchCellTips(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.FetchCellTipsEventArgs) Handles tdbg.FetchCellTips
        e.CellTip = rl3("Su_dung_Alt__V_de_xem_chi_tiet") '"Sử dụng Alt + V để xem chi tiết"
    End Sub

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            'Filter the data 
            FilterChangeGrid(tdbg, sFilter)
            ReLoadTDBGrid()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        Me.Cursor = Cursors.WaitCursor
        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub HotKeyAllV(ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.Alt And e.KeyCode = Keys.V Then 'Hiển thị chi tiết
            tdbgD.Visible = Not tdbgD.Visible
            If tdbgD.Visible Then
                tdbg.Height = Math.Abs(tdbg.Height - tdbgD.Height)
                LoadTDBGridDetail()
            Else
                tdbg.Height = Math.Abs(tdbg.Height + tdbgD.Height)
            End If
        End If
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        HotKeyAllV(e)
        Me.Cursor = Cursors.WaitCursor
        If e.KeyCode = Keys.Enter Then
            If tdbg.FilterActive Then Me.Cursor = Cursors.Default : Exit Sub
            If tsbEdit.Enabled Then
                tsbEdit_Click(sender, Nothing)
            ElseIf tsbView.Enabled Then
                tsbView_Click(sender, Nothing)
            End If
        End If
        HotKeyCtrlVOnGrid(tdbg, e) 'Nhấn Ctrl + V trên lưới 'có trong D99X0000
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Col
            Case COL_OrderNum
                e.Handled = CheckKeyPress(e.KeyChar, True)
            Case COL_IsDepreciated, COL_IsDisposed, COL_IsRevalued, COL_Locked  'Chặn Ctrl + V trên cột Check
                e.Handled = CheckKeyPress(e.KeyChar)
            Case COL_ConvertedAmount, COL_RemainAmount, COL_AmountDepreciation, COL_DepreciatedAmount
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

#End Region

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0402
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 15/11/2011 02:40:59
    '# Modified User: 
    '# Modified Date: 
    '# Description:  Kiểm tra trước khi sửa/Xóa
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0402(ByVal sAssetID As String) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0402 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(sAssetID) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) 'Language, tinyint, NOT NULL
        Return sSQL
    End Function


    Private Sub tdbg_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        If tdbg.RowCount = 0 Then Exit Sub
        If tdbgD.Visible Then LoadTDBGridDetail()

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0301
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 15/11/2011 03:39:30
    '# Modified User: 
    '# Modified Date: 
    '# Description: In main
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0301(ByVal sReportTypeID As String, ByVal sReportID As String) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0301 "
        sSQL &= SQLString(tdbg.Columns(COL_AssetID).Text) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= "N" & SQLString(tdbg.Columns(COL_AssetName).Text) & COMMA 'AssetName, varchar[250], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_BeginUse).Text) & COMMA 'BeginUse, varchar[20], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_BeginDep).Text) & COMMA 'BeginDep, varchar[20], NOT NULL
        sSQL &= SQLMoney(tdbg.Columns(COL_DepreciatedAmount).Text, DxxFormat.D90_ConvertedDecimals) & COMMA 'DepreciatedAmount, money, NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_SetUpFrom).Text) & COMMA 'SetupFrom, varchar[20], NOT NULL
        sSQL &= SQLNumber(tdbg.Columns(COL_DepreciatedPeriod).Text) & COMMA 'DepreciatedPeriod, int, NOT NULL
        sSQL &= SQLNumber(tdbg.Columns(COL_ServiceLife).Text) & COMMA 'ServiceLife, int, NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_ObjectTypeID).Text) & COMMA 'ObjectTypeID, varchar[20], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_ObjectID).Text) & COMMA 'ObjectID, varchar[20], NOT NULL
        sSQL &= "N" & SQLString(tdbg.Columns(COL_ObjectName).Text) & COMMA 'ObjectName, varchar[250], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_EmployeeID).Text) & COMMA 'EmployeeID, varchar[20], NOT NULL
        sSQL &= "N" & SQLString(tdbg.Columns(COL_FullName).Text) & COMMA 'FullName, varchar[250], NOT NULL
        sSQL &= SQLMoney(tdbg.Columns(COL_ConvertedAmount).Text, DxxFormat.D90_ConvertedDecimals) & COMMA 'ConvertedAmount, money, NOT NULL
        sSQL &= SQLMoney(tdbg.Columns(COL_AmountDepreciation).Text, DxxFormat.D90_ConvertedDecimals) & COMMA 'AmountDepreciation, money, NOT NULL
        sSQL &= SQLMoney(tdbg.Columns(COL_RemainAmount).Text, DxxFormat.D90_ConvertedDecimals) & COMMA 'RemainAmount, money, NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(sReportTypeID) & COMMA 'ReportTypeID, varchar[20], NOT NULL
        sSQL &= SQLString(sReportID) & COMMA 'ReportID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode)
        Return sSQL
    End Function


    Private Sub PrintReport(ByVal sReportName As String, ByVal sReportTypeID As String, ByVal sSubReportName As String, ByVal sTitle As String)
        Dim sPathReport As String = ""
        sReportName = GetReportPath(sReportTypeID, sReportName, "", sPathReport) 'Gọi form chọn đường dẫn báo cáo
        If sReportName = "" Then Exit Sub
        Me.Cursor = Cursors.WaitCursor

        Dim report As New D99C1003
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportCaption As String = ""

        Dim sSQL As String = ""
        Dim sSQLSub As String = "Select * From D91V0016 Where DivisionID = " & SQLString(gsDivisionID)

        'Gán giá trị cho sSubReportName và sSQLSub
        UnicodeSubReport(sSubReportName, sSQLSub, gsDivisionID, gbUnicode)
        sReportCaption = sTitle & " - " & sReportName
        'Gán giá trị cho sPathReport
        sSQL = SQLStoreD02P0301(sReportTypeID, sReportName)
        With report
            .OpenConnection(conn)
            ' .AddParameter("?????", "?????")
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(sSQL)
            .PrintReport(sPathReport, sReportCaption)
        End With
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub mnsPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsPrintPrint.Click, tsbPrintPrint.Click, tsmPrintPrint.Click
        PrintReport("D02R0300", Me.Name, "D02R0000", rl3("Danh_muc_hinh_thanh_tai_san_co_dinh"))
    End Sub

    Private Sub mnsPrintReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsPrintReport.Click, tsbPrintReport.Click, tsmPrintReport.Click
        PrintReport("D02R0301", "D02F0301", "D91R0000", rl3("Bien_ban_giao_nhan"))
    End Sub

    Private Sub tdbg_UnboundColumnFetch(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.UnboundColumnFetchEventArgs) Handles tdbg.UnboundColumnFetch
        e.Value = (e.Row + 1).ToString
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P5559
    '# Created User: VANVINH
    '# Created Date: 14/11/2012 11:30:31
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P5559(ByVal StrBatchID As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Cap nhat Khoa phieu" & vbCrLf)
        sSQL &= "Exec D02P5559 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLNumber(1) & COMMA 'Locked, tinyint, NOT NULL
        sSQL &= SQLNumber(1) & COMMA 'TransactionType, int, NOT NULL
        sSQL &= "'('" & SQLString(StrBatchID) & "')'" 'StrBatchID, varchar[8000], NOT NULL
        Return sSQL
    End Function



    Private Sub mnsLockVoucher_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsLockVoucher.Click, tsmLockVoucher.Click
        Dim sSQL As String = ""
        If D99C0008.Msg(rl3("MSG000002"), rl3("Thong_bao"), L3MessageBoxButtons.YesNo, L3MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then '"Bạn có muốn khóa phiếu này không?"
            If tdbg.Columns(COL_ModuleID).Text <> "02" Then
                D99C0008.MsgL3(rl3("Du_lieu_tu_module_khac_chuyen_qua") & Space(1) & rl3("Ban_khong_the_thay_doi_duoc")) 'Dữ liệu từ module khác chuyển qua. Bạn không thể thay đổi được.
                Exit Sub
            End If
            If tdbg.Columns(COL_Period).Text <> giTranMonth.ToString("00") & "/" & giTranYear.ToString Then
                D99C0008.MsgL3(rl3("Du_lieu_khong_thuoc_ky_nay") & Space(1) & rl3("Ban_khong_the_thay_doi_duoc")) 'Dữ liệu không thuộc kỳ này. Bạn không thể thay đổi được.
                Exit Sub
            End If
            'Thay đổi ngày 13/11/2012 theo incident 52440 của Bảo Trân bởi Văn Vinh
            sSQL = SQLStoreD02P5559(tdbg.Columns(COL_BatchID).Text)
            'sSQL = "Update D02T0012 Set "
            'sSQL = sSQL & " Locked = 1,"
            'sSQL = sSQL & " LockedUserID = '" & gsUserID & "',"
            'sSQL = sSQL & " LockedDate = " & SQLDateSave(Now)
            'sSQL = sSQL & " Where DivisionID = " & SQLString(gsDivisionID) & " And TransactionID = '" & tdbg.Columns(COL_TransactionID).Text & "'"
            ExecuteSQLNoTransaction(sSQL)
            LoadTDBGrid(, tdbg.Columns(COL_AssetID).Text)
        End If
    End Sub

    Private Sub tdbg_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbg.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If tdbg.RowCount = 0 Then
                tsmLockVoucher.Enabled = False
                mnsLockVoucher.Enabled = False
                Exit Sub
            End If
            tsmLockVoucher.Enabled = (tdbg.RowCount > 0) And tdbg.Columns(COL_Locked).Text = "0" And (iPer_F5557 >= EnumPermission.Add) And (Not gbClosed)
            mnsLockVoucher.Enabled = tsmLockVoucher.Enabled
            'Thay đổi ngày 13/11/2012 theo incident 52440
            If _esetupFrom = enumSetUpFrom.[NEW] Or _esetupFrom = enumSetUpFrom.CAP Then
                tsmLockVoucher.Enabled = False
                mnsLockVoucher.Enabled = False
            End If
            If tdbg.Columns(COL_SetUpFrom).Text = "NEW" Or tdbg.Columns(COL_SetUpFrom).Text = "CAP" Then
                tsmLockVoucher.Enabled = False
                mnsLockVoucher.Enabled = False
            End If
        End If
    End Sub

    Private Sub LoadTDBCombo()
        LoadCboPeriodReport(tdbcPeriodFrom, tdbcPeriodTo, "D02")
        tdbcPeriodFrom.Text = Format(giTranMonth, "00") & "/" & Format(giTranYear, "0000")
        tdbcPeriodTo.Text = Format(giTranMonth, "00") & "/" & Format(giTranYear, "0000")

        Dim sSQL As String = ""
        'Load tdbcAssetIDFrom
        sSQL = "--Do nguon combo Ma tai san" & vbCrLf
        sSQL &= "SELECT '%' AS AssetID, " & AllName & " AS AssetName, 0 AS DisplayOrder " & vbCrLf
        sSQL &= "UNION ALL " & vbCrLf
        sSQL &= "SELECT DISTINCT N19.AssetID, N19.AssetName" & UnicodeJoin(gbUnicode) & " As AssetName, 1 AS DisplayOrder " & vbCrLf
        sSQL &= "FROM D02N0019 (" & SQLNumber(giTranMonth) & COMMA & SQLNumber(giTranYear) & ") AS N19  " & vbCrLf
        sSQL &= "LEFT JOIN	D02T0001 T01 WITH(NOLOCK) ON T01.AssetID = N19.AssetID  " & vbCrLf
        sSQL &= "WHERE N19.IsCompleted = 1 " & vbCrLf
        sSQL &= "AND N19.Disabled = 0 " & vbCrLf
        sSQL &= "AND N19.DivisionID = " & SQLString(gsDivisionID) & "" & vbCrLf
        sSQL &= "ORDER BY	DisplayOrder, AssetID"
        Dim dtAss As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(tdbcAssetIDFrom, dtAss.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcAssetIDTo, dtAss.DefaultView.ToTable, gbUnicode)
        tdbcAssetIDFrom.SelectedIndex = 0
        tdbcAssetIDTo.SelectedIndex = 0

        'Load tdbcSetUpFrom
        sSQL = "--Do nguon combo Nguon goc hinh thanh" & vbCrLf
        sSQL &= "SELECT	 'ALL' AS SetUpFrom, " & AllName & " AS Description, " & vbCrLf
        sSQL &= "0 AS DisplayOrder" & vbCrLf
        sSQL &= "UNION ALL" & vbCrLf
        sSQL &= "SELECT 'NEW' AS SetUpFrom, N'" & IIf(gbUnicode, rl3("Mua_moi"), ConvertUnicodeToVni(rl3("Mua_moi"))).ToString & "' AS Description, " & vbCrLf
        sSQL &= "1 AS DisplayOrder" & vbCrLf
        sSQL &= "UNION ALL " & vbCrLf
        sSQL &= "SELECT	 'CIP' AS SetUpFrom, N'" & IIf(gbUnicode, rl3("Tu_xay_dung_co_ban"), ConvertUnicodeToVni(rl3("Tu_xay_dung_co_ban"))).ToString & "' AS Description, 2 AS DisplayOrder" & vbCrLf
        sSQL &= "UNION ALL " & vbCrLf
        sSQL &= "SELECT 'BAL' AS SetUpFrom, N'" & IIf(gbUnicode, rL3("Nhap_so_du"), ConvertUnicodeToVni(rL3("Nhap_so_du"))).ToString & "' AS Description, " & vbCrLf
        sSQL &= "3 AS DisplayOrder" & vbCrLf
        sSQL &= "UNION ALL " & vbCrLf
        sSQL &= "SELECT 'CAP' AS SetUpFrom, N'" & IIf(gbUnicode, rL3("Dieu_dong_von"), ConvertUnicodeToVni(rL3("Dieu_dong_von"))).ToString & "' AS Description, " & vbCrLf
        sSQL &= "4 AS DisplayOrder" & vbCrLf
        sSQL &= "ORDER BY DisplayOrder"
        LoadDataSource(tdbcSetUpFrom, sSQL, gbUnicode)
        tdbcSetUpFrom.SelectedIndex = 0

        'Load tdbcObjectTypeID
        Dim dtObjectTypeID As DataTable = ReturnTableObjectTypeID(gbUnicode)
        LoadDataSource(tdbcObjectTypeID, dtObjectTypeID, gbUnicode)
        'Load tdbdOjectID
        'Load tdbdOjectID
        sSQL = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " As ObjectName, ObjectTypeID From Object WITH(NOLOCK) Where Disabled = 0  order by ObjectID" ' and ObjectTypeID=" & SQLString(ID)
        dtObjectID = ReturnDataTable(sSQL)
    End Sub

#Region "Events tdbcObjectTypeID load tdbcObjectID with txtObjectName"

    Private Sub tdbcObjectTypeID_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.Close
        If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 Then
            tdbcObjectTypeID.Text = ""
            tdbcObjectIDFrom.Text = ""
            tdbcObjectIDTo.Text = ""
        End If

    End Sub

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.SelectedValueChanged
        If tdbcObjectTypeID.SelectedValue Is Nothing Then
            LoadtdbcObjectID("-1")
            Exit Sub
        End If
        LoadtdbcObjectID(tdbcObjectTypeID.SelectedValue.ToString())
        tdbcObjectIDFrom.Text = ""
        tdbcObjectIDTo.Text = ""
    End Sub

    Private Sub tdbcObjectIDFrom_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectIDFrom.Close
        If tdbcObjectIDFrom.FindStringExact(tdbcObjectIDFrom.Text) = -1 Then
            tdbcObjectIDFrom.Text = ""
        End If
    End Sub

    Private Sub tdbcObjectIDTo_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectIDTo.Close
        If tdbcObjectIDFrom.FindStringExact(tdbcObjectIDTo.Text) = -1 Then
            tdbcObjectIDTo.Text = ""
        End If
    End Sub

    Private Sub LoadtdbcObjectID(ByVal ID As String)
        LoadDataSource(tdbcObjectIDFrom, ReturnTableFilter(dtObjectID, "ObjectTypeID = " & SQLString(ID), True), gbUnicode)
        LoadDataSource(tdbcObjectIDTo, ReturnTableFilter(dtObjectID, "ObjectTypeID = " & SQLString(ID), True), gbUnicode)
    End Sub
#End Region

    Private Function AllowFilter() As Boolean
        If chkIsPeriod.Checked Then
            If tdbcPeriodFrom.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rl3("Tu_ky"))
                tdbcPeriodFrom.Focus()
                Return False
            End If
            If tdbcPeriodTo.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rl3("Den_ky"))
                tdbcPeriodTo.Focus()
                Return False
            End If
            If Not CheckValidPeriodFromTo(tdbcPeriodFrom, tdbcPeriodTo) Then Return False
        End If
        If chkDate.Checked Then
            If c1dateDateFrom.Value.ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Tu_ngay"))
                c1dateDateFrom.Focus()
                Return False
            End If
            If c1dateDateTo.Value.ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Den_ngay"))
                c1dateDateTo.Focus()
                Return False
            End If
            If Not CheckValidDateFromTo(c1dateDateFrom, c1dateDateTo) Then Return False
        End If
        If tdbcAssetIDFrom.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Tai_san"))
            tdbcAssetIDFrom.Focus()
            Return False
        End If
        If tdbcAssetIDTo.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Tai_san"))
            tdbcAssetIDTo.Focus()
            Return False
        End If
        If tdbcSetUpFrom.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Nguon_goc_hinh_thanh"))
            tdbcSetUpFrom.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub SetBackColorObligatory()
        tdbcAssetIDFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAssetIDTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcSetUpFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcPeriodFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcPeriodTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateDateFrom.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateDateTo.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        btnFilter.Focus()
        If btnFilter.Focused = False Then Exit Sub
        If Not AllowFilter() Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        LoadTDBGrid(True)
        'dtF12 = Nothing
        'CallD99U1111()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub chkDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDate.Click
        If chkDate.Checked Then
            c1dateDateFrom.Value = Date.Today
            c1dateDateTo.Value = Date.Today
        Else
            c1dateDateFrom.Value = ""
            c1dateDateTo.Value = ""
        End If
        ReadOnlyControl(Not (chkDate.Checked), c1dateDateFrom, c1dateDateTo)
        c1dateDateFrom.Enabled = chkDate.Checked
        c1dateDateTo.Enabled = chkDate.Checked
    End Sub

    Private Sub chkIsPeriod_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkIsPeriod.Click
        If chkIsPeriod.Checked Then
            tdbcPeriodFrom.Text = Format(giTranMonth, "00") & "/" & Format(giTranYear, "0000")
            tdbcPeriodTo.Text = Format(giTranMonth, "00") & "/" & Format(giTranYear, "0000")
        Else
            tdbcPeriodFrom.Text = ""
            tdbcPeriodTo.Text = ""
        End If
        ReadOnlyControl(Not (chkIsPeriod.Checked), tdbcPeriodFrom, tdbcPeriodTo)
        tdbcPeriodFrom.Enabled = chkIsPeriod.Checked
        tdbcPeriodTo.Enabled = chkIsPeriod.Checked
    End Sub

#Region "Events tdbcPeriodTo"

    Private Sub tdbcPeriodTo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcPeriodTo.LostFocus
        If tdbcPeriodTo.FindStringExact(tdbcPeriodTo.Text) = -1 Then tdbcPeriodTo.Text = ""
    End Sub

#End Region

#Region "Events tdbcPeriodFrom"

    Private Sub tdbcPeriodFrom_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcPeriodFrom.LostFocus
        If tdbcPeriodFrom.FindStringExact(tdbcPeriodFrom.Text) = -1 Then tdbcPeriodFrom.Text = ""
    End Sub

#End Region

#Region "Events tdbcSetUpFrom"

    Private Sub tdbcSetUpFrom_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSetUpFrom.LostFocus
        If tdbcSetUpFrom.FindStringExact(tdbcSetUpFrom.Text) = -1 Then tdbcSetUpFrom.Text = ""
    End Sub

#End Region

#Region "Events tdbcAssetIDTo"

    Private Sub tdbcAssetIDTo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetIDTo.LostFocus
        If tdbcAssetIDTo.FindStringExact(tdbcAssetIDTo.Text) = -1 Then tdbcAssetIDTo.Text = ""
    End Sub

#End Region

#Region "Events tdbcAssetIDFrom"

    Private Sub tdbcAssetIDFrom_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetIDFrom.LostFocus
        If tdbcAssetIDFrom.FindStringExact(tdbcAssetIDFrom.Text) = -1 Then tdbcAssetIDFrom.Text = ""
    End Sub

#End Region

    Private Sub tdbcName_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSetUpFrom.Close
        tdbcName_Validated(sender, Nothing)
    End Sub

    Private Sub tdbcName_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSetUpFrom.Validated
        Dim tdbc As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        tdbc.Text = tdbc.WillChangeToText
    End Sub

    Private Sub btnF12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnF12.Click
        If usrOption Is Nothing Then Exit Sub 'TH lưới không có cột
        usrOption.Location = New Point(tdbg.Left, btnF12.Top - (usrOption.Height + 7))
        Me.Controls.Add(usrOption)
        usrOption.BringToFront()
        usrOption.Visible = True
    End Sub

    Private Sub tsbAddNewDD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbAddNewDD.Click, tsmAddNewDD.Click, mnsAddNEWDD.Click
        Dim frm As New D02F1009
        With frm
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If gbSavedOK Then
                Dim sKey As String = frm.AssetID
                LoadTDBGrid(True, sKey)
            End If
            .Dispose()
        End With
    End Sub

    Private Sub mnsEditOther_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsEditOther.Click
        If Not CheckThroughPeriod(tdbg.Columns(COL_TranMonth).Text, tdbg.Columns(COL_TranYear).Text) Then Exit Sub
        If L3Bool(tdbg.Columns(COL_Locked).Text) Then
            D99C0008.MsgL3(rL3("MSG000003") & Space(1) & rL3("MSG000023")) 'Phieu_nay_da_duoc_khoa_Ban_khong_duoc_phep_xoa"))
            Exit Sub
        End If
        Select Case convertStringToEnum(tdbg.Columns(COL_SetUpFrom).Text)
            Case enumSetUpFrom.ALL

            Case enumSetUpFrom.[NEW]
                Dim frm As New D02F1001
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .FormState = EnumFormState.FormEditOther

                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.CIP
                Dim frm As New D02F1010
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .D54ProjectID = tdbg.Columns(COL_D54ProjectID).Text
                    .D27PropertyProductID = tdbg.Columns(COL_D27PropertyProductID).Text
                    .FormState = EnumFormState.FormEditOther

                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.BAL
                Dim frm As New D02F1002
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .FormState = EnumFormState.FormEditOther

                    .ShowDialog()
                    .Dispose()
                End With
            Case enumSetUpFrom.CAP
                Dim frm As New D02F1009
                With frm
                    .AssetID = tdbg.Columns(COL_AssetID).Text
                    .FormState = EnumFormState.FormEditOther
                    .ShowDialog()
                    .Dispose()
                End With
        End Select
    End Sub

    '10/4/2017, Phạm Thị Thu : id 96097-[CDS] Phiếu hình thành TSCĐ
    Private Sub tsbPrintEquipmentAttachOfAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbPrintEquipmentAttachOfAsset.Click, tsmPrintEquipmentAttachOfAsset.Click, mnsPrintEquipmentAttachOfAsset.Click
        Me.Cursor = Cursors.WaitCursor
        PrintListAttachFixedAsset("D02R0300B")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub PrintListAttachFixedAsset(ByVal sReportTypeID As String, Optional ByVal ModuleID As String = "02")
        Dim sReportName As String = "D02R0300B"
        Dim sReportPath As String = ""
        Dim sReportTitle As String = rL3("In_thiOt_bÜ_¢Ûnh_kIm_tªi_s¶n_F")
        Dim sCustomReport As String = ""
        Dim file As String = D99D0541.GetReportPathNew(ModuleID, sReportTypeID, sReportName, sCustomReport, sReportPath, sReportTitle)
        If sReportName = "" Then Exit Sub

        Dim sSQL As String = SQLStoreD02P1034()

        Dim dtTable As DataTable = ReturnDataTable(sSQL)
        Dim report As New D99C1003
        Select Case file.ToLower
            Case "rpt"
                Dim conn As New SqlConnection(gsConnectionString)
                With report
                    .OpenConnection(conn)

                    Dim sSQLSub As String = ""
                    sSQLSub = "SELECT CompanyName, CompanyPhone, CompanyFax, AddressLine1, AddressLine2,AddressLine3, AddressLine4, AddressLine5, CompanyAddress" & vbCrLf
                    sSQLSub &= "FROM D91V0016 WHERE DivisionID = " & SQLString(gsDivisionID)
                    Dim sSubReport As String = "D91R0000"
                    UnicodeSubReport(sSubReport, sSQLSub, gsDivisionID, gbUnicode)

                    .AddSub(sSQLSub, sSubReport & ".rpt")
                    .AddMain(sSQL)
                    .PrintReport(sReportPath, sReportTitle & " - " & sReportName)
                End With
            Case Else
                D99D0541.PrintOfficeType(sReportTypeID, sReportName, sReportPath, file, dtTable)
        End Select

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1034
    '# Created User: NGOCTHOAI
    '# Created Date: 10/04/2017 02:07:31
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1034() As String
        Dim sSQL As String = ""
        sSQL &= ("-- do nguon in Danh sach thiet bi dinh kem TSCD " & vbCrlf)
        sSQL &= "Exec D02P1034 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(1) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_AssetID).Text) 'AssetID, varchar[20], NOT NULL
        Return sSQL
    End Function

    '13/5/2017,	Phạm Thị Thu: id 96504-Hỗ trợ tính năng xuất Excel tại màn hình truy vấn hình thành tài sản cố định
    Private Sub tsbExportToExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExportToExcel.Click, tsmExportToExcel.Click, mnsExportToExcel.Click
        CreateTableCaption() 'Tạo table caption để xuất cột trên lưới
        CallShowD99F2222(Me, dtCaptionCols, dtGrid, gsGroupColumns)
    End Sub

    Private Sub CreateTableCaption()
        Dim Arr As New ArrayList
        For i As Integer = 0 To tdbg.Splits.Count - 1
            If tdbg.Splits(i).SplitSize = 0 Then Continue For
            If tdbg.Splits(i).SplitSize = 1 And tdbg.Splits(i).SplitSizeMode = C1.Win.C1TrueDBGrid.SizeModeEnum.Exact Then Continue For
            AddColVisible(tdbg, i, Arr, , False, False, gbUnicode)
        Next
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
    End Sub

    Private Sub tdbcObjectIDFrom_Validated(sender As Object, e As EventArgs) Handles tdbcObjectIDFrom.Validated
        oFilterCombo.FilterCombo(tdbcObjectIDFrom, e)
        If tdbcObjectIDFrom.FindStringExact(tdbcObjectIDFrom.Text) = -1 Then 'Code của sự kiện LostFocus
            tdbcObjectIDFrom.Text = ""
        End If
    End Sub

    Private Sub tdbcObjectIDTo_Validated(sender As Object, e As EventArgs) Handles tdbcObjectIDTo.Validated
        oFilterCombo.FilterCombo(tdbcObjectIDTo, e)
        If tdbcObjectIDTo.FindStringExact(tdbcObjectIDTo.Text) = -1 Then 'Code của sự kiện LostFocus
            tdbcObjectIDTo.Text = ""
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T5010
    '# Created User: 
    '# Created Date: 18/11/2021 04:51:42
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T5010() As String
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        Dim sSQL As String = ""
        sSQL &= ("-- Xoa du lieu " & vbCrLf)
        sSQL &= "Delete From D02T5010 Where AssetID = " & SQLString(tdbg.Columns(COL_AssetID).Text) & " AND DivisionID = " & SQLString(gsDivisionID)
        Return sSQL
    End Function



End Class
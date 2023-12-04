Imports System.Windows.Forms
Imports System
'Imports C1.C1Excel
Public Class D02F2060


#Region "Const of tdbgM"
    Private Const COLM_VoucherID As Integer = 0         ' VoucherID
    Private Const COLM_VoucherNo As Integer = 1         ' Số chứng từ
    Private Const COLM_VoucherTypeID As Integer = 2     ' VoucherTypeID
    Private Const COLM_Description As Integer = 3       ' Ghi chú
    Private Const COLM_TransactionTypeID As Integer = 4 ' Loại kiểm kê
    Private Const COLM_IsInventory As Integer = 5       ' IsInventory
    Private Const COLM_CreateUserID As Integer = 6      ' CreateUserID
    Private Const COLM_CreateDate As Integer = 7        ' CreateDate
    Private Const COLM_LastModifyUserID As Integer = 8  ' LastModifyUserID
    Private Const COLM_LastModifyDate As Integer = 9    ' LastModifyDate
    Private Const COLM_VoucherDate As Integer = 10      ' Ngày lập
    Private Const COLM_EmployeeID As Integer = 11       ' Người lập
#End Region


#Region "Const of tdbgD - Total of Columns: 26"
    Private Const COLD_IsSelected As Integer = 0           ' Chọn
    Private Const COLD_AssetID As Integer = 1              ' Mã tài sản
    Private Const COLD_AssetName As Integer = 2            ' Tên tài sản
    Private Const COLD_AssetTypeID As Integer = 3          ' AssetTypeID
    Private Const COLD_AssetTypeName As Integer = 4        ' Loại tài sản
    Private Const COLD_OVoucherNo As Integer = 5           ' Phiếu hình thành
    Private Const COLD_ObjectTypeID As Integer = 6         ' Loại đối tượng
    Private Const COLD_ObjectID As Integer = 7             ' Mã đối tượng
    Private Const COLD_ObjectName As Integer = 8           ' Tên đối tượng
    Private Const COLD_LocationID As Integer = 9           ' Vị trí
    Private Const COLD_Status As Integer = 10              ' Tình trạng
    Private Const COLD_VoucherID As Integer = 11           ' VoucherID
    Private Const COLD_TransactionID As Integer = 12       ' TransactionID
    Private Const COLD_CurrentCost As Integer = 13          ' Nguyên giá
    Private Const COLD_RemainQTY As Integer = 14           ' Số lượng còn lại
    Private Const COLD_RemainAMT As Integer = 15           ' Giá trị còn lại
    Private Const COLD_InventoryQTY As Integer = 16        ' Số lượng kiểm kê
    Private Const COLD_InventoryAMT As Integer = 17        ' Giá trị kiểm kê
    Private Const COLD_DifferenceQTY As Integer = 18       ' Chênh lệch số lượng
    Private Const COLD_DifferenceRemainAMT As Integer = 19 ' Chênh lệch giá trị 
    Private Const COLD_ALVoucherID As Integer = 20         ' ALVoucherID
    Private Const COLD_ALTransactionID As Integer = 21     ' ALTransactionID
    Private Const COLD_LinkVoucherID As Integer = 22       ' LinkVoucherID
    Private Const COLD_LinkTransactionID As Integer = 23   ' LinkTransactionID
    Private Const COLD_TransDesc As Integer = 24           ' Diễn giải
    Private Const COLD_Notes As Integer = 25               ' Ghi chú
#End Region


    Private dtGrid, dtGridDetail, dtObject, dtIGEMethodID As DataTable
    Dim bKeyPress As Boolean = False
    Dim bAskSave As Boolean = False
    Dim iPerD02F2060 As Integer = -1
    Dim _sVoucherID As String = "" 'Khóa IGE
    Private usrOption As New D99U1111()
    Dim dtF12 As DataTable
    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            _FormState = value
            LoadTDBCombo()
            LoadtdbdDropdown()
            iPerD02F2060 = ReturnPermission("D02F2060")
            iPer_F5558 = ReturnPermission("D02F5558") 'Phan quyen cho VoucherNo
            ResetGridM(False)
            Select Case _FormState
                Case EnumFormState.FormAdd
                Case EnumFormState.FormEdit
                Case EnumFormState.FormView
            End Select
        End Set
    End Property

    Private Sub D02F2060_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        ResetColorGrid(tdbgM, tdbgD)
        gbEnabledUseFind = False
        'LoadTDBGridM()
        LoadLanguage()
        SetShortcutPopupMenuNew(Me, ToolStrip1, ContextMenuStrip1, False)
        SetImageButton(btnSave, btnNotSave, imgButton)
        EnableMenu(False)
        LockStatus()
        CheckIdTextBox(txtVoucherNo)
        tdbgD_LockedColumns()
        InputbyUnicode(Me, gbUnicode)
        SetBackColorObligatory()
        InputDateInTrueDBGrid(tdbgM, COLM_VoucherDate)
        tdbgD_NumberFormat()
        LockControlDetail(True)
        LoadDefault()
        CallD99U1111()
        SetResolutionForm(Me, ContextMenuStrip1)
        Me.Cursor = Cursors.Default
    End Sub

    Private Function AllowFilterDetail() As Boolean
        If Not chkAssetTypeIDTSCD.Checked AndAlso Not chkAssetTypeIDCCDC.Checked Then
            D99C0008.MsgNotYetChoose(rL3("Loai_tai_san_hoac_Cong_cu_dung_cu"))
            chkAssetTypeIDTSCD.Focus()
            Return False
        End If

        If tdbcObjectTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Loai_doi_tuong1"))
            tdbcObjectTypeID.Focus()
            Return False
        End If
        If tdbcObjectIDFrom.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Doi_tuong"))
            tdbcObjectIDFrom.Focus()
            Return False
        End If
        If tdbcObjectIDTo.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Doi_tuong"))
            tdbcObjectIDTo.Focus()
            Return False
        End If
        If tdbcLocationIDFrom.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Vi_tri"))
            tdbcLocationIDFrom.Focus()
            Return False
        End If
        If tdbcLocationIDTo.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Vi_tri"))
            tdbcLocationIDTo.Focus()
            Return False
        End If
        If tdbcAssetIDFrom.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Ma_tai_san"))
            tdbcAssetIDFrom.Focus()
            Return False
        End If
        If tdbcAssetIDTo.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Ma_tai_san"))
            tdbcAssetIDTo.Focus()
            Return False
        End If
        
        Return True
    End Function

    Dim bPressFilter As Boolean = False
    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        btnFilter.Focus()
        bPressFilter = True
        If btnFilter.Focused = False Then Exit Sub
        If Not AllowFilterDetail() Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        LoadTDBGridD(0)
        Me.Cursor = Cursors.Default
    End Sub

#Region "Events tdbcObjectTypeID load tdbcObjectID"

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.SelectedValueChanged
        If tdbcObjectTypeID.SelectedValue Is Nothing OrElse tdbcObjectTypeID.Text = "" Then
            LoadtdbcObjectID("-1")
            tdbcObjectIDFrom.Text = ""
            tdbcObjectIDTo.Text = ""
            Exit Sub
        End If
        If ReturnValueC1Combo(tdbcObjectTypeID) = "%" Then
            ReadOnlyControl(tdbcObjectIDFrom, tdbcObjectIDTo)
            tdbcObjectIDFrom.SelectedIndex = 0
            tdbcObjectIDTo.SelectedIndex = 0
            Exit Sub
        End If
        LoadtdbcAsset()
        LoadtdbcObjectID(tdbcObjectTypeID.SelectedValue.ToString())
    End Sub

    Private Sub tdbcObjectTypeID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.LostFocus
        If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 Then
            tdbcObjectTypeID.Text = ""
            tdbcObjectIDFrom.Text = ""
            tdbcObjectIDTo.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub tdbcObjectID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectIDFrom.LostFocus, tdbcObjectIDTo.LostFocus, chkIsManagement.LostFocus
        If tdbcObjectIDFrom.FindStringExact(tdbcObjectIDFrom.Text) = -1 Then tdbcObjectIDFrom.Text = ""
        If tdbcObjectIDTo.FindStringExact(tdbcObjectIDTo.Text) = -1 Then tdbcObjectIDTo.Text = ""
    End Sub

#End Region

    Private Sub LoadEdit()
        If dtGrid Is Nothing Then Exit Sub 'Chưa đổ nguồn cho lưới
        If dtGrid.Rows.Count = 0 Then Exit Sub 'Chưa đổ nguồn cho lưới
        'dtGridDetail.Clear()
        'dtGridDetail = ReturnDataTable(SQLStoreD02P2061(1))
        LoadTDBGridD(1)
        LoadTDBCVoucherTypeID()
        tdbcVoucherTypeID.SelectedValue = tdbgM.Columns(COLM_VoucherTypeID).Text
        txtVoucherNo.Text = tdbgM.Columns(COLM_VoucherNo).Text
        c1dateVoucherDate.Value = tdbgM.Columns(COLM_VoucherDate).Text
        tdbcEmployeeID.SelectedValue = tdbgM.Columns(COLM_EmployeeID).Text
        txtDescription.Text = tdbgM.Columns(COLM_Description).Text
        bKeyPress = False
        btnFilter.Enabled = False
        btnImport.Enabled = False
        btnExport.Enabled = False
    End Sub

    Private Sub EnableMenu(ByVal bEnabled As Boolean)
        btnSave.Enabled = bEnabled
        btnNotSave.Enabled = bEnabled
        grpMaster.Enabled = Not bEnabled
        tdbgM.Enabled = Not bEnabled
        If bEnabled Then
            CheckMenu("-1", ToolStrip1, -1, False, True, ContextMenuStrip1)
        Else
            CheckMenu(Me.Name, ToolStrip1, tdbgM.RowCount, gbEnabledUseFind, True, ContextMenuStrip1)
        End If
    End Sub

    Private Sub LoadLanguage()
        '================================================================ 
        Me.Text = rl3("Kiem_ke_tai_san") & " - " & Me.Name & UnicodeCaption(gbUnicode) 'KiÓm k£ tªi s¶n
        '================================================================ 
        lblAssetID.Text = rl3("Ma_tai_san") 'Mã tài sản
        lblLocationID.Text = rl3("Vi_tri") 'Vị trí
        lblObjectID.Text = rl3("Doi_tuong") 'Đối tượng
        lblObjectTypeID.Text = rl3("Loai_doi_tuong1") 'Loại đối tượng
        lblData.Text = rl3("Du_lieu") 'Dữ liệu
        lblVoucherNo.Text = rl3("So_chung_tu") 'Số chứng từ
        lblVoucherTypeID.Text = rl3("Loai_chung_tuO") 'Loại chứng từ
        lblteVoucherDate.Text = rl3("Ngay_lap") 'Ngày lập
        lblEmployeeID.Text = rl3("Nguoi_lap") 'Người lập
        lblVoucher.Text = rl3("Chung_tu") 'Chứng từ
        lblNotes.Text = rl3("Ghi_chu") 'Ghi chú
        '================================================================ 
        btnFilterMaster.Text = rl3("Xe_m") 'Xe&m
        btnFilter.Text = rL3("Loc") & " (F5)" 'Lọc (F5)
        btnNotSave.Text = rl3("_Khong_luu") '&Không Lưu
        btnSave.Text = rL3("_Luu") '&Lưu

        btnExport.Text = rL3("Xuat_du_lieu") 'Xuất dữ liệu
        btnImport.Text = rL3("Nhap_du_lieu") 'Nhập dữ liệu
        '================================================================ 
        chkIsPeriod.Text = rl3("Ky") 'Kỳ
        chkIsDate.Text = rl3("Ngay") 'Ngày
        chkAssetTypeIDCCDC.Text = rl3("Cong_cu_dung_cu") 'Công cụ dụng cụ
        chkAssetTypeIDTSCD.Text = rL3("Tai_san_co_dinh") 'Tài sản cố định
        chkIsLiquidated.Text = rL3("Khong_hien_thi_cac_tai_san_da_thanh_ly") 'Không hiển thị các tài sản đã thanh lý
        chkIsManagement.Text = rL3("Hien_thi_tai_san_quan_ly_cua_don_vi") 'Hiển thị tài sản quản lý của đơn vị
        '================================================================ 

        grpMaster.Text = rl3("Thong_tin_kiem_ke") 'Thông tin kiểm kê
        pnlMaster.Text = rL3("Thong_tin_tai_san") 'Thông tin tài sản

        '================================================================ 
        grpMaster.Text = rL3("Thong_tin_kiem_ke") 'Thông tin kiểm kê
        '================================================================ 
        tdbcPeriodTo.Columns("Period").Caption = rl3("Ky") 'Kỳ
        tdbcPeriodFrom.Columns("Period").Caption = rl3("Ky") 'Kỳ
        tdbcAssetIDTo.Columns("AssetID").Caption = rl3("Ma") 'Mã
        tdbcAssetIDTo.Columns("AssetName").Caption = rl3("Ten") 'Tên
        tdbcAssetIDFrom.Columns("AssetID").Caption = rl3("Ma") 'Mã
        tdbcAssetIDFrom.Columns("AssetName").Caption = rl3("Ten") 'Tên
        tdbcLocationIDTo.Columns("LookupID").Caption = rl3("Ma") 'Mã
        tdbcLocationIDTo.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcLocationIDFrom.Columns("LookupID").Caption = rl3("Ma") 'Mã
        tdbcLocationIDFrom.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcObjectIDTo.Columns("ObjectID").Caption = rl3("Ma") 'Mã
        tdbcObjectIDTo.Columns("ObjectName").Caption = rl3("Ten") 'Tên
        tdbcObjectIDFrom.Columns("ObjectID").Caption = rl3("Ma") 'Mã
        tdbcObjectIDFrom.Columns("ObjectName").Caption = rl3("Ten") 'Tên
        tdbcObjectTypeID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã
        tdbcObjectTypeID.Columns("ObjectTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rl3("Loai_phieu") 'Loại phiếu
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rl3("Ten_phieu") 'Tên phiếu
        tdbcEmployeeID.Columns("EmployeeID").Caption = rl3("Ma") 'Mã
        tdbcEmployeeID.Columns("EmployeeName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbdStatus.Columns("Status").Caption = rl3("Ma") 'Mã
        tdbdStatus.Columns("StatusName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbgM.Columns(COLM_VoucherNo).Caption = rl3("So_chung_tu") 'Số chứng từ
        tdbgM.Columns(COLM_VoucherDate).Caption = rl3("Ngay_lap") 'Ngày lập
        tdbgM.Columns(COLM_EmployeeID).Caption = rl3("Nguoi_lap") 'Người lập
        tdbgM.Columns(COLM_Description).Caption = rL3("Ghi_chu") 'Ghi chú
        tdbgM.Columns(COLM_TransactionTypeID).Caption = rL3("Loai_kiem_ke") 'Loại kiểm kê

        '================================================================ 
        tdbgD.Columns(COLD_IsSelected).Caption = rL3("Chon") 'Chọn
        tdbgD.Columns(COLD_AssetID).Caption = rL3("Ma_tai_san") 'Mã tài sản
        tdbgD.Columns(COLD_AssetName).Caption = rL3("Ten_tai_san") 'Tên tài sản
        tdbgD.Columns(COLD_AssetTypeName).Caption = rL3("Loai_tai_sanU") 'Loại tài sản
        tdbgD.Columns(COLD_OVoucherNo).Caption = rL3("Phieu_hinh_thanh") 'Phiếu hình thành
        tdbgD.Columns(COLD_ObjectTypeID).Caption = rL3("Loai_doi_tuong1") 'Loại đối tượng
        tdbgD.Columns(COLD_ObjectID).Caption = rL3("Ma_doi_tuong") 'Mã đối tượng
        tdbgD.Columns(COLD_ObjectName).Caption = rL3("Ten_doi_tuong") 'Tên đối tượng
        tdbgD.Columns(COLD_LocationID).Caption = rL3("Vi_tri") 'Vị trí
        tdbgD.Columns(COLD_Status).Caption = rL3("Tinh_trang") 'Tình trạng
        tdbgD.Columns(COLD_RemainQTY).Caption = rL3("So_luong_con_lai") 'Số lượng còn lạiUYNHKHANH
        tdbgD.Columns(COLD_RemainAMT).Caption = rL3("Gia_tri_con_lai") 'Giá trị còn lại
        tdbgD.Columns(COLD_InventoryQTY).Caption = rL3("So_luong_kiem_ke") 'Số lượng kiểm kê
        tdbgD.Columns(COLD_InventoryAMT).Caption = rL3("Gia_tri_kiem_ke") 'Giá trị kiểm kê
        tdbgD.Columns(COLD_TransDesc).Caption = rL3("Dien_giai") 'Diễn giải
        tdbgD.Columns(COLD_Notes).Caption = rL3("Ghi_chu")
        '================================================================ 
        mnsPrintDetail.Text = rL3("In_chi_tiet") 'In chi tiết
        '================================================================ 
        mnsInventoryAsset.Text = rL3("_Cap_nhat_kiem_ke") '&Cập nhật kiểm kê
        '================================================================   
        tdbgD.Columns(COLD_DifferenceQTY).Caption = rl3("Chenh_lech_so_luong") 'Chênh lệch số lượng
        tdbgD.Columns(COLD_DifferenceRemainAMT).Caption = rL3("Chenh_lech_gia_tri_") 'Chênh lệch giá trị 
        '================================================================ 
        tdbgD.Columns(COLD_CurrentCost).Caption = rL3("Nguyen_gia") 'Nguyên giá
        '================================================================ 
        btnF12.Text = "F12 ( " & rL3("Hien_thi") & " )" 'Hiển thị (F12)


        '================================================================ 
        lblPlanCode.Text = rL3("Ma_ke_hoach") 'Mã kế hoạch
        '================================================================ 
        tdbcPlanCode.Columns("PlanCode").Caption = rL3("Ma") 'Mã

    End Sub


    Private Sub chkIsDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkIsDate.CheckedChanged
        If chkIsDate.Checked Then
            UnReadOnlyControl(True, c1dateDateFrom, c1dateDateTo)
            c1dateDateFrom.Value = Now.Date
            c1dateDateTo.Value = Now.Date
        Else
            ReadOnlyControl(True, c1dateDateFrom, c1dateDateTo)
            c1dateDateFrom.Value = DBNull.Value
            c1dateDateTo.Value = DBNull.Value
        End If
    End Sub

    Private Sub chkIsPeriod_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkIsPeriod.CheckedChanged
        If chkIsPeriod.Checked Then
            UnReadOnlyControl(True, tdbcPeriodFrom, tdbcPeriodTo)
        Else
            ReadOnlyControl(True, tdbcPeriodFrom, tdbcPeriodTo)
            ClearText(tdbcPeriodFrom)
            ClearText(tdbcPeriodTo)
        End If
    End Sub

    Private Sub LoadTDBCombo()
        LoadCboPeriodReport(tdbcPeriodFrom, tdbcPeriodTo, "D02")
        tdbcPeriodFrom.SelectedValue = giTranMonth.ToString("00") & "/" & giTranYear
        tdbcPeriodTo.SelectedValue = giTranMonth.ToString("00") & "/" & giTranYear


        'Do nguon cho loai doi tuong co chua $
        LoadObjectTypeIDAll(tdbcObjectTypeID, gbUnicode)

        Dim sSQL As String = ""

        'Do nguon cho combo Object
        sSQL = "--Do nguon cho object" & vbCrLf
        sSQL &= "Select '%' as ObjectTypeID, '%' as ObjectID, " & AllName & " as ObjectName, 0 as DisplayOrder  " & vbCrLf
        sSQL &= "Union all" & vbCrLf
        sSQL &= "SELECT		ObjectTypeID, ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " AS ObjectName, 1 as DisplayOrder	" & vbCrLf
        sSQL &= "FROM 		Object WITH (NOLOCK) " & vbCrLf
        sSQL &= "WHERE		Disabled = 0" & vbCrLf
        sSQL &= "ORDER BY  	DisplayOrder, ObjectTypeID, ObjectID"
        dtObject = ReturnDataTable(sSQL)
        LoadDataSource(tdbcObjectIDFrom, sSQL, gbUnicode)
        LoadDataSource(tdbcObjectIDTo, sSQL, gbUnicode)
        tdbcObjectTypeID.SelectedValue = "%"

        'Do nguon cho combo Vi tri
        sSQL = "-- Combo Vi tri" & vbCrLf
        sSQL &= "SELECT	 '%' AS LookupID, " & AllName & " AS Description, 0 AS DisplayOrder "
        sSQL &= "UNION ALL "
        sSQL &= "SELECT 	LookupID, Description" & UnicodeJoin(gbUnicode) & "  As Description, DisplayOrder "
        sSQL &= "FROM 	D91T0320 WITH(NOLOCK) "
        sSQL &= "WHERE	 LookupType = 'D02_Position' "
        sSQL &= "AND (DAGroupID ='' Or "
        sSQL &= "DAGroupID  IN (Select DAGroupID "
        sSQL &= "From lemonsys.dbo.D00V0080 "
        sSQL &= "Where UserID = " & SQLString(gsUserID) & " ) Or " & SQLString(gsUserID) & " = 'LEMONADMIN') "
        sSQL &= "ORDER BY DisplayOrder, LookupID	"
        LoadDataSource(tdbcLocationIDFrom, sSQL, gbUnicode)
        LoadDataSource(tdbcLocationIDTo, sSQL, gbUnicode)
        tdbcLocationIDFrom.SelectedValue = "%"
        tdbcLocationIDTo.SelectedValue = "%"

        'Do nguon cho combo ma tai san
        LoadtdbcAsset()

        'Load tdbcVoucherTypeID
        'sSQL = "--Combo Loai chung tu" & vbCrLf
        'sSQL &= "SELECT 		IGEMethodID, IGEMethodName" & UnicodeJoin(gbUnicode) & " As IGEMethodName,  Defaults, FormID, Disabled" & vbCrLf
        'sSQL &= "FROM 		D91T0045 WITH(NOLOCK)" & vbCrLf
        'sSQL &= "WHERE 		ModuleID = '56' " & vbCrLf
        ''sSQL &= "And Disabled = 0 " & vbCrLf
        'sSQL &= "And FormID ='D56F0060' " & vbCrLf
        'sSQL &= "And (DivisionID = " & SQLString(gsDivisionID) & " Or DivisionID = '' )" & vbCrLf
        'sSQL &= " ORDER BY 	IGEMethodID"
        dtIGEMethodID = ReturnDataTable(ReturnTableVoucherTypeID("D02", gsDivisionID, "", gbUnicode))


        'Load EmployeeID
        LoadCboCreateBy(tdbcEmployeeID, gbUnicode)

        sSQL = "SELECT	DISTINCT BatchID AS BatchID,PlanCode AS PlanCode" & vbCrLf
        sSQL &= "FROM			D02T2080 WITH(NOLOCK) " & vbCrLf
        sSQL &= "WHERE		AStatusID = '90'" & vbCrLf
        LoadDataSource(tdbcPlanCode, sSQL, gbUnicode)

    End Sub

    Private Sub LoadtdbcAsset()
        Dim sSQL As String = ""
        'Do nguon cho combo ma tai san
        sSQL = "--Do nguon combo Ma tai san " & vbCrLf
        sSQL &= "SELECT	 '%' AS AssetID, " & AllName & " AS AssetName, 0 AS DisplayOrder  "
        If chkAssetTypeIDTSCD.Checked Then
            sSQL &= "UNION ALL "
            sSQL &= "SELECT DISTINCT N19.AssetID, N19.AssetName" & UnicodeJoin(gbUnicode) & " As AssetName, 1 AS DisplayOrder "
            sSQL &= "FROM 	D02N0019 (3, 2007) AS N19  "
            sSQL &= "LEFT JOIN	D02T0001 T01 WITH(NOLOCK) ON T01.AssetID = N19.AssetID  "
            sSQL &= "WHERE(N19.IsCompleted = 1) "
            sSQL &= "AND N19.Disabled = 0 "
            sSQL &= "AND N19.DivisionID = " & SQLString(gsDivisionID) & " "
        End If
        If chkAssetTypeIDCCDC.Checked Then
            sSQL &= "UNION ALL "
            sSQL &= "Select DISTINCT  A.InventoryID As AssetID, B.InventoryName" & UnicodeJoin(gbUnicode) & "  as AssetName, 2 as DisplayOrder "
            sSQL &= "From D43T2000 A WITH(NOLOCK)"
            sSQL &= "Left join D07T0002 B WITH(NOLOCK) On A.InventoryID=B.InventoryID "
            sSQL &= "Where A.TransactionTypeID IN ('','SD') "
            If ReturnValueC1Combo(tdbcObjectTypeID).ToString <> "%" Then
                sSQL &= "and ObjectTypeID=" & SQLString(ReturnValueC1Combo(tdbcObjectTypeID))
            End If
        End If
        sSQL &= "ORDER BY	DisplayOrder, AssetID "
        LoadDataSource(tdbcAssetIDFrom, sSQL, gbUnicode)
        LoadDataSource(tdbcAssetIDTo, sSQL, gbUnicode)
        tdbcAssetIDFrom.SelectedValue = "%"
        tdbcAssetIDTo.SelectedValue = "%"
    End Sub
    Public Function ReturnTableVoucherTypeID(ByVal sModuleID As String, ByVal DivisionID As String, ByVal sEditTransTypeID As String, Optional ByVal bUseUnicode As Boolean = False) As String
        Dim sSQL As String = "--Do nguon cho combo loai phieu" & vbCrLf
        sSQL &= "Select T01.VoucherTypeID, " & IIf(bUseUnicode, "VoucherTypeNameU", "VoucherTypeName").ToString & " as VoucherTypeName, Auto, S1Type, S1, S2Type, S2, " & vbCrLf
        sSQL &= "S3, S3Type, OutputOrder, OutputLength, Separator, T40.FormID " & vbCrLf
        sSQL &= "From D91T0001 T01 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Left Join D02T0080 T40 WITH(NOLOCK) ON T01.VoucherTypeID = T40.VoucherTypeID" & vbCrLf
        sSQL &= "Where Use" & sModuleID & " = 1 And Disabled = 0 " & vbCrLf
        If DivisionID <> "" Then sSQL &= "AND( VoucherDivisionID='' Or VoucherDivisionID = " & SQLString(DivisionID) & ") " & vbCrLf
        'Load cho trường hợp Sửa, Xem
        If sEditTransTypeID <> "" Then
            sSQL &= "Or T01.VoucherTypeID = " & SQLString(sEditTransTypeID) & vbCrLf
        End If
        sSQL &= "Order By VoucherTypeID"
        Return sSQL
    End Function
    Private Sub LoadTDBCVoucherTypeID()
        LoadDataSource(tdbcVoucherTypeID, dtIGEMethodID, gbUnicode)
    End Sub

    Private Sub LoadtdbcObjectID(ByVal ID As String)
        If dtObject Is Nothing Then Exit Sub
        LoadDataSource(tdbcObjectIDFrom, ReturnTableFilter(dtObject, "ObjectTypeID ='%' or ObjectTypeID = " & SQLString(ID), True), gbUnicode)
        tdbcObjectIDFrom.SelectedIndex = 0
        UnReadOnlyControl(True, tdbcObjectIDFrom)

        LoadDataSource(tdbcObjectIDTo, ReturnTableFilter(dtObject, "ObjectTypeID ='%' or ObjectTypeID = " & SQLString(ID), True), gbUnicode)
        tdbcObjectIDTo.SelectedIndex = 0
        UnReadOnlyControl(True, tdbcObjectIDTo)
    End Sub

    Private Sub LoadDefault()
        'c1dateDateFrom.Value = Now.Date
        'c1dateDateTo.Value = Now.Date
        chkIsPeriod.Checked = True
        c1dateDateFrom.Value = DBNull.Value
        c1dateDateTo.Value = DBNull.Value
        tdbcObjectTypeID.SelectedValue = "%"
        tdbcLocationIDFrom.SelectedValue = "%"
        tdbcLocationIDTo.SelectedValue = "%"
        tdbcAssetIDFrom.SelectedValue = "%"
        tdbcAssetIDTo.SelectedValue = "%"

    End Sub

    Private Sub LoadtdbdDropdown()
        Dim sSQL As String = ""
        sSQL = "--Do nguon cho dropdown tinh trang" & vbCrLf
        sSQL &= "select	convert(tinyint,0) as Status, "
        sSQL &= "case when " & SQLString(gsLanguage) & " = '84'  "
        sSQL &= "then case when " & SQLString(gbUnicode) & " = 0 then 'Coù' else N'Có' end "
        sSQL &= "else 'Pass' "
        sSQL &= "end as StatusName "
        sSQL &= "union all "
        sSQL &= "select	convert(tinyint,1) as Status, "
        sSQL &= "case when " & SQLString(gsLanguage) & " = '84' "
        sSQL &= "then case when " & SQLString(gbUnicode) & " = 0 then 'Khoâng coù' else N'Không có' end "
        sSQL &= "else 'Fail' "
        sSQL &= "end as StatusName "
        LoadDataSource(tdbdStatus, sSQL, gbUnicode)

    End Sub

    Private Sub D02F2060_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt Then
        ElseIf e.Control Then
        Else
            Select Case e.KeyCode
                Case Keys.Enter
                    UseEnterAsTab(Me, True)
                Case Keys.F5
                    btnFilter_Click(sender, Nothing)
                    'If btnFilter.Enabled Then btnFilter_Click(sender, Nothing)
                Case Keys.F11
                    HotKeyF11(Me, tdbgM)
            End Select
        End If
    End Sub

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(tdbgM.Columns(COLM_CreateUserID).Text, tdbgM.Columns(COLM_CreateDate).Text, tdbgM.Columns(COLM_LastModifyUserID).Text, tdbgM.Columns(COLM_LastModifyDate).Text)
    End Sub

    Private Sub LoadTDBGridM(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        If FlagAdd Then
            ' Thêm mới thì gán sFind ="" và gán FilterText =""
            ResetFilter(tdbgM, sFilterM, bRefreshFilterM)
            sFind = ""
        End If
        Dim sSQL As String = SQLStoreD02P2060()
        dtGrid = ReturnDataTable(sSQL)
        'Cách mới theo chuẩn: Tìm kiếm và Liệt kê tất cả luôn luôn sáng Khi(dt.Rows.Count > 0)
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        LoadDataSource(tdbgM, dtGrid, gbUnicode)
        ReLoadTDBGridM()
        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("VoucherID" & "=" & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbgM.Row = dt1.Rows.IndexOf(dr(0)) 'dùng tdbg.Bookmark có thể không đúng
            If Not tdbgM.Focused Then tdbgM.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
        End If
    End Sub

    Private Sub btnFilterMaster_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterMaster.Click
        btnFilterMaster.Focus()
        If btnFilterMaster.Focused = False Then Exit Sub
        If Not AllowFilterMaster() Then Exit Sub
        LoadTDBGridM(True)
    End Sub

    Private Function AllowFilterMaster() As Boolean
        If chkIsPeriod.Checked Then
            If Not CheckValidPeriodFromTo(tdbcPeriodFrom, tdbcPeriodTo) Then Return False
        End If
        If chkIsDate.Checked Then
            If Not CheckValidDateFromTo(c1dateDateFrom, c1dateDateTo) Then Return False
        End If
        Return True
    End Function

    Private Sub ReLoadTDBGridM(Optional ByVal bLoadEdit As Boolean = True)
        Dim strFind As String = sFind
        If sFilterM.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilterM.ToString

        dtGrid.DefaultView.RowFilter = strFind
        ResetGridM()
        If _FormState = EnumFormState.FormAdd Then Exit Sub
        If tdbgM.RowCount = 0 Then
            ClearText(pnlMaster)
            LockControlDetail(True)
            If dtGridDetail IsNot Nothing Then
                dtGridDetail.Clear()
                'LoadDataSource(tdbgD, dtGridDetail, gbUnicode)
            End If
        Else
            LockControlDetail(False)
            _FormState = EnumFormState.FormView
            If bLoadEdit Then
                LoadEdit()
                tdbgM.Focus()
            End If
        End If
    End Sub
    Dim IPerF206 As Integer = ReturnPermission("D02F2060")
    ' Trường hợp tìm kiếm không có dữ liệu thì Khóa Detail lại
    Private Sub LockControlDetail(ByVal bLock As Boolean)
        pnlMaster.Enabled = Not bLock
        pnlFilter.Enabled = Not bLock
        If bLock Then
            If dtGridDetail IsNot Nothing Then dtGridDetail.Clear()
        End If
        If IPerF206 >= 1 Then
            btnF12.Enabled = True
        Else
            btnF12.Enabled = False
        End If
    End Sub

    Private Sub ResetGridM(Optional bCalFooter As Boolean = True)
        CheckMenu(Me.Name, ToolStrip1, tdbgM.RowCount, gbEnabledUseFind, True, ContextMenuStrip1)
        CheckMenuOther() 'Kiem tra sáng mờ menu kiểm kê
        If bCalFooter Then FooterTotalGrid(tdbgM, COLM_VoucherNo)

    End Sub

    Private Sub CheckOtherMenu()

    End Sub


    Private Sub tdbgD_LockedColumns()
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_AssetID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_AssetName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_AssetTypeName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_OVoucherNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_LocationID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Status).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_RemainQTY).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_RemainAMT).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_InventoryQTY).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_InventoryAMT).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_TransDesc).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Notes).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_DifferenceRemainAMT).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_DifferenceQTY).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_CurrentCost).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)

    End Sub


    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, mnsAdd.Click
        If Not CheckVoucherDateInPeriodFormLoad() Then Exit Sub 'Kiểm tra Ngày phiếu với Kỳ kế toán hiện tại
        _FormState = EnumFormState.FormAdd
        _sVoucherID = ""
        LoadAddNew()
        EnableMenu(True)
        LockColumns(False)
        LockStatus()
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, mnsEdit.Click


        If tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" AndAlso L3Bool(tdbgM.Columns(COLM_IsInventory).Text) Then
            D99C0008.Msg(rL3("Phieu_da_nhap_kiem_ke") & " " & rL3("Ban_khong_duoc_phep_suaXoa"))
            LockColumns(True)
            pnlFilter.Enabled = False
            pnlInput.Enabled = False

        Else
            _FormState = EnumFormState.FormEdit

            If tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" Then 'BBKK
                ReadOnlyControl(True, tdbcVoucherTypeID, txtVoucherNo)
                EnableMenu(True)
                'btnFilter.Enabled = True
                pnlFilter.Enabled = True
                btnFilter.Enabled = True
                pnlInput.Enabled = True
                txtDescription.Focus()
                LockColumns(False)
                LockStatus()
            Else 'NKK
                ReadOnlyControl(True, tdbcVoucherTypeID, txtVoucherNo)
                EnableMenu(True)
                pnlFilter.Enabled = False
                pnlInput.Enabled = True
                btnExport.Enabled = True
                btnImport.Enabled = True
                txtVoucherNo.Focus()
                LockColumns(True)
                tdbgD.Splits(0).DisplayColumns(COLD_InventoryQTY).Locked = False
                tdbgD.Splits(0).DisplayColumns(COLD_InventoryAMT).Locked = False
                tdbgD.Splits(0).DisplayColumns(COLD_TransDesc).Locked = False
                tdbgD.Splits(0).DisplayColumns(COLD_InventoryQTY).Style.ResetBackColor()
                tdbgD.Splits(0).DisplayColumns(COLD_InventoryAMT).Style.ResetBackColor()
                tdbgD.Splits(0).DisplayColumns(COLD_TransDesc).Style.ResetBackColor()
            End If

        End If
    End Sub

    Private Sub LockColumns(ByVal block As Boolean) 'Lock cac colum không cho sửa khi đã cập nhật kiểm kê
        If block Then 'Trang thái cập nhật kiểm kê thì khóa cột chọn lại
            tdbgD.Splits(SPLIT0).DisplayColumns(COLD_IsSelected).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        Else
            tdbgD.Splits(SPLIT0).DisplayColumns(COLD_IsSelected).Style.ResetBackColor()
        End If

        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_AssetName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_AssetTypeName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_LocationID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Status).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_RemainQTY).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_RemainAMT).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_DifferenceRemainAMT).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_DifferenceQTY).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)

        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_IsSelected).Locked = block
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_AssetName).Locked = block
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_AssetTypeName).Locked = block
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectTypeID).Locked = block
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectID).Locked = block
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_ObjectName).Locked = block
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_LocationID).Locked = block
        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Status).Locked = block
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_RemainQTY).Locked = block
        'tdbgD.Splits(SPLIT0).DisplayColumns(COLD_RemainAMT).Locked = block


    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, mnsDelete.Click
        If tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" AndAlso L3Bool(tdbgM.Columns(COLM_IsInventory).Text) Then
            D99C0008.Msg(rL3("Phieu_da_nhap_kiem_ke") & " " & rL3("Ban_khong_duoc_phep_suaXoa"))
            LockColumns(True)
            pnlFilter.Enabled = False
            pnlInput.Enabled = False
        Else
            'Hỏi trước khi Xóa
            If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub
            'If Not AllowDelete() Then Exit Sub
            'Thực hiện xóa phiếu
            Dim sSQL As String = ""

            If tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" Then
                sSQL &= SQLStoreD02P2062(1) & vbCrLf
            Else
                sSQL &= SQLStoreD02P2062(3) & vbCrLf
            End If
            sSQL &= SQLDeleteD02T2061() & vbCrLf
            sSQL &= SQLDeleteD02T2060() & vbCrLf
            Dim bRunSQL As Boolean = ExecuteSQL(sSQL)
            If bRunSQL Then
                DeleteOK() 'Thông báo Xóa thành công
                DeleteVoucherNoD91T9111(txtVoucherNo.Text, "D02T2060", "VoucherNo")
                'RunAuditLog("Commission", "03", _sVoucherID, txtVoucherNo.Text, c1dateVoucherDate.Value.ToString)
                'Xử lý load dữ liệu
                DeleteGridEvent(tdbgM, dtGrid, gbEnabledUseFind)
                LoadTDBGridM()
                If dtGrid.Rows.Count = 0 Then
                    tsbAdd_Click(Nothing, Nothing)
                Else
                    LoadEdit()
                End If
            Else
                DeleteNotOK()
            End If
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T2061
    '# Created User: HUỲNH KHANH
    '# Created Date: 24/09/2014 08:22:42
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T2061() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Xoa du lieu trong bang D02T2061" & vbCrLf)
        sSQL &= "Delete From D02T2061"
        sSQL &= " Where VoucherID = " & SQLString(tdbgM.Columns(COLM_VoucherID).Text)
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T2060
    '# Created User: HUỲNH KHANH
    '# Created Date: 24/09/2014 08:23:55
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T2060() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Xoa du lieu bang D02T2060" & vbCrLf)
        sSQL &= "Delete From D02T2060"
        sSQL &= " Where VoucherID = " & SQLString(tdbgM.Columns(COLM_VoucherID).Text)
        Return sSQL
    End Function



    Dim iHeight As Integer = 0 ' Lấy tọa độ Y của chuột click tới
    Private Sub tdbgM_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbgM.MouseClick
        iHeight = e.Location.Y
    End Sub

    Private Sub tdbgM_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbgM.DoubleClick
        If iHeight <= tdbgM.Splits(0).ColumnCaptionHeight Then Exit Sub
        If tdbgM.RowCount <= 0 OrElse tdbgM.FilterActive Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbgM_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbgM.KeyDown
        Me.Cursor = Cursors.WaitCursor
        If e.KeyCode = Keys.Enter Then tdbgM_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbgM, e)
        Me.Cursor = Cursors.Default
    End Sub

#Region "Active Find - List All (Client)"
    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""

    Dim dtCaptionCols As DataTable

    Private Sub tsbFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbFind.Click, mnsFind.Click
        gbEnabledUseFind = True
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbgM.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        'Những cột bắt buộc nhập
        Dim arrColObligatory() As Integer = {COLM_VoucherNo}
        Dim Arr As New ArrayList
        For i As Integer = 0 To tdbgM.Splits.Count - 1
            AddColVisible(tdbgM, i, Arr, arrColObligatory, False, False, gbUnicode)
        Next
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbgM, Arr)
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me.Name, "0", gbUnicode)
    End Sub

    Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
        If ResultWhereClause Is Nothing Or ResultWhereClause.ToString = "" Then Exit Sub
        sFind = ResultWhereClause.ToString()
        ReLoadTDBGridM()
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, mnsListAll.Click
        sFind = ""
        ResetFilter(tdbgM, sFilterM, bRefreshFilterM)
        ReLoadTDBGridM()
    End Sub

#End Region

    Dim sFilterM As New System.Text.StringBuilder()
    Dim bRefreshFilterM As Boolean = False
    Private Sub tdbgM_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbgM.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilterM Then Exit Sub
            FilterChangeGrid(tdbgM, sFilterM) 'Nếu có Lọc khi In
            ReLoadTDBGridM()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub



    Private Sub tdbgM_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbgM.KeyPress
        If tdbgM.Columns(tdbgM.Col).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then
            e.Handled = CheckKeyPress(e.KeyChar)
        ElseIf tdbgM.Splits(tdbgM.SplitIndex).DisplayColumns(tdbgM.Col).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far Then
            e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End If
        'Select Case tdbgM.Col
        '    Case COL_OrderNum 'Chặn nhập liệu trên cột STT tăng tự động trong code
        '        e.Handled = CheckKeyPress(e.KeyChar, True)
        'End Select
    End Sub

    Private Sub tdbgM_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbgM.RowColChange
        If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub

        CheckMenuOther() 'Kiem tra sáng mờ menu kiểm kê
        LockStatus()
        'Neu luoi co 1 dong thi k can chay su kien nay
        If tdbgM.RowCount <= 1 Then Exit Sub
        'Neu o thanh Filter thi k kiem tra va chay su kien RowColChange
        If tdbgM.FilterActive Then
            bKeyPress = False
            Exit Sub
        End If
        If tdbgM.Columns(COLM_VoucherID).Tag Is Nothing OrElse tdbgM.Columns(COLM_VoucherID).Text <> tdbgM.Columns(COLM_VoucherID).Tag.ToString Then
            LoadEdit()
        End If
    End Sub

    Private Sub CheckMenuOther()
        If tdbgM.RowCount > 0 AndAlso iPerD02F2060 >= 2 AndAlso (tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" AndAlso Not L3Bool(tdbgM.Columns(COLM_IsInventory).Text)) Then
            mnsInventoryAsset.Enabled = (gbClosed = False) '17/7/2017, Phạm Thị Thu: id 99923-Lỗi khi khóa sổ không thực hiện được nghiệp vụ truy vấn D02
        Else
            mnsInventoryAsset.Enabled = False
        End If
        If tdbgM.RowCount > 0 AndAlso tdbgM.Columns(COLM_TransactionTypeID).Text = "NKK" Then
            mnsInventoryAssetDelete.Enabled = (gbClosed = False) '17/7/2017, Phạm Thị Thu: id 99923-Lỗi khi khóa sổ không thực hiện được nghiệp vụ truy vấn D02
        Else
            mnsInventoryAssetDelete.Enabled = False
        End If
        mnsPrintDetail.Enabled = tdbgM.RowCount > 0 AndAlso iPerD02F2060 >= 1 '17/7/2017, Phạm Thị Thu: id 99923-Lỗi khi khóa sổ không thực hiện được nghiệp vụ truy vấn D02
    End Sub

    Private Sub tdbgM_AfterSort(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.FilterEventArgs) Handles tdbgM.AfterSort
        If tdbgM.FilterActive Then Exit Sub
        LoadEdit()
    End Sub



#Region "SQL"
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2060
    '# Created User: HUỲNH KHANH
    '# Created Date: 23/09/2014 01:47:19
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2060() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon cho luoi 1" & vbCrLf)
        sSQL &= "Exec D02P2060 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLNumber(IsTime()) & COMMA 'IsTime, tinyint, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcPeriodFrom, "TranMonth")) & COMMA 'FromMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcPeriodTo, "TranMonth")) & COMMA 'ToMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcPeriodFrom, "TranYear")) & COMMA 'FromYear, int, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcPeriodTo, "TranYear")) & COMMA 'ToYear, int, NOT NULL
        sSQL &= SQLDateSave(c1dateDateFrom.Value) & COMMA 'FromDate, datetime, NOT NULL
        sSQL &= SQLDateSave(c1dateDateTo.Value) & COMMA 'ToDate, datetime, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function

    Private Function IsTime() As Integer
        If chkIsPeriod.Checked Then
            If chkIsDate.Checked Then
                Return 3
            Else
                Return 1
            End If
        Else
            If chkIsDate.Checked Then
                Return 2
            Else
                Return 0
            End If
        End If
    End Function

    Private _savedOK As Boolean
    Public WriteOnly Property savedOK() As Boolean
        Set(ByVal Value As Boolean)
            _savedOK = Value
        End Set
    End Property

    Private Sub D02F2060_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If _FormState = EnumFormState.FormEdit Then
            If Not _savedOK Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            If btnSave.Enabled Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        End If
    End Sub

    Dim sFilterD As New System.Text.StringBuilder()
    Dim bRefreshFilterD As Boolean = False
    Private Sub LoadTDBGridD(ByVal iMode As Byte, Optional ByVal dtG As DataTable = Nothing, Optional ByVal sKey As String = "")
        ResetFilter(tdbgD, sFilterD, bRefreshFilterD)
        Dim sSQL As String = SQLStoreD02P2061(iMode)
        'If Not bPressFilter Then
        '    dtGridDetail = ReturnDataTable(sSQL)
        'Else
        '    If dtGridDetail IsNot Nothing Then
        '        dtGridDetail.Merge(ReturnDataTable(sSQL))
        '    Else
        '        dtGridDetail = ReturnDataTable(sSQL)
        '    End If
        'End If

        Dim dt As DataTable = ReturnDataTable(SQLStoreD02P2061(iMode))
        'If dtGridDetail IsNot Nothing Then dtGridDetail.PrimaryKey = Nothing
        If _FormState = EnumFormState.FormView Then
            dtGridDetail = dt
        Else
            If dt.Rows.Count > 0 Then
                If dtGridDetail Is Nothing OrElse dtGridDetail.Rows.Count = 0 Then
                    dtGridDetail = dt.Copy
                Else
                    dtGridDetail.DefaultView.RowFilter = "IsSelected = 1"
                    dtGridDetail = dtGridDetail.DefaultView.ToTable
                    dt.PrimaryKey = New DataColumn() {dt.Columns("AssetID"), dt.Columns("ALVoucherID"), dt.Columns("ALTransactionID")}
                    dtGridDetail.Merge(dt, True, MissingSchemaAction.AddWithKey)
                End If
            Else
                If dtGridDetail IsNot Nothing Then
                    dtGridDetail.DefaultView.RowFilter = "IsSelected = 1"
                    dtGridDetail = dtGridDetail.DefaultView.ToTable
                End If

            End If
        End If
        gbEnabledUseFind = dtGridDetail.Rows.Count > 0
        LoadDataSource(tdbgD, dtGridDetail, gbUnicode)
        ReLoadTDBGridD()
        If sKey <> "" Then
            Dim dt1 As DataTable = dtGridDetail.DefaultView.ToTable
            Dim dr() As DataRow = dtGridDetail.Select("AssetID" & "=" & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbgM.Row = dt1.Rows.IndexOf(dr(0)) 'dùng tdbg.Bookmark có thể không đúng
            If Not tdbgD.Focused Then tdbgD.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
        End If
    End Sub

    Private Sub ReLoadTDBGridD(Optional ByVal bLoadEdit As Boolean = True)
        Dim strFind As String = ""
        If sFilterD.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilterD.ToString
        'If strFind <> "" Then
        '    strFind &= " or " & COLD_IsSelected & "=1"
        'End If
        dtGridDetail.DefaultView.RowFilter = strFind

        SUMFooterD()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2061
    '# Created User: HUỲNH KHANH
    '# Created Date: 23/09/2014 02:29:22
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2061(ByVal iMode As Integer) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon cho luoi 2" & vbCrLf)
        sSQL &= "Exec D02P2061 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcObjectTypeID)) & COMMA 'ObjectTypeID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcObjectIDFrom)) & COMMA 'ObjectIDFrom, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcObjectIDTo)) & COMMA 'ObjectIDTo, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcLocationIDFrom)) & COMMA 'LocationIDFrom, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcLocationIDTo)) & COMMA 'LocationIDTo, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetIDFrom)) & COMMA 'AssetIDFrom, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetIDTo)) & COMMA 'AssetIDTo, varchar[20], NOT NULL
        sSQL &= SQLNumber(AssetTypeID()) & COMMA 'AssetTypeID, varchar[20], NOT NULL
        sSQL &= SQLString(VoucherID(iMode)) & COMMA 'VoucherNo, varchar[20], NOT NULL //cai nay se truyen la voucherID, anh Nam noi
        sSQL &= SQLNumber(iMode) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString("") & COMMA 'ManagementObjTypeID, varchar[50], NOT NULL
        sSQL &= SQLString("") & COMMA 'ManagementObjIDFrom, varchar[50], NOT NULL
        sSQL &= SQLString("") & COMMA 'ManagementObjIDTo, varchar[50], NOT NULL
        sSQL &= SQLNumber(chkIsLiquidated.Checked) & COMMA 'IsLiquidated, tinyint, NOT NULL
        '30/10/2019, Lê Thị Phú Hà:id 123377-Kiểm kê TSCĐ -> Lọc tài sản thuộc bộ phận quản lý của đơn vị đó
        sSQL &= SQLNumber(chkIsManagement.Checked) & COMMA 'IsManagement, tinyint, NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcPlanCode, "BatchID"))
        Return sSQL
    End Function


    Private Function AssetTypeID() As Integer
        If chkAssetTypeIDTSCD.Checked And Not chkAssetTypeIDCCDC.Checked Then
            Return 0
        ElseIf chkAssetTypeIDCCDC.Checked And Not chkAssetTypeIDTSCD.Checked Then
            Return 1
        ElseIf chkAssetTypeIDTSCD.Checked And chkAssetTypeIDCCDC.Checked Then
            Return 2
        Else
            Return 3
        End If
    End Function

    Private Function VoucherID(ByVal iMode As Integer) As String
        If iMode = 0 Then
            Return ""
        Else
            Return tdbgM.Columns(COLM_VoucherID).Text
        End If
    End Function

    Private Sub SUMFooterD()
        FooterTotalGrid(tdbgD, COLD_AssetID)
        'FooterSumNew(tdbgD, COL_xxxxx, COL_xxxxx)
    End Sub

    Private Sub LoadAddNew()
        ClearText(pnlMaster)
        chkAssetTypeIDTSCD.Checked = True
        If dtGridDetail IsNot Nothing Then
            dtGridDetail.Clear()
        End If
        LockControlDetail(False)
        UnReadOnlyControl(True, tdbcVoucherTypeID)
        btnFilter.Enabled = True
        c1dateVoucherDate.Value = Now.Date
        LoadTDBCVoucherTypeID()
        GetTextCreateBy(tdbcEmployeeID)
        tdbcObjectTypeID.SelectedValue = "%"
        tdbcLocationIDFrom.SelectedValue = "%"
        tdbcLocationIDTo.SelectedValue = "%"
        tdbcAssetIDFrom.SelectedValue = "%"
        tdbcAssetIDTo.SelectedValue = "%"
        pnlInput.Enabled = True
        For i As Integer = 0 To dtIGEMethodID.Rows.Count - 1
            If dtIGEMethodID.Rows(i).Item("FormID").ToString = "D02F2060" Then
                Dim sFormID As String = dtIGEMethodID.Rows(i).Item("VoucherTypeID").ToString
                tdbcVoucherTypeID.Text = sFormID
                Exit Sub
            End If
        Next
    End Sub

#End Region

    Private Sub tdbgD_BeforeColEdit(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColEditEventArgs) Handles tdbgD.BeforeColEdit
        If e.ColIndex = COLD_Status Then
            If L3Bool(tdbgD.Columns(COLD_IsSelected).Value) = False Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub tdbgD_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgD.ComboSelect
        tdbgD.UpdateData()
    End Sub


    Private Sub tdbgD_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbgD.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        'Select Case e.ColIndex
        '    Case COLD_Status
        '        If Not L3IsNumeric(tdbgD.Columns(e.ColIndex).Text, XXXXXX) Then e.Cancel = True
        'End Select
    End Sub


    Private Sub tdbgD_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgD.AfterColUpdate
        '--- Gán giá trị cột sau khi tính toán và giá trị phụ thuộc từ Dropdown
        'Select Case e.ColIndex
        '    Case COLD_IsSelect
        '    Case COLD_AssetID
        '    Case COLD_AssetName
        '    Case COLD_AssetTypeName
        '    Case COLD_ObjectTypeID
        '    Case COLD_ObjectID
        '    Case COLD_ObjectName
        '    Case COLD_LocationID
        '    Case COLD_Status
        '    Case COLD_VoucherID
        '    Case COLD_TransactionID
        'End Select
        'tdbgD.UpdateData()
        'SUMFooterD()
    End Sub


    Private Sub tdbgD_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbgD.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        '--- Đổ nguồn cho các Dropdown phụ thuộc

        'Select Case tdbgD.Col
        '    Case COLD_Status
        '        tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Status).Button = (L3Bool(tdbgD(tdbgD.Row, COLD_IsSelected)) = True And (tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK"))
        '        tdbgD.UpdateData()
        'End Select

    End Sub

    Private Sub tdbgD_FetchCellStyle(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.FetchCellStyleEventArgs) Handles tdbgD.FetchCellStyle
        'If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormEdit Then
        'tdbgD.Splits(0).DisplayColumns(COLD_Status).Locked = False
        'Select Case e.Col
        '    Case COLD_Status
        '        If L3Bool(tdbgD(e.Row, COLD_IsSelected)) And tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" Then
        '            e.CellStyle.BackColor = COLOR_BACKCOLOROBLIGATORY
        '            e.CellStyle.Locked = False
        '        Else
        '            e.CellStyle.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        '            e.CellStyle.Locked = True
        '        End If
        'End Select
        'End If
    End Sub


    Private Sub tdbgD_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbgD.FilterChange
        Try
            If (dtGridDetail Is Nothing) Then Exit Sub
            If bRefreshFilterD Then Exit Sub
            FilterChangeGrid(tdbgD, sFilterD) 'Nếu có Lọc khi In
            ReLoadTDBGridD()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub tdbgD_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbgD.KeyDown
        If e.Control And e.KeyCode = Keys.S Then HeadClick(tdbgD.Col)
        Me.Cursor = Cursors.WaitCursor
        HotKeyCtrlVOnGrid(tdbgD, e)
        'If e.Control And e.KeyCode = Keys.S Then HeadClick(tdbgD.Col)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbgD_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbgD.KeyPress

    End Sub

    Private Sub SetBackColorObligatory()
        'tdbcPeriodTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        'tdbcPeriodFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        'c1dateDateFrom.BackColor = COLOR_BACKCOLOROBLIGATORY
        'c1dateDateTo.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectIDFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectIDTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcLocationIDFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcLocationIDTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAssetIDFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAssetIDTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtVoucherNo.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcEmployeeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateVoucherDate.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub SetReturnFormView()
        _FormState = EnumFormState.FormView
        EnableMenu(False)
        If tdbgM.RowCount = 0 Then
            ClearText(pnlMaster)
            LockControlDetail(True)
        Else
            LoadEdit()
            tdbgM.Focus()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        btnSave.Focus()
        If btnSave.Focused = False Then Exit Sub
        btnSave.Focus()
        If btnSave.Focused = False Then Exit Sub
        'Hỏi trước khi lưu
        If bAskSave Then 'Nhấn từ nút Lưu
            If AskSave() = Windows.Forms.DialogResult.No Then
                SetReturnFormView()
                Exit Sub
            End If

        Else 'Nhấn từ nút Không lưu
            bAskSave = True
        End If
        SaveData(sender)
    End Sub

    Private Function SaveData(ByVal sender As System.Object) As Boolean

        tdbgD.UpdateData()
        _savedOK = False
        If Not AllowSave() Then Return False
        Me.Cursor = Cursors.WaitCursor

        If Not CheckVoucherDateInPeriod(c1dateVoucherDate.Text) Then
            c1dateVoucherDate.Focus()
            Me.Cursor = Cursors.Default
            Exit Function
        End If

        btnSave.Enabled = False
        btnNotSave.Enabled = False

        'Thực hiện quy trình lưu dư liệu
        Dim sSQL As New StringBuilder("")
        Me.Cursor = Cursors.WaitCursor
        Select Case _FormState
            Case EnumFormState.FormAdd
                _sVoucherID = CreateIGE("D02T2060", "VoucherID", "02", "AC", gsStringKey)
                If tdbcVoucherTypeID.Columns("Auto").Text = "1" And bEditVoucherNo = False Then 'Sinh tự động và không nhấn F2
                    txtVoucherNo.Text = CreateIGEVoucherNoNew(tdbcVoucherTypeID, "D02T2060", _sVoucherID)
                Else 'Không sinh tự động hay có nhấn F2
                    If bEditVoucherNo = False Then
                        If CheckDuplicateVoucherNoNew(D02, "D02T2060", _sVoucherID, txtVoucherNo.Text) = True Then btnSave.Enabled = True : _sVoucherID = "" : Me.Cursor = Cursors.Default : Exit Function
                    Else 'Có nhấn F2 để sửa số phiếu
                        InsertD02T5558(_sVoucherID, sOldVoucherNo, txtVoucherNo.Text)
                    End If
                    InsertVoucherNoD91T9111(txtVoucherNo.Text, "D02T2060", _sVoucherID)
                End If
                
                sSQL.Append(SQLInsertD02T2060().ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T2061s().ToString & vbCrLf)
                sSQL.Append(SQLStoreD02P2062(0).ToString & vbCrLf)

                ''****************************************************************
            Case EnumFormState.FormEdit
                sSQL.Append(SQLDeleteD02T2061() & vbCrLf & SQLDeleteD02T2060() & vbCrLf)
                sSQL.Append(SQLInsertD02T2060().ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T2061s().ToString & vbCrLf)
                If tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" AndAlso Not L3Bool(tdbgM.Columns(COLM_IsInventory).Text) Then 'BBKK
                    sSQL.Append(SQLStoreD02P2062(0).ToString & vbCrLf)
                Else 'NKK
                    sSQL.Append(SQLStoreD02P2062(2).ToString & vbCrLf)
                End If

            Case EnumFormState.FormOther 'Cập nhật kiểm kê
                _sVoucherID = CreateIGE("D02T2060", "VoucherID", "02", "AC", gsStringKey)
                If tdbcVoucherTypeID.Columns("Auto").Text = "1" And bEditVoucherNo = False Then 'Sinh tự động và không nhấn F2
                    txtVoucherNo.Text = CreateIGEVoucherNoNew(tdbcVoucherTypeID, "D02T2060", _sVoucherID)
                Else 'Không sinh tự động hay có nhấn F2
                    If bEditVoucherNo = False Then
                        If CheckDuplicateVoucherNoNew(D02, "D02T2060", _sVoucherID, txtVoucherNo.Text) = True Then btnSave.Enabled = True : _sVoucherID = "" : Me.Cursor = Cursors.Default : Exit Function
                    Else 'Có nhấn F2 để sửa số phiếu
                        InsertD02T5558(_sVoucherID, sOldVoucherNo, txtVoucherNo.Text)
                    End If
                    InsertVoucherNoD91T9111(txtVoucherNo.Text, "D02T2060", _sVoucherID)
                End If

                'sSQL.Append(SQLUpdateD02T2061s_1().ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T2060().ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T2061s().ToString & vbCrLf)
                sSQL.Append(SQLStoreD02P2062(2).ToString)
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _savedOK = True
            bKeyPress = False
            bPressFilter = False
            'If bEditVoucherNo Then
            '    ExecuteSQLNoTransaction(SQLInsertD05T5558().ToString)
            'End If
            Select Case _FormState
                Case EnumFormState.FormAdd
                    LoadTDBGridM(True, _sVoucherID)
                    btnImport.Enabled = False 'Nhập
                    btnExport.Enabled = True 'Xuất
                Case EnumFormState.FormEdit
                    LoadTDBGridM(, tdbgM.Columns(COLM_VoucherID).Text)
                    btnImport.Enabled = False 'Nhập
                    btnExport.Enabled = True 'Xuất
                Case EnumFormState.FormOther
                    LoadTDBGridM(True, _sVoucherID)
                    btnImport.Enabled = True 'Nhập
                    btnExport.Enabled = True 'Xuất
            End Select
            ReadOnlyControl(tdbcVoucherTypeID, txtVoucherNo)
            SetReturnFormView()
            tdbgD.Splits(0).DisplayColumns(COLD_InventoryQTY).Locked = True
            tdbgD.Splits(0).DisplayColumns(COLD_InventoryAMT).Locked = True
            tdbgD.Splits(0).DisplayColumns(COLD_TransDesc).Locked = True
            tdbgD.Splits(0).DisplayColumns(COLD_InventoryQTY).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
            tdbgD.Splits(0).DisplayColumns(COLD_InventoryAMT).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
            tdbgD.Splits(0).DisplayColumns(COLD_TransDesc).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)


        Else
            If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormOther Then
                DeleteVoucherNoD91T9111_Transaction(txtVoucherNo.Text, "D02T2060", "VoucherNo", tdbcVoucherTypeID, bEditVoucherNo)
            End If
            SaveNotOK()
            btnSave.Enabled = False
            btnNotSave.Enabled = False
            Return False
        End If
        'Phải để ở đây để get bEditVoucherNo chạy SQl
        bEditVoucherNo = False
        sOldVoucherNo = ""
        bFirstF2 = False
        LockStatus()
        Return True
    End Function

    Private Function AllowSave() As Boolean

        If tdbcVoucherTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Loai_chung_tuO"))
            tdbcVoucherTypeID.Focus()
            Return False
        End If
        If txtVoucherNo.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("So_chung_tu"))
            txtVoucherNo.Focus()
            Return False
        End If
        If c1dateVoucherDate.Value.ToString = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ngay_lap"))
            c1dateVoucherDate.Focus()
            Return False
        End If
        If tdbcEmployeeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Nguoi_lap"))
            tdbcEmployeeID.Focus()
            Return False
        End If
        tdbgD.UpdateData()
        If tdbgD.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbgD.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbgD.RowCount - 1
            If L3Bool(tdbgD(i, COLD_IsSelected)) And tdbgD(i, COLD_Status).ToString = "" Then
                D99C0008.MsgNotYetEnter(tdbgD.Columns(COLD_Status).Caption)
                tdbgD.Focus()
                tdbgD.SplitIndex = 0
                tdbgD.Col = COLD_Status
                tdbgD.Bookmark = i
                Return False
            End If
        Next
        Return True
    End Function




    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T2060
    '# Created User: HUỲNH KHANH
    '# Created Date: 24/09/2014 09:40:49
    '#---------------------------------------------------------------------------------------------------
    'Private Function SQLInsertD02T2060() As StringBuilder
    '    Dim sSQL As New StringBuilder
    '    sSQL.Append("-- Luu master" & vbCrlf)
    '    sSQL.Append("Insert Into D02T2060(")
    '    sSQL.Append("VoucherID, VoucherNo, VoucherTypeID, VoucherDate, EmployeeID, ")
    '    sSQL.Append("TranMonth, TranYear, Description, DescriptionU, CreateUserID, ")
    '    sSQL.Append("CreateDate, LastModifyUserID, LastModifyDate")
    '    sSQL.Append(") Values(" & vbCrlf)
    '    sSQL.Append(SQLString(_sVoucherID) & COMMA) 'VoucherID, varchar[20], NOT NULL
    '    sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[20], NOT NULL
    '    sSQL.Append(SQLString(ReturnValueC1Combo(tdbcVoucherTypeID).ToString) & COMMA) 'VoucherTypeID, varchar[20], NOT NULL
    '    sSQL.Append(SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'VoucherDate, datetime, NOT NULL
    '    sSQL.Append(SQLString(ReturnValueC1Combo(tdbcEmployeeID).ToString) & COMMA) 'EmployeeID, varchar[50], NOT NULL
    '    sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NOT NULL
    '    sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, int, NOT NULL
    '    sSQL.Append(SQLStringUnicode(txtDescription, False) & COMMA) 'Description, varchar[500], NOT NULL
    '    sSQL.Append(SQLStringUnicode(txtDescription, True) & COMMA) 'DescriptionU, nvarchar[500], NOT NULL
    '    sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
    '    sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NOT NULL
    '    sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
    '    sSQL.Append("GetDate()") 'LastModifyDate, datetime, NOT NULL
    '    sSQL.Append(")")

    '    Return sSQL
    'End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T2060
    '# Created User: HUỲNH KHANH
    '# Created Date: 02/10/2014 02:31:45
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T2060() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- -- Luu master" & vbCrLf)
        sSQL.Append("Insert Into D02T2060(")
        sSQL.Append("DivisionID, VoucherID, VoucherNo, VoucherTypeID, VoucherDate, ")
        sSQL.Append("EmployeeID, TranMonth, TranYear, DescriptionU, ")
        sSQL.Append("CreateUserID, CreateDate, LastModifyUserID, LastModifyDate,TransactionTypeID, LinkBatchID ")
        sSQL.Append(") Values(" & vbCrLf)
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
        If _FormState = EnumFormState.FormEdit Then
            sSQL.Append(SQLString(tdbgM.Columns(COLM_VoucherID).Text) & COMMA) 'VoucherID, varchar[20], NOT NULL
        Else
            sSQL.Append(SQLString(_sVoucherID) & COMMA) 'VoucherID, varchar[20], NOT NULL
        End If

        sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[50], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcVoucherTypeID).ToString) & COMMA) 'VoucherTypeID, varchar[20], NOT NULL
        sSQL.Append(SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'VoucherDate, datetime, NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcEmployeeID).ToString) & COMMA) 'EmployeeID, varchar[50], NOT NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, int, NOT NULL
        sSQL.Append(SQLStringUnicode(txtDescription, True) & COMMA) 'DescriptionU, nvarchar[1000], NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NOT NULL
        If _FormState = EnumFormState.FormAdd Then
            sSQL.Append(SQLString("BBKK")) 'LastModifyDate, datetime, NOT NULLelse
        ElseIf _FormState = EnumFormState.FormEdit Then
            sSQL.Append(SQLString(tdbgM.Columns(COLM_TransactionTypeID).Text)) 'LastModifyDate, datetime, NOT NULL
        ElseIf _FormState = EnumFormState.FormOther Then
            sSQL.Append(SQLString("NKK")) 'LastModifyDate, datetime, NOT NULLelse
        End If
        sSQL.Append(COMMA & SQLString(ReturnValueC1Combo(tdbcPlanCode, "BatchID")))
        sSQL.Append(")")

        Return sSQL
    End Function



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T2061s
    '# Created User: HUỲNH KHANH
    '# Created Date: 24/09/2014 09:44:15
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T2061s() As StringBuilder
        Dim sRet As New StringBuilder
        'Sinh IGE chi tiết
        Dim sTransactionID As String = ""
        Dim iFirstTrans As Long = 0
        Dim iCountIGE As Integer = 0
        tdbgD.UpdateData()
        iCountIGE = dtGridDetail.Select("IsSelected =1").Length
        '---------------------------------
        Dim dr() As DataRow = dtGridDetail.Select("IsSelected =1")
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To dr.Length - 1
            If sSQL.ToString = "" And sRet.ToString = "" Then sSQL.Append("-- Luu luoi chi tiet" & vbCrLf)
            If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormOther Then
                sTransactionID = CreateIGENewS("D02T2061", "TransactionID", "02", "AD", gsStringKey, sTransactionID, iCountIGE, iFirstTrans)
                tdbgD(i, COLD_TransactionID) = sTransactionID
            End If
            If _FormState = EnumFormState.FormEdit And tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" Then
                sTransactionID = CreateIGENewS("D02T2061", "TransactionID", "02", "AD", gsStringKey, sTransactionID, iCountIGE, iFirstTrans)
                tdbgD(i, COLD_TransactionID) = sTransactionID
            End If
            sSQL.Append("Insert Into D02T2061(")
            sSQL.Append("VoucherID, TransactionID, TransDescU, AssetID, ")
            sSQL.Append("AssetTypeID, ObjectTypeID, ObjectID, LocationID, Status, ")
            sSQL.Append("ALVoucherID, ALTransactionID, RemainQTY, RemainAMT,InventoryQTY,InventoryAMT,LinkVoucherID, LinkTransactionID,OVoucherNo, NotesU")
            sSQL.Append(") Values(" & vbCrLf)
            If _FormState = EnumFormState.FormEdit Then
                sSQL.Append(SQLString(tdbgM.Columns(COLM_VoucherID).Text) & COMMA) 'VoucherID, varchar[20], NOT NULL
            Else
                sSQL.Append(SQLString(_sVoucherID) & COMMA) 'VoucherID, varchar[20], NOT NULL
            End If

            sSQL.Append(SQLString(dr(i).Item("TransactionID")) & COMMA) 'TransactionID, varchar[20], NOT NULL
            sSQL.Append(SQLStringUnicode(dr(i).Item("TransDesc"), gbUnicode, True) & COMMA) 'TransDescU, nvarchar[1000], NOT NULL
            sSQL.Append(SQLString(dr(i).Item("AssetID")) & COMMA) 'AssetID, varchar[50], NOT NULL
            sSQL.Append(SQLNumber(dr(i).Item("AssetTypeID")) & COMMA) 'AssetTypeID, varchar[50], NOT NULL
            sSQL.Append(SQLString(dr(i).Item("ObjectTypeID")) & COMMA) 'ObjectTypeID, varchar[20], NOT NULL
            sSQL.Append(SQLString(dr(i).Item("ObjectID")) & COMMA) 'ObjectID, varchar[20], NOT NULL
            sSQL.Append(SQLString(dr(i).Item("LocationID")) & COMMA) 'LocationID, varchar[50], NOT NULL
            sSQL.Append(SQLNumber(dr(i).Item("Status")) & COMMA) 'Status, tinyint, NOT NULL
            sSQL.Append(SQLString(dr(i).Item("ALVoucherID")) & COMMA) 'ALVoucherID, varchar[20], NOT NULL
            sSQL.Append(SQLString(dr(i).Item("ALTransactionID")) & COMMA) 'ALTransactionID, varchar[20], NOT NULL
            sSQL.Append(SQLMoney(dr(i).Item("RemainQTY"), DxxFormat.D07_QuantityDecimals) & COMMA) 'RemainQTY, decimal, NOT NULL
            sSQL.Append(SQLMoney(dr(i).Item("RemainAMT"), DxxFormat.D90_ConvertedDecimals) & COMMA) 'RemainAMT, decimal, NOT NULL
            sSQL.Append(SQLMoney(dr(i).Item("InventoryQTY"), DxxFormat.D07_QuantityDecimals) & COMMA) 'RemainQTY, decimal, NOT NULL
            sSQL.Append(SQLMoney(dr(i).Item("InventoryAMT"), DxxFormat.D90_ConvertedDecimals) & COMMA) 'RemainAMT, decimal, NOT NULL
            sSQL.Append(SQLString(dr(i).Item("LinkVoucherID")) & COMMA) 'LinkVoucherID
            sSQL.Append(SQLString(dr(i).Item("LinkTransactionID")) & COMMA) 'LinkTransactionID
            sSQL.Append(SQLString(dr(i).Item("OVoucherNo")) & COMMA) 'OVoucherNo
            sSQL.Append(SQLStringUnicode(dr(i).Item("Notes"), True, gbUnicode)) 'NotesU, nvarchar[500], NOT NULL
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T2061s
    '# Created User: HUỲNH KHANH
    '# Chạy trong menu xoa kiem ke
    '# Created Date: 11/10/2014 11:16:05
    '#---------------------------------------------------------------------------------------------------
    'Private Function SQLUpdateD02T2061s() As StringBuilder
    '    Dim sRet As New StringBuilder
    '    Dim sSQL As New StringBuilder
    '    For i As Integer = 0 To tdbgD.RowCount - 1
    '        If i = 0 Then sSQL.Append("-- Update cap nhat kiem ke" & vbCrlf)
    '        sSQL.Append("Update D02T2061 Set ")
    '        'sSQL.Append("VoucherID = " & SQLString(tdbgD(i, COLD_VoucherID)) & COMMA) 'varchar[20], NOT NULL
    '        sSQL.Append("InventoryQTY = " & SQLMoney("0") & COMMA) 'decimal, NOT NULL
    '        sSQL.Append("InventoryAMT = " & SQLMoney("0") & COMMA) 'decimal, NOT NULL
    '        sSQL.Append("IsInventory = " & SQLNumber("0")) 'int, NOT NULL
    '        sSQL.Append(" Where VoucherID=" & SQLString(tdbgM.Columns(COLM_VoucherID).Text))
    '        sRet.Append(sSQL.tostring & vbCrLf)
    '        sSQL.Remove(0, sSQL.Length)
    '    Next
    '    Return sRet
    'End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T2061s
    '# Created User: HUỲNH KHANH
    '#Chay trong menu Cập nhật kiểm kê
    '# Created Date: 11/10/2014 11:23:44
    '#---------------------------------------------------------------------------------------------------
    'Private Function SQLUpdateD02T2061s_1() As StringBuilder
    '    Dim sRet As New StringBuilder
    '    Dim sSQL As New StringBuilder
    '    For i As Integer = 0 To tdbgD.RowCount - 1
    '        If i = 0 Then sSQL.Append("-- Cap nhat kiem ke" & vbCrLf)
    '        sSQL.Append("Update D02T2061 Set ")
    '        'sSQL.Append("VoucherID = " & SQLString(tdbgD(i, COLD_VoucherID)) & COMMA) 'varchar[20], NOT NULL
    '        'sSQL.Append("TransactionID = " & SQLString(tdbgD(i, COLD_TransactionID)) & COMMA) 'varchar[20], NOT NULL
    '        sSQL.Append("TransDesc = " & SQLStringUnicode(tdbgD(i, COLD_TransDesc), gbUnicode, False) & COMMA) 'varchar[500], NOT NULL
    '        sSQL.Append("TransDescU = " & SQLStringUnicode(tdbgD(i, COLD_TransDesc), gbUnicode, True) & COMMA) 'nvarchar[1000], NOT NULL
    '        sSQL.Append("InventoryQTY = " & SQLMoney(tdbgD(i, COLD_InventoryQTY), DxxFormat.D07_QuantityDecimals) & COMMA) 'decimal, NOT NULL
    '        sSQL.Append("InventoryAMT = " & SQLMoney(tdbgD(i, COLD_InventoryAMT), DxxFormat.D90_ConvertedDecimals) & COMMA) 'decimal, NOT NULL
    '        sSQL.Append("IsInventory = " & SQLNumber("1")) 'int, NOT NULL
    '        sSQL.Append(" Where VoucherID = " & SQLString(tdbgM.Columns(COLM_VoucherID).Text) & " And TransactionID = " & SQLString(tdbgD.Columns(COLD_TransactionID).Text))
    '        sRet.Append(sSQL.ToString & vbCrLf)
    '        sSQL.Remove(0, sSQL.Length)
    '    Next
    '    Return sRet
    'End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2062
    '# Created User: HUỲNH KHANH
    '# Created Date: 13/10/2014 05:08:01
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2062(ByVal iMode As Integer) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Chạy store D02P2062" & vbCrLf)
        sSQL &= "Exec D02P2062 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormOther Then
            sSQL &= SQLString(_sVoucherID) & COMMA 'VoucherID, varchar[50], NOT NULL
        Else
            sSQL &= SQLString(tdbgM.Columns(COLM_VoucherID).Text) & COMMA 'VoucherID, varchar[50], NOT NULL
        End If

        sSQL &= SQLNumber(iMode) 'Mode, int, NOT NULL
        Return sSQL
    End Function


    Private Sub btnNotSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNotSave.Click
        
        If _FormState = EnumFormState.FormAdd AndAlso txtVoucherNo.Text = "" Then
            If tdbgM.RowCount > 0 Then
                'LoadEdit()
            End If
            GoTo 1
        End If
        If AskMsgBeforeRowChange() Then
            bAskSave = False
            If Not SaveData(sender) Then
                Exit Sub
            End If
        Else
            'LoadEdit()
            If _FormState = EnumFormState.FormEdit Then
                LoadTDBGridM(, tdbgM.Columns(COLM_VoucherID).Text)
            End If
        End If
1:
        bPressFilter = False
        SetReturnFormView()
    End Sub

#Region "Events tdbcVoucherTypeID"

    Dim sOldVoucherNo As String = "" 'Lưu lại số phiếu cũ
    Dim bEditVoucherNo As Boolean = False '= True: có nhấn F2; = False: không 
    Dim bFirstF2 As Boolean = False 'Nhấn F2 lần đầu tiên 
    Dim iPer_F5558 As Integer = 0 'Phân quyền cho Sửa số phiếu tại Form_load iPer_F5558 = ReturnPermission(DxxF5558)
    Private Sub txtVoucherNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVoucherNo.KeyDown
        If e.KeyCode <> Keys.F2 Then Exit Sub
        If tdbcVoucherTypeID.Text = "" Or txtVoucherNo.Text = "" Then Exit Sub
        If _FormState = EnumFormState.FormAdd And btnSave.Enabled = False Then Exit Sub
        If _FormState = EnumFormState.FormEdit And iPer_F5558 <= 2 Then Exit Sub
        If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormEdit Or _FormState = EnumFormState.FormEditOther Then '  'Cho sửa Số phiếu ở trạng thái Thêm mới hay Sửa
            If bFirstF2 = False Then
                sOldVoucherNo = txtVoucherNo.Text
                bFirstF2 = True
            End If
            Dim bSavedOk As Boolean = False
            'Dim frm As New D91F5558
            'With frm
            '    .FormName = "D91F5558"
            '    .FormPermission = "D02F5558" 'Màn hình phân quyền
            '    .ModuleID = D02 'Mã module hiện tại, VD: D07
            '    .TableName = "D02T2060" 'Tên bảng chứa số phiếu
            '    '.VoucherTypeID = ReturnValueC1Combo(tdbcVoucherTypeID).ToString
            '    If _FormState = EnumFormState.FormAdd Then
            '        .VoucherID = "" 'Khóa sinh IGE là rỗng
            '    ElseIf _FormState = EnumFormState.FormEdit Then
            '        .VoucherID = tdbgM.Columns(COLM_VoucherID).Text 'Khóa sinh IGE
            '    End If
            '    .VoucherNo = txtVoucherNo.Text  'Số phiếu cần sửa
            '    .Mode = "0" ' Tùy theo Module, mặc định là 0
            '    .KeyID01 = ""
            '    .KeyID03 = ""
            '    .KeyID04 = ""
            '    .KeyID05 = ""
            '    .ShowDialog()
            '    If .Output02 <> "" Then
            '        txtVoucherNo.Text = .Output02 'Giá trị trả về Số phiếu mới
            '        ReadOnlyControl(txtVoucherNo) 'Lock text Số phiếu
            '        bEditVoucherNo = True 'Đã nhấn F2
            '        _savedOK = True
            '    End If
            '    .Dispose()
            'End With

            Dim arrPro() As StructureProperties = Nothing
            SetProperties(arrPro, "FormIDPermission", "D02F5558")
            SetProperties(arrPro, "VoucherTypeID", ReturnValueC1Combo(tdbcVoucherTypeID))
            If _FormState = EnumFormState.FormAdd Then
                SetProperties(arrPro, "VoucherID", "")
            ElseIf _FormState = EnumFormState.FormEdit Then
                SetProperties(arrPro, "VoucherID", tdbgM.Columns(COLM_VoucherID).Text)
            End If
            SetProperties(arrPro, "Mode", 0)
            SetProperties(arrPro, "KeyID01", "")
            SetProperties(arrPro, "TableName", "D02T2060")
            SetProperties(arrPro, "ModuleID", D02)
            SetProperties(arrPro, "OldVoucherNo", txtVoucherNo.Text)
            SetProperties(arrPro, "KeyID02", "")
            SetProperties(arrPro, "KeyID03", "")
            SetProperties(arrPro, "KeyID04", "")
            SetProperties(arrPro, "KeyID05", "")
            Dim frm As Form = CallFormShowDialog("D91D0640", "D91F5558", arrPro)
            Dim sNew As String = GetProperties(frm, "NewVoucherNo").ToString
            If sNew <> "" Then
                txtVoucherNo.Text = sNew 'Giá trị trả về Số phiếu mới
                ReadOnlyControl(txtVoucherNo) 'Lock text Số phiếu
                bEditVoucherNo = True 'Đã nhấn F2
                _savedOK = True
            End If
        End If
    End Sub

    Private Sub tdbcVoucherTypeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.LostFocus
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then tdbcVoucherTypeID.Text = "" : txtVoucherNo.Text = ""
    End Sub

    Private Sub tdbcVoucherTypeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.Close
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then
            tdbcVoucherTypeID.Text = ""
            txtVoucherNo.Text = ""
        End If
    End Sub

    Private Sub tdbcVoucherTypeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.SelectedValueChanged
        bEditVoucherNo = False
        bFirstF2 = False
        If tdbcVoucherTypeID.SelectedValue Is Nothing OrElse tdbcVoucherTypeID.Text = "" Then
            txtVoucherNo.Text = ""
            ReadOnlyControl(txtVoucherNo)
            Exit Sub
        End If
        If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormOther Then
            If tdbcVoucherTypeID.Columns("Auto").Text = "1" Then 'Sinh tự động
                txtVoucherNo.Text = CreateIGEVoucherNo(tdbcVoucherTypeID, False)
                ReadOnlyControl(txtVoucherNo)
            Else 'Không sinh tự động
                txtVoucherNo.Text = ""
                UnReadOnlyControl(txtVoucherNo, True)
            End If

        End If
    End Sub
#End Region


    'Private Sub tsbExportToExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExportToExcel.Click, mnsExportToExcel.Click
    '    Dim sReportTypeID As String = "D02F2060"
    '    Dim sReportName As String = "" '"DXXRXXXX"
    '    Dim sReportPath As String = ""
    '    Dim sReportTitle As String = "" 'Thêm biến
    '    Dim sCustomReport As String = "" '= tdbcTranTypeID.Columns("InvoiceForm").Text
    '    Try
    '        Dim file As String = GetReportPathNew("02", sReportTypeID, sReportName, sCustomReport, sReportPath, sReportTitle)
    '        If sReportName = "" Then Exit Sub
    '        Select Case file.ToLower
    '            Case "xls", "xlsx"
    '                Me.Cursor = Cursors.WaitCursor
    '                Dim sPathFile As String = GetObjectFile(sReportTypeID, sReportName, file, sReportPath)
    '                If sPathFile = "" Then Exit Select
    '                MyExcel(dtGridDetail, sPathFile, file, False)
    '        End Select
    '    Catch ex As Exception

    '    Finally
    '        Me.Cursor = Cursors.Default
    '    End Try
    'End Sub

    Private Sub tsmExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbExportToExcel.Click, mnsExportToExcel.Click
        'Lưới không có nút Hiển thị
        'Nếu lưới không có Group thì mở dòng code If dtCaptionCols Is Nothing Then
        'và truyền đối số cuối cùng là False vào hàm AddColVisible
        'If dtCaptionCols Is Nothing orelse dtCaptionCols.Rows.Count < 1 Then
        Dim arrColObligatory() As Integer = {}
        Dim Arr As New ArrayList
        AddColVisible(tdbgD, SPLIT0, Arr, arrColObligatory, , , gbUnicode)
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbgD, Arr)
        'End If
        'Form trong DLL
        ''CallShowD99F2222(Me, ResetTableByGrid(usrOption, dtCaptionCols.DefaultView.ToTable), dtFind, gsGroupColumns)'Nếu có sử dụng F12 cũ D09U1111
        CallShowD99F2222(Me, dtCaptionCols, dtGridDetail, gsGroupColumns)

    End Sub




    Private Sub tdbgD_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbgD.Columns(COLD_RemainQTY).DataField, DxxFormat.D07_QuantityDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgD.Columns(COLD_RemainAMT).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgD.Columns(COLD_InventoryQTY).DataField, DxxFormat.D07_QuantityDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgD.Columns(COLD_InventoryAMT).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgD.Columns(COLD_DifferenceQTY).DataField, DxxFormat.D07_QuantityDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgD.Columns(COLD_DifferenceRemainAMT).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgD.Columns(COLD_CurrentCost).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)

        InputNumber(tdbgD, arr)
    End Sub

    Dim bIsInventory As Boolean = False
    Private Sub mnsInventoryAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsInventoryAsset.Click
        _FormState = EnumFormState.FormOther
        LoadNewInventory()
    End Sub

    Private Sub LoadNewInventory() 'tạo mới cập nhật nhật kiểm kê
        btnSave.Enabled = True
        btnNotSave.Enabled = True
        pnlFilter.Enabled = False
        'pnlInput.Enabled = False
        btnImport.Enabled = True
        btnExport.Enabled = True
        tdbgM.Enabled = False
        ReadOnlyControl(pnlFilter, pnlMaster)
        UnReadOnlyControl(True, tdbcVoucherTypeID)
        LockColumns(True)
        tdbgD.Splits(0).DisplayColumns(COLD_InventoryQTY).Locked = False
        tdbgD.Splits(0).DisplayColumns(COLD_InventoryAMT).Locked = False
        tdbgD.Splits(0).DisplayColumns(COLD_TransDesc).Locked = False
        tdbgD.Splits(0).DisplayColumns(COLD_InventoryQTY).Style.ResetBackColor()
        tdbgD.Splits(0).DisplayColumns(COLD_InventoryAMT).Style.ResetBackColor()
        tdbgD.Splits(0).DisplayColumns(COLD_TransDesc).Style.ResetBackColor()

        If tdbgM.Columns(COLM_VoucherID).Tag Is Nothing OrElse tdbgM.Columns(COLM_VoucherID).Text <> tdbgM.Columns(COLM_VoucherID).Tag.ToString Then
            If dtGrid Is Nothing Then Exit Sub 'Chưa đổ nguồn cho lưới
            If dtGrid.Rows.Count = 0 Then Exit Sub 'Chưa đổ nguồn cho lưới

            dtGridDetail = ReturnDataTable(SQLStoreD02P2061(2))
            LoadTDBGridD(2)
            LoadTDBCVoucherTypeID()
            tdbcVoucherTypeID.SelectedValue = "-1"
            txtVoucherNo.Text = ""
            c1dateVoucherDate.Value = Now.Date
            tdbcEmployeeID.SelectedValue = tdbgM.Columns(COLM_EmployeeID).Text
            txtDescription.Text = ""
        End If
    End Sub

    Private Sub LockStatus()
        If _FormState = EnumFormState.FormEdit Then
            If tdbgM.Columns(COLM_TransactionTypeID).Text = "BBKK" Then
                tdbgD.Splits(0).DisplayColumns(COLD_Status).Locked = False
                tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Status).Button = True
                tdbgD.Splits(0).DisplayColumns(COLD_Status).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
            End If
            If tdbgM.Columns(COLM_TransactionTypeID).Text = "NKK" Then
                tdbgD.Splits(0).DisplayColumns(COLD_Status).Locked = True
                tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Status).Button = False
                tdbgD.Splits(0).DisplayColumns(COLD_Status).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            tdbgD.Splits(0).DisplayColumns(COLD_Status).Locked = False
            tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Status).Button = True
            tdbgD.Splits(0).DisplayColumns(COLD_Status).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        Else
            tdbgD.Splits(0).DisplayColumns(COLD_Status).Locked = True
            tdbgD.Splits(SPLIT0).DisplayColumns(COLD_Status).Button = False
            tdbgD.Splits(0).DisplayColumns(COLD_Status).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        End If
        
    End Sub

    Private Sub mnsInventoryAssetDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsInventoryAssetDelete.Click
        'Dim sSQL As String = SQLUpdateD02T2061s().ToString & vbCrLf
        'sSQL &= SQLStoreD02P2062(3).ToString
        'Dim bRunSQL As Boolean = ExecuteSQL(sSQL)
        'If bRunSQL Then
        '    D99C0008.Msg(rL3("Da_xoa_kiem_ke"))
        '    LoadTDBGridM(, _sVoucherID)
        'End If
    End Sub


#Region "ExportExcel"

    'Dim EXL As Object = CreateObject("Excel.Application")
    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        '        Dim fileName As String = Application.StartupPath + "\ExportData.xls"
        '        ' step 1: create a new workbook
        '        Dim C1XLBook1 As New C1XLBook()
        '        ' step 2: write content into some cells
        '        Dim sheet As XLSheet = C1XLBook1.Sheets(0)

        '        Dim i, j As Integer
        '        Dim k, l As Integer
        '        k = 1
        '        Dim sFileNames As String = "VoucherID,TransactionID,AssetID, OVoucherNo,ObjectID,ObjectTypeID,LocationID,RemainQTY,RemainAMT, InventoryQTY,InventoryAMT"
        '        Dim style1 As New XLStyle(C1XLBook1)
        '        style1.Font = New Font("Tahoma", 10, FontStyle.Regular)
        '        style1.BackColor = Color.RoyalBlue
        '        style1.ForeColor = Color.White
        '        For j = 0 To tdbgD.Columns.Count - 1

        '            If (sFileNames.Contains(tdbgD.Columns(j).DataField)) Then
        '                sheet(0, k).Value = tdbgD.Columns(j).DataField
        '                sheet(0, k).Style = style1
        '                k = k + 1
        '            End If
        '        Next
        '        k = 1
        '        For i = 0 To tdbgD.RowCount - 1
        '            l = 1
        '            For j = 0 To tdbgD.Columns.Count - 1
        '                If (sFileNames.Contains(tdbgD.Columns(j).DataField) AndAlso L3Bool(tdbgD(i, COLD_IsSelected))) Then
        '                    sheet(k, l).Value = tdbgD(i, j)
        '                    l = l + 1
        '                End If
        '            Next j
        '            k = k + 1
        '        Next i
        '        AutoSizeColumns(sheet, C1XLBook1)
        '        ' step 3: save the file

        'ErrorOpenFile:
        '        Try
        '            C1XLBook1.Save(fileName)
        '            EXL.Workbooks.Open(fileName)
        '            EXL.Visible = True
        '        Catch ex As Exception
        '            If CloseProcessWindow(fileName) Then GoTo ErrorOpenFile
        '        End Try
        '        'System.Diagnostics.Process.Start(fileName)

        '        Me.Cursor = Cursors.Default
        ' Dim sFileNames As String = "VoucherID - an,TransactionID-an,AssetID-hien, OVoucherNo,ObjectID,ObjectTypeID,LocationID,RemainQTY,RemainAMT, InventoryQTY,InventoryAMT"
        Me.Cursor = Cursors.WaitCursor
        Dim arrColAlwaysShow() As String = {"VoucherID", "TransactionID"}
        Dim arrColAlwaysHide() As String = {"IsSelected", "Status", "AssetName", "transDesc", "ObjectName", "AssetTypeName"}
        ExportToExcelFromGrid(tdbgD, "ExportData.xls", , arrColAlwaysShow, arrColAlwaysHide)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click

        'Me.Cursor = Cursors.WaitCursor
        'dtImport = dtGridDetail.Copy
        'If dtImport IsNot Nothing Then dtImport.Clear()
        'If Not ImportExcelToGrid(dtImport) Then Exit Sub
        'For i As Integer = 0 To dtGridDetail.Rows.Count - 1
        '    For j As Integer = 0 To dtImport.Rows.Count - 1
        '        If dtGridDetail.Rows(i).Item("AssetID").ToString = dtImport.Rows(j).Item("AssetID").ToString AndAlso dtGridDetail.Rows(i).Item("VoucherID").ToString = dtImport.Rows(j).Item("VoucherID").ToString AndAlso dtGridDetail.Rows(i).Item("TransactionID").ToString = dtImport.Rows(j).Item("TransactionID").ToString Then
        '            dtGridDetail.Rows(i).Item("InventoryQTY") = dtImport.Rows(i).Item("InventoryQTY")
        '            dtGridDetail.Rows(i).Item("InventoryAMT") = dtImport.Rows(i).Item("InventoryAMT")
        '        End If
        '    Next
        'Next
        'LoadDataSource(tdbgD, dtGridDetail, gbUnicode)
        'ReLoadTDBGridD()
        'Me.Cursor = Cursors.Default

        Me.Cursor = Cursors.WaitCursor
        Dim dtImport As DataTable = dtGridDetail.Clone
        If dtImport IsNot Nothing Then dtImport.Clear()
        Dim sOutputPath As String = ""
        dtImport.Columns("ALVoucherID").AllowDBNull = True
        dtImport.Columns("ALTransactionID").AllowDBNull = True
        dtImport.Columns("AssetID").AllowDBNull = True

        If ImportExcelToGrid(dtImport, , sOutputPath) Then
            Me.Cursor = Cursors.Default
            For i As Integer = 0 To dtGridDetail.Rows.Count - 1
                For j As Integer = 0 To dtImport.Rows.Count - 1
                    If dtGridDetail.Rows(i).Item("AssetID").ToString = dtImport.Rows(j).Item("AssetID").ToString AndAlso dtGridDetail.Rows(i).Item("VoucherID").ToString = dtImport.Rows(j).Item("VoucherID").ToString AndAlso dtGridDetail.Rows(i).Item("TransactionID").ToString = dtImport.Rows(j).Item("TransactionID").ToString Then
                        dtGridDetail.Rows(i).Item("InventoryQTY") = dtImport.Rows(i).Item("InventoryQTY")
                        dtGridDetail.Rows(i).Item("InventoryAMT") = dtImport.Rows(i).Item("InventoryAMT")
                        dtGridDetail.Rows(i).Item("TransDesc") = dtImport.Rows(i).Item("TransDesc")
                    End If
                Next
            Next
            LoadDataSource(tdbgD, dtGridDetail, gbUnicode)
            ReLoadTDBGridD()
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
    End Sub

    'Private Sub AutoSizeColumns(ByVal sheet As XLSheet, ByVal C1XLBook1 As C1XLBook)
    '    Dim iRowStart As Integer = 0
    '    Using g As Graphics = Graphics.FromHwnd(IntPtr.Zero)
    '        Dim r As Integer, c As Integer
    '        For c = 0 To sheet.Columns.Count - 1
    '            Dim colWidth As Integer = -1
    '            'For r = 0 To sheet.Rows.Count - 1
    '            For r = iRowStart To sheet.Rows.Count - 1
    '                Dim value As Object = sheet(r, c).Value
    '                If Not (value Is Nothing) Then
    '                    ' get value (unformatted at this point)
    '                    ' get font (default or style)
    '                    Dim font As Font = C1XLBook1.DefaultFont
    '                    Dim s As XLStyle = sheet(r, c).Style
    '                    If Not (s Is Nothing) Then
    '                        If Not (s.Font Is Nothing) Then
    '                            font = s.Font
    '                        End If
    '                    End If
    '                    ' measure string (add a little tolerance)
    '                    Dim sz As Size
    '                    If Not IsDBNull(value) Then
    '                        sz = System.Drawing.Size.Ceiling(g.MeasureString(CStr(value) + "XX", font))
    '                    End If
    '                    ' keep widest so far
    '                    If sz.Width > colWidth Then
    '                        colWidth = sz.Width
    '                    End If
    '                End If
    '                ' done measuring, set column width
    '                If colWidth > -1 Then
    '                    sheet.Columns(c).Width = C1XLBook.PixelsToTwips(colWidth)
    '                End If
    '            Next
    '        Next
    '    End Using
    'End Sub

    'Private Function CloseProcessWindow(ByVal fileName As String, Optional ByVal bShowMessage As Boolean = True) As Boolean
    '    Dim bClosed As Boolean = False
    '    Try
    '        For Each wbExcel As Object In EXL.Workbooks
    '            If wbExcel.FullName = fileName Then
    '                If bShowMessage Then
    '                    If (D99C0008.MsgAsk(rL3("Ban_phai_dong_File") & Space(1) & fileName.Substring(fileName.LastIndexOf("\") + 1) & Space(1) & rL3("truoc_khi_xuat_Excel") & "." & vbCrLf & rL3("Ban_co_muon_dong_khong")) = Windows.Forms.DialogResult.Yes) Then
    '                        wbExcel.Save()
    '                        wbExcel.Close()
    '                        If EXL.Workbooks.Count = 0 Then
    '                            EXL.Visible = False
    '                        End If
    '                        Return True
    '                    Else
    '                        Return False
    '                    End If
    '                Else
    '                    wbExcel.Save()
    '                    wbExcel.Close()
    '                    If EXL.Workbooks.Count = 0 Then
    '                        EXL.Visible = False
    '                    End If
    '                    Return True
    '                End If
    '            End If
    '            bClosed = True
    '        Next
    '    Catch ex As Exception

    '    End Try
    '    'Doan code dung de dong file Excel mo san (khong phai do Chuong trinh mo)
    '    If Not bClosed Then
    '        Dim p As System.Diagnostics.Process = Nothing
    '        Dim sWindowName As String = "Microsoft Excel - ExportData.xls"
    '        Try
    '            For Each pr As Process In Process.GetProcessesByName("EXCEL")
    '                'Update 05/04/2013
    '                If pr.MainWindowTitle.Contains(sWindowName) OrElse pr.MainWindowTitle = sWindowName.Substring(0, sWindowName.LastIndexOf(".")) Then
    '                    If p Is Nothing Then
    '                        p = pr
    '                    ElseIf p.StartTime < pr.StartTime Then
    '                        p = pr
    '                    End If
    '                End If
    '            Next
    '            If p IsNot Nothing Then
    '                'Update 05/04/2013
    '                Me.BringToFront()
    '                Me.Activate()
    '                If (D99C0008.MsgAsk(rL3("Ban_phai_dong_File") & Space(1) & fileName.Substring(fileName.LastIndexOf("\") + 1) & Space(1) & rL3("truoc_khi_xuat_Excel") & "." & vbCrLf & rL3("Ban_co_muon_dong_khong")) = Windows.Forms.DialogResult.Yes) Then
    '                    p.Kill()
    '                    Return True
    '                Else
    '                    Return False
    '                End If
    '            End If
    '            Return False
    '        Catch ex As Exception
    '        End Try
    '    End If

    'End Function

#End Region
    Dim bSelect As Boolean = False 'Mặc định Uncheck - tùy thuộc dữ liệu database
    Private Sub HeadClick(ByVal iCol As Integer)
        If tdbgD.RowCount <= 0 Then Exit Sub
        Select Case iCol
            Case COLD_IsSelected
                L3HeadClick(tdbgD, iCol, bSelect)
        End Select
    End Sub

    Private Sub tdbgD_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgD.HeadClick
        HeadClick(e.ColIndex)
    End Sub

    Private Sub mnsPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsPrint.Click, tsbPrint.Click
        Me.Cursor = Cursors.WaitCursor
        Print(Me, Me.Name)
        Me.Cursor = Cursors.Default
    End Sub

    Dim report As D99C2003
    Private Sub Print(ByVal form As Form, ByVal sReportTypeID As String, Optional ByVal ModuleID As String = "02")
        Dim sReportName As String = ""
        Dim sReportPath As String = ""
        Dim sReportTitle As String = ""
        Dim sCustomReport As String = ""
        Dim file As String = D99D0541.GetReportPathNew(ModuleID, sReportTypeID, sReportName, sCustomReport, sReportPath, sReportTitle)
        If sReportName = "" Then Exit Sub

        Dim sSQL As String = ""
        Select Case file.ToLower
            Case "rpt"
                printReport(sReportName, sReportPath, sReportTitle, sSQL) ' ID : 262792
            Case Else
                D99D0541.PrintOfficeType(sReportTypeID, sReportName, sReportPath, file, dtGridDetail)
        End Select
    End Sub


    Private Sub printReport(ByVal sReportName As String, ByVal sReportPath As String, ByVal sReportCaption As String, ByVal sSQL As String)
        If Not AllowNewD99C2003(report, Me) Then Exit Sub
        Dim conn As New SqlConnection(gsConnectionString)
        With report
            .OpenConnection(conn)
            Dim sSQLSub As String = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
            Dim sSubReport As String = "D02R0000"
            UnicodeSubReport(sSubReport, sSQLSub, gsDivisionID, gbUnicode)
            .AddSub(sSQLSub, sSubReport & ".rpt")
            .AddMain(dtGridDetail)
            .PrintReport(sReportPath, sReportCaption & " - " & sReportName)
        End With
    End Sub

    Private Sub mnsPrintDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsPrintDetail.Click
        Dim frm As New D02F2061
        frm.VoucherID = tdbgM.Columns(COLM_VoucherID).Text
        frm.ShowDialog()
        frm.Dispose()
    End Sub

    Private Sub btnF12_Click(sender As Object, e As EventArgs) Handles btnF12.Click
        If usrOption Is Nothing Then Exit Sub 'TH lưới không có cột
        'usrOption.Location = tdbgD.Location 
        usrOption.Location = New Point(tdbgD.Location.X + grpMaster.Width + 10, 160)
        Me.Controls.Add(usrOption)
        usrOption.Height = tdbgD.Height
        usrOption.BringToFront()
        usrOption.Visible = True
    End Sub
    Private Sub CallD99U1111()
        Dim arrColObligatory() As Object = {COLD_AssetID, COLD_IsSelected}
        usrOption.AddColVisible(tdbgD, dtF12, arrColObligatory)
        If usrOption IsNot Nothing Then usrOption.Dispose()
        usrOption = New D99U1111(Me, tdbgD, dtF12)
        usrOption.Anchor = CType(EnumAnchorStyles.TopLeftBottom, System.Windows.Forms.AnchorStyles)
    End Sub
End Class
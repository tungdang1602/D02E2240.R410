Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text
Imports System.IO
Imports System

Public Class D02F1001

#Region "Const of tdbg"
    Private Const COL_TransactionID As Integer = 0     ' TransactionID
    Private Const COL_ProjectID As Integer = 1         ' ProjectID
    Private Const COL_PropertyProductID As Integer = 2 ' PropertyProductID
    Private Const COL_BatchID As Integer = 3           ' BatchID
    Private Const COL_CipNo As Integer = 4             ' CipNo
    Private Const COL_ModuleID As Integer = 5          ' ModuleID
    Private Const COL_Status As Integer = 6            ' Status
    Private Const COL_Choose As Integer = 7            ' Chọn
    Private Const COL_VoucherTypeID As Integer = 8     ' Loại phiếu
    Private Const COL_VoucherNo As Integer = 9         ' Số phiếu
    Private Const COL_VoucherDate As Integer = 10      ' Ngày phiếu
    Private Const COL_RefDate As Integer = 11          ' Ngày hóa đơn
    Private Const COL_SeriNo As Integer = 12           ' Số Sêri
    Private Const COL_RefNo As Integer = 13            ' Số hóa đơn
    Private Const COL_ObjectTypeID As Integer = 14     ' Loại đối tượng
    Private Const COL_ObjectID As Integer = 15         ' Mã đối tượng
    Private Const COL_ObjectName As Integer = 16       ' Tên đối tượng
    Private Const COL_Description As Integer = 17      ' Diễn giải
    Private Const COL_DebitAccountID As Integer = 18   ' Tài khoản nợ
    Private Const COL_CreditAccountID As Integer = 19  ' Tài khoản có
    Private Const COL_CurrencyID As Integer = 20       ' Loại tiền
    Private Const COL_ExchangeRate As Integer = 21     ' Tỷ giá
    Private Const COL_OriginalAmount As Integer = 22   ' Số tiền nguyên tệ
    Private Const COL_ConvertedAmount As Integer = 23  ' Số tiền quy đổi
    Private Const COL_VATGroupID As Integer = 24       ' Nhóm thuế
    Private Const COL_VATTypeID As Integer = 25        ' Loại thuế
    Private Const COL_VATNo As Integer = 26            ' Mã số thuế
    Private Const COL_SourceID As Integer = 27         ' Nguồn vốn
    Private Const COL_DivisionID As Integer = 28       ' DivisionID
    Private Const COL_TranMonth As Integer = 29        ' TranMonth
    Private Const COL_TranYear As Integer = 30         ' TranYear
    Private Const COL_Internal As Integer = 31         ' Internal
    Private Const COL_Str01 As Integer = 32            ' Str01
    Private Const COL_Str02 As Integer = 33            ' Str02
    Private Const COL_Str03 As Integer = 34            ' Str03
    Private Const COL_Str04 As Integer = 35            ' Str04
    Private Const COL_Str05 As Integer = 36            ' Str05
    Private Const COL_Num01 As Integer = 37            ' Num01
    Private Const COL_Num02 As Integer = 38            ' Num02
    Private Const COL_Num03 As Integer = 39            ' Num03
    Private Const COL_Num04 As Integer = 40            ' Num04
    Private Const COL_Num05 As Integer = 41            ' Num05
    Private Const COL_Date01 As Integer = 42           ' Date01
    Private Const COL_Date02 As Integer = 43           ' Date02
    Private Const COL_Date03 As Integer = 44           ' Date03
    Private Const COL_Date04 As Integer = 45           ' Date04
    Private Const COL_Date05 As Integer = 46           ' Date05
    Private Const COL_Ana01ID As Integer = 47          ' Ana01ID
    Private Const COL_Ana02ID As Integer = 48          ' Ana02ID
    Private Const COL_Ana03ID As Integer = 49          ' Ana03ID
    Private Const COL_Ana04ID As Integer = 50          ' Ana04ID
    Private Const COL_Ana05ID As Integer = 51          ' Ana05ID
    Private Const COL_Ana06ID As Integer = 52          ' Ana06ID
    Private Const COL_Ana07ID As Integer = 53          ' Ana07ID
    Private Const COL_Ana08ID As Integer = 54          ' Ana08ID
    Private Const COL_Ana09ID As Integer = 55          ' Ana09ID
    Private Const COL_Ana10ID As Integer = 56          ' Ana10ID
#End Region

#Region "Const of tdbg2 - Total of Columns: 59"
    Private Const COL2_TransactionID As Integer = 0     ' TransactionID
    Private Const COL2_ProjectID As Integer = 1         ' ProjectID
    Private Const COL2_PropertyProductID As Integer = 2 ' PropertyProductID
    Private Const COL2_BatchID As Integer = 3           ' BatchID
    Private Const COL2_CipNo As Integer = 4             ' CipNo
    Private Const COL2_ModuleID As Integer = 5          ' ModuleID
    Private Const COL2_Status As Integer = 6            ' Status
    Private Const COL2_Choose As Integer = 7            ' Chi phí
    Private Const COL2_IsNotAllocate As Integer = 8     ' Không tính KH
    Private Const COL2_SourceID As Integer = 9          ' Nguồn vốn
    Private Const COL2_VoucherTypeID As Integer = 10    ' Loại phiếu
    Private Const COL2_VoucherNo As Integer = 11        ' Số phiếu
    Private Const COL2_VoucherDate As Integer = 12      ' Ngày phiếu
    Private Const COL2_RefDate As Integer = 13          ' Ngày hóa đơn
    Private Const COL2_SeriNo As Integer = 14           ' Số Sêri
    Private Const COL2_RefNo As Integer = 15            ' Số hóa đơn
    Private Const COL2_ObjectTypeID As Integer = 16     ' Loại đối tượng
    Private Const COL2_ObjectID As Integer = 17         ' Mã đối tượng
    Private Const COL2_ObjectName As Integer = 18       ' Tên đối tượng
    Private Const COL2_Description As Integer = 19      ' Diễn giải
    Private Const COL2_DebitAccountID As Integer = 20   ' Tài khoản nợ
    Private Const COL2_CreditAccountID As Integer = 21  ' Tài khoản có
    Private Const COL2_CurrencyID As Integer = 22       ' Loại tiền
    Private Const COL2_ExchangeRate As Integer = 23     ' Tỷ giá
    Private Const COL2_OriginalAmount As Integer = 24   ' Số tiền nguyên tệ
    Private Const COL2_ConvertedAmount As Integer = 25  ' Số tiền quy đổi
    Private Const COL2_VATGroupID As Integer = 26       ' Nhóm thuế
    Private Const COL2_VATTypeID As Integer = 27        ' Loại thuế
    Private Const COL2_VATNo As Integer = 28            ' Mã số thuế
    Private Const COL2_DivisionID As Integer = 29       ' DivisionID
    Private Const COL2_TranMonth As Integer = 30        ' TranMonth
    Private Const COL2_TranYear As Integer = 31         ' TranYear
    Private Const COL2_Internal As Integer = 32         ' Internal
    Private Const COL2_Str01 As Integer = 33            ' Str01
    Private Const COL2_Str02 As Integer = 34            ' Str02
    Private Const COL2_Str03 As Integer = 35            ' Str03
    Private Const COL2_Str04 As Integer = 36            ' Str04
    Private Const COL2_Str05 As Integer = 37            ' Str05
    Private Const COL2_Num01 As Integer = 38            ' Num01
    Private Const COL2_Num02 As Integer = 39            ' Num02
    Private Const COL2_Num03 As Integer = 40            ' Num03
    Private Const COL2_Num04 As Integer = 41            ' Num04
    Private Const COL2_Num05 As Integer = 42            ' Num05
    Private Const COL2_Date01 As Integer = 43           ' Date01
    Private Const COL2_Date02 As Integer = 44           ' Date02
    Private Const COL2_Date03 As Integer = 45           ' Date03
    Private Const COL2_Date04 As Integer = 46           ' Date04
    Private Const COL2_Date05 As Integer = 47           ' Date05
    Private Const COL2_Ana01ID As Integer = 48          ' Ana01ID
    Private Const COL2_Ana02ID As Integer = 49          ' Ana02ID
    Private Const COL2_Ana03ID As Integer = 50          ' Ana03ID
    Private Const COL2_Ana04ID As Integer = 51          ' Ana04ID
    Private Const COL2_Ana05ID As Integer = 52          ' Ana05ID
    Private Const COL2_Ana06ID As Integer = 53          ' Ana06ID
    Private Const COL2_Ana07ID As Integer = 54          ' Ana07ID
    Private Const COL2_Ana08ID As Integer = 55          ' Ana08ID
    Private Const COL2_Ana09ID As Integer = 56          ' Ana09ID
    Private Const COL2_Ana10ID As Integer = 57          ' Ana10ID
    Private Const COL2_TaskID As Integer = 58           ' TaskID
#End Region

#Region "Const of tdbg3"
    Private Const COL3_AssignmentID As String = "AssignmentID"     ' Mã phân bổ
    Private Const COL3_AssignmentName As String = "AssignmentName" ' Tên phân bổ
    Private Const COL3_DebitAccountID As String = "DebitAccountID" ' TK Nợ
    Private Const COL3_PercentAmount As String = "PercentAmount"   ' Tỷ lệ
    Private Const COL3_Extend As String = "Extend"                 ' Extend
    Private Const COL3_HistoryID As String = "HistoryID"           ' HistoryID
#End Region
    Dim iPerD02F0100 As Integer = ReturnPermission("D02F0100") 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
    Dim iPerD02F0087 As Integer = ReturnPermission("D02F0087") 'ID : 224617 - BỔ SUNG Cho phép gọi màn hình THIẾT LẬP DANH MỤC TÀI SẢN CỐ ĐỊNH tại bước hình thành TS
    '-------Biến khai báo cho khoản mục
    Dim bUseAna As Boolean 'Kiểm tra có sử dụng Khoản mục không, để set thuộc tính Enabled nút Khoản mục 
    Dim sAuditCode As String = "PurAsset"
    Dim sCreateUserID As String
    Dim sCreateDate As String
    'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b1:
    Dim _setupFrom As String = "NEW"
    Private _assetID As String = ""
    Dim clsFilterCombo As Lemon3.Controls.FilterCombo
    Dim clsFilterDropdown As Lemon3.Controls.FilterDropdown
    Dim dtObjectID As DataTable
    Dim dtManagementID As DataTable

    Public Property AssetID() As String
        Get
            Return _assetID
        End Get
        Set(ByVal Value As String)
            _assetID = Value
        End Set
    End Property
    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            _FormState = value
            '--- Chuẩn Khoản mục b2: Lấy caption cho 10 khoản mục
            bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, tdbg.Splits.Count - 1, True, gbUnicode)
            '--- Chuẩn Khoản mục b21: D91 có sử dụng Khoản mục
            If bUseAna Then
                LoadCaptionAnatdbg2()
            Else
                tdbg.RemoveHorizontalSplit(tdbg.Splits.Count - 1)
                '3/3/2017, id 95439-Lỗi chọn mã phân bổ khi sử dụng combo theo dạng mới
                tdbg.Splits(1).SplitSize = tdbg.Splits(1).SplitSize + 9 'cộng thêm size của split vừa xóa

                tdbg2.RemoveHorizontalSplit(tdbg2.Splits.Count - 1)
            End If
            '------------------------------------
            'Lưới 1 không thấy hiển thị Thông tin phụ
            Dim bUseSubInfo As Boolean = LoadCaptionSubInfo() 'load caption các thông tin phụ
            If Not bUseSubInfo Then
                tdbg2.RemoveHorizontalSplit(SPLIT2)
            End If

            '3/3/2017, id 95439-Lỗi chọn mã phân bổ khi sử dụng combo theo dạng mới
            If bUseAna = False And bUseSubInfo = False Then
                tdbg2.Splits(1).SplitSize = tdbg2.Splits(1).SplitSize + 12 'cộng thêm size của 2 split vừa xóa
            ElseIf bUseAna = False Or bUseSubInfo = False Then
                tdbg2.Splits(1).SplitSize = tdbg2.Splits(1).SplitSize + 6 'cộng thêm size của split vừa xóa
            End If

            LoadTDBCombo()
            LoadTDBDropDown()
            'LoadTDBComboPropertyProductID()
            clsFilterCombo = New Lemon3.Controls.FilterCombo
            clsFilterCombo.CheckD91 = False 'Giá trị mặc định True: kiểm tra theo DxxFormat.LoadFormNotINV. Ngược lại luôn luôn Filter dạng mới (dùng cho Novaland)
            clsFilterCombo.AddPairObject(tdbcObjectTypeID, tdbcObjectID) 'Đã bổ sung cột Loại ĐT
            clsFilterCombo.AddPairObject(tdbcManagementObjTypeID, tdbcManagementObjID) 'Đã bổ sung cột Loại ĐT

            clsFilterCombo.UseFilterComboObjectID()
            clsFilterCombo.UseFilterCombo(tdbcAssetID, tdbcEmployeeID)
            clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
            clsFilterCombo.LoadtdbcObjectID(tdbcManagementObjID, dtManagementID, ReturnValueC1Combo(tdbcManagementObjTypeID))

            ' Dim dic As New Dictionary(Of String, String)
            ' dic.Add(tdbdBudgetID.Name, "Note")'Ví dụ Cần lấy cột Note trong tdbdBudgetID. Nếu lấy Name, hoặc Description hoặc cột 1 thì không cần truyền
            clsFilterDropdown = New Lemon3.Controls.FilterDropdown()
            'clsFilterDropdown.SingleLine = True'Mặc đinh False. Chọn nhiều dòng gắn lại dữ liệu cho 1 dòng và cách nhau bằng ; (sử dụng Tài khoản D90F1110)
            clsFilterDropdown.CheckD91 = True 'Giá trị mặc định True: kiểm tra theo DxxFormat.LoadFormNotINV. Ngược lại luôn luôn Filter dạng mới (dùng cho Novaland)
            ' clsFilterDropdown.DicDDName = dic
            clsFilterDropdown.UseFilterDropdown(tdbg3, COL3_AssignmentID)
            'clsFilterDropdown.UseFilterDropdown(tdbg2, COL2_ObjectID)'Nếu dùng nhiều lưới

            txtServiceLife.Tag = txtServiceLife.Text
            txtPercentage.Tag = ""

            Select Case _FormState
                Case EnumFormState.FormAdd
                    LoadAddNew()
                Case EnumFormState.FormEdit
                    LoadEdit()
                Case EnumFormState.FormEditOther
                    LoadEdit()
                    LockCtrlEditOther()
                Case EnumFormState.FormView
                    LoadEdit()
                    btnSave.Enabled = False
            End Select
        End Set
    End Property

    Private Sub LockCtrlEditOther()
        grpAssetID.Enabled = False
        tabMain.Enabled = False
        btnNext.Enabled = False
        ReadOnlyAll(grpFinancialInfo, c1dateUseDate, c1dateAssetDate, c1dateDepDate)
    End Sub

    Private Sub LoadCaptionAnatdbg2()
        For i As Integer = 1 To 10
            Dim sField As String = "Ana01ID".Replace("01", i.ToString("00"))
            tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(sField).Locked = True
            tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(sField).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)

            tdbg2.Columns(sField).Caption = tdbg.Columns(sField).Caption
            tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(sField).Locked = True
            tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(sField).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
            tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(sField).Visible = tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(sField).Visible
            tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(sField).HeadingStyle.Font = FontUnicode(gbUnicode)

            tdbg2.Columns(sField).Tag = tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(sField).Visible
        Next
    End Sub

    Private Sub D02F1001_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '12/10/2020, id 144622-Tài sản cố định_Lỗi chưa cảnh báo khi lưu
        If _FormState = EnumFormState.FormEdit Then
            If Not gbSavedOK Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            If (tdbcAssetID.Text <> "" Or tdbcAssetAccountID.Text <> "" Or tdbcDepAccountID.Text <> "" Or tdbcObjectTypeID.Text <> "" Or tdbcObjectID.Text <> "" Or tdbcEmployeeID.Text <> "" Or txtServiceLife.Text <> "" Or c1dateBeginUsing.Text <> "" Or c1dateBeginDep.Text <> "" Or c1dateDepDate.Text <> "") Then
                If Not gbSavedOK Then
                    If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub D02F1002_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me)
                Exit Sub
            Case Keys.F11
                HotKeyF11(Me, tdbg)
                Exit Sub
        End Select

        'If e.Control And e.KeyCode = Keys.F1 Then
        '    btnHotKeys_Click(Nothing, Nothing)
        '    Exit Sub
        'End If

        If e.Alt Then
            Select Case e.KeyCode
                Case Keys.D1, Keys.NumPad1
                    tabMain.SelectedTab = tabPage1
                    tdbg.Focus()
                Case Keys.D2, Keys.NumPad2
                    tabMain.SelectedTab = tabPage2
                    tdbg2.Focus()
                Case Keys.D3, Keys.NumPad3
                    tabMain.SelectedTab = TabPage3
                    tdbg3.Focus()
            End Select
        End If
    End Sub

    Dim sOldPropertyProductID As String
    Private Sub D02F1002_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        gbEnabledUseFind = False
        gbSavedOK = False
        InputbyUnicode(Me, gbUnicode)
        Loadlanguage()
        SetBackColorObligatory()
        'mnuFind
        mnsFind.Image = My.Resources.find
        mnsFind.Text = rL3("Tim__kiem") 'Tìm &kiếm
        CType(mnsFind, System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)

        'mnuListAll
        mnsListAll.Image = My.Resources.ListAll
        mnsListAll.Text = rL3("_Liet_ke_tat_ca") '&Liệt kế tất cả
        CType(mnsListAll, System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)

        'Nhập Ngày
        InputDateInTrueDBGrid(tdbg, COL_RefDate)
        InputDateInTrueDBGrid(tdbg2, COL2_RefDate, COL2_Date01, COL2_Date02, COL2_Date03, COL2_Date04, COL2_Date05)
        '*************
        tdbg_LockedColumns()
        tdbg2_LockedColumns()
        tdbg_NumberFormat()
        'LoadTDBGrid1()
        LoadTDBGrid2()
        LoadTDBGrid3()
        Dim arrPercent() As FormatColumn = Nothing
        AddPercentColumns(arrPercent, COL3_PercentAmount)
        InputNumber(tdbg3, arrPercent)
        '*****************
        ResetColorGrid(tdbg, 0, tdbg.Splits.Count - 1)
        ResetColorGrid(tdbg2, 0, tdbg2.Splits.Count - 1)
        ResetFooterGrid(tdbg3, 0, tdbg3.Splits.Count - 1)
        tdbg.Splits(0).MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder

        tdbg2.Splits(0).MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
        tdbg2.Splits(1).MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor
        tdbg2.Splits(1).DisplayColumns(COL2_IsNotAllocate).Visible = D02Systems.UseProperty
        If tdbg2.Splits.Count >= 2 Then tdbg2.Splits(1).MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor
        ResetSplitDividerSize(tdbg)
        ResetSplitDividerSize(tdbg2)
        If Not D02Systems.CIPforPropertyProduct Then
            pnlProperty.Visible = False
            txtAssetName.Width = txtAssetName.Width + pnlProperty.Width
        End If
        sOldPropertyProductID = tdbcPropertyProductID.Text
        LockControlByAsset() '23/1/2018, Phạm Thị Thu: id 105627-Mặc định tài khoản TSCD không đươc phép sửa khi chọn mã TSCD đi hình thành
        LockServiceLife() '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        SetResolutionForm(Me, ContextMenuStrip1)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg_NumberFormat()
        'tdbg.Columns(COL_ExchangeRate).NumberFormat = DxxFormat.ExchangeRateDecimals
        'tdbg.Columns(COL_OriginalAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        'tdbg.Columns(COL_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals

        'tdbg2.Columns(COL2_ExchangeRate).NumberFormat = DxxFormat.ExchangeRateDecimals
        'tdbg2.Columns(COL2_OriginalAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        'tdbg2.Columns(COL2_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg_NumberFormat1()
        tdbg2_NumberFormat2()
    End Sub

    Private Sub tdbg_NumberFormat1()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg.Columns(COL_ExchangeRate).DataField, DxxFormat.ExchangeRateDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_OriginalAmount).DataField, DxxFormat.DecimalPlaces, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_ConvertedAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg, arr)
    End Sub

    Private Sub tdbg2_NumberFormat2()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg2.Columns(COL2_ExchangeRate).DataField, DxxFormat.ExchangeRateDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg2.Columns(COL2_OriginalAmount).DataField, DxxFormat.DecimalPlaces, 28, 8)
        AddDecimalColumns(arr, tdbg2.Columns(COL2_ConvertedAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)

        InputNumber(tdbg2, arr)
    End Sub



    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub SetBackColorObligatory()
        tdbcAssetID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAssetAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDepAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        If D02Systems.ObligatoryReceiver Then
            tdbcEmployeeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        End If
        'txtServiceLife.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateBeginUsing.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateBeginDep.BackColor = COLOR_BACKCOLOROBLIGATORY
        If D02Systems.IsCalDepByDate = True Then c1dateDepDate.BackColor = COLOR_BACKCOLOROBLIGATORY '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        tdbg2.Splits(SPLIT1).DisplayColumns(COL2_OriginalAmount).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbg2.Splits(SPLIT1).DisplayColumns(COL2_SourceID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY

        If D02Systems.IsObligatoryManagement Then
            tdbcManagementObjTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
            tdbcManagementObjID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        End If
        'tdbg3.Splits(0).DisplayColumns(COL2_AssignmentID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub LoadtdbcAssetID()
        Dim sUnicode As String = ""
        If gbUnicode Then sUnicode = "U"
        Dim strSQL As String = ""
        strSQL = ""
        strSQL &= "SELECT '+' As AssetID, N'<Thêm Mới>'  As AssetName,'' As AssetAccountID, '' As DepAccountID, " & vbCrLf 'ID : 224617 - BỔ SUNG Cho phép gọi màn hình THIẾT LẬP DANH MỤC TÀI SẢN CỐ ĐỊNH tại bước hình thành TS
        strSQL &= "'' As ObjectTypeID,'' As ObjectID,'' As EmployeeID,'' As  FullName,0 As ConvertedAmount,0 As Percentage, " & vbCrLf
        strSQL &= " 0 As ServiceLife,0 As AmountDepreciation ,Null As AssetDate,'' As  MethodID,'' As MethodEndID,  " & vbCrLf
        strSQL &= "  null as BeginUsing,  " & vbCrLf
        strSQL &= " null as BeginDep,  " & vbCrLf
        strSQL &= "0 As  DepreciatedPeriod,0 As DepreciatedAmount,0 As IsCompleted,0 As IsRevalued,0 As IsDisposed,  " & vbCrLf
        strSQL &= "0 As AssignmentTypeID,null As DeprTableName, null As UseDate,0 As IntCode,null As DepDate, " & vbCrLf
        strSQL &= "'' As ManagementObjTypeID, '' As ManagementObjID" & vbCrLf
        strSQL &= "UNION" & vbCrLf
        strSQL &= "SELECT   T1.AssetID, T1.AssetName" & sUnicode & " as AssetName, T1.AssetAccountID, T1.DepAccountID, " & vbCrLf
        strSQL &= " ObjectTypeID, ObjectID, EmployeeID, FullName" & sUnicode & " as FullName, T1.ConvertedAmount, T1.Percentage, " & vbCrLf
        strSQL &= " ServiceLife, T1.AmountDepreciation ,T1.AssetDate,A.Description" & sUnicode & " as MethodID, B.Description" & sUnicode & " as MethodEndID,  " & vbCrLf
        strSQL &= "  Case when UseMonth <10 then '0' else '' end + ltrim(str(UseMonth)) + '/' + ltrim(str(UseYear)) as BeginUsing, " & vbCrLf
        strSQL &= "  Case when DepMonth <10 then '0' else '' end + ltrim(str(DepMonth)) + '/' + ltrim(str(DepYear)) as BeginDep, " & vbCrLf
        strSQL &= "  T1.DepreciatedPeriod,T1.DepreciatedAmount, IsCompleted, IsRevalued, IsDisposed, " & vbCrLf
        strSQL &= " T1.AssignmentTypeID, C.DeprTableName, T1.UseDate, A.IntCode, T1.DepDate, " & vbCrLf
        strSQL &= " T1.ManagementObjTypeID, T1.ManagementObjID" & vbCrLf
        strSQL &= " FROM D02T0001 T1 WITH(NOLOCK) Inner Join D02T8000 A WITH(NOLOCK) On T1.MethodID = A.IntCode " & vbCrLf
        strSQL &= " Inner Join D02T8000 B WITH(NOLOCK) On T1.MethodEndID = B.IntCode" & vbCrLf
        strSQL &= " Left Join D02T0070 C WITH(NOLOCK) On T1.DeprTableID = C.DeprTableID " & vbCrLf
        strSQL &= " WHERE A.Language =" & SQLString(gsLanguage) & "   And A.ModuleID = '02' And A.Type = 0 " & vbCrLf
        strSQL &= "  And B.Language = " & SQLString(gsLanguage) & "  And B.ModuleID = '02' And B.Type = 1  " & vbCrLf
        strSQL &= " AND T1.DivisionID= " & SQLString(gsDivisionID) & vbCrLf

        If _FormState = EnumFormState.FormAdd Then strSQL &= "And  IsCompleted = 0 " & vbCrLf
        strSQL = strSQL & "ORDER BY AssetID "

        LoadDataSource(tdbcAssetID, strSQL, gbUnicode)
        tdbcAssetID.Splits(0).DisplayColumns(0).Visible = True
        tdbcAssetID.Splits(0).DisplayColumns(2).Visible = False
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcAssetID
        LoadtdbcAssetID()
        'Load tdbcAssetAccount
        'Load tdbcDepAccount
        LoadDataSource(tdbcAssetAccountID, ReturnTableAccountID("AccountStatus = 0 AND GroupID='7'", gbUnicode), gbUnicode)
        LoadDataSource(tdbcDepAccountID, ReturnTableAccountID("AccountStatus = 0 AND GroupID='19'", gbUnicode), gbUnicode)
        'Load tdbcObjectTypeID/ tdbdObjectTypeID
        Dim dtObjectTypeID As DataTable = ReturnTableObjectTypeID(gbUnicode)
        LoadDataSource(tdbcObjectTypeID, dtObjectTypeID, gbUnicode)
        dtObjectID = ReturnDataTable(LoadObjectID)

        Dim dtManagementObjTypeID As DataTable = ReturnTableObjectTypeID(gbUnicode)

        LoadDataSource(tdbcManagementObjTypeID, dtManagementObjTypeID, gbUnicode)
        dtManagementID = ReturnDataTable(LoadObjectID)
        'Load tdbcEmployeeID
        sSQL = "Select ObjectID as EmployeeID, ObjectName" & UnicodeJoin(gbUnicode) & " as EmployeeName From Object WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where ObjectTypeID='NV' Order By ObjectID"
        LoadDataSource(tdbcEmployeeID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTDBComboPropertyProductID(Optional ByVal sPropertyProductID As String = "")
        Dim sSQL As String = ""
        sSQL &= " SELECT  DISTINCT 	PropertyProductID "
        sSQL &= " FROM 			D02T0012  WITH(NOLOCK)  "
        sSQL &= " WHERE 			PropertyProductID <>'' AND AssetID = '' "
        Dim dtTable As DataTable = ReturnDataTable(sSQL)
        If _FormState = EnumFormState.FormAdd Then
            LoadDataSource(tdbcPropertyProductID, dtTable, gbUnicode)
        Else 'Trường hợp xem sửa thì load đúng 1 dòng thôi rồi disable luôn
            LoadDataSource(tdbcPropertyProductID, ReturnTableFilter(dtTable, "PropertyProductID = " & SQLString(sPropertyProductID), True), gbUnicode)
        End If

    End Sub

    'Private Function LoadObjectID(ByVal ID As String) As String
    '    Dim sSQL As String = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) " & vbCrLf
    '    sSQL &= "Where ObjectTypeID=" & SQLString(ID) & vbCrLf
    '    sSQL &= " Order By ObjectID "
    '    Return sSQL
    'End Function

    'Khanh bỏ điều kiện where đi
    Private Function LoadObjectID() As String
        Dim sSQL As String = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) " & vbCrLf
        sSQL &= " Order By ObjectID "
        Return sSQL
    End Function

    Private Sub LoadTDBDropDown()
        Dim sSQL As String = ""
        'Load tdbdSourceID 
        sSQL = "SELECT 		SourceID, SourceName" & UnicodeJoin(gbUnicode) & " as SourceName " & vbCrLf
        sSQL &= "FROM 		D02T0013 WITH(NOLOCK)  " & vbCrLf
        sSQL &= "WHERE 		Disabled = 0 " & vbCrLf
        sSQL &= "ORDER BY 	SourceID" & vbCrLf
        LoadDataSource(tdbdSourceID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadtdbdAssignmentID()
        Dim sSQL As String = ""
        Dim sUnicode As String = UnicodeJoin(gbUnicode)
        'Load tdbdAssignmentID
        sSQL = "Select '+' As AssignmentID, N'<Thêm Mới>'  As AssignmentName,'' As DebitAccountID, '' As Extend " & vbCrLf
        sSQL &= "UNION" & vbCrLf
        sSQL &= "SELECT 	AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " as AssignmentName, DebitAccountID, Extend " & vbCrLf
        sSQL &= "FROM 		D02T0002 WITH(NOLOCK)  " & vbCrLf
        sSQL &= "WHERE 		Disabled = 0 " & vbCrLf
        If L3Int(ReturnValueC1Combo(tdbcAssetID, "AssignmentTypeID")) = 2 Then
            sSQL &= "AND (Extend = 1 OR Extend = 2)" & vbCrLf
        Else
            sSQL &= "AND Extend = 0" & vbCrLf
        End If
        sSQL &= "ORDER BY 	AssignmentID" & vbCrLf
        LoadDataSource(tdbdAssignmentID, sSQL, gbUnicode)



    End Sub

#Region "Events of tdbc"

#Region "Events tdbcAssetID with txtAssetName"

    Private Sub tdbcAssetID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetID.KeyDown
        If e.KeyCode <> Keys.F2 Then Exit Sub
        Try
            'If clsFilterCombo.IsNewFilter Then Exit Sub ' TH filter dạng mới thì F2 gọi D99F5555 đã có sẵn
            Dim arrPro() As StructureProperties = Nothing
            SetProperties(arrPro, "InListID", "25")
            SetProperties(arrPro, "InWhere", "( AssetName Like '%%' ) And IsCompleted = 0  And DivisionID = " & SQLString(gsDivisionID))
            SetProperties(arrPro, "WhereValue", tdbcAssetID.Text)
            Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6010", arrPro)
            Dim sKey As String = GetProperties(frm, "Output01").ToString
            If sKey <> "" Then
                'Load dữ liệu
                tdbcAssetID.SelectedValue = sKey
            End If
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try
    End Sub

    Dim bLoadGrid1 As Boolean = True

    Private Sub tdbcAssetID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetID.SelectedValueChanged
        If tdbcAssetID.SelectedValue Is Nothing Then txtAssetName.Text = "" : Exit Sub
        LoadTDBGrid2(True)
        c1dateAssetDate.Value = Now.Date
        If (Not IsDBNull(tdbcAssetID.Columns("AssetDate").Text)) And tdbcAssetID.Columns("AssetDate").Text <> "" Then c1dateAssetDate.Value = tdbcAssetID.Columns("AssetDate").Text
        txtAssetName.Text = tdbcAssetID.Columns("AssetName").Value.ToString
        bLoadGrid1 = False
        tdbcAssetAccountID.SelectedValue = tdbcAssetID.Columns("AssetAccountID").Value
        tdbcDepAccountID.SelectedValue = tdbcAssetID.Columns("DepAccountID").Value
        bLoadGrid1 = True
        If _FormState = EnumFormState.FormAdd Then
            tdbcObjectTypeID.SelectedValue = tdbcAssetID.Columns("ObjectTypeID").Value
            'LoadDataSource(tdbcObjectID, LoadObjectID(ReturnValueC1Combo(tdbcObjectTypeID, "").ToString), gbUnicode)
            clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
            tdbcObjectID.SelectedValue = tdbcAssetID.Columns("ObjectID").Value.ToString
        End If
        tdbcEmployeeID.SelectedValue = tdbcAssetID.Columns("EmployeeID").Value.ToString
        txtPercentage.Text = SQLNumber(tdbcAssetID.Columns("Percentage").Value.ToString, DxxFormat.D08_RatioDecimals)
        txtPercentage.Tag = txtPercentage.Text
        txtServiceLife.Text = SQLNumber(tdbcAssetID.Columns("ServiceLife").Value.ToString, DxxFormat.DefaultNumber0)
        txtServiceLife.Tag = txtServiceLife.Text
        txtMethodID.Text = tdbcAssetID.Columns("MethodID").Text.ToString
        txtMethodEndID.Text = tdbcAssetID.Columns("MethodEndID").Value.ToString
        txtDeprTableName.Text = tdbcAssetID.Columns("DeprTableName").Value.ToString

        c1dateBeginDep.Value = tdbcAssetID.Columns("BeginDep").Value.ToString
        c1dateBeginUsing.Value = tdbcAssetID.Columns("BeginUsing").Value.ToString
        txtDepreciatedPeriod.Text = tdbcAssetID.Columns("DepreciatedPeriod").Value.ToString

        tdbcManagementObjTypeID.SelectedValue = tdbcAssetID.Columns("ManagementObjTypeID").Value
        tdbcManagementObjID.SelectedValue = tdbcAssetID.Columns("ManagementObjID").Value

        c1dateDepDate.Value = tdbcAssetID.Columns("DepDate").Text '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định

        Select Case L3Int(ReturnValueC1Combo(tdbcAssetID, "AssignmentTypeID"))
            Case 0
                tdbg3.Splits(0).DisplayColumns(COL3_AssignmentID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
                tdbg3.Splits(0).DisplayColumns(COL3_AssignmentName).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
            Case Else
                tdbg3.Splits(0).DisplayColumns(COL3_AssignmentID).Style.ResetBackColor()
                tdbg3.Splits(0).DisplayColumns(COL3_AssignmentName).Style.ResetBackColor()
        End Select

        LoadTDBGrid1()
        LoadtdbdAssignmentID()
        '***********

    End Sub

    '23/1/2018, Phạm Thị Thu: id 105627-Mặc định tài khoản TSCD không đươc phép sửa khi chọn mã TSCD đi hình thành
    Private Sub LockControlByAsset()
        Dim dt As DataTable = ReturnDataTable(SQLStoreD02P0022)
        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                DisabledControl(grpAssetID, tdbcAssetID.Name, dt.Rows(i).Item("Control").ToString, L3Bool(dt.Rows(i).Item("Value").ToString))
            Next
        End If
    End Sub

    '23/1/2018, Phạm Thị Thu: id 105627-Mặc định tài khoản TSCD không đươc phép sửa khi chọn mã TSCD đi hình thành
    Private Sub DisabledControl(ByVal ctrlParent As Control, ByVal sAssetControl As String, sField As String, bValue As Boolean)
        If ctrlParent.HasChildren Then
            Dim iCount As Integer = ctrlParent.Controls.Count
            For i As Integer = 0 To iCount - 1
                Dim ctrlChild As Control = ctrlParent.Controls(i)
                Dim ctrlName As String = ""
                If ctrlChild.Name.StartsWith("tdbc") Then
                    ctrlName = "tdbc" & sField
                    If sAssetControl <> ctrlChild.Name AndAlso ctrlChild.Name = ctrlName AndAlso bValue Then
                        ReadOnlyControl(ctrlChild)
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0022
    '# Created User: NGOCTHOAI
    '# Created Date: 23/01/2018 09:29:35
    '23/1/2018, Phạm Thị Thu: id 105627-Mặc định tài khoản TSCD không đươc phép sửa khi chọn mã TSCD đi hình thành
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0022() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Khoa control theo Ma tai san " & vbCrlf)
        sSQL &= "Exec D02P0022 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[25], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[2], NOT NULL
        sSQL &= SQLString(Me.Name) 'FormID, varchar[20], NOT NULL
        Return sSQL
    End Function



    'Private Sub tdbcAssetID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetID.LostFocus
    '    'If tdbcAssetID.Text = "" Then Exit Sub
    '    'If tdbcAssetID.FindStringExact(tdbcAssetID.Text) = -1 Then
    '    '    tdbcAssetID.SelectedValue = ""
    '    '    tdbcAssetID.Text = ""
    '    '    D99C0008.MsgL3(rL3("Ban_phai_chon_ma_trong_danh_sach"))
    '    'End If
    '    ShowD02F0087()
    'End Sub

    Private Sub tdbcAssetID_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetID.Validated
        clsFilterCombo.FilterCombo(tdbcAssetID, e)
        If tdbcAssetID.Text = "" Then Exit Sub
        If tdbcAssetID.FindStringExact(tdbcAssetID.Text) = -1 OrElse tdbcAssetID.SelectedValue Is Nothing Then
            tdbcAssetID.SelectedValue = ""
            tdbcAssetID.Text = ""
            D99C0008.MsgL3(rL3("Ban_phai_chon_ma_trong_danh_sach"))
        End If
        ShowD02F0087() 'ID : 224617 - BỔ SUNG Cho phép gọi màn hình THIẾT LẬP DANH MỤC TÀI SẢN CỐ ĐỊNH tại bước hình thành TS
    End Sub
#End Region

#Region "Events tdbcObjectTypeID load tdbcObjectID with txtObjectName"

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.SelectedValueChanged
        'If tdbcObjectTypeID.SelectedValue Is Nothing OrElse tdbcObjectTypeID.Text = "" Then
        '    LoadDataSource(tdbcObjectID, LoadObjectID(""), gbUnicode)
        '    Exit Sub
        'End If
        'LoadDataSource(tdbcObjectID, LoadObjectID(tdbcObjectTypeID.SelectedValue.ToString()), gbUnicode)
        clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
        tdbcObjectID.Text = ""
    End Sub

    Private Sub tdbcObjectTypeID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.LostFocus
        'If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 OrElse tdbcObjectTypeID.SelectedValue Is Nothing Then
        '    tdbcObjectTypeID.Text = ""
        '    LoadDataSource(tdbcObjectID, LoadObjectID(""), gbUnicode)
        '    tdbcObjectID.Text = ""
        '    Exit Sub
        'End If
    End Sub

    Private Sub tdbcObjectTypeID_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.Validated
        clsFilterCombo.FilterCombo(tdbcObjectTypeID, e)
        If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 OrElse tdbcObjectTypeID.SelectedValue Is Nothing Then
            tdbcObjectTypeID.Text = ""
            'LoadDataSource(tdbcObjectID, LoadObjectID(""), gbUnicode)
            tdbcObjectID.Text = ""
            Exit Sub
        End If
    End Sub

    'Private Sub tdbcObjectID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcObjectID.KeyDown
    '    If e.Control And e.KeyCode = Keys.R Then LoadDataSource(tdbcObjectID, LoadObjectID(ReturnValueC1Combo(tdbcObjectTypeID).ToString), gbUnicode)
    'End Sub

    Private Sub tdbcObjectID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID.SelectedValueChanged
        If tdbcObjectID.SelectedValue Is Nothing Then
            txtObjectName.Text = ""
        Else
            txtObjectName.Text = tdbcObjectID.Columns(2).Value.ToString
        End If
    End Sub

    Private Sub tdbcObjectID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID.LostFocus
        'If tdbcObjectID.FindStringExact(tdbcObjectID.Text) = -1 Then tdbcObjectID.Text = ""
    End Sub

    Private Sub tdbcObjectID_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcObjectID.Validated
        clsFilterCombo.FilterCombo(tdbcObjectID, e)
        If tdbcObjectID.FindStringExact(tdbcObjectID.Text) = -1 Then tdbcObjectID.Text = ""
    End Sub

#End Region

#Region "Event tdbcManagementTypeOBjID"
    Private Sub tdbcManagementObjTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObjTypeID.SelectedValueChanged
        clsFilterCombo.LoadtdbcObjectID(tdbcManagementObjID, dtManagementID, ReturnValueC1Combo(tdbcManagementObjTypeID))
        tdbcManagementObjID.Text = ""
    End Sub
    Private Sub tdbcManagementObjTypeID_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcManagementObjTypeID.Validated
        clsFilterCombo.FilterCombo(tdbcManagementObjTypeID, e)
        If tdbcManagementObjTypeID.FindStringExact(tdbcManagementObjTypeID.Text) = -1 OrElse tdbcManagementObjTypeID.SelectedValue Is Nothing Then
            tdbcManagementObjTypeID.Text = ""
            'LoadDataSource(tdbcObjectID, LoadObjectID(""), gbUnicode)
            tdbcManagementObjID.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub tdbcManagementObjID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObjID.SelectedValueChanged
        If tdbcManagementObjID.SelectedValue Is Nothing Then
            txtManagementObjName.Text = ""
        Else
            txtManagementObjName.Text = tdbcManagementObjID.Columns(2).Value.ToString
        End If
    End Sub
#End Region

#Region "Events tdbcEmployeeID with txtEmployeeName"

    Private Sub tdbcEmployeeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcEmployeeID.SelectedValueChanged
        If tdbcEmployeeID.SelectedValue Is Nothing Then
            txtEmployeeName.Text = ""
        Else
            txtEmployeeName.Text = tdbcEmployeeID.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcEmployeeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcEmployeeID.LostFocus
        'If tdbcEmployeeID.FindStringExact(tdbcEmployeeID.Text) = -1 Then
        '    tdbcEmployeeID.Text = ""
        'End If
    End Sub

    Private Sub tdbcEmployeeID_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcEmployeeID.Validated
        clsFilterCombo.FilterCombo(tdbcEmployeeID, e)
        If tdbcEmployeeID.FindStringExact(tdbcEmployeeID.Text) = -1 Then
            tdbcEmployeeID.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcAssetAccount và tdbcDepAccountID"

    Private Sub tdbcAssetAccount_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetAccountID.LostFocus
        If tdbcAssetAccountID.FindStringExact(tdbcAssetAccountID.Text) = -1 Then tdbcAssetAccountID.Text = ""
    End Sub

    Private Sub tdbcAssetAccount_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetAccountID.SelectedValueChanged, tdbcDepAccountID.SelectedValueChanged
        If Not bLoadGrid1 Then Exit Sub
        LoadTDBGrid2(True)
        LoadTDBGrid1()
    End Sub

    Private Sub tdbcDepAccountID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDepAccountID.LostFocus
        If tdbcDepAccountID.FindStringExact(tdbcDepAccountID.Text) = -1 Then tdbcDepAccountID.Text = ""
    End Sub
#End Region
#End Region

    Dim sPropertyProductID As String = ""
    Private Sub LoadMasterAssetID()
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0500(_assetID, _setupFrom)
        Dim dtMaster As DataTable = ReturnDataTable(sSQL)
        If dtMaster.Rows.Count > 0 Then
            With dtMaster.Rows(0)
                sPropertyProductID = .Item("D27PropertyProductID").ToString
                tdbcPropertyProductID.Text = sPropertyProductID

                'LoadTDBComboPropertyProductID(sPropertyProductID)
                tdbcPropertyProductID.SelectedValue = sPropertyProductID
                tdbcAssetID.SelectedValue = _assetID
                tdbcObjectTypeID.SelectedValue = .Item("ObjectTypeID")
                tdbcObjectTypeID.Tag = tdbcObjectTypeID.SelectedValue
                tdbcObjectTypeID_LostFocus(Nothing, Nothing)
                tdbcObjectID.SelectedValue = .Item("ObjectID")
                tdbcEmployeeID.SelectedValue = .Item("EmployeeID").ToString
                '************
                txtConvertedAmount.Text = SQLNumber(.Item("ConvertedAmount").ToString, DxxFormat.D90_ConvertedDecimals)
                txtAmountDepreciation.Text = SQLNumber(.Item("AmountDepreciation").ToString, DxxFormat.D90_ConvertedDecimals)
                txtDepreciateAmount.Text = SQLNumber(.Item("DepreciatedAmount").ToString, DxxFormat.D90_ConvertedDecimals) ' 17/12/2013 id 62172 
                txtRemainAmount.Text = SQLNumber(.Item("RemainAmount").ToString, DxxFormat.D90_ConvertedDecimals)
                'txtMethodID.Text = tdbcAssetID.Columns("MethodID").Text
                If L3Int(.Item("MethodID")) = 1 Then
                    lblMethodEndID.Text = rL3("Bang_khau_hao")
                Else
                    lblMethodEndID.Text = rL3("KH_ky_cuoi")
                End If
                'txtMethodEndID.Text = .Item("MethodEndID").ToString
                'txtDeprTableName.Text = tdbcAssetID.Columns("DeprTableName").Value.ToString
                txtServiceLife.Text = SQLNumber(.Item("ServiceLife").ToString, DxxFormat.DefaultNumber0)
                txtServiceLife.Tag = txtServiceLife.Text
                txtDepreciatedPeriod.Text = SQLNumber(.Item("DepreciatedPeriod").ToString, DxxFormat.DefaultNumber0)
                txtPercentage.Text = SQLNumber(.Item("Percentage").ToString, DxxFormat.D08_RatioDecimals)
                txtPercentage.Tag = txtPercentage.Text
                c1dateBeginUsing.Value = .Item("BeginUse").ToString
                c1dateBeginDep.Value = .Item("BeginDep").ToString
                c1dateUseDate.Value = .Item("UseDate").ToString
                c1dateAssetDate.Value = .Item("AssetDate").ToString
                c1dateDepDate.Value = .Item("DepDate").ToString '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
                sCreateUserID = .Item("CreateUserID").ToString
                sCreateDate = .Item("CreateDate").ToString

            End With
            ReadOnlyControl(tdbcPropertyProductID)
        End If
    End Sub

    Dim dtGrid1 As DataTable

    Private Sub LoadTDBGrid1(Optional ByVal bClear As Boolean = False)
        If bClear Then
            If dtGrid1 IsNot Nothing Then dtGrid1.Clear()
        Else
            Dim sUnicode As String = UnicodeJoin(gbUnicode)

            Dim sSQL As String = ""
            sSQL = "Select convert(bit,0) as Choose, TransactionID, DivisionID, D02T0012.ModuleID, AssetID, VoucherTypeID, VoucherNo, VoucherDate," & _
                    "TranMonth, TranYear, Case When IsNull(RTrim(LTrim(ItemName" & sUnicode & ")),'')='' Then Description" & sUnicode & " Else ItemName" & sUnicode & " End As Description," & _
                    "CurrencyID , ExchangeRate, DebitAccountID, CreditAccountID," & _
                    "OriginalAmount, ConvertedAmount, D02T0012.Status, RefNo, RefDate, SeriNo, D02T0012.SourceID, D02T0012.SourceID as OSourceID," & _
                    "D02T0012.CreateUserID, D02T0012.CreateDate, D02T0012.LastModifyUserID, D02T0012.LastModifyDate," & _
                    "D02T0012.ObjectTypeID, D02T0012.ObjectID, BatchID, O.ObjectName" & sUnicode & " as ObjectName, D02T0012.VATNo, D02T0012.VATGroupID, D02T0012.VATTypeID, D02T0012.Internal," & _
                    "Ana01ID, Ana02ID, Ana03ID, Ana04ID, Ana05ID,Ana06ID, Ana07ID, Ana08ID, Ana09ID , Ana10ID,Convert(bit,IsNotAllocate) As IsNotAllocate ,ProjectID, PropertyProductID,TaskID " & vbCrLf
            sSQL &= "From D02T0012 WITH (NOLOCK) " & vbCrLf
            sSQL &= " LEFT JOIN Object O WITH(NOLOCK) ON D02T0012.ObjectTypeID = O.ObjectTypeID	AND D02T0012.ObjectID=O.ObjectID" & vbCrLf
            sSQL &= "Where DivisionID=" & SQLString(gsDivisionID) & " And  TranMonth + TranYear * 100 <= " & giTranMonth & "+" & giTranYear & "*100" & _
             " And Isnull(AssetID,'') = ''  And D02T0012.Status = 0 AND isnull(D02T0012.TransactionTypeID, '') <> 'KH'" & _
             " AND ISNULL(CipID,'') = ''"
            If tdbcAssetAccountID.Text <> "" Or tdbcDepAccountID.Text <> "" Then
                sSQL &= "  And (DebitAccountID = " & SQLString(ReturnValueC1Combo(tdbcAssetAccountID)) & "  OR (CreditAccountID = " & SQLString(ReturnValueC1Combo(tdbcAssetAccountID)) & " AND DebitAccountID = " & SQLString(ReturnValueC1Combo(tdbcDepAccountID)) & ") OR CreditAccountID = " & SQLString(ReturnValueC1Combo(tdbcDepAccountID)) & ")"
            End If
            Dim sTransactions As String = GetTransactionsTDBG2()
            If sTransactions <> "" Then sSQL &= "  And isnull(TransactionID,'') not  in (" & sTransactions & ")"
            sSQL &= " AND ISNULL(CipID,'') = '' " & vbCrLf
            sSQL &= " AND PropertyProductID = " & SQLString(tdbcPropertyProductID.Text) & vbCrLf
            sSQL &= " ORDER BY VoucherDate"
            dtGrid1 = ReturnDataTable(sSQL)
        End If
        gbEnabledUseFind = dtGrid1.Rows.Count > 0
        LoadDataSource(tdbg, dtGrid1, gbUnicode)

        ReLoadTDBGrid1()
    End Sub

    Private Sub ResetGrid1()
        mnsFind.Enabled = tdbg.RowCount > 0 OrElse gbEnabledUseFind
        mnsListAll.Enabled = mnsFind.Enabled
        FooterTotalGrid(tdbg, COL_VoucherNo)
        FooterSumNew(tdbg, COL_OriginalAmount, COL_ConvertedAmount)
    End Sub


#Region "Events of Grid1"

    Dim bRefreshFilter As Boolean = False 'Cờ bật set FilterText =""
    Dim sFilter1 As New StringBuilder

    Private Sub tdbg_AfterColUpdate(sender As Object, e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.AfterColUpdate
        If e.ColIndex = COL_Choose Then
            SetValueObject() '9/9/2021, Lê Thị Diễm Vi:id 178252-BV_Mặc định bộ phận tiếp nhận theo đối tượng trên phiếu xuất kho
        End If
    End Sub

    Private Sub SetValueObject()
        '9/9/2021, Lê Thị Diễm Vi:id 178252-BV_Mặc định bộ phận tiếp nhận theo đối tượng trên phiếu xuất kho
        If ReturnValueC1Combo(tdbcAssetID, "ObjectTypeID") <> "" And ReturnValueC1Combo(tdbcAssetID, "ObjectID") <> "" Then Exit Sub

        Dim dr() As DataRow = dtGrid1.Select("Choose = True")
        If dr.Length > 0 Then
            tdbcObjectTypeID.SelectedValue = dr(0).Item("ObjectTypeID").ToString
            If tdbcObjectTypeID.Text <> "" Then
                clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
                tdbcObjectID.SelectedValue = dr(0).Item("ObjectID").ToString
            End If
        End If
    End Sub

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid1 Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            'Filter the data 
            FilterChangeGrid(tdbg, sFilter1) 'Nếu có Lọc khi In
            ReLoadTDBGrid1()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Dim bSelected As Boolean = False
    Dim bSelected2 As Boolean = False
    Private Sub HeadClick(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sField As String, ByRef bSelect As Boolean)
        Select Case sField
            Case "Choose"
                c1Grid.AllowSort = False
                L3HeadClick(c1Grid, sField, bSelect)
            Case Else
                c1Grid.AllowSort = True
        End Select
    End Sub

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.HeadClick
        HeadClick(tdbg, e.Column.DataColumn.DataField, bSelected)
        SetValueObject() '9/9/2021, Lê Thị Diễm Vi:id 178252-BV_Mặc định bộ phận tiếp nhận theo đối tượng trên phiếu xuất kho
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.Control And e.KeyCode = Keys.S Then
            HeadClick(tdbg, tdbg.Columns(tdbg.Col).DataField, bSelected)
            SetValueObject() '9/9/2021, Lê Thị Diễm Vi:id 178252-BV_Mặc định bộ phận tiếp nhận theo đối tượng trên phiếu xuất kho
            Exit Sub
        End If
        HotKeyCtrlVOnGrid(tdbg, e) 'Đã bổ sung D99X0000
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Col
            Case COL_Choose 'Chặn Ctrl + V trên cột Check
                e.Handled = CheckKeyPress(e.KeyChar)
                'Case COL_ConvertedAmount, COL_OriginalAmount
                '    e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub
#End Region

#Region "Events of Grid2"

    Dim bRefreshFilter2 As Boolean = False 'Cờ bật set FilterText =""
    Dim sFilter2 As New StringBuilder

    Private Sub tdbg2_BeforeColEdit(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColEditEventArgs) Handles tdbg2.BeforeColEdit
        Select Case e.ColIndex
            Case IndexOfColumn(tdbg2, "Str01") To IndexOfColumn(tdbg2, "Str05")
                e.Cancel = tdbg2.Columns(e.ColIndex).DataWidth = 0
        End Select
    End Sub

    Private Sub tdbg2_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.AfterColUpdate
        Select Case e.ColIndex
            Case COL2_OriginalAmount
                FooterSumNew(tdbg2, COL2_OriginalAmount, COL2_ConvertedAmount)
            Case COL2_IsNotAllocate
                'Sửa tính ConvertedAmount theo Incident 	74480
                tdbg2.UpdateData()
                If dtGrid2 IsNot Nothing Then
                    Dim sConvertedAmount As Double = Number(txtConvertedAmount.Text) - Number(dtGrid2.Compute("SUM([ConvertedAmount])", "IsNotAllocate = 1 And DebitAccountID = " & SQLString(ReturnValueC1Combo(tdbcAssetAccountID).ToString)))
                    Dim dServiceLife As Double = Number(IIf(Number(txtServiceLife.Text) = 0, 1, Number(txtServiceLife.Text)))
                    txtDepreciateAmount.Text = SQLNumber(sConvertedAmount / dServiceLife, DxxFormat.D90_ConvertedDecimals)
                Else
                    txtDepreciateAmount.Text = SQLNumber(0, DxxFormat.D90_ConvertedDecimals)
                End If

        End Select
        'LockServiceLife()
    End Sub

    Private Sub LockServiceLife()
        '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        If D02Systems.IsCalPeriodByRate = False Then
            txtServiceLife.BackColor = COLOR_BACKCOLOROBLIGATORY
            ReadOnlyControl(txtPercentage)
        Else
            ReadOnlyControl(txtServiceLife)
            txtPercentage.BackColor = COLOR_BACKCOLOROBLIGATORY
        End If
    End Sub

    Private Sub tdbg2_AfterDelete(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg2.AfterDelete
        ResetGrid2()
        dtGrid2.AcceptChanges() 'Xóa các dòng đã nhấn Delete
    End Sub

    Private Sub tdbg2_BeforeColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg2.BeforeColUpdate
        Select Case e.ColIndex
            Case COL2_SourceID
                If tdbg2.Columns(e.ColIndex).Text <> tdbg2.Columns(e.ColIndex).DropDown.Columns(0).Text Then
                    tdbg2.Columns(e.ColIndex).Text = ""
                End If
        End Select
    End Sub

    Private Sub tdbg2_BeforeDelete(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.CancelEventArgs) Handles tdbg2.BeforeDelete
        dtGrid2.AcceptChanges() 'Xóa các dòng đã nhấn Delete
        Dim dr As DataRow = dtGrid2.Rows(tdbg2.Row)
        'Gán lại giá trị lúc load database
        dr.Item(COL2_Choose) = 0
        'dr.Item(COL2_SourceID) = dr.Item("O" & COL2_SourceID)
        dr.Item(COL2_SourceID) = dr.Item(COL2_SourceID) ' ID : 263039

        dtGrid1.ImportRow(dr)
        dtGrid1.DefaultView.Sort = "VoucherDate"
    End Sub

    Private Sub tdbg2_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.ComboSelect
        tdbg2.UpdateData()
    End Sub

    Private Sub tdbg2_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg2.FilterChange
        Try
            If (dtGrid2 Is Nothing) Then Exit Sub
            If bRefreshFilter2 Then Exit Sub 'set FilterText ="" thì thoát
            'Filter the data 
            FilterChangeGrid(tdbg2, sFilter2) 'Nếu có Lọc khi In
            ReLoadTDBGrid2()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub HeadClick2(ByVal index As Integer)
        Select Case index
            Case COL2_SourceID, COL2_Str01, COL2_Date01
                CopyColumns(tdbg2, index, tdbg2.Columns(index).Text, tdbg2.Row)
            Case COL2_Choose
                HeadClick(tdbg2, "Choose", bSelected2)
        End Select
    End Sub

    Private Sub tdbg2_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.HeadClick
        HeadClick2(e.ColIndex)
    End Sub

    Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
        If e.Control And e.KeyCode = Keys.S Then
            HeadClick2(tdbg2.Col)
            Exit Sub
        End If
        HotKeyCtrlVOnGrid(tdbg2, e) 'Đã bổ sung D99X0000
    End Sub

    Private Sub tdbg2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg2.KeyPress
        Select Case tdbg2.Col
            Case COL2_Choose 'Chặn Ctrl + V trên cột Check
                e.Handled = CheckKeyPress(e.KeyChar)
                'Case COL2_ConvertedAmount, COL2_OriginalAmount
                '    e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
                'Case IndexOfColumn(tdbg2, COL_Str01) To IndexOfColumn(tdbg2, COL_Str05)
                '    e.Handled = tdbg2.Columns(tdbg2.Col).DataWidth = 0
        End Select
    End Sub

    Private Sub ReLoadTDBGrid2()
        dtGrid2.DefaultView.RowFilter = sFilter2.ToString
        ResetGrid2()
    End Sub
#End Region

    Private Function GetTransactionsTDBG2() As String
        'Dim sTransactions As String = ""
        'For i As Integer = 0 To tdbg2.RowCount - 1
        '    If sTransactions <> "" Then sTransactions &= ","
        '    sTransactions &= SQLString(sTransactions)
        'Next
        'Return sTransactions
        'ID : 251209 : Fix lỗi điều kiện getTransactions bị lỗi dẫn đến tràn bộ nhớ. Double giá trị ('','''','''','''','','','''','''',''','')
        Dim sTransactions As String = ""
        Dim sTransactionID As String = ""
        For i As Integer = 0 To tdbg2.RowCount - 1
            sTransactionID = tdbg2(i, COL2_TransactionID).ToString
            If sTransactionID <> "" Then
                If i <> 0 Then
                    sTransactions &= ","
                    sTransactions &= SQLString(sTransactionID)
                End If
                sTransactions &= SQLString(sTransactionID)
            End If
            SQLString(sTransactions)
        Next

        Return sTransactions
    End Function

    Dim dtGrid2 As DataTable
    Private Sub LoadTDBGrid2(Optional ByVal bClear As Boolean = False)
        If bClear Then
            If dtGrid2 IsNot Nothing Then dtGrid2.Clear()
        Else
            Dim sUnicode As String = UnicodeJoin(gbUnicode)
            Dim sSQL As String = "Select CONVERT(bit, case when TransMode='CP' then 1 else 0 end) as Choose, TransactionID, DivisionID, ModuleID, SplitNo,AssetID,VoucherTypeID,VoucherNo,VoucherDate,TranMonth,TranYear," & _
                    "TransactionDate,Description" & sUnicode & " as Description,CurrencyID,ExchangeRate,DebitAccountID,CreditAccountID,OriginalAmount,ConvertedAmount," & _
                    "Status,TransactionTypeID,RefNo,RefDate,Disabled,CreateUserID,CreateDate,LastModifyUserID,LastModifyDate," & _
                    "SeriNo,ObjectTypeID,ObjectID,BatchID,VATObjectTypeID,VATObjectID,ObjectName" & sUnicode & " as ObjectName,VATNo,VATGroupID,VATTypeID," & _
                    "Ana01ID,Ana02ID,Ana03ID,Ana04ID,Ana05ID,Ana06ID,Ana07ID,Ana08ID,Ana09ID,Ana10ID," & _
                    "CipID,Notes" & sUnicode & " as Notes,AssignmentID,NormID,Posted,SourceID, SourceID as OSourceID, DebitObjectTypeID,DebitObjectID,CreditObjectTypeID,CreditObjectID," & _
                    "SplitBatchID,Internal,DeprTableID,Str01" & sUnicode & " as Str01,Str02" & sUnicode & " as Str02,Str03" & sUnicode & " as Str03,Str04" & sUnicode & " as Str04,Str05" & sUnicode & " as Str05,Num01,Num02,Num03,Num04,Num05,Date01,Date02,Date03," & _
                    "Date04,Date05,PeriodID,ItemName" & sUnicode & " as ItemName,GroupID,TransMode,Cancel,Convert(bit,IsNotAllocate) As IsNotAllocate,ProjectID, PropertyProductID " & vbCrLf
            sSQL &= " From D02T0012 WITH(NOLOCK) " & vbCrLf
            sSQL &= "Where TransactionTypeID in('MM', 'SDMM') And AssetID= " & SQLString(tdbcAssetID.Text) & " Order by TransactionID"
            dtGrid2 = ReturnDataTable(sSQL)
        End If
        LoadDataSource(tdbg2, dtGrid2, gbUnicode)
        ResetGrid2()
    End Sub

    Private Sub ResetGrid2()
        FooterTotalGrid(tdbg2, COL2_VoucherNo)
        FooterSumNew(tdbg2, COL2_OriginalAmount, COL2_ConvertedAmount)
        Dim dDepreciatedAmount As Double = 0
        Dim dConvertAmount As Double = CallTotalAmount(dDepreciatedAmount)
        txtConvertedAmount.Text = Format(dConvertAmount, DxxFormat.D90_ConvertedDecimals)
        txtAmountDepreciation.Text = Format(dDepreciatedAmount, DxxFormat.D90_ConvertedDecimals)
        txtServiceLife_TextChanged(Nothing, Nothing)
    End Sub

    Dim dtGrid3 As DataTable = Nothing

    Private Sub LoadTDBGrid3(Optional ByVal bClear As Boolean = False)
        If bClear Then
            If dtGrid3 IsNot Nothing Then dtGrid3.Clear()
        Else
            'Incident 79529 bỏ  BeginYear and BeginMonth
            Dim strSQL As String = ""
            strSQL = " Select D02T5000.HistoryID, D02T5000.AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " as AssignmentName, DebitAccountID, PercentAmount/100 AS PercentAmount, D02T0002.Extend"
            strSQL &= " From D02T5000 WITH(NOLOCK) Inner join D02T0002 WITH(NOLOCK) On D02T5000.AssignmentID = D02T0002.AssignmentID"
            strSQL &= " Inner join D02T0001 WITH(NOLOCK) ON D02T5000.AssetID = D02T0001.AssetID " & vbCrLf
            strSQL &= " Inner join D02V0041 WITH(NOLOCK) On D02T0001.AssignmentTypeID = D02V0041.AssignmentTypeID  " & vbCrLf
            strSQL &= " Where D02T5000.HistoryTypeID = 'AS' "
            strSQL &= " And D02T5000.DivisionID = " & SQLString(gsDivisionID)
            strSQL &= " And BatchID = " & SQLString(tdbcAssetID.Text) & " And D02T5000.AssetID = " & SQLString(tdbcAssetID.Text) & vbCrLf
            strSQL &= " Order by HistoryID"
            dtGrid3 = ReturnDataTable(strSQL)
        End If
        LoadDataSource(tdbg3, dtGrid3, gbUnicode)
        FooterTotalGrid(tdbg3, COL3_AssignmentID)
    End Sub

    Private Sub LoadAddNew()
        btnNext.Enabled = False
        btnSave.Enabled = True
        c1dateBeginUsing.Value = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        c1dateBeginDep.Value = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        c1dateUseDate.Value = Date.Today
        c1dateAssetDate.Value = Date.Today
        c1dateDepDate.Value = Date.Today '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        LoadTDBComboPropertyProductID()
    End Sub

    Private Sub LoadEdit()
        btnNext.Visible = False
        btnSave.Left = btnNext.Left
        ReadOnlyControl(tdbcAssetID, tdbcAssetAccountID, tdbcDepAccountID)
        LoadMasterAssetID()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0015
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 22/12/2011 08:16:41
    '# Modified User: 
    '# Modified Date: 
    '# Description: Load caption Thông tin phụ
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0015() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0015 "
        sSQL &= SQLString("D02T0012") & COMMA 'TableName, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function



    'Private Function LoadCaptionSubInfo() As Boolean
    '    Dim bUseSubInfo As Boolean = False
    '    Dim dtCaption As DataTable = ReturnDataTable(SQLStoreD02P0015)
    '    If dtCaption.Rows.Count = 0 Then Return False
    '    Dim arr() As FormatColumn = Nothing
    '    For i As Integer = 0 To dtCaption.Rows.Count - 1
    '        Dim sField As String = dtCaption.Rows(i).Item("DataID").ToString.Replace("Ana", "")

    '        If tdbg1.Columns.IndexOf(tdbg1.Columns(sField)) = -1 Then Continue For
    '        tdbg1.Columns(sField).Caption = dtCaption.Rows(i).Item("Data" & gsLanguage).ToString
    '        tdbg1.Splits(SPLIT1).DisplayColumns(sField).HeadingStyle.Font = FontUnicode(gbUnicode)
    '        tdbg1.Splits(SPLIT1).DisplayColumns(sField).Visible = CBool(dtCaption.Rows(i).Item("Disabled"))
    '        If tdbg1.Splits(SPLIT1).DisplayColumns(sField).Visible Then
    '            bUseSubInfo = True
    '            Select Case L3Int(dtCaption.Rows(i).Item("DataType"))
    '                Case 0 'Số
    '                    AddNumberColumns(arr, SqlDbType.Money, tdbg1.Columns(sField).DataField, "N" & L3Int(dtCaption.Rows(i).Item("DecimalNum")))
    '                Case 1 'Chuỗi
    '                    tdbg1.Columns(sField).DataWidth = L3Int(dtCaption.Rows(i).Item("DecimalNum"))
    '                Case 2 'Ngày
    '            End Select
    '        End If
    '    Next
    '    InputNumber(tdbg1, arr)
    '    '''''''''''''''''''''''''''''''''''''''''''''
    '    Return bUseSubInfo
    'End Function

    Private Function LoadCaptionSubInfo() As Boolean
        Dim bUseSubInfo As Boolean = False
        Dim dtCaption As DataTable = ReturnDataTable(SQLStoreD02P0015)
        If dtCaption.Rows.Count = 0 Then Return False
        Dim arr() As FormatColumn = Nothing
        For i As Integer = 0 To dtCaption.Rows.Count - 1
            Dim sField As String = dtCaption.Rows(i).Item("DataID").ToString.Replace("Ana", "")

            If tdbg.Columns.IndexOf(tdbg.Columns(sField)) = -1 Then Continue For
            ' tdbg.Columns(sField).Caption = dtCaption.Rows(i).Item("Description").ToString
            tdbg.Columns(sField).Caption = dtCaption.Rows(i).Item("Data" & gsLanguage).ToString
            'tdbg.Splits(SPLIT1).DisplayColumns(sField).HeadingStyle.Font = FontUnicode(gbUnicode)
            'tdbg.Splits(SPLIT1).DisplayColumns(sField).Visible = Not CBool(dtCaption.Rows(i).Item("Disabled"))


            If tdbg2.Columns.IndexOf(tdbg2.Columns(sField)) = -1 Then Continue For
            tdbg2.Columns(sField).Caption = dtCaption.Rows(i).Item("Data" & gsLanguage).ToString
            tdbg2.Splits(tdbg2.Splits.Count - 2).DisplayColumns(sField).HeadingStyle.Font = FontUnicode(gbUnicode)
            tdbg2.Splits(tdbg2.Splits.Count - 2).DisplayColumns(sField).Visible = CBool(dtCaption.Rows(i).Item("Disabled"))
            If tdbg2.Splits(tdbg2.Splits.Count - 2).DisplayColumns(sField).Visible Then
                bUseSubInfo = True
                Select Case L3Int(dtCaption.Rows(i).Item("DataType"))
                    Case 0 'Số
                        AddNumberColumns(arr, SqlDbType.Money, tdbg2.Columns(sField).DataField, "N" & L3Int(dtCaption.Rows(i).Item("DecimalNum")))
                    Case 1 'Chuỗi
                        tdbg2.Columns(sField).DataWidth = L3Int(dtCaption.Rows(i).Item("DecimalNum"))
                    Case 2 'Ngày
                End Select
            End If
        Next
        If arr IsNot Nothing Then InputNumber(tdbg2, arr)
        '''''''''''''''''''''''''''''''''''''''''''''
        Return bUseSubInfo
    End Function

#Region "Events of lưới phân bổ"
    Private Sub tdbg3_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg3.ComboSelect
        tdbg3.UpdateData()
    End Sub

    Private Sub tdbg3_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg3.AfterColUpdate
        Select Case e.Column.DataColumn.DataField
            'Case COL3_AssignmentID
            '    If tdbg3.Columns(COL3_AssignmentID).Text = "" Then Exit Select
            '    tdbg3.Columns(COL3_AssignmentName).Text = tdbdAssignmentID.Columns("AssignmentName").Text
            '    tdbg3.Columns(COL3_DebitAccountID).Text = tdbdAssignmentID.Columns("DebitAccountID").Text
            '    tdbg3.Columns(COL3_Extend).Text = tdbdAssignmentID.Columns("Extend").Text
            '    FooterTotalGrid(tdbg3, COL3_AssignmentID)

            Case COL3_AssignmentID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg3, e.Column.DataColumn.DataField)
                If tdbd Is Nothing Then Exit Select
                If clsFilterDropdown.IsNewFilter Then
                    'Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg3, e, tdbd)
                    'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                    If tdbg3.Columns(e.ColIndex).Text = "+" Then
                        tdbg3.Columns(e.ColIndex).Text = ""
                        ShowD02F0101(COL3_AssignmentID)
                    End If
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdownMulti(tdbg3, e, tdbd)
                    AfterColUpdate(e.ColIndex, dr)
                    'Exit Sub
                Else ' Nhập liệu dạng cũ (xổ dropdown)
                    'Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg3.Columns(e.ColIndex).Text))
                    'AfterColUpdate(e.ColIndex, row)
                    'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                    If tdbg3.Columns(e.ColIndex).Text = "+" Then
                        tdbg3.Columns(e.ColIndex).Text = ""
                        ShowD02F0101(COL3_AssignmentID)
                    End If
                    Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg3.Columns(e.ColIndex).Text))
                    AfterColUpdate(e.ColIndex, row)


                End If
        End Select
    End Sub

    Private Sub AfterColUpdate(ByVal iCol As Integer, ByVal dr() As DataRow)
        Dim iRow As Integer = tdbg3.Row
        If dr Is Nothing OrElse dr.Length = 0 Then
            Dim row As DataRow = Nothing
            AfterColUpdate(iCol, row)
        ElseIf dr.Length = 1 Then
            If tdbg3.Bookmark <> tdbg3.Row AndAlso tdbg3.RowCount = tdbg3.Row Then 'Đang đứng dòng mới
                Dim dr1 As DataRow = dtGrid3.NewRow
                dtGrid3.Rows.InsertAt(dr1, tdbg3.Row)
                SetDefaultValues(tdbg3, dr1) 'Bổ sung set giá trị mặc định 19/08/2015
                tdbg3.Bookmark = tdbg3.Row
            End If
            AfterColUpdate(iCol, dr(0))
        Else
            For Each row As DataRow In dr
                tdbg3.Bookmark = iRow
                tdbg3.Row = iRow
                AfterColUpdate(iCol, row)
                tdbg3.UpdateData()
                iRow += 1
            Next
            tdbg3.Focus()
        End If
    End Sub

    Private Sub AfterColUpdate(ByVal iCol As Integer, ByVal dr As DataRow)
        'Gán lại các giá trị phụ thuộc vào Dropdown
        Select Case iCol
            Case IndexOfColumn(tdbg3, COL3_AssignmentID)
                If dr Is Nothing OrElse dr.Item("AssignmentID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    tdbg3.Columns(COL3_AssignmentID).Text = ""
                    tdbg3.Columns(COL3_AssignmentName).Text = ""
                    tdbg3.Columns(COL3_DebitAccountID).Text = ""
                    tdbg3.Columns(COL3_Extend).Text = ""
                    Exit Sub
                End If
                tdbg3.Columns(COL3_AssignmentID).Text = dr.Item("AssignmentID").ToString
                tdbg3.Columns(COL3_AssignmentName).Text = dr.Item("AssignmentName").ToString
                tdbg3.Columns(COL3_DebitAccountID).Text = dr.Item("DebitAccountID").ToString
                tdbg3.Columns(COL3_Extend).Text = dr.Item("Extend").ToString
                FooterTotalGrid(tdbg3, COL3_AssignmentID)
        End Select
    End Sub

    Private Sub tdbg3_ButtonClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg3.ButtonClick
        If clsFilterDropdown.IsNewFilter = False Then Exit Sub
        If tdbg3.AllowUpdate = False Then Exit Sub
        If tdbg3.Splits(tdbg3.SplitIndex).DisplayColumns(e.ColIndex).Locked Then Exit Sub
        Select Case e.ColIndex
            Case IndexOfColumn(tdbg3, COL3_AssignmentID)
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg3, tdbg3.Columns(e.ColIndex).DataField)
                'If tdbd Is Nothing Then Exit Select
                'Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg3, e, tdbd)
                'If dr Is Nothing Then Exit Sub
                'AfterColUpdate(e.ColIndex, dr)
                If tdbd Is Nothing Then Exit Select
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdownMulti(tdbg3, e, tdbd)
                If dr Is Nothing Then Exit Sub
                'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                If dr(0).Item("AssignmentID").ToString = "+" Then 
                    tdbg3.Columns(COL3_AssignmentID).Text = ""
                    ShowD02F0101(COL3_AssignmentID)
                    Dim row As DataRow = ReturnDataRow(tdbdAssignmentID, "AssignmentID=" & SQLString(tdbg3.Columns(COL3_AssignmentID).Text))
                    AfterColUpdate(tdbg3.Col, row)
                Else
                    AfterColUpdate(tdbg3.Col, dr)
                End If
        End Select
    End Sub

    Private Sub tdbg3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg3.KeyDown
        If clsFilterDropdown.CheckKeydownFilterDropdown(tdbg3, e) Then
            Select Case tdbg3.Col
                Case IndexOfColumn(tdbg3, COL3_AssignmentID)
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg3, tdbg3.Columns(tdbg3.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdownMulti(tdbg3, e, tdbd)
                    If dr Is Nothing Then Exit Sub
                    If dr(0).Item("AssignmentID").ToString = "+" Then 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                        tdbg3.Columns(COL3_AssignmentID).Text = ""
                        ShowD02F0101(COL3_AssignmentID)
                        Dim row As DataRow = ReturnDataRow(COL3_AssignmentID, "AssignmentID=" & SQLString(tdbg3.Columns(COL3_AssignmentID).Text))
                        AfterColUpdate(tdbg3.Col, row)
                    Else
                        AfterColUpdate(tdbg3.Col, dr)
                    End If
                    Exit Sub
            End Select
        End If
    End Sub

    Private Sub tdbg3_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg3.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        Select Case e.Column.DataColumn.DataField
            Case COL3_AssignmentID
                If clsFilterDropdown.IsNewFilter Then Exit Sub 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                If tdbg3.Columns(COL3_AssignmentID).Text <> tdbdAssignmentID.Columns("AssignmentID").Text Then
                    tdbg3.Columns(COL3_AssignmentID).Text = ""
                    tdbg3.Columns(COL3_AssignmentName).Text = ""
                    tdbg3.Columns(COL3_DebitAccountID).Text = ""
                    tdbg3.Columns(COL3_Extend).Text = ""
                End If
        End Select
    End Sub
#End Region

#Region "Events of Textbox"

    Private Sub txtConvertedAmount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDepreciateAmount.TextChanged, txtRemainAmount.TextChanged
        Dim txtAmount As TextBox = CType(sender, TextBox)
        txtAmount.Text = Format(Number(txtAmount.Text), DxxFormat.D90_ConvertedDecimals)
    End Sub

    Private Sub txtAmountDepreciation_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmountDepreciation.TextChanged, txtConvertedAmount.TextChanged
        Dim txtAmount As TextBox = CType(sender, TextBox)
        txtAmount.Text = Format(Number(txtAmount.Text), DxxFormat.D90_ConvertedDecimals)
        'RemainAmount = ConvertedAmount - AmountDepreciation
        txtRemainAmount.Text = Format(Number(txtConvertedAmount.Text) - Number(txtAmountDepreciation.Text), DxxFormat.D90_ConvertedDecimals)
    End Sub

    Private Sub txtPercentage_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPercentage.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtPercentage_Validated(sender As Object, e As EventArgs) Handles txtPercentage.Validated
        '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        If D02Systems.IsCalPeriodByRate = True Then
            If Number(txtPercentage.Tag.ToString) = Number(txtPercentage.Text) Then Exit Sub

            txtPercentage.Text = SQLNumber(txtPercentage.Text, DxxFormat.D08_RatioDecimals)

            Dim dblServiceLife As Double = Number(txtConvertedAmount.Text) / (Number(txtConvertedAmount.Text) * Number(txtPercentage.Text) / 100) * 12
            If dblServiceLife <> 0 Then
                txtServiceLife.Text = SQLNumber(Math.Floor(dblServiceLife), DxxFormat.DefaultNumber0)
            Else
                txtServiceLife.Text = "0"
            End If

            'Tính lại gtri Mức khấu hao (txtDepreciateAmount)
            txtDepreciateAmount.Text = SQLNumber(Number(txtConvertedAmount.Text) / (Number(txtConvertedAmount.Text) / (Number(txtConvertedAmount.Text) * Number(txtPercentage.Text) / 100) * 12), DxxFormat.D90_ConvertedDecimals)

            txtPercentage.Tag = txtPercentage.Text
            txtServiceLife.Tag = txtServiceLife.Text
        End If
    End Sub

    Private Sub txtServiceLife_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMethodID.KeyPress, txtMethodEndID.KeyPress, txtServiceLife.KeyPress, txtDepreciatedPeriod.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDotSign)
    End Sub

    Private Sub txtServiceLife_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtServiceLife.TextChanged
        If D02Systems.IsCalPeriodByRate = True Then Exit Sub '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ

        'Lấy ServiceLife
        Dim dServiceLife As Double = Number(IIf(Number(txtServiceLife.Text) = 0, 1, Number(txtServiceLife.Text)))
        'DepreciatedAmount = ConvertedAmount / ServiceLife
        tdbg2.UpdateData()
        'Sửa tính ConvertedAmount theo Incident 	74480
        Dim sConvertedAmount As Double = 0
        If dtGrid2 IsNot Nothing Then
            sConvertedAmount = Number(txtConvertedAmount.Text) - Number(dtGrid2.Compute("SUM([ConvertedAmount])", "IsNotAllocate = 1 And DebitAccountID = " & SQLString(ReturnValueC1Combo(tdbcAssetAccountID).ToString)))
        End If

        txtDepreciateAmount.Text = SQLNumber(sConvertedAmount / dServiceLife, DxxFormat.D90_ConvertedDecimals)
        'Percentage = 100 / ServiceLife
        txtPercentage.Text = SQLNumber(100 / dServiceLife, DxxFormat.DefaultNumber2)
        txtServiceLife.Text = Format(Number(txtServiceLife.Text), DxxFormat.DefaultNumber0)

        txtServiceLife.Tag = txtServiceLife.Text
        txtPercentage.Tag = txtPercentage.Text
    End Sub

    Private Sub txtServiceLife_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtServiceLife.Validated
        If D02Systems.IsCalPeriodByRate = True Then Exit Sub '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ

        If Number(txtServiceLife.Tag.ToString) = Number(txtServiceLife.Text) Then Exit Sub

        If Number(txtConvertedAmount.Text) <> 0 And Number(txtAmountDepreciation.Text) <> 0 And Number(txtDepreciateAmount.Text) <> 0 And L3Int(ReturnValueC1Combo(tdbcAssetID, "IntCode")) = 0 Then
            If Number(txtConvertedAmount.Text) < 0.001 Then Exit Sub
            txtDepreciatedPeriod.Text = (Number(txtAmountDepreciation.Text) * L3Int(txtDepreciatedPeriod.Text) / Number(txtConvertedAmount.Text)).ToString
        End If
        If txtDepreciatedPeriod.Text <> "" Then
            If Number(txtServiceLife.Text) < Number(txtDepreciatedPeriod.Text) Then
                D99C0008.MsgL3("Số kỳ khấu hao không được nhỏ hơn Số kỳ đã khấu hao.")
                txtServiceLife.Focus()
                txtServiceLife.SelectAll()
                Exit Sub
            End If
        End If
        'Tính lại gtri Mức khấu hao (txtDepreciateAmount)
        If Number(txtServiceLife.Text) = 0 Or txtServiceLife.Text = "" Then
            'Sửa tính ConvertedAmount theo Incident 	74480
            Dim sConvertedAmount As Double = Number(txtConvertedAmount.Text) - Number(dtGrid2.Compute("SUM([ConvertedAmount])", "IsNotAllocate = 1 And DebitAccountID = " & SQLString(ReturnValueC1Combo(tdbcAssetAccountID).ToString)))
            txtDepreciateAmount.Text = SQLNumber(sConvertedAmount, DxxFormat.D90_ConvertedDecimals)
            Exit Sub
        End If

    End Sub

    Private Sub txtDepreciateAmount_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDepreciateAmount.Validated
        'If Not IsNumeric(txtDepreciateAmount.Text) Then
        '    txtDepreciateAmount.Text = ""
        'End If
        'If txtDepreciateAmount.Text <> "" And Number(txtDepreciateAmount.Text) > Number(txtRemainAmount.Text) Then
        '    txtDepreciateAmount.Text = ""
        'End If
        ''Tính lại gtri Tỷ lệ khấu hao (txtPercentage)
        'If txtServiceLife.Text = "" Or Number(txtServiceLife.Text) = 0 Then
        '    txtPercentage.Text = "0"
        'Else
        '    txtPercentage.Text = SQLNumber(100 / Number(txtServiceLife.Text), DxxFormat.D08_RatioDecimals)
        'End If
    End Sub
#End Region

    Private Sub tdbg3_LockedColumns()
        tdbg3.Splits(SPLIT0).DisplayColumns(COL3_AssignmentName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg3.Splits(SPLIT0).DisplayColumns(COL3_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub


    Private Sub tdbg2_LockedColumns()
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_VoucherTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_VoucherNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_VoucherDate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_RefDate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_SeriNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_RefNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_ObjectTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_ObjectID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_ObjectName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_Description).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_CreditAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_CurrencyID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_ExchangeRate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_OriginalAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_ConvertedAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_VATGroupID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_VATTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(tdbg2.Splits.Count - 1).DisplayColumns(COL2_VATNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    Private Sub tdbg_LockedColumns()
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_VoucherTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_VoucherNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_VoucherDate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_RefDate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_SeriNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_RefNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_ObjectTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_ObjectID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_ObjectName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_Description).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_CreditAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_CurrencyID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_ExchangeRate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_OriginalAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_ConvertedAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_VATGroupID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_VATTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_VATNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(tdbg.Splits.Count - 1).DisplayColumns(COL_SourceID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub


    Dim _historyIDMaster As String = "" 'Sinh IGE cho master D02T5000

    Private Function SQLUpdateD02T0012_Delete() As StringBuilder
        Dim sRet As New StringBuilder
        Dim strSQL As String = "SELECT * FROM D02T0012 WITH(NOLOCK)  WHERE(AssetID = " & SQLString(ReturnValueC1Combo(tdbcAssetID)) & ") And " & _
                                " TransactionTypeID in ('MM', '') And (DivisionID = '" & gsDivisionID & "') "
        Dim dtTemp As DataTable = ReturnDataTable(strSQL)
        If dtTemp.Rows.Count = 0 Then Return sRet
        Dim sSQL As New StringBuilder
        'Lấy những dòng không tồn tại trên lưới Chọn bút toán
        For i As Integer = 0 To dtTemp.Rows.Count - 1
            Dim dr() As DataRow = dtGrid2.Select("TransactionID" & " = '' And " & "TransactionID" & " is not NULL And " & "TransactionID" & " = " & SQLString(dtTemp.Rows(i).Item("TransactionID")), "")
            If dr.Length = 0 Then
                sSQL.Append(" Update D02T0012  Set Status=0, AssetID = '' , TransactionTypeID = '',IsNotAllocate=0 ")
                sSQL.Append(" Where TransactionID=" & SQLString(dtTemp.Rows(i).Item("TransactionID")))
                sRet.Append(sSQL)
                sSQL = New StringBuilder
            End If
        Next
        Return sRet
    End Function

    '   sSQL = sSQL & "Update D02T5000 "
    'sSQL = sSQL & "Set BeginMonth=" & IIf(IsNull(txtBeginDep.Text) Or (txtBeginDep.Text) = "", 0, "" & CInt(Left(txtBeginDep.Text, 2)) & "")
    'sSQL = sSQL & " ,BeginYear =" & IIf(IsNull(txtBeginDep.Text) Or (txtBeginDep.Text) = "", 0, "" & CInt(Right(txtBeginDep.Text, 4)) & "")
    'sSQL = sSQL & " ,IsStopDepreciation=0"
    'sSQL = sSQL & " Where AssetID='" & SmoothQuote(strAssetIDBoundText) & "'"
    'sSQL = sSQL & " And   BatchID='" & SmoothQuote(strAssetIDBoundText) & "'"
    'sSQL = sSQL & " And   HistoryTypeID='SD'"

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T5000
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 03/01/2012 02:57:48
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T5000(ByVal sHistoryTypeID As String) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T5000 Set ")

        If sHistoryTypeID = "SU" Then
            sSQL.Append("BeginMonth = " & SQLNumber(Strings.Left(c1dateBeginDep.Text, 2)) & COMMA) 'tinyint, NOT NULL
            sSQL.Append("BeginYear = " & SQLNumber(Strings.Right(c1dateBeginDep.Text, 4)) & COMMA) 'smallint, NOT NULL
            sSQL.Append("IsStopUse = 0" & COMMA)
        ElseIf sHistoryTypeID = "SD" Then
            sSQL.Append("BeginMonth = " & SQLNumber(Strings.Left(c1dateBeginDep.Text, 2)) & COMMA) 'tinyint, NOT NULL
            sSQL.Append("BeginYear = " & SQLNumber(Strings.Right(c1dateBeginDep.Text, 4)) & COMMA) 'smallint, NOT NULL
            sSQL.Append("IsStopDepreciation=0" & COMMA)
        Else
            ' update 15/10/2013 id 60527 
            sSQL.Append("BeginMonth = " & SQLNumber(Strings.Left(c1dateBeginUsing.Text, 2)) & COMMA) 'tinyint, NOT NULL
            sSQL.Append("BeginYear = " & SQLNumber(Strings.Right(c1dateBeginUsing.Text, 4)) & COMMA) 'smallint, NOT NULL
            sSQL.Append(" ServiceLife=" & SQLNumber(txtServiceLife.Text) & COMMA)
        End If
        ' update 15/10/2013 id 60527 
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) ', NOT NULL
        sSQL.Append("LastModifyDate = GetDate()") ', NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("AssetID = " & SQLString(_assetID))
        sSQL.Append(" And BatchID=" & SQLString(_assetID))
        sSQL.Append(" And HistoryTypeID=" & SQLString(sHistoryTypeID))
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P5555
    '# Created User: KIM LONG
    '# Created Date: 17/05/2016 10:55:36
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P5555() As String
        Dim sSQL As String = ""
        sSQL &= ("-- EXEC D02P5555 " & vbCrlf)
        sSQL &= "Exec D02P5555 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(Me.Name) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[200], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Laguage, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetID)) 'AssetID, varchar[20], NOT NULL
        Return sSQL
    End Function

    Dim bCheckIsManagement As Boolean = False
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        tdbg.UpdateData()
        tdbg2.UpdateData()
        tdbg3.UpdateData()

        If Not AllowSave() Then Exit Sub
        btnSave.Enabled = False
        btnClose.Enabled = False
        gbSavedOK = False
       
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        _historyIDMaster = CreateIGE("D02T5000", "HistoryID", "02", "HB", gsStringKey)
        Select Case _FormState
            Case EnumFormState.FormAdd
                sSQL.Append(SQLUpdateD02T0001.ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0012s.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000(_historyIDMaster, giTranMonth, giTranYear, "OB", tdbcObjectTypeID.Text, tdbcObjectID.Text, tdbcEmployeeID.Text, txtEmployeeName.Text).ToString & vbCrLf)
                Dim _historyIDMasterNew As String = CreateIGE("D02T5000", "HistoryID", "02", "HB", gsStringKey)
                If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" Then
                    bCheckIsManagement = True
                    sSQL.Append(SQLInsertD02T5000(_historyIDMasterNew, giTranMonth, giTranYear, "OB", tdbcObjectTypeID.Text, tdbcObjectID.Text, tdbcEmployeeID.Text, txtEmployeeName.Text).ToString & vbCrLf)
                End If
                sSQL.Append(SQLInsertD02T5000s().ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000_3().ToString & vbCrLf)
                'L§u lÜch sõ thanh lü
                Dim sHistory As String = CreateIGE("D02T5000", "HistoryID", "02", "HB", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistory, giTranMonth, giTranYear, "IL").ToString & vbCrLf)
                'Update D02T0012
                sSQL.Append(SQLUpdateD02T0012s_tdbg)

                '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
                Dim sHistoryAAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HB", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistoryAAC, giTranMonth, giTranYear, "AAC", , , , , , , , , , , tdbcAssetAccountID.Text).ToString)
                Dim sHistoryDAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HB", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistoryDAC, giTranMonth, giTranYear, "DAC", , , , , , , , , , , , tdbcDepAccountID.Text).ToString)
                sSQL.Append(SQLInsertD02T5010(sHistoryAAC, "AAC", tdbcAssetAccountID.Text))
                sSQL.Append(SQLInsertD02T5010(sHistoryDAC, "DAC", tdbcDepAccountID.Text))
            Case EnumFormState.FormEdit, EnumFormState.FormEditOther
                sSQL.Append(SQLUpdateD02T0001.ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0012_Delete().ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0012s.ToString & vbCrLf)
                sSQL.Append(SQLDeleteD02T5000() & vbCrLf)
                sSQL.Append(SQLInsertD02T5000(_historyIDMaster, giTranMonth, giTranYear, "OB", tdbcObjectTypeID.Text, tdbcObjectID.Text, tdbcEmployeeID.Text, txtEmployeeName.Text).ToString & vbCrLf)
                Dim _historyIDMasterNew As String = CreateIGE("D02T5000", "HistoryID", "02", "HB", gsStringKey)
                If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" Then
                    bCheckIsManagement = True
                    sSQL.Append(SQLInsertD02T5000(_historyIDMasterNew, giTranMonth, giTranYear, "OB", tdbcObjectTypeID.Text, tdbcObjectID.Text, tdbcEmployeeID.Text, txtEmployeeName.Text).ToString & vbCrLf)
                End If
                sSQL.Append(SQLInsertD02T5000s().ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T5000("SU").ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T5000("SD").ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T5000("SL").ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0012s_tdbg)

                '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
                Dim sHistoryAAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HB", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistoryAAC, giTranMonth, giTranYear, "AAC", , , , , , , , , , , tdbcAssetAccountID.Text).ToString)
                Dim sHistoryDAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HB", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistoryDAC, giTranMonth, giTranYear, "DAC", , , , , , , , , , , , tdbcDepAccountID.Text).ToString)
                sSQL.Append(SQLInsertD02T5010(sHistoryAAC, "AAC", tdbcAssetAccountID.Text))
                sSQL.Append(SQLInsertD02T5010(sHistoryDAC, "DAC", tdbcDepAccountID.Text))
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then

            '******
            Dim dr() As DataRow = dtGrid2.Select("ObjectTypeID" & "<>'' or " & "ObjectTypeID" & " is not null")
            Dim sObjectTypeID As String = "", sObjectID As String = ""
            If dr.Length > 0 Then
                sObjectTypeID = dr(0).Item("ObjectTypeID").ToString
                sObjectID = dr(0).Item("ObjectID").ToString
            End If

            Dim iUpdate As Integer
            Dim _dt As DataTable = ReturnDataTable(SQLStoreD02P5555())

            'ID 86210 17.05.2016
            If _dt.Rows.Count > 0 Then
                iUpdate = 1
                If L3String(_dt.Rows(0)("Status")) <> "0" Then
                    If L3Bool(_dt.Rows(0)("MsgAsk")) Then
                        Dim bRunMsk As DialogResult = D99C0008.MsgAsk(L3String(_dt.Rows(0)("Message")), MessageBoxDefaultButton.Button1)
                        If bRunMsk = Windows.Forms.DialogResult.Yes Then iUpdate = 1 : GoTo 1
                        iUpdate = 0
                    Else
                        D99C0008.Msg(L3String(_dt.Rows(0)("Message")))
                        iUpdate = 0
                    End If
                End If
            End If
1:
            ExecuteSQLNoTransaction(SQLStoreD02P1001(iUpdate, sObjectID))

            SaveOK()
            gbSavedOK = True
            btnClose.Enabled = True

            '******************************************************

            tdbg2.UpdateData()
            dr = dtGrid2.Select("ProjectID <> '' And PropertyProductID <> ''")
            For i As Integer = 0 To dr.Length - 1
                ExecuteSQL(SQLStoreD02P2035(dr(i).Item("TransactionID").ToString))
            Next

            '******
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _assetID = tdbcAssetID.Text
                    'ExecuteAuditLog(sAuditCode, "01", tdbcAssetID.Text, txtAssetName.Text)
                    Lemon3.D91.RunAuditLog("02", sAuditCode, "01", tdbcAssetID.Text, txtAssetName.Text)
                    btnNext.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit, EnumFormState.FormEditOther
                    'ExecuteAuditLog(sAuditCode, "02", tdbcAssetID.Text, txtAssetName.Text)
                    Lemon3.D91.RunAuditLog("02", sAuditCode, "02", tdbcAssetID.Text, txtAssetName.Text)
                    btnSave.Enabled = True
                    btnClose.Focus()
            End Select
            LoadTDBGrid1()

        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1001
    '# Created User: KIM LONG
    '# Created Date: 17/05/2016 11:28:34
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1001(ByVal iUpdate As Integer, ByVal sObjectID As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- --cap nhat nha cung cap cho tai san" & vbCrLf)
        sSQL &= "Exec D02P1001 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(Me.Name) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[200], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Laguage, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(iUpdate) & COMMA 'IsUpdate, tinyint, NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= SQLString(sObjectID) 'SupplierID, varchar[20], NOT NULL
        Return sSQL
    End Function



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2035
    '# Created User: HUỲNH KHANH
    '# Created Date: 24/02/2015 10:15:39
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2035(ByVal sTransactionID As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Update du lieu cho cac module sau khi luu hoan tat" & vbCrLf)
        sSQL &= "Exec D02P2035 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= SQLString(sTransactionID) & COMMA 'TransactionID, varchar[50], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(D02) & COMMA 'ModuleID, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function



    Dim arrCipID As New StringBuilder 'Danh sách các CipID trên lưới 1

    Private Function AllowSave() As Boolean
        If tdbcAssetID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Ma_tai_san"))
            tdbcAssetID.Focus()
            Return False
        End If

        If tdbcAssetAccountID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Tai_khoan_TS"))
            tdbcAssetAccountID.Focus()
            Return False
        End If
        If tdbcDepAccountID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Tai_khoan_KHU"))
            tdbcDepAccountID.Focus()
            Return False
        End If
        If tdbcObjectTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Bo_phan_quan_ly"))
            tabMain.SelectedTab = tabPage1
            tdbcObjectTypeID.Focus()
            Return False
        End If
        If tdbcObjectID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Bo_phan_quan_ly"))
            tabMain.SelectedTab = tabPage1
            tdbcObjectID.Focus()
            Return False
        End If

        If D02Systems.ObligatoryReceiver AndAlso tdbcEmployeeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Nguoi_tiep_nhan"))
            tabMain.SelectedTab = tabPage1
            tdbcEmployeeID.Focus()
            Return False
        End If
        If Number(txtConvertedAmount.Text) < Number(txtAmountDepreciation.Text) Then
            'Mức khấu hao không được lớn hơn nguyên giá
            D99C0008.MsgL3(rL3("Muc_khau_hao_khong_duoc_lon_hon_nguyen_gia"))
            tabMain.SelectedTab = tabPage1
            If txtConvertedAmount.Enabled Then txtConvertedAmount.Focus()
            Return False
        End If

        'If Val(txtAmountDepreciation.Text) < 0 Then
        '    MsgBox(Language(MSG_NegativeAmountDepreciation), vbCritical + vbInformation, Language(MSG_MESSAGE))
        '    SSTab1.Tab = 0
        '    '.SetFocus
        '    TestInput = False
        '    Exit Function
        'End If
        Dim dr() As DataRow = dtGrid2.Select("IsNotAllocate = 0")
        If dr.Length > 0 Then
            If Number(txtServiceLife.Text.Trim) = 0 Then
                D99C0008.MsgNotYetEnter(rL3("So_ky_khau_hao"))
                txtServiceLife.Focus()
                Return False
            End If

        End If
        'If Number(txtDepreciatedPeriod.Text.Trim) = 0 Then
        '    D99C0008.MsgNotYetEnter(rl3("So_ky_da_khau_hao"))
        '    txtDepreciatedPeriod.Focus()
        '    Return False
        'End If
        If Number(txtServiceLife.Text) < Number(txtDepreciatedPeriod.Text) Then
            D99C0008.MsgL3(rL3("So_ky_da_khau_hao_khong_duoc_lon_hon_So_ky_khau_hao"))
            txtDepreciatedPeriod.Focus()
            Return False
        End If

        '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        If D02Systems.IsCalPeriodByRate = True Then
            If Number(txtPercentage.Text) = 0 Then
                D99C0008.MsgNotYetEnter(rL3("Ty_le_khau_hao_%"))
                txtPercentage.Focus()
                Return False
            End If
        End If

        If Number(txtPercentage.Text) > 100 Then
            D99C0008.MsgL3(rL3("Ty_le_khau_hao_phai_nho_hon_100%"))
            txtPercentage.SelectionStart = 0
            txtPercentage.SelectionLength = txtPercentage.Text.Length
            txtPercentage.Focus()
            Return False
        End If

        ''If txtDepreciatedPeriod.Text.Trim.Length > MaxInt Then
        ''    D99C0008.MsgL3(rl3("So_vuot_qua_gioi_han"))
        ''    txtDepreciatedPeriod.Focus()
        ''    Return False
        ''End If
        If c1dateBeginUsing.Value.ToString = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ky_bat_dau_su_dung"))
            c1dateBeginUsing.Focus()
            Return False
        End If
        If c1dateBeginDep.Value.ToString = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ky_bat_dau_KH"))
            c1dateBeginDep.Focus()
            Return False
        End If

        Dim bPeriod As Double = Year(CDate(c1dateBeginDep.Value)) * 100 + Month(CDate(c1dateBeginDep.Value))
        Dim uPeriod As Double = Year(CDate(c1dateBeginUsing.Value)) * 100 + Month(CDate(c1dateBeginUsing.Value))
        If bPeriod < uPeriod Then
            D99C0008.MsgL3(rL3("Ky_bat_dau_tinh_khau_hao_phai_lon_hon_hoac_bang_ky_su_dung"))
            c1dateBeginDep.Focus()
            Return False
        End If
        Dim curPeriod As Double = giTranYear * 100 + giTranMonth
        If bPeriod < curPeriod Then
            D99C0008.MsgL3(rL3("Ky_bat_dau_tinh_khau_hao_phai_lon_hon_hoac_bang_ky_hinh_thanh"))
            c1dateBeginDep.Focus()
            Return False
        End If
        If uPeriod < curPeriod Then
            D99C0008.MsgL3(rL3("Ky_su_dung_phai_lon_hon_hoac_bang_ky_hinh_thanh"))
            c1dateBeginUsing.Focus()
            Return False
        End If

        '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        If D02Systems.IsCalDepByDate = True Then
            If c1dateDepDate.Text = "" Then
                D99C0008.MsgNotYetEnter(lblDepDate.Text)
                c1dateDepDate.Focus()
                Return False
            End If
        End If
        If c1dateDepDate.Text <> "" And c1dateBeginDep.Text <> "" Then
            If CDate(c1dateDepDate.Value).Year <> CDate(c1dateBeginDep.Value).Year OrElse CDate(c1dateDepDate.Value).Month <> CDate(c1dateBeginDep.Value).Month Then
                D99C0008.MsgL3(rL3("Ngay_khau_hao_phai_nam_trong_ky_khau_hao"), L3MessageBoxIcon.Exclamation)
                c1dateDepDate.Focus()
                Return False
            End If
        End If

        If tdbg2.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tabMain.SelectedTab = tabPage2
            tdbg2.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbg2.RowCount - 1
            If tdbg2(i, COL2_OriginalAmount).ToString = "" Then
                D99C0008.MsgNotYetEnter(rL3("So_tien_nguyen_te"))
                tabMain.SelectedTab = tabPage2
                tdbg2.Focus()
                tdbg2.SplitIndex = SPLIT1
                tdbg2.Col = COL2_OriginalAmount
                tdbg2.Bookmark = i
                Return False
            End If
            If tdbg2(i, COL2_SourceID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rL3("Nguon_von"))
                tabMain.SelectedTab = tabPage2
                tdbg2.Focus()
                tdbg2.SplitIndex = SPLIT1
                tdbg2.Col = COL2_SourceID
                tdbg2.Bookmark = i
                Return False
            End If
        Next
        'Kh¤ng câ TK Ní nªo ¢§íc ¢Ünh kho¶n giçng TK tªi s¶n!
        Dim dr2() As DataRow = dtGrid2.Select("DebitAccountID" & "= " & SQLString(tdbcAssetAccountID.Text))
        If dr2.Length = 0 Then
            D99C0008.MsgL3(rL3("Khong_co_TK_no_nao_duoc_dinh_khoan_giong_TK_tai_san"))
            tabMain.SelectedTab = tabPage2
            tdbg2.Focus()
            tdbg2.SplitIndex = SPLIT1
            tdbg2.Col = COL2_DebitAccountID
            tdbg2.Row = 0
            Return False
        End If
        'Nguy£n giÀ ¢Ünh kho¶n kh¤ng bÂng nguy£n giÀ cïa tªi s¶n!
        Dim dConvertedAmo As Double = Number(Format(CallTotalAmount(), DxxFormat.D90_ConvertedDecimals))
        If dConvertedAmo <> Number(txtConvertedAmount.Text) Then
            D99C0008.MsgL3(rL3("Nguyen_gia_dinh_khoan_khong_bang_nguyen_gia_cua_tai_san"))
            tabMain.SelectedTab = tabPage2
            tdbg2.Focus()
            tdbg2.SplitIndex = SPLIT1
            tdbg2.Col = COL2_ConvertedAmount
            tdbg2.Row = 0
            Return False
        End If

        '***************************
        'Tabpage 3
        If tdbg3.RowCount = 0 Then
            D99C0008.MsgNotYetEnter(rL3("Ma_phan_bo"))
            tabMain.SelectedTab = TabPage3
            tdbg3.Focus()
            tdbg3.SplitIndex = SPLIT0
            tdbg3.Col = IndexOfColumn(tdbg3, COL3_AssignmentID)
            tdbg3.Bookmark = 0
            Return False
        End If
        If tdbcAssetID.Columns("AssignmentTypeID").Text = "" Then Return True

        Select Case L3Int(ReturnValueC1Combo(tdbcAssetID, "AssignmentTypeID"))
            Case 0
                For i As Integer = 0 To tdbg3.RowCount - 1
                    If tdbg3(i, COL3_AssignmentID).ToString = "" Then
                        D99C0008.MsgNotYetEnter(rL3("Ma_phan_bo"))
                        tabMain.SelectedTab = TabPage3
                        tdbg3.Focus()
                        tdbg3.SplitIndex = SPLIT0
                        tdbg3.Col = IndexOfColumn(tdbg3, COL3_AssignmentID)
                        tdbg3.Bookmark = i
                        Return False
                    End If
                    If tdbg3(i, COL3_AssignmentName).ToString = "" Then
                        D99C0008.MsgNotYetEnter(rL3("Ten_phan_bo"))
                        tabMain.SelectedTab = TabPage3
                        tdbg3.Focus()
                        tdbg3.SplitIndex = SPLIT0
                        tdbg3.Col = IndexOfColumn(tdbg3, COL3_AssignmentName)
                        tdbg3.Bookmark = i
                        Return False
                    End If
                Next
                Dim dTotalPercent As Double = Number(dtGrid3.Compute("SUM(" & COL3_PercentAmount & ")", ""))
                If dTotalPercent <> 1 Then '~100%
                    D99C0008.MsgL3(rL3("Tong_ty_le_phai_bang_100U"))
                    tabMain.SelectedTab = TabPage3
                    tdbg3.Focus()
                    tdbg3.SplitIndex = SPLIT0
                    tdbg3.Col = IndexOfColumn(tdbg3, COL3_PercentAmount)
                    tdbg3.Bookmark = 0
                    Return False
                End If
            Case 2
                Dim dt As DataTable = dtGrid3.DefaultView.ToTable
                Dim drExt1() As DataRow = dt.Select(COL3_Extend & "=1")
                If drExt1.Length = 0 Then
                    D99C0008.MsgL3(rL3("Ten_phan_bo"))
                    tabMain.SelectedTab = TabPage3
                    tdbg3.Focus()
                    tdbg3.SplitIndex = SPLIT0
                    tdbg3.Col = IndexOfColumn(tdbg3, COL3_AssignmentName)
                    tdbg3.Bookmark = dt.Rows.IndexOf(drExt1(0))
                    Return False
                End If
                Dim drExt2() As DataRow = dt.Select(COL3_Extend & "=2")
                If drExt2.Length <> 1 Then
                    D99C0008.MsgL3(rL3("Ten_phan_bo_chenh_lech"))
                    tabMain.SelectedTab = TabPage3
                    tdbg3.Focus()
                    tdbg3.SplitIndex = SPLIT0
                    tdbg3.Col = IndexOfColumn(tdbg3, COL3_AssignmentName)
                    tdbg3.Bookmark = dt.Rows.IndexOf(drExt2(0))
                    Return False
                End If
        End Select
        If D02Systems.IsObligatoryManagement Then
            If tdbcManagementObjTypeID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rL3("Bo_phan_quan_ly"))
                tabMain.SelectedTab = tabPage1
                tdbcManagementObjTypeID.Focus()
                Return False
            End If
            If tdbcManagementObjID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rL3("Bo_phan_quan_ly"))
                tabMain.SelectedTab = tabPage1
                tdbcManagementObjID.Focus()
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub btnHotKeys_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHotKeys.Click
        Dim f As New D02F7777
        With f
            .CallShowForm(Me.Name)
            .ShowDialog()
        End With
    End Sub


    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        LoadtdbcAssetID()
        'Có load lại các combo k?
        tdbcAssetID.SelectedValue = ""
        tdbcAssetAccountID.SelectedValue = ""
        tdbcDepAccountID.SelectedValue = ""
        tdbcEmployeeID.SelectedValue = ""
        tdbcObjectTypeID.SelectedValue = ""
        tdbcObjectID.SelectedValue = ""
        _assetID = ""
        LoadAddNew()
        tabMain.SelectedTab = tabPage1
        LoadTDBGrid1(True)
        LoadTDBGrid2(True)
        LoadTDBGrid3(True)
        txtDepreciateAmount.Text = ""
        tdbcAssetID.Focus()
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Cap_nhat_thong_tin_mua_moi_TSCD_-_D02F1001") & UnicodeCaption(gbUnicode) 'CËp nhËt th¤ng tin mua mìi TSC˜ - D02F1001
        '================================================================ 
        lblDepAccountID.Text = rL3("Tai_khoan_KHU") 'Tài khoản KH
        lblAssetID.Text = rL3("Ma_tai_san") 'Mã tài sản

        lblEmployeeID.Text = rL3("Nguoi_tiep_nhan") 'Người tiếp nhận
        lblAssetAccount.Text = rL3("Tai_khoan_TS") 'Tài khoản TS
        lblMethodID.Text = rL3("Phuong_phap_KH") 'Phương pháp KH
        lblMethodEndID.Text = rL3("Khau_hao_ky_cuoi") 'Khấu hao kỳ cuối
        lblDeprTableName.Text = rL3("Bang_khau_hao") 'Bảng khấu hao
        lblConvertedAmount.Text = rL3("Nguyen_gia") 'Nguyên giá
        lblAmountDepreciation.Text = rL3("Hao_mon_luy_ke") 'Hao mòn luỹ kế
        lblRemainAmount.Text = rL3("Gia_tri_con_lai") 'Giá trị còn lại
        lblServiceLife.Text = rL3("So_ky_khau_hao") 'Số kỳ khấu hao
        lblDepreciatedPeriod.Text = rL3("So_ky_da_khau_hao") 'Số kỳ đã khấu hao
        lblPercentage.Text = rL3("Ty_le_khau_hao") & " %" 'Tỷ lệ khấu hao %
        lblteUseDate.Text = rL3("Ngay_su_dung") 'Ngày sử dụng
        lblteAssetDate.Text = rL3("Ngay_tiep_nhan") 'Ngày tiếp nhận
        lblDepreciateAmount.Text = rL3("Muc_khau_hao") 'Mức khấu hao
        lblteBeginUsing.Text = rL3("Ky_bat_dau_su_dung") 'Kỳ bắt đầu sử dụng
        lblteBeginDep.Text = rL3("Ky_bat_dau_KH") 'Kỳ bắt đầu KH
        lblPropertyProductID.Text = rL3("Ma_BDS") 'Mã BĐS
        lblObjectTypeID.Text = rL3("Bo_phan_tiep_nhan") 'Bộ phận tiếp nhận
        lblDepDate.Text = rL3("Ngay_bat_dau_khau_hao") 'Ngày bắt đầu khấu hao
        '================================================================ 
        btnChoose.Text = "&" & rL3("Chon") 'Chọn
        btnHotKeys.Text = rL3("Phim_nong") 'Phím nóng
        btnSave.Text = rL3("_Luu") '&Lưu
        btnNext.Text = rL3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rL3("Do_ng") 'Đó&ng
        '================================================================ 
        grpAssetID.Text = rL3("Ma_tai_san") 'Mã tài sản
        grpFinancialInfo.Text = rL3("Thong_tin_tai_chinh") 'Thông tin tài chính
        '================================================================ 
        tabPage1.Text = "1. " & rL3("Chon_but_toan") 'Chọn bút toán
        tabPage2.Text = "2. " & rL3("But_toan_hinh_thanh") 'Bút toán hình thành
        TabPage3.Text = "3. " & rL3("Phan_bo_khau_hao") 'Phân bổ khấu hao
        '================================================================ 
        tdbcDepAccountID.Columns("AccountID").Caption = rL3("Ma") 'Mã
        tdbcDepAccountID.Columns("AccountName").Caption = rL3("Ten") 'Tên
        tdbcAssetAccountID.Columns("AccountID").Caption = rL3("Ma") 'Mã
        tdbcAssetAccountID.Columns("AccountName").Caption = rL3("Ten") 'Tên
        tdbcEmployeeID.Columns("EmployeeID").Caption = rL3("Ma") 'Mã
        tdbcEmployeeID.Columns("EmployeeName").Caption = rL3("Ten") 'Tên
        tdbcObjectID.Columns("ObjectTypeID").Caption = rL3("Loai_DT") 'Loại ĐT
        tdbcObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên

        tdbcObjectTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcObjectTypeID.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcAssetID.Columns("AssetID").Caption = rL3("Ma") 'Mã
        tdbcAssetID.Columns("AssetName").Caption = rL3("Ten") 'Tên
        tdbcPropertyProductID.Columns("PropertyProductID").Caption = rL3("Ma_BDS") 'Mã BĐS
        '================================================================ 
        tdbdSourceID.Columns("SourceID").Caption = rL3("Ma") 'Mã
        tdbdSourceID.Columns("SourceName").Caption = rL3("Ten") 'Tên
        tdbdAssignmentID.Columns("AssignmentID").Caption = rL3("Ma") 'Mã
        tdbdAssignmentID.Columns("AssignmentName").Caption = rL3("Ten") 'Tên
        '================================================================ 
        tdbg.Columns("Choose").Caption = rL3("Chon") 'Chọn
        tdbg.Columns("VoucherTypeID").Caption = rL3("Loai_phieu") 'Loại phiếu
        tdbg.Columns("VoucherNo").Caption = rL3("So_phieu") 'Số phiếu
        tdbg.Columns("VoucherDate").Caption = rL3("Ngay_phieu") 'Ngày phiếu
        tdbg.Columns("RefDate").Caption = rL3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg.Columns("SeriNo").Caption = rL3("So_Seri") 'Số Sêri
        tdbg.Columns("RefNo").Caption = rL3("So_hoa_don") 'Số hóa đơn
        tdbg.Columns("ObjectTypeID").Caption = rL3("Loai_doi_tuong") 'Loại đối tượng
        tdbg.Columns("ObjectID").Caption = rL3("Ma_doi_tuong") 'Mã đối tượng
        tdbg.Columns("ObjectName").Caption = rL3("Ten_doi_tuong") 'Tên đối tượng
        tdbg.Columns("Description").Caption = rL3("Dien_giai") 'Diễn giải
        tdbg.Columns("DebitAccountID").Caption = rL3("Tai_khoan_no") 'Tài khoản nợ
        tdbg.Columns("CreditAccountID").Caption = rL3("Tai_khoan_co") 'Tài khoản có
        tdbg.Columns("CurrencyID").Caption = rL3("Loai_tien") 'Loại tiền
        tdbg.Columns("ExchangeRate").Caption = rL3("Ty_gia") 'Tỷ giá
        tdbg.Columns("OriginalAmount").Caption = rL3("So_tien_nguyen_te") 'Số tiền nguyên tệ
        tdbg.Columns("ConvertedAmount").Caption = rL3("So_tien_quy_doi") 'Số tiền quy đổi
        tdbg.Columns("VATGroupID").Caption = rL3("Nhom_thue") 'Nhóm thuế
        tdbg.Columns("VATTypeID").Caption = rL3("Loai_hoa_don") 'Loại hóa đơn rl3("Loai_thue")
        tdbg.Columns("VATNo").Caption = rL3("Ma_so_thue") 'Mã số thuế
        tdbg.Columns("SourceID").Caption = rL3("Nguon_von") 'Nguồn vốn

        '================================================================ 
        tdbg2.Columns(COL2_Choose).Caption = rL3("Chi_phi") 'Chi phí
        tdbg2.Columns(COL2_IsNotAllocate).Caption = rL3("Khong_tinh_KH") 'Không tính KH
        tdbg2.Columns(COL2_SourceID).Caption = rL3("Nguon_von") 'Nguồn vốn
        tdbg2.Columns(COL2_VoucherTypeID).Caption = rL3("Loai_phieu") 'Loại phiếu
        tdbg2.Columns(COL2_VoucherNo).Caption = rL3("So_phieu") 'Số phiếu
        tdbg2.Columns(COL2_VoucherDate).Caption = rL3("Ngay_phieu") 'Ngày phiếu
        tdbg2.Columns(COL2_RefDate).Caption = rL3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg2.Columns(COL2_SeriNo).Caption = rL3("So_Seri") 'Số Sêri
        tdbg2.Columns(COL2_RefNo).Caption = rL3("So_hoa_don") 'Số hóa đơn
        tdbg2.Columns(COL2_ObjectTypeID).Caption = rL3("Loai_doi_tuong1") 'Loại đối tượng
        tdbg2.Columns(COL2_ObjectID).Caption = rL3("Ma_doi_tuong") 'Mã đối tượng
        tdbg2.Columns(COL2_ObjectName).Caption = rL3("Ten_doi_tuong") 'Tên đối tượng
        tdbg2.Columns(COL2_Description).Caption = rL3("Dien_giai") 'Diễn giải
        tdbg2.Columns(COL2_DebitAccountID).Caption = rL3("Tai_khoan_no") 'Tài khoản nợ
        tdbg2.Columns(COL2_CreditAccountID).Caption = rL3("Tai_khoan_co") 'Tài khoản có
        tdbg2.Columns(COL2_CurrencyID).Caption = rL3("Loai_tien") 'Loại tiền
        tdbg2.Columns(COL2_ExchangeRate).Caption = rL3("Ty_gia") 'Tỷ giá
        tdbg2.Columns(COL2_OriginalAmount).Caption = rL3("So_tien_nguyen_te") 'Số tiền nguyên tệ
        tdbg2.Columns(COL2_ConvertedAmount).Caption = rL3("So_tien_quy_doi") 'Số tiền quy đổi
        tdbg2.Columns(COL2_VATGroupID).Caption = rL3("Nhom_thue") 'Nhóm thuế
        tdbg2.Columns(COL2_VATTypeID).Caption = rL3("Loai_thue") 'Loại thuế
        tdbg2.Columns(COL2_VATNo).Caption = rL3("Ma_so_thue") 'Mã số thuế


        tdbg3.Columns("AssignmentID").Caption = rL3("Ma_phan_bo") 'Mã phân bổ
        tdbg3.Columns("AssignmentName").Caption = rL3("Ten_phan_bo") 'Tên phân bổ
        tdbg3.Columns("DebitAccountID").Caption = rL3("TK_no") 'TK Nợ
        tdbg3.Columns("PercentAmount").Caption = rL3("Ty_le") 'Tỷ lệ
        '================================================================ 
        lblManagementObjTypeID.Text = rL3("Bo_phan_quan_ly") 'Bộ phận quản lý
        '================================================================ 
        tdbcManagementObjTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcManagementObjTypeID.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcManagementObjID.Columns("ObjectTypeID").Caption = rL3("Loai_DT") 'Loại ĐT
        tdbcManagementObjID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcManagementObjID.Columns("ObjectName").Caption = rL3("Ten") 'Tên

    End Sub


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0012s
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 03/01/2012 10:32:32
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0012s_tdbg() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbg.RowCount - 1
            sSQL.Append(" Update D02T0012  Set ")
            If L3Bool(tdbg(i, COL_Choose)) Then
                sSQL.Append("  TransMode = 'CP'  ")
            Else
                sSQL.Append("  TransMode = 'GV' ")
            End If
            sSQL.Append(" Where TransactionID = " & SQLString(tdbg(i, COL_TransactionID)) & vbCrLf)

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL = New StringBuilder
        Next
        Return sRet
    End Function

    Private Function SQLUpdateD02T0012s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbg2.RowCount - 1
            sSQL.Append("Update D02T0012 Set ")
            sSQL.Append("Status = 1,")
            sSQL.Append("AssetID = " & SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA) 'varchar[20], NULL
            If tdbg2(i, COL2_DebitAccountID).ToString = ReturnValueC1Combo(tdbcDepAccountID).ToString And tdbg2(i, COL2_CreditAccountID).ToString = ReturnValueC1Combo(tdbcAssetAccountID).ToString Then
                sSQL.Append("TransactionTypeID = " & SQLString("SDMM") & COMMA) 'varchar[20], NULL
            Else
                sSQL.Append("TransactionTypeID = " & SQLString("MM") & COMMA) 'varchar[20], NULL
            End If
            If L3Bool(tdbg2(i, COL2_Choose)) Then
                sSQL.Append("TransMode = " & SQLString("CP") & COMMA) 'varchar[20], NOT NULL
            Else
                sSQL.Append("TransMode = " & SQLString("GV") & COMMA) 'varchar[20], NOT NULL
            End If
            sSQL.Append("SourceID = " & SQLString(tdbg2(i, COL2_SourceID)) & COMMA) 'varchar[20], NULL
            sSQL.Append(vbCrLf)
            sSQL.Append("Str01U = " & SQLStringUnicode(tdbg2(i, COL2_Str01), gbUnicode, True) & COMMA) 'varchar[1000], NULL
            sSQL.Append("Str02U = " & SQLStringUnicode(tdbg2(i, COL2_Str02), gbUnicode, True) & COMMA) 'varchar[1000], NULL
            sSQL.Append("Str03U = " & SQLStringUnicode(tdbg2(i, COL2_Str03), gbUnicode, True) & COMMA) 'varchar[1000], NULL
            sSQL.Append("Str04U = " & SQLStringUnicode(tdbg2(i, COL2_Str04), gbUnicode, True) & COMMA) 'varchar[1000], NULL
            sSQL.Append("Str05U = " & SQLStringUnicode(tdbg2(i, COL2_Str05), gbUnicode, True) & COMMA) 'varchar[1000], NULL
            sSQL.Append("Num01 = " & SQLMoney(tdbg2(i, COL2_Num01), DxxFormat.DefaultNumber2) & COMMA) 'money, NULL
            sSQL.Append("Num02 = " & SQLMoney(tdbg2(i, COL2_Num02), DxxFormat.DefaultNumber2) & COMMA) 'money, NULL
            sSQL.Append("Num03 = " & SQLMoney(tdbg2(i, COL2_Num03), DxxFormat.DefaultNumber2) & COMMA) 'money, NULL
            sSQL.Append("Num04 = " & SQLMoney(tdbg2(i, COL2_Num04), DxxFormat.DefaultNumber2) & COMMA) 'money, NULL
            sSQL.Append("Num05 = " & SQLMoney(tdbg2(i, COL2_Num05), DxxFormat.DefaultNumber2) & COMMA) 'money, NULL
            sSQL.Append("Date01 = " & SQLDateSave(tdbg2(i, COL2_Date01)) & COMMA) 'datetime, NULL
            sSQL.Append("Date02 = " & SQLDateSave(tdbg2(i, COL2_Date02)) & COMMA) 'datetime, NULL
            sSQL.Append("Date03 = " & SQLDateSave(tdbg2(i, COL2_Date03)) & COMMA) 'datetime, NULL
            sSQL.Append("Date04 = " & SQLDateSave(tdbg2(i, COL2_Date04)) & COMMA) 'datetime, NULL
            sSQL.Append("Date05 = " & SQLDateSave(tdbg2(i, COL2_Date05)) & COMMA) 'datetime, NULL
            sSQL.Append("IsNotAllocate = " & SQLNumber(tdbg2(i, COL2_IsNotAllocate))) 'datetime, NULL

            sSQL.Append(" Where ")
            sSQL.Append("TransactionID = " & SQLString(tdbg2(i, COL2_TransactionID)) & " And ")
            sSQL.Append("ModuleID = " & SQLString(tdbg2(i, COL2_ModuleID)))

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0100
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 28/11/2011 03:31:12
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0100() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0100 Set Status = 2 ")
        sSQL.Append("Where A.CipID in (" & arrCipID.ToString & ")")
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0001
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 28/11/2011 03:32:23
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0001 Set ")
        sSQL.Append("AssetAccountID = " & SQLString(ReturnValueC1Combo(tdbcAssetAccountID)) & COMMA) 'varchar[20], NULL
        sSQL.Append("DepAccountID = " & SQLString(ReturnValueC1Combo(tdbcDepAccountID)) & COMMA) 'varchar[20], NULL
        sSQL.Append("ObjectTypeID = " & SQLString(ReturnValueC1Combo(tdbcObjectTypeID)) & COMMA) 'varchar[20], NULL
        sSQL.Append("ObjectID = " & SQLString(ReturnValueC1Combo(tdbcObjectID)) & COMMA) 'varchar[20], NULL
        sSQL.Append("EmployeeID = " & SQLString(ReturnValueC1Combo(tdbcEmployeeID)) & COMMA) 'varchar[20], NULL
        sSQL.Append(vbCrLf)
        sSQL.Append("FullNameU = " & SQLStringUnicode(txtEmployeeName.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("ConvertedAmount = " & SQLNumber(txtConvertedAmount.Text) & COMMA) 'int, NULL
        sSQL.Append("ServiceLife = " & SQLNumber(txtServiceLife.Text) & COMMA) 'int, NULL
        sSQL.Append("Percentage = " & SQLMoney(txtPercentage.Text, DxxFormat.DefaultNumber2) & COMMA) 'money, NULL
        sSQL.Append("AmountDepreciation = " & SQLMoney(txtAmountDepreciation.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append(vbCrLf)
        sSQL.Append("AssetDate = " & SQLDateSave(c1dateAssetDate.Text) & COMMA) 'datetime, NULL
        '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        sSQL.Append("DepDate = " & SQLDateSave(c1dateDepDate.Text) & COMMA) 'datetime, NULL

        sSQL.Append("UseDate = " & SQLDateSave(c1dateUseDate.Text) & COMMA) 'datetime, NULL
        sSQL.Append("DepreciatedAmount = " & SQLMoney(txtDepreciateAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("RemainAmount = " & SQLMoney(txtRemainAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("DepreciatedPeriod = " & SQLNumber(txtDepreciatedPeriod.Text) & COMMA) 'int, NULL
        sSQL.Append(vbCrLf)
        sSQL.Append("UseMonth = " & SQLNumber(Strings.Left(c1dateBeginUsing.Text, 2)) & COMMA) 'tinyint, NULL
        sSQL.Append("UseYear = " & SQLNumber(Strings.Right(c1dateBeginUsing.Text, 4)) & COMMA) 'smallint, NULL
        sSQL.Append("TranMonth = " & SQLNumber(giTranMonth) & COMMA) 'tinyint, NULL
        sSQL.Append("TranYear = " & SQLNumber(giTranYear) & COMMA) 'smallint, NULL
        sSQL.Append("DepMonth = " & SQLNumber(Strings.Left(c1dateBeginDep.Text, 2)) & COMMA) 'tinyint, NULL
        sSQL.Append("DepYear = " & SQLNumber(Strings.Right(c1dateBeginDep.Text, 4)) & COMMA) 'smallint, NULL
        sSQL.Append(vbCrLf)
        sSQL.Append("IsCompleted = 1,") 'IsRevalued =0, IsDisposed = 0,
        sSQL.Append("SetUpFrom = " & SQLString(_setupFrom) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("LastModifyDate = GetDate()")  'datetime, NOT NULL
        If tdbg2.RowCount > 0 Then
            sSQL.Append(COMMA & "D54ProjectID = " & SQLString(tdbg2(0, COL_ProjectID)) & COMMA) 'varchar[50], NOT NULL
            sSQL.Append("D27PropertyProductID = " & SQLString(tdbg2(0, COL_PropertyProductID))) 'varchar[50], NOT NULL
        End If
        'ID 87726 07.06.2016
        sSQL.Append(COMMA & "D54TaskID  = " & SQLString(tdbg2(0, COL2_TaskID)))  'TaskID, NOT NULL
        '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        sSQL.Append(COMMA & "IsCalPeriodByRate = " & SQLNumber(D02Systems.IsCalPeriodByRate) & vbCrLf) 'int, NULL
        If ReturnValueC1Combo(tdbcManagementObjID) <> "" Then
            sSQL.Append(COMMA & "ManagementObjTypeID = " & SQLString(ReturnValueC1Combo(tdbcManagementObjTypeID)))
            sSQL.Append(COMMA & "ManagementObjID = " & SQLString(ReturnValueC1Combo(tdbcManagementObjID)))
        End If
        sSQL.Append(" Where ")
        sSQL.Append("AssetID = " & SQLString(ReturnValueC1Combo(tdbcAssetID)))

        Return sSQL
    End Function



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5000
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 28/11/2011 03:54:52
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T5000(ByVal sHistoryID As String, ByVal iBeginMonth As Object, ByVal iBeginYear As Object,
        ByVal sHistoryTypeID As Object, Optional ByVal ObjectTypeID As Object = "", Optional ByVal ObjectID As Object = "",
        Optional ByVal EmployeeID As Object = "", Optional ByVal FullName As Object = "",
        Optional ByVal IsStopDepreciation As Object = "", Optional ByVal IsStopUse As Object = "", Optional ByVal ServiceLife As Object = "",
        Optional ByVal PercentAmount As Object = "", Optional ByVal AssignmentID As Object = "",
        Optional ByVal mode As Integer = 0, Optional AssetAccountID As String = "", Optional DepAccountID As String = "") As StringBuilder


        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T5000(")
        sSQL.Append("HistoryID, DivisionID, AssetID, BatchID,  ")
        sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, HistoryTypeID, Status, InstanceID,IsLiquidated,")
        sSQL.Append("ObjectTypeID, ObjectID, EmployeeID, FullNameU,  ")
        sSQL.Append(" IsStopDepreciation, IsStopUse, ServiceLife,PercentAmount, AssignmentID, ")
        sSQL.Append(" CreateDate, CreateUserID, LastModifyUserID, LastModifyDate, AssetAccountID, DepAccountID, ManagementObjTypeID, ManagementObjID, IsManagement")
        'OperateID,SourceID,ConvertedAmount,  
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(sHistoryID) & COMMA) 'HistoryID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA) 'AssetID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA) 'BatchID, varchar[20], NOT NULL
        sSQL.Append(SQLNumber(iBeginMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(iBeginYear) & COMMA) 'BeginYear, smallint, NOT NULL
        sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, smallint, NOT NULL
        sSQL.Append(SQLString(sHistoryTypeID) & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
        sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NULL
        sSQL.Append(SQLNumber(0) & COMMA) 'Giá trị default 'InstanceID, tinyint, NULL
        sSQL.Append(SQLNumber(0) & COMMA) 'Giá trị default =0'IsLiquidated, tinyint, NOT NULL
        If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" AndAlso L3String(sHistoryTypeID) = "OB" AndAlso bCheckIsManagement = True Then
            sSQL.Append(SQLString("") & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString("") & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString("") & COMMA) 'EmployeeID, varchar[20], NULL
            sSQL.Append(SQLStringUnicode("") & COMMA) 'EmployeeID, varchar[20], NULL
        Else
            sSQL.Append(SQLString(ObjectTypeID) & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(ObjectID) & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString(EmployeeID) & COMMA) 'EmployeeID, varchar[20], NULL
            sSQL.Append(SQLStringUnicode(FullName, gbUnicode, True) & COMMA) 'FullNameU, nvarchar, NOT NULL
        End If
        sSQL.Append(IIf(IsStopDepreciation.ToString = "", "NULL", SQLNumber(IsStopDepreciation)).ToString & COMMA) 'IsStopDepreciation, tinyint, NULL
        sSQL.Append(IIf(IsStopUse.ToString = "", "NULL", SQLNumber(IsStopUse)).ToString & COMMA) 'IsStopUse, tinyint, NULL
        sSQL.Append(IIf(ServiceLife.ToString = "", "NULL", SQLNumber(ServiceLife)).ToString & COMMA) 'ServiceLife, int, NULL
        sSQL.Append(IIf(PercentAmount.ToString = "", "NULL", SQLMoney(Number(PercentAmount) * 100, DxxFormat.DefaultNumber2)).ToString & COMMA) 'PercentAmount, money, NULL
        sSQL.Append(SQLString(AssignmentID) & COMMA) 'AssignmentID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NOT NULL
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        sSQL.Append(SQLString(AssetAccountID) & COMMA) 'AssetAccountID, varchar[20], NULL
        sSQL.Append(SQLString(DepAccountID) & COMMA) 'DepAccountID, varchar[20], NULL
        'sSQL.Append(SQLString(?????) & COMMA) 'OperateID, varchar[20], NULL
        'sSQL.Append(SQLString(?????) & COMMA) 'SourceID, varchar[20], NULL
        'sSQL.Append(SQLMoney(?????, DxxFormat.?????) & COMMA) 'ConvertedAmount, money, NULL
        'sSQL.Append(SQLString(?????) & COMMA) 'GroupID, varchar[20], NOT NULL
        'sSQL.Append(SQLString(?????) & COMMA) 'AssetWHID, varchar[50], NOT NULL
        If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" AndAlso L3String(sHistoryTypeID) = "OB" AndAlso bCheckIsManagement = True Then
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcManagementObjTypeID)) & COMMA)
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcManagementObjID)) & COMMA)
            sSQL.Append(SQLNumber(1))
            bCheckIsManagement = False
        Else
            sSQL.Append(SQLString("") & COMMA)
            sSQL.Append(SQLString("") & COMMA)
            sSQL.Append(SQLNumber(0))
        End If
        sSQL.Append(") " & vbCrLf)

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5010
    '# Created User: 
    '# Created Date: 17/11/2021 05:02:02
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T5010(sHistoryID As String, sHistoryTypeID As String, sAccountID As String) As StringBuilder
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Luu " & vbCrLf)
        sSQL.Append("Insert Into D02T5010(")
        sSQL.Append("HistoryID, DivisionID, BatchID, HistoryTypeID, AssetID, " & vbCrLf)
        sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, GroupID, " & vbCrLf)
        sSQL.Append("AccountID")
        sSQL.Append(") Values(" & vbCrLf)
        sSQL.Append(SQLString(sHistoryID) & COMMA) 'HistoryID, varchar[50], NOT NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[50], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA) 'BatchID, varchar[50], NOT NULL
        sSQL.Append(SQLString(sHistoryTypeID) & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA & vbCrLf) 'AssetID, varchar[50], NOT NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'BeginYear, int, NOT NULL
        sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, int, NOT NULL
        sSQL.Append(SQLString("") & COMMA & vbCrLf) 'GroupID, varchar[50], NOT NULL
        sSQL.Append(SQLString(sAccountID)) 'AccountID, varchar[50], NOT NULL
        sSQL.Append(") " & vbCrLf)

        Return sSQL
    End Function


    Private Function SQLInsertD02T5000s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim iCount As Long = tdbg3.RowCount
        Dim iFirstHis As Long = 0
        Dim sHistoryID As String = ""
        Dim sSQL As New StringBuilder

        For i As Integer = 0 To tdbg3.RowCount - 1
            sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HB", gsStringKey, sHistoryID, iCount, iFirstHis)
            tdbg3(i, COL3_HistoryID) = sHistoryID

            sRet.Append(SQLInsertD02T5000(sHistoryID, giTranMonth, giTranYear, "AS", , , , , , , , tdbg3(i, COL3_PercentAmount), tdbg3(i, COL3_AssignmentID)).ToString & vbCrLf)
            sSQL = New StringBuilder
        Next
        Return sRet
    End Function

    Private Function SQLInsertD02T5000_3() As StringBuilder
        Dim sRet As New StringBuilder
        Dim iCount As Long = 3
        Dim iFirstHis As Long = 0
        Dim sHistoryID As String = ""
        Dim sSQL As New StringBuilder

        sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HB", gsStringKey, sHistoryID, iCount, iFirstHis)
        sRet.Append(SQLInsertD02T5000(sHistoryID, Strings.Left(c1dateBeginDep.Text, 2), Strings.Right(c1dateBeginDep.Text, 4), "SD", , , , , 0).ToString & vbCrLf)

        sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HB", gsStringKey, sHistoryID, iCount, iFirstHis)
        sRet.Append(SQLInsertD02T5000(sHistoryID, Strings.Left(c1dateBeginUsing.Text, 2), Strings.Right(c1dateBeginUsing.Text, 4), "SU", , , , , , 0).ToString & vbCrLf)

        sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HB", gsStringKey, sHistoryID, iCount, iFirstHis)
        sRet.Append(SQLInsertD02T5000(sHistoryID, Strings.Left(c1dateBeginUsing.Text, 2), Strings.Right(c1dateBeginUsing.Text, 4), "SL", , , , , , , txtServiceLife.Text).ToString & vbCrLf)
        Return sRet
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD95P0105
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 29/11/2011 07:59:42
    '# Modified User: 
    '# Modified Date: 
    '# Description: Kiểm tra trước khi lưu
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD95P0105(ByVal Serial As Object, ByVal InvoiceNum As Object) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D95P0105 "
        sSQL &= SQLString(Serial) & COMMA 'Serial, varchar[20], NOT NULL
        sSQL &= SQLString(InvoiceNum) 'InvoiceNum, varchar[20], NOT NULL
        Return sSQL
    End Function

    '    '#---------------------------------------------------------------------------------------------------
    '    '# Title: SQLInsertD02T0012s
    '    '# Created User: Nguyễn Thị Ánh
    '    '# Created Date: 29/11/2011 08:01:46
    '    '# Modified User: 
    '    '# Modified Date: 
    '    '# Description: 
    '    '#---------------------------------------------------------------------------------------------------
    '    Dim dTotalConvertedAmount As Double = 0
    '    Private Function SQLInsertD02T0012s() As StringBuilder
    '        Dim sRet As New StringBuilder
    '        Dim iCount As Long = tdbg1.RowCount - 1
    '        Dim iFirstTran As Long = 0
    '        Dim sTransactionID As String = ""
    '        Dim sSQL As New StringBuilder

    '        For i As Integer = 0 To tdbg1.RowCount - 1
    '            If iUseInvoiceCodeD02 = 0 Then GoTo 1
    '            If tdbg1(i, COL1_SerialNo).ToString <> "" And tdbg1(i, COL1_RefNo).ToString <> "" And IsNumeric(tdbg1(i, COL1_SerialNo)) Then
    '                Dim strSQL As String = SQLStoreD95P0105(tdbg1(i, COL1_SerialNo), tdbg1(i, COL1_RefNo))
    '                Dim dtTemp As DataTable = ReturnDataTable(strSQL)
    '                If dtTemp.Rows.Count = 0 Then GoTo 1
    '                If dtTemp.Rows(0).Item("Status").ToString <> "0" Then GoTo 1

    '                Dim f As New D95M0240
    '                f.ID01 = tdbg1(i, COL1_SerialNo).ToString
    '                f.ID02 = tdbg1(i, COL1_RefNo).ToString
    '                f.FormActive = "D95F0131"
    '                f.ShowDialog()
    '                Dim bClose As Boolean = f.bClose
    '                f.Dispose()

    '                If Not bClose Then GoTo 1
    '                dtTemp = ReturnDataTable(strSQL)
    '                If dtTemp.Rows.Count = 0 Then GoTo 1 '"BÁn ch§a ¢Ünh nghÚa mÉu hâa ¢¥n cho Sç hâa ¢¥n nªy!"
    '                If dtTemp.Rows(0).Item("Status").ToString = "0" Then D99C0008.MsgL3("Bạn chưa định nghĩa mẫu hóa đơn cho Số hóa đơn này!")
    '            End If
    '1:
    '            sTransactionID = CreateIGENewS("D02T0012", "TransactionID", "02", "TH", gsStringKey, sTransactionID, iCount, iFirstTran)
    '            tdbg1(i, COL1_TransactionID) = sTransactionID

    '            sSQL.Append("Insert Into D02T0012(")
    '            sSQL.Append("TransactionID, DivisionID, ModuleID, SplitNo, AssetID, ")
    '            sSQL.Append("VoucherTypeID, VoucherNo, VoucherDate, TranMonth, TranYear, ")
    '            sSQL.Append("TransactionDate, Description, CurrencyID, ExchangeRate, DebitAccountID, ")
    '            sSQL.Append("CreditAccountID, OriginalAmount, ConvertedAmount, Status, TransactionTypeID, ")
    '            sSQL.Append("RefNo, RefDate, Disabled, CreateUserID, CreateDate, ")
    '            sSQL.Append("LastModifyUserID, LastModifyDate, SeriNo, ObjectTypeID, ObjectID, BatchID,")
    '            sSQL.Append("Ana01ID, Ana02ID, Ana03ID,Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID,")
    '            sSQL.Append(" CipID, Notes,  SourceID,")
    '            sSQL.Append("Str01, Str02, Str03, Str04, Str05, Num01, Num02, Num03, Num04, Num05, Date01, Date02, Date03, Date04, Date05,")
    '            sSQL.Append("DescriptionU,  NotesU, Str01U, Str02U, Str03U, Str04U, Str05U") 'ObjectNameU,ItemNameU
    '            sSQL.Append(") Values(")
    '            sSQL.Append(SQLString(tdbg1(i, COL1_TransactionID)) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
    '            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
    '            sSQL.Append(SQLString("02") & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL
    '            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA) 'AssetID, varchar[20], NULL
    '            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcVoucherTypeID)) & COMMA) 'VoucherTypeID, varchar[20], NULL
    '            sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[50], NULL
    '            sSQL.Append(SQLDateSave(c1dateVoucherDate.Text) & COMMA) 'VoucherDate, datetime, NULL
    '            sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
    '            sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Description), gbUnicode, False) & COMMA) 'Description, varchar[500], NULL
    '            sSQL.Append(SQLString(txtCurrenyID.Text) & COMMA) 'CurrencyID, varchar[20], NOT NULL
    '            sSQL.Append(SQLMoney(txtExchangRate.Text, DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_DebitAccountID)) & COMMA) 'DebitAccountID, varchar[20], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_CreditAccountID)) & COMMA) 'CreditAccountID, varchar[20], NULL
    '            sSQL.Append(SQLMoney(tdbg1(i, COL1_OriginalAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'OriginalAmount, money, NULL
    '            sSQL.Append(SQLMoney(tdbg1(i, COL1_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL
    '            dTotalConvertedAmount += Number(tdbg1(i, COL1_ConvertedAmount))
    '            sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NOT NULL
    '            sSQL.Append(SQLString("XDCB") & COMMA) 'TransactionTypeID, varchar[20], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_RefNo)) & COMMA) 'RefNo, varchar[50], NULL
    '            sSQL.Append(SQLDateSave(tdbg1(i, COL1_RefDate)) & COMMA) 'RefDate, datetime, NULL
    '            sSQL.Append(SQLNumber(0) & COMMA) 'Disabled, bit, NOT NULL
    '            sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
    '            sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
    '            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
    '            sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_SerialNo)) & COMMA) 'SeriNo, varchar[20], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_ObjectTypeID)) & COMMA) 'ObjectTypeID, varchar[20], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_ObjectID)) & COMMA) 'ObjectID, varchar[20], NULL
    '            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA) 'BatchID, varchar[20], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana01ID)) & COMMA) 'Ana01ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana02ID)) & COMMA) 'Ana02ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana03ID)) & COMMA) 'Ana03ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana04ID)) & COMMA) 'Ana04ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana05ID)) & COMMA) 'Ana05ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana06ID)) & COMMA) 'Ana06ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana07ID)) & COMMA) 'Ana07ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana08ID)) & COMMA) 'Ana08ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana09ID)) & COMMA) 'Ana09ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_Ana10ID)) & COMMA) 'Ana10ID, varchar[50], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_CipID)) & COMMA) 'CipID, varchar[20], NULL
    '            sSQL.Append(SQLStringUnicode(txtNotes.Text, gbUnicode, False) & COMMA) 'Notes, varchar[250], NULL
    '            sSQL.Append(SQLString(tdbg1(i, COL1_SourceID)) & COMMA) 'SourceID, varchar[20], NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str01), gbUnicode, False) & COMMA) 'Str01, varchar[1000], NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str02), gbUnicode, False) & COMMA) 'Str02, varchar[1000], NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str03), gbUnicode, False) & COMMA) 'Str03, varchar[1000], NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str04), gbUnicode, False) & COMMA) 'Str04, varchar[1000], NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str05), gbUnicode, False) & COMMA) 'Str05, varchar[1000], NULL
    '            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num01), DxxFormat.DefaultNumber2) & COMMA) 'Num01, money, NULL
    '            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num02), DxxFormat.DefaultNumber2) & COMMA) 'Num02, money, NULL
    '            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num03), DxxFormat.DefaultNumber2) & COMMA) 'Num03, money, NULL
    '            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num04), DxxFormat.DefaultNumber2) & COMMA) 'Num04, money, NULL
    '            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num05), DxxFormat.DefaultNumber2) & COMMA) 'Num05, money, NULL
    '            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date01)) & COMMA) 'Date01, datetime, NULL
    '            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date02)) & COMMA) 'Date02, datetime, NULL
    '            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date03)) & COMMA) 'Date03, datetime, NULL
    '            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date04)) & COMMA) 'Date04, datetime, NULL
    '            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date05)) & COMMA) 'Date05, datetime, NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Description), gbUnicode, True) & COMMA) 'DescriptionU, nvarchar, NOT NULL
    '            sSQL.Append(SQLStringUnicode(txtNotes.Text, gbUnicode, True) & COMMA) 'NotesU, nvarchar, NOT NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str01), gbUnicode, True) & COMMA) 'Str01U, nvarchar, NOT NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str02), gbUnicode, True) & COMMA) 'Str02U, nvarchar, NOT NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str03), gbUnicode, True) & COMMA) 'Str03U, nvarchar, NOT NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str04), gbUnicode, True) & COMMA) 'Str04U, nvarchar, NOT NULL
    '            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str05), gbUnicode, True)) 'Str05U, nvarchar, NOT NULL

    '            sSQL.Append(")")

    '            sRet.Append(sSQL.ToString & vbCrLf)
    '            sSQL.Remove(0, sSQL.Length)
    '        Next
    '        Return sRet
    '    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0012
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 29/11/2011 10:08:10
    '# Modified User: 
    '# Modified Date: 
    '# Description: Lưu Sửa
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0012() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T0012"
        sSQL &= " Where "
        sSQL &= "AssetID = " & SQLString(_assetID) & " And "
        sSQL &= "DivisionID = " & SQLString(gsDivisionID) & " And "
        sSQL &= "ModuleID = '02'" & " And "
        sSQL &= "IsNull(TransactionTypeID,'') in ('XDCB')"
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T5000
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 29/11/2011 10:09:03
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T5000() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T5000"
        sSQL &= " Where "
        sSQL &= "AssetID = " & SQLString(_assetID) & " And DivisionID = " & SQLString(gsDivisionID)
        sSQL &= " AND HistoryTypeID in ('OB' , 'SC' , 'AS', 'AAC', 'DAC' ) "
        Return sSQL
    End Function

    'Dim iUseInvoiceCodeD02 As Integer = 0
    'Private Function GetUseInvoiceCodeD02() As Integer
    '    Dim strSQL As String = "SELECT UseInvoiceCodeD02 From D95T0000 "
    '    Dim dtTemp As DataTable = ReturnDataTable(strSQL)
    '    If dtTemp.Rows.Count = 0 Then Exit Function
    '    Return L3Int(dtTemp.Rows(0).Item("UseInvoiceCodeD02"))
    'End Function

#Region "Active Find - List All (Client)"
    Dim dtCaptionCols As DataTable

    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""

    Private Sub mnsFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsFind.Click
        gbEnabledUseFind = True
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        'Những cột bắt buộc nhập
        Dim Arr As New ArrayList
        For i As Integer = 0 To tdbg.Splits.Count - 1
            AddColVisible(tdbg, i, Arr, , False, False, gbUnicode)
        Next
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr, New Integer() {COL_OriginalAmount, COL_ConvertedAmount})
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me.Name, "0", gbUnicode)
    End Sub

    Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
        If ResultWhereClause Is Nothing Or ResultWhereClause.ToString = "" Then Exit Sub
        sFind = ResultWhereClause.ToString()
        ReLoadTDBGrid1()
    End Sub

    Private Sub mnsListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter1, bRefreshFilter)
        ReLoadTDBGrid1()
    End Sub

    Private Sub ReLoadTDBGrid1()
        Dim strFind As String = sFind
        If sFilter1.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter1.ToString
        dtGrid1.DefaultView.RowFilter = strFind
        ResetGrid1()
    End Sub
#End Region

    Private Sub btnChoose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChoose.Click
        If tdbg.RowCount = 0 Then Exit Sub
        tdbg.UpdateData()
        Dim dr() As DataRow = dtGrid1.Select("Choose" & "=1")
        If dr.Length = 0 Then
            D99C0008.MsgNotYetChoose(rL3("Phieu"))
            tdbg.Focus()
            tdbg.SplitIndex = 0
            tdbg.Row = 0
            tdbg.Col = COL_Choose
            Exit Sub
        End If
        For i As Integer = dr.Length - 1 To 0 Step -1
            dr(i).Item("Choose") = 0
            dtGrid2.ImportRow(dr(i))
            dtGrid1.Rows.Remove(dr(i))
        Next
        ResetGrid1()
        ResetGrid2()
        tabMain.SelectedTab = tabPage2
    End Sub

    Public Function CallTotalAmount(Optional ByRef dDepreciatedAmount As Double = 0) As Double
        Dim dConvertAmount As Double = 0
        'Dim dDepreciatedAmount As Double = 0
        For i As Integer = 0 To tdbg2.RowCount - 1
            'Tính Nguyên tệ
            If tdbg2(i, COL2_DebitAccountID).ToString = ReturnValueC1Combo(tdbcAssetAccountID).ToString Then
                dConvertAmount += Number(tdbg2(i, COL2_ConvertedAmount))
                'ElseIf tdbg2(i, COL2_CreditAccountID).ToString = ReturnValueC1Combo(tdbcAssetAccount).ToString Then
                '    dConvertAmount -= Number(tdbg2(i, COL2_ConvertedAmount))
            End If
            'Tính khấu hao
            If tdbg2(i, COL2_DebitAccountID).ToString = ReturnValueC1Combo(tdbcDepAccountID).ToString Then
                dDepreciatedAmount -= Number(tdbg2(i, COL2_ConvertedAmount))
            ElseIf tdbg2(i, COL2_CreditAccountID).ToString = ReturnValueC1Combo(tdbcDepAccountID).ToString Then
                dDepreciatedAmount += Number(tdbg2(i, COL2_ConvertedAmount))
            End If
            dDepreciatedAmount = Math.Abs(dDepreciatedAmount)
        Next

        'txtConvertedAmount.Text = Format(dConvertAmount, DxxFormat.D90_ConvertedDecimals)
        'txtAmountDepreciation.Text = Format(dDepreciatedAmount, DxxFormat.D90_ConvertedDecimals)
        Return dConvertAmount
    End Function

    Private Sub tdbg2_OnAddNew(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbg2.OnAddNew
        tdbg2.Columns(COL2_IsNotAllocate).Value = 0
    End Sub


#Region "Events tdbcPropertyProductID load tdbdCipID"

    Private Sub tdbcPropertyProductID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcPropertyProductID.SelectedValueChanged
        If tdbcPropertyProductID.SelectedValue Is Nothing OrElse tdbcPropertyProductID.Text = "" Then
            tdbcPropertyProductID.Text = ""
        End If
        'ID 87726 7.06.2016
        If tdbcPropertyProductID.FindStringExact(tdbcPropertyProductID.Text) = -1 Then
            tdbcPropertyProductID.Text = ""
        End If

        If _FormState <> EnumFormState.FormView AndAlso tdbg2.RowCount > 0 AndAlso sOldPropertyProductID <> tdbcPropertyProductID.Text Then
            If D99C0008.MsgAsk(rL3("Du_lieu_tren_luoi_se_bi_xoa_Ban_co_muon_thuc_hien_ko")) = Windows.Forms.DialogResult.Yes Then
                LoadTDBGrid1()
                dtGrid2.Clear()
            Else
                tdbcPropertyProductID.SelectedValue = sOldPropertyProductID
            End If
        Else
            LoadTDBGrid1()
        End If
        sOldPropertyProductID = tdbcPropertyProductID.Text
        tabMain.SelectedTab = tabPage1
    End Sub


    'ID 87726 7.06.2016
    'Private Sub tdbcPropertyProductID_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles tdbcPropertyProductID.Validating
    '    If tdbcPropertyProductID.FindStringExact(tdbcPropertyProductID.Text) = -1 Then
    '        tdbcPropertyProductID.Text = ""
    '    End If

    '    If _FormState <> EnumFormState.FormView AndAlso tdbg2.RowCount > 0 AndAlso sOldPropertyProductID <> tdbcPropertyProductID.Text Then
    '        If D99C0008.MsgAsk(rL3("Du_lieu_tren_luoi_se_bi_xoa_Ban_co_muon_thuc_hien_ko")) = Windows.Forms.DialogResult.Yes Then
    '            LoadTDBGrid1()
    '            dtGrid2.Clear()
    '        Else
    '            tdbcPropertyProductID.SelectedValue = sOldPropertyProductID
    '        End If
    '    Else
    '        LoadTDBGrid1()
    '    End If
    '    sOldPropertyProductID = tdbcPropertyProductID.Text
    '    tabMain.SelectedTab = tabPage1
    'End Sub
#End Region

    'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
#Region "Gọi đến D02F0101"
    Private Sub ShowD02F0101(ByVal iCol As String)
        If iPerD02F0100 < 2 Then
            D99C0008.MsgL3(rL3("Ban_khong_co_quyen_them_moi_phan_bo_khau_hao")) 'BÁn kh¤ng câ quyÒn th£m mìi
            Exit Sub
        End If

        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "FormIDPermission", "D02F0101")
        Dim frm As Form = CallFormShowDialog("D02D1240", "D02F0101", arrPro)
        If L3Bool(GetProperties(frm, "SavedOK")) Then
            LoadtdbdAssignmentID()
            tdbg3.Columns(iCol).Text = L3String(GetProperties(frm, "AssignmentID"))
        End If
    End Sub
#End Region
    'ID : 224617 - BỔ SUNG Cho phép gọi màn hình THIẾT LẬP DANH MỤC TÀI SẢN CỐ ĐỊNH tại bước hình thành TS
#Region "Gọi đến D02F0087"
    Dim bFormD02F0087 As Boolean = False


    Private Sub ShowD02F0087()


        Dim sMethodID As String = ""
        Dim sSQLD91T1001_SaveLastKey As String = ""
        If tdbcAssetID.Text = "+" Then
            If iPerD02F0087 < 2 Then
                D99C0008.MsgL3(rL3("Ban_khong_co_quyen_tao_ma_tu_dong")) 'BÁn kh¤ng câ quyÒn th£m mìi
                Exit Sub
            End If


            'If D02Systems.AssetAuto = 2 And D02Systems.IsShowFormAutoCreate = True Then
            '    ' Gọi đến màn hình D02F0087 khi Nhấn button Chọn ở màn hình D02F0087 sẽ gọi tới màn hình D02F1031
            '    Dim arrPro() As StructureProperties = Nothing
            '    SetProperties(arrPro, "FormIDPermission", "D02F0087")
            '    Dim frm As Form = CallFormShowDialog("D02D1040", "D02F0087", arrPro)
            '    AssetID = GetProperties(frm, "sAssetID").ToString ' Tạo sAssetID để nhận AssetID được trả về từ Form gọi
            '    sMethodID = GetProperties(frm, "sMethodID").ToString
            '    sSQLD91T1001_SaveLastKey = GetProperties(frm, "sSQLD91T1001_SaveLastKey").ToString
            '    bFormD02F0087 = True
            'End If
            If D02Systems.AssetAuto = 2 And D02Systems.IsShowFormAutoCreate = True Then
                Dim arrPro() As StructureProperties = Nothing
                SetProperties(arrPro, "FormIDPermission", "D02F0087")
                Dim frm As Form = CallFormShowDialog("D02D1040", "D02F0087", arrPro)
                AssetID = GetProperties(frm, "sAssetID").ToString ' Tạo sAssetID để nhận AssetID được trả về từ Form gọi
                sMethodID = GetProperties(frm, "sMethodID").ToString
                sSQLD91T1001_SaveLastKey = GetProperties(frm, "sSQLD91T1001_SaveLastKey").ToString
                bFormD02F0087 = True

            End If
            ' Lấy giá trị từ màn hình D02F0087 để gọi tới màn hình D02F1031
            Dim arr() As StructureProperties = Nothing
            SetProperties(arr, "bFormD02F0087", bFormD02F0087)
            SetProperties(arr, "sAssetID", AssetID)
            SetProperties(arr, "sMethodID", sMethodID)
            SetProperties(arr, "sSQLD91T1001_SaveLastKey", sSQLD91T1001_SaveLastKey)
            SetProperties(arr, "FormState", EnumFormState.FormAdd)
            Dim frm2 As Form = CallFormShowDialog("D02D1040", "D02F1031", arr)
            If L3Bool(GetProperties(frm2, "SavedOK")) Then
                'Khi lưu thành công Load lại combo và hiển thị mã mới được thêm lên combo
                LoadtdbcAssetID()
                tdbcAssetID.SelectedValue = L3String(GetProperties(frm2, "AssetID_D02F1031"))
            End If
        End If
        AssetID = ""
        sMethodID = ""
        bFormD02F0087 = False

    End Sub

#End Region
End Class
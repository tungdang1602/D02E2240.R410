Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text
Imports System.IO
Imports System

Public Class D02F1002

#Region "Const of tdbg1"
    Private Const COL1_BatchID As Integer = 0            ' BatchID
    Private Const COL1_TransactionID As Integer = 1      ' TransactionID
    Private Const COL1_DataID As Integer = 2             ' DataID
    Private Const COL1_DataName As Integer = 3           ' Dữ liệu
    Private Const COL1_IsNotAllocate As Integer = 4      ' Không tính KH
    Private Const COL1_Description As Integer = 5        ' Diễn giải
    Private Const COL1_ObjectTypeID As Integer = 6       ' Loại đối tượng
    Private Const COL1_ObjectID As Integer = 7           ' Đối tượng
    Private Const COL1_AccountID As Integer = 8          ' Mã tài khoản
    Private Const COL1_ConvertedAmount As Integer = 9    ' Số tiền
    Private Const COL1_SourceID As Integer = 10          ' Nguồn hình thành
    Private Const COL1_Str01 As Integer = 11             ' Văn bản 01
    Private Const COL1_Str02 As Integer = 12             ' Văn bản 02
    Private Const COL1_Str03 As Integer = 13             ' Văn bản 03
    Private Const COL1_Str04 As Integer = 14             ' Văn bản 04
    Private Const COL1_Str05 As Integer = 15             ' Văn bản 05
    Private Const COL1_Num01 As Integer = 16             ' Số 01
    Private Const COL1_Num02 As Integer = 17             ' Số 02
    Private Const COL1_Num03 As Integer = 18             ' Số 03
    Private Const COL1_Num04 As Integer = 19             ' Số 04
    Private Const COL1_Num05 As Integer = 20             ' Số 05
    Private Const COL1_Date01 As Integer = 21            ' Ngày 01
    Private Const COL1_Date02 As Integer = 22            ' Ngày 02
    Private Const COL1_Date03 As Integer = 23            ' Ngày 03
    Private Const COL1_Date04 As Integer = 24            ' Ngày 04
    Private Const COL1_Date05 As Integer = 25            ' Ngày 05
    Private Const COL1_Ana01ID As Integer = 26           ' Khoản mục 01
    Private Const COL1_Ana02ID As Integer = 27           ' Khoản mục 02
    Private Const COL1_Ana03ID As Integer = 28           ' Khoản mục 03
    Private Const COL1_Ana04ID As Integer = 29           ' Khoản mục 04
    Private Const COL1_Ana05ID As Integer = 30           ' Khoản mục 05
    Private Const COL1_Ana06ID As Integer = 31           ' Khoản mục 06
    Private Const COL1_Ana07ID As Integer = 32           ' Khoản mục 07
    Private Const COL1_Ana08ID As Integer = 33           ' Khoản mục 08
    Private Const COL1_Ana09ID As Integer = 34           ' Khoản mục 09
    Private Const COL1_Ana10ID As Integer = 35           ' Khoản mục 10
    Private Const COL1_TransactionTypeID As Integer = 36 ' TransactionTypeID
#End Region


#Region "Const of tdbg2"
    Private Const COL2_AssignmentID As Integer = 0   ' Mã phân bổ
    Private Const COL2_AssignmentName As Integer = 1 ' Tên phân bổ
    Private Const COL2_DebitAccountID As Integer = 2 ' TK Nợ
    Private Const COL2_PercentAmount As Integer = 3  ' Tỷ lệ
    Private Const COL2_Extend As Integer = 4         ' Extend
    Private Const COL2_HistoryID As Integer = 5      ' HistoryID
#End Region

    '---Kiểm tra khoản mục theo chuẩn gồm 7 bước
    '--- Chuẩn Khoản mục b1: Khai báo biến
    '-------Biến khai báo cho khoản mục
    Private Const SplitAna As Int16 = 2 ' Ghi nhận Khoản mục chứa ở Split nào
    Dim bUseAna As Boolean 'Kiểm tra có sử dụng Khoản mục không, để set thuộc tính Enabled nút Khoản mục 
    Dim iDisplayAnaCol As Integer = 0 ' Cột Khoản mục đầu tiên được hiển thị, khi nhấn nút Khoản mục thì Focus đến cột đó
    Dim xCheckAna(9) As Boolean 'Khởi động tại Form_load: Ghi lại việc kiểm tra lần đầu Lưu, khi nhấn Lưu lần thứ 2 thì không cần kiểm tra nữa
    '------------------------------------------------------------------------------------------------
    Dim dtAssignmentID, dtManagementID As DataTable
    Dim sAuditCode As String = "Opening02"
    Dim sCreateUserID As String
    Dim sCreateDate As String
    Dim sGetDate As String
    'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b1:
    Dim sEditVoucherTypeID As String = ""
    Private _assetID As String = ""
    Dim clsFilterCombo As Lemon3.Controls.FilterCombo
    Dim clsFilterDropdown As Lemon3.Controls.FilterDropdown


    Public Property AssetID() As String
        Get
            Return _assetID
        End Get
        Set(ByVal Value As String)
            _assetID = Value
        End Set
    End Property

    Dim dtObjectID, dtAnaCaption As DataTable
    Private _FormState As EnumFormState

    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            _FormState = value
            '--- Chuẩn Khoản mục b2: Lấy caption cho 10 khoản mục
            bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg1, COL1_Ana01ID, SplitAna, , gbUnicode, dtAnaCaption)
            SetNewXaCheckAna()
            '--- Chuẩn Khoản mục b21: D91 có sử dụng Khoản mục
            If bUseAna Then
                iDisplayAnaCol = 1
            Else
                tdbg1.RemoveHorizontalSplit(SplitAna)
            End If

            '------------------------------------
            'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b2:
            LoadTDBCombo()
            LoadTDBDropDown()
            'GetCaptionDescription()
            clsFilterCombo = New Lemon3.Controls.FilterCombo
            clsFilterCombo.CheckD91 = True 'Giá trị mặc định True: kiểm tra theo DxxFormat.LoadFormNotINV. Ngược lại luôn luôn Filter dạng mới (dùng cho Novaland)
            clsFilterCombo.AddPairObject(tdbcObjectTypeID, tdbcObjectID) 'Đã bổ sung cột Loại ĐT
            clsFilterCombo.AddPairObject(tdbcManagementObjTypeID, tdbcManagementObjID)
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
            clsFilterDropdown.UseFilterDropdown(tdbg1, COL1_Ana01ID, COL1_Ana02ID, COL1_Ana03ID, COL1_Ana04ID, COL1_Ana05ID, COL1_Ana06ID, COL1_Ana07ID, COL1_Ana08ID, COL1_Ana09ID, COL1_Ana10ID)
            clsFilterDropdown.UseFilterDropdown(tdbg2, COL2_AssignmentID) 'Nếu dùng nhiều lưới

            txtServiceLife.Tag = ""
            txtPercentage.Tag = ""

            Select Case _FormState
                Case EnumFormState.FormAdd
                    ' LoadTDBCombo()
                    'LoadTDBDropDown()
                    LoadAddNew()
                    btnSave.Enabled = True
                    btnNext.Enabled = False
                Case EnumFormState.FormEdit
                    LoadtdbcEmployeeID()
                    'LoadTDBDropDown()
                    LoadEdit()
                    btnSave.Enabled = True
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                Case EnumFormState.FormEditOther
                    LoadtdbcEmployeeID()
                    'LoadTDBDropDown()
                    LoadEdit()
                    btnSave.Enabled = True
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LockCtrlEditOther()
                Case EnumFormState.FormView
                    'LoadTDBDropDown()

                    LoadEdit()
                    btnSave.Enabled = False
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
            End Select
        End Set
    End Property

    Private Sub LockCtrlEditOther()
        grpAssetID.Enabled = False
        tabMain.Enabled = False
        btnNext.Enabled = False
        chkPosted.Enabled = False
        ReadOnlyAll(grpFinancialInfo, c1dateUseDate, c1dateAssetDate, c1dateDepDate)
    End Sub

    Private Sub SetNewXaCheckAna()
        Dim i As Integer
        For i = 0 To 9
            xCheckAna(i) = False
        Next i
    End Sub

    Private Sub tdbg2_LockedColumns()
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_AssignmentName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    Private Sub D02F1002_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me)
                Exit Sub
            Case Keys.F11
                HotKeyF11(Me, tdbg1, SPLIT0, COL1_DataID)
                Exit Sub
        End Select

        If e.Control And e.KeyCode = Keys.F1 Then
            btnHotKeys_Click(Nothing, Nothing)
            Exit Sub
        End If

        If e.Alt Then
            Select Case e.KeyCode
                Case Keys.D1, Keys.NumPad1
                    tabMain.SelectedTab = tabPage1
                    c1dateVoucherDate.Focus()
                Case Keys.D2, Keys.NumPad2
                    tabMain.SelectedTab = tabPage2
                    tdbg2.Focus()
            End Select
        End If
    End Sub

    Private Sub D02F1002_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave

    End Sub

    Private Sub D02F1002_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        iPer_F5558 = ReturnPermission("D02F5558")
        gbSavedOK = False
        InputbyUnicode(Me, gbUnicode)
        Loadlanguage()
        Dim bUseInfo As Boolean = LoadCaptionSubInfo()
        If Not bUseInfo Then tdbg1.RemoveHorizontalSplit(1)
        SetBackColorObligatory()
        InputDateInTrueDBGrid(tdbg1, COL1_Date01, COL1_Date02, COL1_Date03, COL1_Date04, COL1_Date05)
        tdbg1_LockedColumns()
        tdbg2_LockedColumns()
        'tdbg1_NumberFormat()
        tdbg2_NumberFormat()
        'iLastCol1 = CountCol(tdbg1, SPLIT2)
        'iLastCol2 = CountCol(tdbg2, SPLIT0)
        Dim arr() As FormatColumn = Nothing
        AddNumberColumns(arr, SqlDbType.Money, tdbg1.Columns(COL1_ConvertedAmount).DataField, "N" & DxxFormat.iD90_ConvertedDecimals)
        InputNumber(tdbg1, arr)
        LockServiceLife() '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub tdbg1_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.HeadClick
        If tdbg1.Col = COL1_ConvertedAmount Or tdbg1.Col = COL1_AccountID Then Exit Sub
        If tdbg1.Col = COL1_DataName Then
            CopyColumnsCustom(tdbg1, tdbg1.Col, tdbg1.Row, 2, tdbg1(tdbg1.Row, tdbg1.Col).ToString)
        Else
            CopyColumns(tdbg1, tdbg1.Col, tdbg1(tdbg1.Row, tdbg1.Col).ToString, tdbg1.Row)
        End If
    End Sub

    Private Sub CopyColumnsCustom(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ColCopy As Integer, ByVal RowCopy As Integer, ByVal ColumnCount As Integer, ByVal sValue As String)
        Dim i, j As Integer
        Try
            If c1Grid.RowCount < 2 Then Exit Sub
            If ColumnCount = 1 Then ' Copy trong 1 cot
                CopyColumns(c1Grid, ColCopy, sValue, RowCopy)
            ElseIf ColumnCount > 1 Then ' Copy nhieu cot lien quan
                c1Grid.UpdateData()
                sValue = c1Grid(RowCopy, ColCopy).ToString
                Dim Flag As DialogResult

                Flag = MessageBox.Show(rl3("Copy_cot_du_lieu_cho") & vbCrLf & rl3("____-_Tat_ca_cac_cot_(nhan_Yes)") & vbCrLf & rl3("____-_Nhung_dong_con_trong_(nhan_No)"), MsgAnnouncement, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                If Flag = Windows.Forms.DialogResult.No Then ' Copy nhung dong con trong
                    For i = RowCopy + 1 To c1Grid.RowCount - 1
                        j = 1
                        If c1Grid(i, ColCopy).ToString = "" OrElse c1Grid(i, ColCopy).ToString = MaskFormatDateShort OrElse c1Grid(i, ColCopy).ToString = MaskFormatDate OrElse (L3IsNumeric(c1Grid(i, ColCopy).ToString) And Val(c1Grid(i, ColCopy).ToString) = 0) Then
                            c1Grid(i, ColCopy) = sValue
                            While j < ColumnCount
                                c1Grid(i, ColCopy + j) = c1Grid(RowCopy, ColCopy + j)
                                j += 1
                            End While
                            c1Grid(i, COL1_DataID) = c1Grid(RowCopy, COL1_DataID)
                            c1Grid(i, COL1_AccountID) = c1Grid(RowCopy, COL1_AccountID)
                        End If
                    Next
                ElseIf Flag = Windows.Forms.DialogResult.Yes Then ' Copy hết
                    For i = RowCopy + 1 To c1Grid.RowCount - 1
                        j = 1
                        c1Grid(i, ColCopy) = sValue
                        While j < ColumnCount
                            c1Grid(i, ColCopy + j) = c1Grid(RowCopy, ColCopy + j)
                            j += 1
                        End While
                        c1Grid(i, COL1_DataID) = c1Grid(RowCopy, COL1_DataID)
                        c1Grid(i, COL1_AccountID) = c1Grid(RowCopy, COL1_AccountID)
                    Next
                    'c1Grid(0, ColCopy) = sValue
                Else
                    Exit Sub
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    'Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
    '    'If e.KeyCode = Keys.Enter Then
    '    '    If tdbg1.Col = iLastCol1 Then
    '    '        HotKeyEnterGrid(tdbg1, COL1_DataName, e)
    '    '        Exit Sub
    '    '    End If
    '    'End If
    '    If e.KeyCode = Keys.F7 Then
    '        If tdbg1.Splits(tdbg1.SplitIndex).DisplayColumns(tdbg1.Col).Locked = False Then
    '            HotKeyF7Custom(tdbg1)
    '            Exit Sub
    '        End If
    '    End If
    '    If e.KeyCode = Keys.F8 Then
    '        If tdbg1.Splits(tdbg1.SplitIndex).DisplayColumns(tdbg1.Col).Locked = False Then
    '            HotKeyF8(tdbg1)
    '            Exit Sub
    '        End If
    '    End If
    '    HotKeyDownGrid(e, tdbg1, COL1_DataName, SPLIT0, SPLIT2)
    'End Sub

    Private Sub HotKeyF7Custom(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        Try
            If c1Grid.RowCount < 1 Then Exit Sub
            If c1Grid.Splits(c1Grid.SplitIndex).DisplayColumns.Item(c1Grid.Col).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far Then ' Số
                If c1Grid(c1Grid.Row, c1Grid.Col).ToString = "" OrElse Val(c1Grid(c1Grid.Row, c1Grid.Col).ToString) = 0 Then
                    c1Grid.Columns(c1Grid.Col).Text = c1Grid(c1Grid.Row - 1, c1Grid.Col).ToString()
                    c1Grid.UpdateData()
                End If
            Else ' Chuỗi hoặc Ngày
                If c1Grid(c1Grid.Row, c1Grid.Col).ToString = "" OrElse c1Grid(c1Grid.Row, c1Grid.Col).ToString = MaskFormatDateShort OrElse c1Grid(c1Grid.Row, c1Grid.Col).ToString = MaskFormatDate Then
                    c1Grid.Columns(c1Grid.Col).Text = c1Grid(c1Grid.Row - 1, c1Grid.Col).ToString()
                    If c1Grid.Col = COL1_DataName Then
                        c1Grid.Columns(COL1_DataID).Text = c1Grid(c1Grid.Row - 1, COL1_DataID).ToString()
                        c1Grid.Columns(COL1_Description).Text = c1Grid(c1Grid.Row - 1, COL1_Description).ToString()
                        c1Grid.Columns(COL1_AccountID).Text = c1Grid(c1Grid.Row - 1, COL1_AccountID).ToString()
                    ElseIf c1Grid.Col = COL1_ObjectTypeID Then
                        c1Grid.Columns(COL1_ObjectID).Text = c1Grid(c1Grid.Row - 1, COL1_ObjectID).ToString()
                    End If
                    c1Grid.UpdateData()
                End If
            End If

        Catch ex As Exception
            D99C0008.Msg("Lỗi HotKeyF7: " & ex.Message)
        End Try
    End Sub

    Private Sub tdbg1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg1.KeyPress
        Select Case tdbg1.Col
            Case COL1_Str01 To COL1_Str05
                e.Handled = tdbg1.Columns(tdbg1.Col).DataWidth = 0
        End Select
        'Select Case tdbg1.Col
        '    '    Case COL1_ConvertedAmount
        '    '        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        '    Case COL1_Num01
        '        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        '    Case COL1_Num02
        '        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        '    Case COL1_Num03
        '        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        '    Case COL1_Num04
        '        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        '    Case COL1_Num05
        '        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)

        'End Select
    End Sub

    Private Sub tdbg1_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.ComboSelect
        Select Case e.ColIndex
            Case COL1_DataName
                tdbg1.Columns(COL1_DataID).Text = tdbdDataID.Columns("DataID").Text
                If tdbg1.Columns(COL1_DataID).Text = "0" Then
                    If (gbUnicode) Then
                        tdbg1.Columns(COL1_Description).Text = rl3("Nguyen_gia_tai_san_co_dinh_nhap_so_duU")
                    Else
                        tdbg1.Columns(COL1_Description).Text = rl3("Nguyen_gia_tai_san_co_dinh_nhap_so_du")
                    End If
                ElseIf tdbg1.Columns(COL1_DataID).Text = "1" Then
                    If (gbUnicode) Then
                        tdbg1.Columns(COL1_Description).Text = rl3("Khau_hao_luy_ke_tai_san_co_dinh_nhap_so_duU")
                    Else
                        tdbg1.Columns(COL1_Description).Text = rl3("Khau_hao_luy_ke_tai_san_co_dinh_nhap_so_du")
                    End If

                End If
                If tdbcAssetID.Text <> "" And tdbg1.Columns(COL1_DataID).Text <> "" Then
                    If tdbg1.Columns(COL1_DataID).Text = "0" Then
                        tdbg1.Columns(COL1_AccountID).Text = tdbcAssetID.Columns("AssetAccountID").Value.ToString
                    Else
                        tdbg1.Columns(COL1_AccountID).Text = tdbcAssetID.Columns("DepAccountID").Value.ToString
                    End If
                Else
                    tdbg1.Columns(COL1_AccountID).Text = ""
                End If
        End Select
    End Sub

    Private Sub tdbg1_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg1.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        '--- Đổ nguồn cho các Dropdown phụ thuộc
        Select Case tdbg1.Col
            Case COL1_ObjectID
                LoadtdbdObjectID(tdbg1(tdbg1.Row, COL1_ObjectTypeID).ToString)
        End Select
    End Sub

    Private Sub tdbg1_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg1.BeforeColUpdate
        Select Case e.ColIndex
            Case COL1_DataName
                If tdbg1.Columns(COL1_DataName).Text <> tdbdDataID.Columns("DataName").Text Then
                    tdbg1.Columns(COL1_DataID).Text = ""
                    tdbg1.Columns(COL1_DataName).Text = ""
                    tdbg1.Columns(COL1_Description).Text = ""
                    tdbg1.Columns(COL1_AccountID).Text = ""
                End If
            Case COL1_ObjectTypeID
                If tdbg1.Columns(COL1_ObjectTypeID).Text <> tdbdObjectTypeID.Columns("ObjectTypeID").Text Then
                    tdbg1.Columns(COL1_ObjectTypeID).Text = ""
                    tdbg1.Columns(COL1_ObjectID).Text = ""
                End If
            Case COL1_ObjectID
                If tdbg1.Columns(COL1_ObjectID).Text <> tdbdObjectID.Columns("ObjectID").Text Then
                    tdbg1.Columns(COL1_ObjectID).Text = ""
                End If
                'Case COL1_Ana01ID
                '    If tdbg1.Columns(COL1_Ana01ID).Text <> tdbdAna01ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(0) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana01ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana01ID).Text.Length > giArrAnaLength(0) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana01ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana01ID).Text = tdbg1.Columns(COL1_Ana01ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana02ID
                '    If tdbg1.Columns(COL1_Ana02ID).Text <> tdbdAna02ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(1) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana02ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana02ID).Text.Length > giArrAnaLength(1) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana02ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana02ID).Text = tdbg1.Columns(COL1_Ana02ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana03ID
                '    If tdbg1.Columns(COL1_Ana03ID).Text <> tdbdAna03ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(2) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana03ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana03ID).Text.Length > giArrAnaLength(2) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana03ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana03ID).Text = tdbg1.Columns(COL1_Ana03ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana04ID
                '    If tdbg1.Columns(COL1_Ana04ID).Text <> tdbdAna04ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(3) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana04ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana04ID).Text.Length > giArrAnaLength(3) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana04ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana04ID).Text = tdbg1.Columns(COL1_Ana04ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana05ID
                '    If tdbg1.Columns(COL1_Ana05ID).Text <> tdbdAna05ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(4) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana05ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana05ID).Text.Length > giArrAnaLength(4) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana05ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana05ID).Text = tdbg1.Columns(COL1_Ana05ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana06ID
                '    If tdbg1.Columns(COL1_Ana06ID).Text <> tdbdAna06ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(5) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana06ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana06ID).Text.Length > giArrAnaLength(5) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana06ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana06ID).Text = tdbg1.Columns(COL1_Ana06ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana07ID
                '    If tdbg1.Columns(COL1_Ana07ID).Text <> tdbdAna07ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(6) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana07ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana07ID).Text.Length > giArrAnaLength(6) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana07ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana07ID).Text = tdbg1.Columns(COL1_Ana07ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana08ID
                '    If tdbg1.Columns(COL1_Ana08ID).Text <> tdbdAna08ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(7) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana08ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana08ID).Text.Length > giArrAnaLength(7) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana08ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana08ID).Text = tdbg1.Columns(COL1_Ana08ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana09ID
                '    If tdbg1.Columns(COL1_Ana09ID).Text <> tdbdAna09ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(8) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana09ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana09ID).Text.Length > giArrAnaLength(8) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana09ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana09ID).Text = tdbg1.Columns(COL1_Ana09ID).Value.ToString
                '            End If
                '        End If
                '    End If
                'Case COL1_Ana10ID
                '    If tdbg1.Columns(COL1_Ana10ID).Text <> tdbdAna10ID.Columns("AnaID").Text Then
                '        If gbArrAnaValidate(9) Then 'Kiểm tra nhập trong danh sách
                '            tdbg1.Columns(COL1_Ana10ID).Text = ""
                '        Else
                '            If tdbg1.Columns(COL1_Ana10ID).Text.Length > giArrAnaLength(9) Then ' Kiểm tra chiều dài nhập vào
                '                tdbg1.Columns(COL1_Ana10ID).Text = ""
                '            Else
                '                tdbg1.Columns(COL1_Ana10ID).Text = tdbg1.Columns(COL1_Ana10ID).Value.ToString
                '            End If
                '        End If
                '    End If

            Case COL1_Ana01ID To COL1_Ana10ID 'Có nhập ngoài danh sách không bỏ
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg1.Columns(e.ColIndex).Text <> tdbg1.Columns(e.ColIndex).DropDown.Columns("AnaID").Text Then
                    CheckAfterColUpdateAna(tdbg1, COL1_Ana01ID, e.ColIndex, dtAnaCaption) 'tham khảo hàm viết phía dưới
                End If
        End Select

    End Sub

    Private Sub tdbg1_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.AfterColUpdate

        Select Case e.ColIndex
            Case COL1_DataName
                If tdbg1.Columns(COL1_ConvertedAmount).Text <> "" Then
                    If tdbg1.Columns(COL1_DataID).Text = "0" Then ' Tính lại gtri txt Nguyên giá
                        'txtConvertedAmount.Text = SQLNumber(Sum(COL1_DataID, "0"), DxxFormat.D90_ConvertedDecimals)
                        'txtAmountDepreciation.Text = "0"
                    ElseIf tdbg1.Columns(COL1_DataID).Text = "1" Then ' Tính lại gtri txt Hao mòn lũy kế
                        'txtConvertedAmount.Text = "0"
                        ' txtAmountDepreciation.Text = SQLNumber(Sum(COL1_DataID, "1"), DxxFormat.D90_ConvertedDecimals)
                    Else
                        tdbg1.Columns(COL1_ConvertedAmount).Text = ""

                    End If
                    txtConvertedAmount.Text = SQLNumber(Sum(COL1_DataID, "0"), DxxFormat.D90_ConvertedDecimals)
                    txtAmountDepreciation.Text = SQLNumber(Sum(COL1_DataID, "1"), DxxFormat.D90_ConvertedDecimals)
                    'Tính giá trị còn lại:
                    txtRemainAmount.Text = SQLNumber(Number(txtConvertedAmount.Text) - Number(txtAmountDepreciation.Text), DxxFormat.D90_ConvertedDecimals)
                End If
            Case COL1_ConvertedAmount
                If tdbg1.Columns(COL1_DataID).Text = "0" Then ' Tính lại gtri txt Nguyên giá
                    txtConvertedAmount.Text = SQLNumber(Sum(COL1_DataID, "0"), DxxFormat.D90_ConvertedDecimals)
                ElseIf tdbg1.Columns(COL1_DataID).Text = "1" Then ' Tính lại gtri txt Hao mòn lũy kế
                    txtAmountDepreciation.Text = SQLNumber(Sum(COL1_DataID, "1"), DxxFormat.D90_ConvertedDecimals)
                End If
                'Tính giá trị còn lại:
                txtRemainAmount.Text = SQLNumber(Number(txtConvertedAmount.Text) - Number(txtAmountDepreciation.Text), DxxFormat.D90_ConvertedDecimals)
            Case COL1_Ana01ID To COL1_Ana10ID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg1, e.Column.DataColumn.DataField)
                If tdbd Is Nothing Then Exit Select
                If clsFilterDropdown.IsNewFilter Then
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg1, e, tdbd)
                    AfterColUpdate_tdbg1(e.ColIndex, dr)
                    Exit Sub
                Else ' Nhập liệu dạng cũ (xổ dropdown)
                    Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg1.Columns(e.ColIndex).Text))
                    AfterColUpdate_tdbg1(e.ColIndex, row)
                End If
        End Select
        tdbg1.UpdateData()
        'LockServiceLife()
    End Sub

    Private Sub AfterColUpdate_tdbg1(ByVal iCol As Integer, ByVal dr() As DataRow)
        Dim iRow As Integer = tdbg1.Row
        If dr Is Nothing OrElse dr.Length = 0 Then
            Dim row As DataRow = Nothing
            AfterColUpdate_tdbg1(iCol, row)
        ElseIf dr.Length = 1 Then
            If tdbg1.Bookmark <> tdbg1.Row AndAlso tdbg1.RowCount = tdbg1.Row Then 'Đang đứng dòng mới
                Dim dr1 As DataRow = dtGrid1.NewRow
                dtGrid1.Rows.InsertAt(dr1, tdbg1.Row)
                SetDefaultValues(tdbg1, dr1) 'Bổ sung set giá trị mặc định 19/08/2015
                tdbg1.Bookmark = tdbg1.Row
            End If
            AfterColUpdate_tdbg1(iCol, dr(0))
        Else
            For Each row As DataRow In dr
                tdbg1.Bookmark = iRow
                tdbg1.Row = iRow
                AfterColUpdate_tdbg1(iCol, row)
                tdbg1.UpdateData()
                iRow += 1
            Next
            tdbg1.Focus()
        End If
    End Sub


    Private Sub AfterColUpdate_tdbg1(ByVal iCol As Integer, ByVal dr As DataRow)
        'Gán lại các giá trị phụ thuộc vào Dropdown
        Select Case iCol
            Case COL1_Ana01ID To COL1_Ana10ID
                If dr Is Nothing OrElse dr.Item("AnaID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    CheckAfterColUpdateAna(tdbg1, COL1_Ana01ID, iCol, dtAnaCaption) 'tham khảo hàm viết phía dưới
                    Exit Sub
                End If
                tdbg1.Columns(iCol).Text = dr.Item("AnaID").ToString
        End Select
    End Sub

    Private Sub tdbg1_ButtonClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.ButtonClick
        If clsFilterDropdown.IsNewFilter = False Then Exit Sub
        If tdbg1.AllowUpdate = False Then Exit Sub
        If tdbg1.Splits(tdbg1.SplitIndex).DisplayColumns(tdbg1.Col).Locked Then Exit Sub
        Select Case tdbg1.Col

            Case COL1_Ana01ID To COL1_Ana10ID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg1, tdbg1.Columns(tdbg1.Col).DataField)
                If tdbd Is Nothing Then Exit Select
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg1, e, tdbd)
                If dr Is Nothing Then Exit Sub
                AfterColUpdate_tdbg1(tdbg1.Col, dr)
        End Select
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        If e.KeyCode = Keys.F7 Then
            If tdbg1.Splits(tdbg1.SplitIndex).DisplayColumns(tdbg1.Col).Locked = False Then
                HotKeyF7Custom(tdbg1)
                Exit Sub
            End If
        End If
        If e.KeyCode = Keys.F8 Then
            If tdbg1.Splits(tdbg1.SplitIndex).DisplayColumns(tdbg1.Col).Locked = False Then
                HotKeyF8(tdbg1)
                Exit Sub
            End If
        End If
        HotKeyDownGrid(e, tdbg1, COL1_DataName, SPLIT0, SPLIT2)

        If clsFilterDropdown.CheckKeydownFilterDropdown(tdbg1, e) Then
            Select Case tdbg1.Col
                Case COL1_Ana01ID To COL1_Ana10ID
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg1, tdbg1.Columns(tdbg1.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg1, e, tdbd)
                    If dr Is Nothing Then Exit Sub
                    AfterColUpdate_tdbg1(tdbg1.Col, dr)
                    Exit Sub
            End Select

        End If
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

    Private Function Sum(ByVal iCol As Integer, ByVal sValue As String) As Double
        Dim dSum As Double = 0
        Dim i As Integer
        For i = 0 To tdbg1.RowCount - 1
            If tdbg1(i, iCol).ToString = sValue Then
                dSum += Number(tdbg1(i, COL1_ConvertedAmount).ToString)
            End If
        Next
        Return dSum
    End Function

    Private Sub SetBackColorObligatory()
        tdbcAssetID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        If D02Systems.ObligatoryReceiver Then
            tdbcEmployeeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        End If
        c1dateVoucherDate.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtVoucherNo.BackColor = COLOR_BACKCOLOROBLIGATORY
        'txtServiceLife.BackColor = COLOR_BACKCOLOROBLIGATORY
        txtDepreciatedPeriod.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateBeginUsing.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateBeginDep.BackColor = COLOR_BACKCOLOROBLIGATORY
        If D02Systems.IsCalDepByDate = True Then c1dateDepDate.BackColor = COLOR_BACKCOLOROBLIGATORY '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        If D02Systems.IsObligatoryManagement Then
            tdbcManagementObjTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
            tdbcManagementObjID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        End If
    End Sub

    Private Sub LoadTDBCAssetID()
        Dim sSQL As String = ""

        sSQL = "SELECT 	T1.AssetID, T1.AssetName" & UnicodeJoin(gbUnicode) & " as AssetName, T1.AssetAccountID, T1.DepAccountID, ObjectTypeID, ObjectID, EmployeeID, " & vbCrLf
        sSQL &= "FullName" & UnicodeJoin(gbUnicode) & " as FullName, T1.ConvertedAmount, T1.Percentage, ServiceLife, T1.AmountDepreciation, T1.AssetDate, A.Description" & UnicodeJoin(gbUnicode) & " AS MethodID," & vbCrLf
        sSQL &= " B.Description" & UnicodeJoin(gbUnicode) & " AS MethodEndID, REPLACE( STR(UseMonth,2), ' ', '0') + '/' + STR(UseYear,4) AS BeginUsing," & vbCrLf
        sSQL &= "REPLACE( STR(DepMonth,2), ' ', '0') + '/' + STR(DepYear,4) AS BeginDep,T1.DepreciatedPeriod, T1.DepreciatedAmount," & vbCrLf
        sSQL &= "  IsCompleted,IsRevalued, IsDisposed, T1.AssignmentTypeID, C.DeprTableName, A.IntCode, T1.DepDate " & vbCrLf
        sSQL &= ", ManagementObjTypeID, ManagementObjID" & vbCrLf
        sSQL &= "FROM 		D02T0001 T1  WITH(NOLOCK) " & vbCrLf
        sSQL &= "INNER JOIN D02T8000 A WITH(NOLOCK) ON T1.MethodID = A.IntCode" & vbCrLf
        sSQL &= "INNER JOIN D02T8000 B WITH(NOLOCK) ON T1.MethodEndID = B.IntCode" & vbCrLf
        sSQL &= "LEFT JOIN 	D02T0070 C WITH(NOLOCK) ON 	T1.DeprTableID = C.DeprTableID" & vbCrLf
        sSQL &= " WHERE 	T1.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= " AND		A.Language = " & SQLString(gsLanguage)
        sSQL &= " AND		A.ModuleID = '02' AND A.Type = 0" & vbCrLf
        sSQL &= " AND		B.Language = " & SQLString(gsLanguage) & vbCrLf
        sSQL &= " AND		B.ModuleID = '02' AND B.Type = 1" & vbCrLf
        If _FormState = EnumFormState.FormAdd Then sSQL &= " AND IsCompleted = 0" & vbCrLf
        sSQL &= "ORDER BY 		T1.ASSETID, T1.AssetName		" & vbCrLf
        LoadDataSource(tdbcAssetID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcAssetID
        LoadTDBCAssetID()
        'Load tdbcObjectTypeID
        LoadObjectTypeID(tdbcObjectTypeID, gbUnicode)
        'Load tdbcObjectID
        sSQL = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) Order By ObjectID "
        dtObjectID = ReturnDataTable(sSQL)
        'Load tdbcEmployeeID
        LoadtdbcEmployeeID()
        'Load tdbcVoucherTypeID
        'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b5:
        LoadVoucherTypeID(tdbcVoucherTypeID, D02, sEditVoucherTypeID, gbUnicode)
        LoadObjectTypeID(tdbcManagementObjTypeID, gbUnicode)
        sSQL = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) Order By ObjectID "
        dtManagementID = ReturnDataTable(sSQL)
    End Sub

    Private Sub LoadtdbcEmployeeID()
        Dim sSQL As String
        sSQL = "Select ObjectID as EmployeeID, ObjectName" & UnicodeJoin(gbUnicode) & " as EmployeeName From Object WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where ObjectTypeID='NV' Order By ObjectID"
        LoadDataSource(tdbcEmployeeID, sSQL, gbUnicode)
    End Sub

    'Private Sub LoadtdbcObjectID(ByVal ID As String)
    '    LoadDataSource(tdbcObjectID, ReturnTableFilter(dtObjectID, "ObjectTypeID=" & SQLString(ID)), gbUnicode)
    'End Sub

    Private Sub LoadTDBDropDown()
        Dim sSQL As String = ""
        'Load tdbdDataID

        Dim dr As DataRow
        Dim dc As DataColumn
        Dim dtDataID As New DataTable
        dc = New DataColumn("DataID", Type.GetType("System.String"))
        dtDataID.Columns.Add(dc)
        dc = New DataColumn("DataName", Type.GetType("System.String"))
        dtDataID.Columns.Add(dc)
        dr = dtDataID.NewRow
        dr("DataID") = "0"
        If gbUnicode Or geLanguage = EnumLanguage.English Then
            dr("DataName") = rl3("Nguyen_gia")
            dtDataID.Rows.Add(dr)
            dr = dtDataID.NewRow
            dr("DataID") = "1"
            dr("DataName") = rl3("Hao_mon_luy_ke")
            dtDataID.Rows.Add(dr)
        Else
            dr("DataName") = "Nguyeân giaù"
            dtDataID.Rows.Add(dr)
            dr = dtDataID.NewRow
            dr("DataID") = "1"
            dr("DataName") = "Hao moøn luõy keá"
            dtDataID.Rows.Add(dr)
        End If


        LoadDataSource(tdbdDataID, dtDataID, gbUnicode)
        'Load tdbdObjectTypeID
        LoadObjectTypeID(tdbdObjectTypeID)
        'Load tdbdSourceID 
        sSQL = "SELECT 		SourceID, SourceName" & UnicodeJoin(gbUnicode) & " as SourceName " & vbCrLf
        sSQL &= "FROM 		D02T0013 WITH(NOLOCK)  " & vbCrLf
        sSQL &= "WHERE 		Disabled = 0 " & vbCrLf
        sSQL &= "ORDER BY 	SourceID" & vbCrLf
        LoadDataSource(tdbdSourceID, sSQL, gbUnicode)
        'Load tdbdAssignmentID
        sSQL = "SELECT 		AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " as AssignmentName, DebitAccountID, Extend " & vbCrLf
        sSQL &= "FROM 		D02T0002 WITH(NOLOCK)  " & vbCrLf
        sSQL &= "WHERE 		Disabled = 0 " & vbCrLf
        sSQL &= "ORDER BY 	AssignmentID" & vbCrLf

        dtAssignmentID = ReturnDataTable(sSQL)
        '--- Chuẩn Khoản mục b3: Load 10 khoản mục
        LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg1, COL1_Ana01ID, gbUnicode)
    End Sub

    Private Sub LoadtdbdAssignmentID(ByVal sAssignmentTypeID As String)
        If sAssignmentTypeID = "2" Then
            LoadDataSource(tdbdAssignmentID, ReturnTableFilter(dtAssignmentID, "Extend= 1 Or Extend=2"), gbUnicode)
        Else
            LoadDataSource(tdbdAssignmentID, ReturnTableFilter(dtAssignmentID, "Extend=0"), gbUnicode)
        End If
    End Sub

    Private Sub LoadtdbdObjectID(ByVal ID As String)
        LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObjectID, "ObjectTypeID=" & SQLString(ID)), gbUnicode)
    End Sub

#Region "Events tdbcAssetID with txtAssetName"

    Private Sub tdbcAssetID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetID.SelectedValueChanged
        If tdbcAssetID.SelectedValue Is Nothing Then
            txtAssetName.Text = ""
        Else

            txtAssetName.Text = tdbcAssetID.Columns("AssetName").Value.ToString
            tdbcObjectTypeID.Text = tdbcAssetID.Columns("ObjectTypeID").Value.ToString
            clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
            tdbcObjectID.SelectedValue = tdbcAssetID.Columns("ObjectID").Value.ToString
            tdbcEmployeeID.Text = tdbcAssetID.Columns("EmployeeID").Value.ToString
            txtEmployeeName.Text = tdbcAssetID.Columns("FullName").Value.ToString
            txtConvertedAmount.Text = SQLNumber(tdbcAssetID.Columns("ConvertedAmount").Value.ToString, DxxFormat.D90_ConvertedDecimals)
            txtPercentage.Text = SQLNumber(tdbcAssetID.Columns("Percentage").Value.ToString, DxxFormat.D08_RatioDecimals)
            txtPercentage.Tag = txtPercentage.Text
            txtServiceLife.Text = SQLNumber(tdbcAssetID.Columns("ServiceLife").Value.ToString, DxxFormat.DefaultNumber0)
            txtServiceLife.Tag = txtServiceLife.Text
            txtAmountDepreciation.Text = SQLNumber(tdbcAssetID.Columns("AmountDepreciation").Value.ToString, DxxFormat.D90_ConvertedDecimals)
            txtDepreciateAmount.Text = SQLNumber(tdbcAssetID.Columns("DepreciatedAmount").Value.ToString, DxxFormat.D90_ConvertedDecimals)
            txtRemainAmount.Text = SQLNumber(Number(txtConvertedAmount.Text) - Number(txtAmountDepreciation.Text), DxxFormat.D90_ConvertedDecimals)

            txtMethodID.Text = tdbcAssetID.Columns("MethodID").Text.ToString
            txtMethodEndID.Text = tdbcAssetID.Columns("MethodEndID").Value.ToString
            txtDeprTableName.Text = tdbcAssetID.Columns("DeprTableName").Value.ToString

            c1dateBeginDep.Value = tdbcAssetID.Columns("BeginDep").Value.ToString
            c1dateBeginUsing.Value = tdbcAssetID.Columns("BeginUsing").Value.ToString
            txtDepreciatedPeriod.Text = tdbcAssetID.Columns("DepreciatedPeriod").Value.ToString
            LoadTDBGAssignmentDefault()
            txtCurrenyID.Text = DxxFormat.BaseCurrencyID
            LoadtdbdAssignmentID(tdbcAssetID.Columns("AssignmentTypeID").Text)
            c1dateDepDate.Value = tdbcAssetID.Columns("DepDate").Text '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định

            tdbcManagementObjTypeID.SelectedValue = tdbcAssetID.Columns("ManagementObjTypeID").Value
            tdbcManagementObjID.SelectedValue = tdbcAssetID.Columns("ManagementObjID").Value
        End If
    End Sub

    Private Sub tdbcAssetID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetID.LostFocus
        'If tdbcAssetID.FindStringExact(tdbcAssetID.Text) = -1 Then
        '    tdbcAssetID.Text = ""
        'End If
    End Sub

    Private Sub tdbcAssetID_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetID.Validated
        clsFilterCombo.FilterCombo(tdbcAssetID, e)
        If tdbcAssetID.FindStringExact(tdbcAssetID.Text) = -1 Then
            tdbcAssetID.Text = ""
        End If

    End Sub

#End Region

#Region "Events tdbcObjectTypeID load tdbcObjectID with txtObjectName"

    Private Sub tdbcObjectTypeID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.GotFocus
        'Dùng phím Enter
        tdbcObjectTypeID.Tag = tdbcObjectTypeID.Text
    End Sub

    Private Sub tdbcObjectTypeID_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcObjectTypeID.MouseDown
        'Di chuyển chuột
        tdbcObjectTypeID.Tag = tdbcObjectTypeID.Text
    End Sub

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.SelectedValueChanged
        clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
        tdbcObjectID.Text = ""
    End Sub

    Private Sub tdbcObjectTypeID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.LostFocus
        'If tdbcObjectTypeID.Tag.ToString = "" And tdbcObjectTypeID.Text = "" Then Exit Sub
        'If tdbcObjectTypeID.Tag.ToString = tdbcObjectTypeID.Text And tdbcObjectTypeID.SelectedValue IsNot Nothing Then Exit Sub
        'If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 OrElse tdbcObjectTypeID.SelectedValue Is Nothing Then
        '    tdbcObjectTypeID.Text = ""
        '    LoadtdbcObjectID("-1")
        '    tdbcObjectID.Text = ""
        '    Exit Sub
        'End If
        'LoadtdbcObjectID(tdbcObjectTypeID.SelectedValue.ToString())
        'tdbcObjectID.Text = ""
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

    '#Region "Events tdbcVoucherTypeID with txtVoucherNo"

    '    Private Sub tdbcVoucherTypeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.SelectedValueChanged
    '        If tdbcVoucherTypeID.SelectedValue Is Nothing Then
    '            txtVoucherNo.Text = ""
    '        Else
    '            GetVoucherNo(tdbcVoucherTypeID, txtVoucherNo, btnSetNewKey)
    '            Dim sValue As String = rl3("Nhap_so_du_tai_san_co_dinh")
    '            If geLanguage = EnumLanguage.Vietnamese And gbUnicode = False Then sValue = ConvertUnicodeToVni(sValue)
    '            If txtDescription.Text.Trim = "" Then txtDescription.Text = sValue
    '        End If
    '    End Sub

    '    Private Sub tdbcVoucherTypeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.LostFocus
    '        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then
    '            tdbcVoucherTypeID.Text = ""
    '        End If
    '        'GetVoucherNo(tdbcVoucherTypeID, txtVoucherNo, btnSetNewKey)
    '    End Sub

    '#End Region

    '    Public Sub GetVoucherNo(ByVal tdbcVoucherTypeID As C1.Win.C1List.C1Combo, ByVal txtVoucherNo As TextBox, ByVal btnSetNewKey As Windows.Forms.Button)
    '        If tdbcVoucherTypeID.Text <> "" Then
    '            If tdbcVoucherTypeID.Columns("Auto").Text = "0" Then 'Không tạo mã tự động
    '                txtVoucherNo.ReadOnly = False
    '                txtVoucherNo.TabStop = True
    '                btnSetNewKey.Enabled = False
    '                txtVoucherNo.Text = ""
    '            Else
    '                gnNewLastKey = 0
    '                txtVoucherNo.ReadOnly = True
    '                txtVoucherNo.TabStop = False
    '                btnSetNewKey.Enabled = True

    '                txtVoucherNo.Text = CreateIGEVoucherNo(tdbcVoucherTypeID, False)
    '            End If
    '        End If
    '    End Sub

    '    Private Sub btnSetNewKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSetNewKey.Click
    '        GetNewVoucherNo(tdbcVoucherTypeID, txtVoucherNo)
    '    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0500
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 05/10/2009 09:28:45
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0500() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0500 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString("BAL") & COMMA 'SetUpFrom, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString("D02.AssetID=" & SQLString(_assetID)) & COMMA 'strFind, varchar[8000], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode)
        Return sSQL
    End Function

    Private Sub LoadMasterAssetID()
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0500()
        Dim dtMaster As DataTable = ReturnDataTable(sSQL)
        If dtMaster.Rows.Count > 0 Then
            With dtMaster.Rows(0)
                tdbcAssetID.Text = .Item("AssetID").ToString
                txtAssetName.Text = .Item("AssetName").ToString
                tdbcObjectTypeID.Text = .Item("ObjectTypeID").ToString
                tdbcObjectID.Text = .Item("ObjectID").ToString
                txtObjectName.Text = .Item("ObjectName").ToString
                tdbcEmployeeID.Text = .Item("EmployeeID").ToString
                txtConvertedAmount.Text = SQLNumber(.Item("ConvertedAmount").ToString, DxxFormat.D90_ConvertedDecimals)
                txtAmountDepreciation.Text = SQLNumber(.Item("AmountDepreciation").ToString, DxxFormat.D90_ConvertedDecimals)
                txtRemainAmount.Text = SQLNumber(.Item("RemainAmount").ToString, DxxFormat.D90_ConvertedDecimals)
                txtMethodID.Text = tdbcAssetID.Columns("MethodID").Text
                txtMethodEndID.Text = tdbcAssetID.Columns("MethodEndID").Value.ToString
                txtDeprTableName.Text = tdbcAssetID.Columns("DeprTableName").Value.ToString
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
        End If
    End Sub

    Dim dtGrid1 As DataTable
    Private Sub LoadTDBGrid1()
        Dim sSQL As New StringBuilder("")
        sSQL.Append("SELECT 	TransactionID,	TransactionTypeID, Description" & UnicodeJoin(gbUnicode) & " as Description, ObjectTypeID,  ObjectID,  " & vbCrLf)
        sSQL.Append("IsNull(DebitAccountID, CreditAccountID) as AccountID,  " & vbCrLf)
        sSQL.Append("			ConvertedAmount, SourceID, " & vbCrLf)
        sSQL.Append("Ana01ID,Ana02ID,Ana03ID, Ana04ID,Ana05ID," & vbCrLf)
        sSQL.Append("Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID, " & vbCrLf)
        sSQL.Append("			Str01" & UnicodeJoin(gbUnicode) & " as Str01 , Str02" & UnicodeJoin(gbUnicode) & " as Str02, Str03" & UnicodeJoin(gbUnicode) & " as Str03, Str04" & UnicodeJoin(gbUnicode) & " as Str04, Str05" & UnicodeJoin(gbUnicode) & " as Str05, " & vbCrLf)
        sSQL.Append("Num01, Num02, Num03, Num04, Num05 , " & vbCrLf)
        sSQL.Append("Date01, Date02, Date03, Date04, Date05, '' AS DataID, '' As DataName , Convert(bit, IsNotAllocate) as IsNotAllocate " & vbCrLf)
        sSQL.Append("FROM 		D02T0012 WITH(NOLOCK)  " & vbCrLf)
        sSQL.Append("WHERE 		IsNull(TransactionTypeID, '') IN ('SD','SDKH') " & vbCrLf)
        sSQL.Append("AND 	DivisionID =  " & SQLString(gsDivisionID) & vbCrLf)
        sSQL.Append("AND 	AssetID = " & SQLString(_assetID) & vbCrLf)
        dtGrid1 = ReturnDataTable(sSQL.ToString)
        LoadDataSource(tdbg1, dtGrid1, gbUnicode)

        For i As Integer = 0 To dtGrid1.Rows.Count - 1
            If dtGrid1.Rows(i).Item("TransactionTypeID").ToString = "SD" Then
                tdbg1(i, COL1_DataID) = "0"
                If (gbUnicode) Then
                    tdbg1(i, COL1_DataName) = rL3("Nguyen_gia")
                Else
                    tdbg1(i, COL1_DataName) = rL3("Nguyen_giaV")
                End If

            ElseIf dtGrid1.Rows(i).Item("TransactionTypeID").ToString = "SDKH" Then
                tdbg1(i, COL1_DataID) = "1"
                If (gbUnicode) Then
                    tdbg1(i, COL1_DataName) = rL3("Hao_mon_luy_ke")

                Else
                    tdbg1(i, COL1_DataName) = rL3("Hao_mon_luy_ke_V")
                End If

            End If
        Next
    End Sub

    Private Sub LoadTDBGrid2()
        Dim sSQL As New StringBuilder()
        sSQL.Append("SELECT 		D02T5000.HistoryID, D02T5000.AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " as AssignmentName, " & vbCrLf)
        sSQL.Append("DebitAccountID, PercentAmount, D02T0002.Extend" & vbCrLf)
        sSQL.Append("FROM 		D02T5000 WITH(NOLOCK) " & vbCrLf)
        sSQL.Append("INNER JOIN 		D02T0002 WITH(NOLOCK) " & vbCrLf)
        sSQL.Append("ON 		D02T5000.AssignmentID = D02T0002.AssignmentID " & vbCrLf)
        sSQL.Append("WHERE 		HistoryTypeID = 'AS' " & vbCrLf)
        'sSQL.Append("AND 	BeginYear =" & giTranYear & vbCrLf) 'Bo BeginYear,BeginMonth theo Incident 79529
        'sSQL.Append("AND 	BeginMonth =" & giTranMonth & vbCrLf)
        sSQL.Append("AND 	DivisionID = " & SQLString(gsDivisionID) & vbCrLf)
        sSQL.Append("AND 	BatchID = " & SQLString(tdbcAssetID.Text) & vbCrLf)
        sSQL.Append("AND 	AssetID = " & SQLString(tdbcAssetID.Text) & vbCrLf)
        sSQL.Append("ORDER BY 	HistoryID" & vbCrLf)
        Dim dtGrid2 As DataTable = ReturnDataTable(sSQL.ToString)
        LoadDataSource(tdbg2, dtGrid2, gbUnicode)
    End Sub

    Private Sub LoadAddNew()
        _assetID = ""
        ClearText(Me)
        c1dateVoucherDate.Value = Now.Date.Date
        tabMain.SelectedTab = tabPage1
        chkPosted.Checked = True
        LoadTDBGrid1()
        LoadTDBGrid2()
        c1dateBeginUsing.Value = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        c1dateBeginDep.Value = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        c1dateUseDate.Value = Date.Today
        c1dateAssetDate.Value = Date.Today
        c1dateDepDate.Value = Date.Today '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        tdbcAssetID.Focus()
    End Sub

    Private Sub LoadEdit()
        tdbcAssetID.ReadOnly = True
        tdbcVoucherTypeID.ReadOnly = True
        txtVoucherNo.ReadOnly = True
        LoadMasterAssetID()
        LoadInfoMasterTab1()
        LoadTDBGrid1()
        LoadTDBGrid2()

    End Sub

    Private Sub LoadInfoMasterTab1()
        Dim sSQL As New StringBuilder()
        sSQL.Append("SELECT 	TOP 1 	VoucherDate,VoucherTypeID, VoucherNo,  CurrencyID, Notes" & UnicodeJoin(gbUnicode) & " as Notes, Posted  " & vbCrLf)
        sSQL.Append("FROM 		D02T0012 WITH(NOLOCK)  " & vbCrLf)
        sSQL.Append("WHERE 		IsNull(TransactionTypeID, '') IN ('SD','SDKH') " & vbCrLf)
        sSQL.Append("AND 		DivisionID = " & SQLString(gsDivisionID) & vbCrLf)
        sSQL.Append("AND 		AssetID =" & SQLString(_assetID) & vbCrLf)
        Dim dt1 As DataTable = ReturnDataTable(sSQL.ToString)
        If dt1.Rows.Count > 0 Then
            With dt1.Rows(0)
                'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b3:
                sEditVoucherTypeID = .Item("VoucherTypeID").ToString
                'LoadTDBCombo()
                '---------------------------------------------------------
                c1dateVoucherDate.Value = .Item("VoucherDate").ToString
                tdbcVoucherTypeID.Text = .Item("VoucherTypeID").ToString
                txtVoucherNo.Text = .Item("VoucherNo").ToString
                txtCurrenyID.Text = .Item("CurrencyID").ToString
                txtDescription.Text = .Item("Notes").ToString
                chkPosted.Checked = CBool(.Item("Posted").ToString)
            End With
        End If
    End Sub

    'Private Sub tdbg1_NumberFormat()
    '    'Dim arr() As FormatColumn = Nothing
    '    'AddNumberColumns(arr, SqlDbType.Money, tdbg1.Columns(COL1_ConvertedAmount).DataField, "N" & DxxFormat.iD90_ConvertedDecimals)
    '    'InputNumber(tdbg1, arr)
    '    'tdbg1.Columns(COL1_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    '    'tdbg1.Columns(COL1_Num01).NumberFormat = DxxFormat.DefaultNumber2
    '    'tdbg1.Columns(COL1_Num02).NumberFormat = DxxFormat.DefaultNumber2
    '    'tdbg1.Columns(COL1_Num03).NumberFormat = DxxFormat.DefaultNumber2
    '    'tdbg1.Columns(COL1_Num04).NumberFormat = DxxFormat.DefaultNumber2
    '    'tdbg1.Columns(COL1_Num05).NumberFormat = DxxFormat.DefaultNumber2
    'End Sub

    Private Sub tdbg2_NumberFormat()
        'tdbg2.Columns(COL2_PercentAmount).NumberFormat = DxxFormat.DefaultNumber2
        tdbg1_NumberFormat1()
        tdbg2_NumberFormat2()
    End Sub

    Private Sub tdbg1_NumberFormat1()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg1.Columns(COL1_ConvertedAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg1, arr)
    End Sub

    Private Sub tdbg2_NumberFormat2()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg2.Columns(COL2_PercentAmount).DataField, DxxFormat.DefaultNumber2, 28, 8)
        InputNumber(tdbg2, arr)
    End Sub



    'Private Sub GetCaptionDescription()
    '    Dim dtCaption As DataTable
    '    Dim iInDex As Integer = COL1_Str01
    '    Dim bNotAllDisabled As Boolean = False
    '    Dim sLang As String = ""

    '    'If (geLanguage = EnumLanguage.Vietnamese) Then
    '    '    If (gbUnicode) Then
    '    '        sLang = "84U"
    '    '    Else
    '    '        sLang = "84"
    '    '    End If
    '    'Else
    '    '    If (gbUnicode) Then
    '    '        sLang = "01U"
    '    '    Else
    '    '        sLang = "01"
    '    '    End If
    '    'End If
    '    Dim sSQL As New StringBuilder()
    '    sSQL.Append("SELECT		Data" & gsLanguage & UnicodeJoin(gbUnicode) & " as Data," & vbCrLf)
    '    sSQL.Append("Description" & UnicodeJoin(gbUnicode) & " as Description, Disabled, DataID " & vbCrLf)
    '    sSQL.Append("FROM 		D02T0003  " & vbCrLf)
    '    sSQL.Append("WHERE 		DataID LIKE 'AnaStr%'" & vbCrLf)
    '    sSQL.Append("UNION ALL" & vbCrLf)
    '    sSQL.Append("SELECT		Data" & sLang & " as Data," & vbCrLf)
    '    sSQL.Append("Description" & UnicodeJoin(gbUnicode) & " as Description, Disabled, DataID " & vbCrLf)
    '    sSQL.Append("FROM 		D02T0003  " & vbCrLf)
    '    sSQL.Append("WHERE		DataID LIKE 'AnaNum%'" & vbCrLf)
    '    sSQL.Append("UNION ALL" & vbCrLf)
    '    sSQL.Append("SELECT		Data" & sLang & " as Data," & vbCrLf)
    '    sSQL.Append("Description" & UnicodeJoin(gbUnicode) & " as Description, Disabled, DataID " & vbCrLf)
    '    sSQL.Append("FROM 		D02T0003  " & vbCrLf)
    '    sSQL.Append("WHERE 		DataID LIKE 'AnaDate%'" & vbCrLf)
    '    dtCaption = ReturnDataTable(sSQL.ToString)
    '    If dtCaption.Rows.Count > 0 Then
    '        For i As Integer = 0 To dtCaption.Rows.Count - 1
    '            With dtCaption.Rows(i)
    '                tdbg1.Columns(iInDex).Caption = .Item("Description").ToString
    '                tdbg1.Splits(SPLIT1).DisplayColumns(iInDex).HeadingStyle.Font = FontUnicode(gbUnicode)
    '                tdbg1.Splits(SPLIT1).DisplayColumns(iInDex).Visible = Not CBool(.Item("Disabled"))
    '                iInDex += 1
    '            End With
    '        Next
    '        iInDex = COL1_Str01
    '        For i As Integer = 0 To dtCaption.Rows.Count - 1
    '            If tdbg1.Splits(SPLIT1).DisplayColumns(iInDex).Visible = True Then
    '                bNotAllDisabled = True
    '                Exit For
    '            Else
    '                iInDex += 1
    '            End If
    '        Next
    '        ''''''''''''''''''''''''''''''''''''''''''''

    '    End If
    '    If bNotAllDisabled = False Then
    '        tdbg1.Splits(SPLIT1).SplitSize = 0
    '        tdbg1.Splits(SPLIT1).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.None
    '        tdbg1.Splits(SPLIT2).SplitSize = 160
    '        'tdbg1.Splits(SPLIT2).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.Always
    '    Else
    '        tdbg1.Splits(SPLIT1).SplitSize = 200
    '        tdbg1.Splits(SPLIT1).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.Always
    '        tdbg1.Splits(SPLIT2).SplitSize = 160
    '    End If
    'End Sub

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

    Private Function LoadCaptionSubInfo() As Boolean
        Dim bUseSubInfo As Boolean = False
        Dim dtCaption As DataTable = ReturnDataTable(SQLStoreD02P0015)
        If dtCaption.Rows.Count = 0 Then Return False
        Dim arr() As FormatColumn = Nothing
        For i As Integer = 0 To dtCaption.Rows.Count - 1
            Dim sField As String = dtCaption.Rows(i).Item("DataID").ToString.Replace("Ana", "")

            If tdbg1.Columns.IndexOf(tdbg1.Columns(sField)) = -1 Then Continue For
            tdbg1.Columns(sField).Caption = dtCaption.Rows(i).Item("Data" & gsLanguage).ToString
            tdbg1.Splits(SPLIT1).DisplayColumns(sField).HeadingStyle.Font = FontUnicode(gbUnicode)
            tdbg1.Splits(SPLIT1).DisplayColumns(sField).Visible = CBool(dtCaption.Rows(i).Item("Disabled"))
            If tdbg1.Splits(SPLIT1).DisplayColumns(sField).Visible Then
                bUseSubInfo = True
                Select Case L3Int(dtCaption.Rows(i).Item("DataType"))
                    Case 0 'Số
                        AddNumberColumns(arr, SqlDbType.Money, tdbg1.Columns(sField).DataField, "N" & L3Int(dtCaption.Rows(i).Item("DecimalNum")))
                    Case 1 'Chuỗi
                        tdbg1.Columns(sField).DataWidth = L3Int(dtCaption.Rows(i).Item("DecimalNum"))
                    Case 2 'Ngày
                End Select
            End If
        Next
        If arr IsNot Nothing Then InputNumber(tdbg1, arr)
        '''''''''''''''''''''''''''''''''''''''''''''
        Return bUseSubInfo
    End Function

    Private Sub c1dateBeginDep_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles c1dateBeginDep.KeyDown
        e.Handled = False
    End Sub

    Private Sub c1dateBeginDep_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles c1dateBeginDep.KeyPress
        e.Handled = False
    End Sub

    Private Sub c1dateBeginUsing_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles c1dateBeginUsing.KeyDown
        e.Handled = False
    End Sub

    Private Sub c1dateBeginUsing_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles c1dateBeginUsing.KeyPress
        e.Handled = False
    End Sub

    Private Sub LoadTDBGAssignmentDefault()
        Dim sSQL As New StringBuilder()
        'sSQL.Append("SELECT '' as DataID,'' as DataName, '' as Description, '' as AccountID, '' as ObjectTypeID, '' as ObjectID, D02T5000.HistoryID, D02T5000.SourceID, SourceName, " & vbCrLf)
        'sSQL.Append("'' as Str01,'' as Str02,'' as Str03,'' as Str04,'' as Str05,'' as Num01,'' as Num02,'' as Num03,'' as Num04,'' as Num05, " & vbCrLf)
        'sSQL.Append("'' as Date01,'' as Date02,'' as Date03,'' as Date04,'' as Date05,'' as Ana01ID, '' as Ana02ID, '' as Ana03ID, '' as Ana04ID, '' as Ana05ID, '' as Ana06ID, " & vbCrLf)
        'sSQL.Append("'' as Ana05ID, '' as Ana06ID, '' as Ana07ID, '' as Ana08ID, '' as Ana09ID, '' as Ana10ID, ConvertedAmount, PercentAmount " & vbCrLf)
        'sSQL.Append("FROM 		D02T5000 " & vbCrLf)
        'sSQL.Append("INNER JOIN 	D02T0013 " & vbCrLf)
        'sSQL.Append("ON 		D02T5000.SourceID = D02T0013.SourceID  " & vbCrLf)
        'sSQL.Append("WHERE 		1 = 0" & vbCrLf)
        'Dim dtDefaultGrid1 As DataTable = ReturnDataTable(sSQL.ToString)
        'LoadDataSource(tdbg1, dtDefaultGrid1, gbUnicode)
        'sSQL.Remove(0, sSQL.Length)
        If dtGrid1 IsNot Nothing Then dtGrid1.Clear()

        sSQL.Append("" & vbCrLf)
        sSQL.Append("SELECT 		D02T5000.HistoryID, D02T5000.AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " as AssignmentName," & vbCrLf)
        sSQL.Append(" 			DebitAccountID, PercentAmount, D02T0002.Extend" & vbCrLf)
        sSQL.Append("FROM 		D02T5000 WITH(NOLOCK) " & vbCrLf)
        sSQL.Append("INNER JOIN 	D02T0002 WITH(NOLOCK) " & vbCrLf)
        sSQL.Append("ON 		D02T5000.AssignmentID = D02T0002.AssignmentID " & vbCrLf)
        sSQL.Append("WHERE 		1 = 0" & vbCrLf)

        Dim dtDefaultGrid2 As DataTable = ReturnDataTable(sSQL.ToString)
        LoadDataSource(tdbg2, dtDefaultGrid2, gbUnicode)
    End Sub

    Private Sub tdbg2_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.AfterColUpdate
        Select Case e.ColIndex
            Case COL2_AssignmentID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, e.Column.DataColumn.DataField)
                If tdbd Is Nothing Then Exit Select
                If clsFilterDropdown.IsNewFilter Then
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                    AfterColUpdate_tdbg2(e.ColIndex, dr)
                    Exit Sub
                Else ' Nhập liệu dạng cũ (xổ dropdown)
                    Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg2.Columns(e.ColIndex).Text))
                    AfterColUpdate_tdbg2(e.ColIndex, row)
                End If
        End Select
    End Sub

    Private Sub AfterColUpdate_tdbg2(ByVal iCol As Integer, ByVal dr() As DataRow)
        Dim iRow As Integer = tdbg2.Row
        If dr Is Nothing OrElse dr.Length = 0 Then
            Dim row As DataRow = Nothing
            AfterColUpdate_tdbg2(iCol, row)
        ElseIf dr.Length = 1 Then
            If tdbg2.Bookmark <> tdbg2.Row AndAlso tdbg2.RowCount = tdbg2.Row Then 'Đang đứng dòng mới
                Dim dtGrid2 As DataTable = CType(tdbg2.DataSource, DataTable)
                Dim dr1 As DataRow = dtGrid2.NewRow
                dtGrid2.Rows.InsertAt(dr1, tdbg2.Row)
                SetDefaultValues(tdbg2, dr1) 'Bổ sung set giá trị mặc định 19/08/2015
                tdbg2.Bookmark = tdbg2.Row
            End If
            AfterColUpdate_tdbg2(iCol, dr(0))
        Else
            For Each row As DataRow In dr
                tdbg2.Bookmark = iRow
                tdbg2.Row = iRow
                AfterColUpdate_tdbg2(iCol, row)
                tdbg2.UpdateData()
                iRow += 1
            Next
            tdbg2.Focus()
        End If
    End Sub

    Private Sub tdbg2_ButtonClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.ButtonClick
        If clsFilterDropdown.IsNewFilter = False Then Exit Sub
        If tdbg2.AllowUpdate = False Then Exit Sub
        If tdbg2.Splits(tdbg2.SplitIndex).DisplayColumns(tdbg2.Col).Locked Then Exit Sub
        Select Case tdbg2.Col
            Case COL2_AssignmentID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, tdbg2.Columns(tdbg2.Col).DataField)
                If tdbd Is Nothing Then Exit Select
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                If dr Is Nothing Then Exit Sub
                AfterColUpdate_tdbg2(tdbg2.Col, dr)
        End Select
    End Sub

    Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
        If clsFilterDropdown.CheckKeydownFilterDropdown(tdbg2, e) Then
            Select Case tdbg2.Col
                Case COL2_AssignmentID
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, tdbg2.Columns(tdbg2.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdownMulti(tdbg2, e, tdbd)
                    If dr Is Nothing Then Exit Sub
                    AfterColUpdate_tdbg2(tdbg2.Col, dr)
                    Exit Sub
            End Select
        End If
    End Sub

    Private Sub AfterColUpdate_tdbg2(ByVal iCol As Integer, ByVal dr As DataRow)
        'Gán lại các giá trị phụ thuộc vào Dropdown
        Select Case iCol
            Case COL2_AssignmentID
                If dr Is Nothing OrElse dr.Item("AssignmentID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    tdbg2.Columns(COL2_AssignmentID).Text = ""
                    tdbg2.Columns(COL2_AssignmentName).Text = ""
                    tdbg2.Columns(COL2_DebitAccountID).Text = ""
                    tdbg2.Columns(COL2_Extend).Text = ""
                    Exit Sub
                End If
                tdbg2.Columns(COL2_AssignmentID).Text = dr.Item("AssignmentID").ToString
                tdbg2.Columns(COL2_AssignmentName).Text = dr.Item("AssignmentName").ToString
                tdbg2.Columns(COL2_DebitAccountID).Text = dr.Item("DebitAccountID").ToString
                tdbg2.Columns(COL2_Extend).Text = dr.Item("Extend").ToString
        End Select
    End Sub

    'Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
    '    'If e.KeyCode = Keys.Enter Then
    '    '    If tdbg2.Col = iLastCol2 Then
    '    '        HotKeyEnterGrid(tdbg2, COL2_AssignmentID, e)
    '    '        Exit Sub
    '    '    End If
    '    'End If
    'End Sub

    Private Sub tdbg2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg2.KeyPress
        '--- Chỉ cho nhập số
        Select Case tdbg2.Col
            Case COL2_PercentAmount
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub tdbg2_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.ComboSelect
        '--- Gán giá trị phụ thuộc từ Dropdown
        Select Case e.ColIndex
            Case COL2_AssignmentID
                tdbg2.Columns(COL2_AssignmentName).Text = tdbdAssignmentID.Columns("AssignmentName").Text
                tdbg2.Columns(COL2_DebitAccountID).Text = tdbdAssignmentID.Columns("DebitAccountID").Text
                tdbg2.Columns(COL2_Extend).Text = tdbdAssignmentID.Columns("Extend").Text
        End Select
    End Sub

    Private Sub tdbg2_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg2.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        Select Case e.ColIndex
            Case COL2_AssignmentID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg2.Columns(COL2_AssignmentID).Text <> tdbdAssignmentID.Columns("AssignmentID").Text Then
                    tdbg2.Columns(COL2_AssignmentID).Text = ""
                    tdbg2.Columns(COL2_AssignmentName).Text = ""
                    tdbg2.Columns(COL2_DebitAccountID).Text = ""
                    tdbg2.Columns(COL2_Extend).Text = ""
                End If

            Case COL2_PercentAmount
                If Not L3IsNumeric(tdbg2.Columns(COL2_PercentAmount).Text, EnumDataType.Money) Then e.Cancel = True
        End Select
    End Sub

    Private Sub txtServiceLife_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtServiceLife.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtDepreciatedPeriod_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDepreciatedPeriod.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtDepreciateAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDepreciateAmount.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtServiceLife_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtServiceLife.Validated
        If D02Systems.IsCalPeriodByRate = True Then Exit Sub '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ

        If Number(txtServiceLife.Tag.ToString) = Number(txtServiceLife.Text) Then Exit Sub

        txtServiceLife.Text = SQLNumber(txtServiceLife.Text, DxxFormat.DefaultNumber0)
        'Tính lại gtri Mức khấu hao (txtDepreciateAmount)
        If Number(txtServiceLife.Text) = 0 Or txtServiceLife.Text = "" Then
            txtDepreciateAmount.Text = SQLNumber(txtConvertedAmount.Text, DxxFormat.D90_ConvertedDecimals)
            Exit Sub
        End If
        Dim dMau As Double = Number(IIf(Number(txtServiceLife.Text) = 0, 1, Number(txtServiceLife.Text)))
        txtDepreciateAmount.Text = SQLNumber((Number(txtConvertedAmount.Text)) / dMau, DxxFormat.D90_ConvertedDecimals)
        txtPercentage.Text = SQLNumber(100 / Number(txtServiceLife.Text), DxxFormat.D08_RatioDecimals)
        txtPercentage.Tag = txtPercentage.Text
        txtServiceLife.Tag = txtServiceLife.Text
    End Sub

    Private Sub txtDepreciateAmount_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDepreciateAmount.Validated
        If Not IsNumeric(txtDepreciateAmount.Text) Then
            txtDepreciateAmount.Text = ""
        End If
        txtDepreciateAmount.Text = SQLNumber(txtDepreciateAmount.Text, DxxFormat.D90_ConvertedDecimals)
        If txtDepreciateAmount.Text <> "" And Number(txtDepreciateAmount.Text) > Number(txtRemainAmount.Text) Then
            txtDepreciateAmount.Text = ""
        End If
        'Tính lại gtri Tỷ lệ khấu hao (txtPercentage)
        If txtServiceLife.Text = "" Or Number(txtServiceLife.Text) = 0 Then
            txtPercentage.Text = "0"
        Else
            txtPercentage.Text = SQLNumber(100 / Number(txtServiceLife.Text), DxxFormat.D08_RatioDecimals)
        End If
        txtPercentage.Tag = txtPercentage.Text
    End Sub

    Private Sub txtDepreciatedPeriod_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDepreciatedPeriod.Validated
        txtDepreciatedPeriod.Text = SQLNumber(txtDepreciatedPeriod.Text, DxxFormat.DefaultNumber0)
    End Sub

    Private Sub tdbg1_LockedColumns()
        tdbg1.Splits(SPLIT0).DisplayColumns(COL1_AccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub
    Dim bCheckIsManagement As Boolean = False
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub

        tdbg1.UpdateData()
        tdbg2.UpdateData()

        If Not AllowSave() Then Exit Sub
      
        sGetDate = SetGetDateSQL()
        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)
        If CheckVoucherDateInPeriod(c1dateVoucherDate.Text) = False Then
            tabMain.SelectedTab = tabPage1
            c1dateVoucherDate.Focus()
            Exit Sub
        End If
        btnSave.Enabled = False
        btnClose.Enabled = False
        gbSavedOK = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd

                Dim sTransactionID As String = ""
                Dim iFirstTrans As Long = 0
                sTransactionID = CreateIGENewS("D02T0012", "TransactionID", "02", "TJ", gsStringKey, sTransactionID, tdbg1.RowCount, iFirstTrans)
                tdbg1(0, COL1_TransactionID) = sTransactionID
                '****************************************
                'Kiểm tra phiếu theo kiểu mới
                'Kiểm tra phiếu
                If tdbcVoucherTypeID.Columns("Auto").Text = "1" And bEditVoucherNo = False Then 'Sinh tự động và không nhấn F2
                    txtVoucherNo.Text = CreateIGEVoucherNoNew(tdbcVoucherTypeID, "D02T0012", sTransactionID)
                Else 'Không sinh tự động hay có nhấn F2
                    If bEditVoucherNo = False Then
                        'Kiểm tra trùng Số phiếu
                        If CheckDuplicateVoucherNoNew(D02, "D02T0012", sTransactionID, txtVoucherNo.Text) = True Then btnSave.Enabled = True : btnClose.Enabled = True : Me.Cursor = Cursors.Default : Exit Sub
                    Else 'Có nhấn F2 để sửa số phiếu
                        'Insert Số phiếu vào bảng D40T5558
                        InsertD02T5558(sTransactionID, sOldVoucherNo, txtVoucherNo.Text)
                    End If
                    'Insert VoucherNo vào bảng D91T9111
                    InsertVoucherNoD91T9111(txtVoucherNo.Text, "D02T0012", sTransactionID)
                End If
                bEditVoucherNo = False
                sOldVoucherNo = ""
                bFirstF2 = False
                '****************************************
                sSQL.Append(SQLUpdateD02T0001_1.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000_ManageHistory.ToString & vbCrLf)
                If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" Then
                    sSQL.Append(SQLInsertD02T5000_ManageHistory.ToString & vbCrLf)
                End If
                sSQL.Append(SQLInsertD02T5000s_AssignHistory.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000_EffectHistory.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000_LiquiDateHistory.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T0012s(sTransactionID, iFirstTrans).ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0001_2.ToString & vbCrLf)

                '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
                Dim sHistoryIDAAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HC", gsStringKey)
                sSQL.Append(SQLInsertD02T5000_AssetAccount(sHistoryIDAAC).ToString)
                Dim sHistoryIDDAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HC", gsStringKey)
                sSQL.Append(SQLInsertD02T5000_DepAccount(sHistoryIDDAC).ToString)
                sSQL.Append(SQLInsertD02T5010(sHistoryIDAAC, "AAC", ReturnValueC1Combo(tdbcAssetID, "AssetAccountID")))
                sSQL.Append(SQLInsertD02T5010(sHistoryIDDAC, "DAC", ReturnValueC1Combo(tdbcAssetID, "DepAccountID")))
            Case EnumFormState.FormEdit, EnumFormState.FormEditOther
                sSQL.Append(SQLDeleteD02T5000.ToString & vbCrLf)
                sSQL.Append(SQLDeleteD02T0012.ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0001_1.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000_ManageHistory.ToString & vbCrLf)
                If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" Then
                    sSQL.Append(SQLInsertD02T5000_ManageHistory.ToString & vbCrLf)
                End If
                sSQL.Append(SQLInsertD02T5000s_AssignHistory.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000_EffectHistory.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000_LiquiDateHistory.ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T0012s().ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0001_2.ToString & vbCrLf)

                '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
                Dim sHistoryIDAAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HC", gsStringKey)
                sSQL.Append(SQLInsertD02T5000_AssetAccount(sHistoryIDAAC).ToString)
                Dim sHistoryIDDAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HC", gsStringKey)
                sSQL.Append(SQLInsertD02T5000_DepAccount(sHistoryIDDAC).ToString)
                sSQL.Append(SQLInsertD02T5010(sHistoryIDAAC, "AAC", ReturnValueC1Combo(tdbcAssetID, "AssetAccountID")))
                sSQL.Append(SQLInsertD02T5010(sHistoryIDDAC, "DAC", ReturnValueC1Combo(tdbcAssetID, "DepAccountID")))
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            gbSavedOK = True
            btnClose.Enabled = True
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
        Else
            If _FormState = EnumFormState.FormAdd Then
                DeleteVoucherNoD91T9111_Transaction(txtVoucherNo.Text, "D02T0012", "VoucherNo", tdbcVoucherTypeID, bEditVoucherNo)
            End If
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    Private Function AllowSave() As Boolean
        Dim bIsOriginalAmount As Boolean = False

        If tdbcAssetID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Ma_tai_san"))
            tdbcAssetID.Focus()
            Return False
        End If
        If tdbcObjectTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Bo_phan_quan_ly"))
            tdbcObjectTypeID.Focus()
            Return False
        End If
        If tdbcObjectID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Bo_phan_quan_ly"))
            tdbcObjectID.Focus()
            Return False
        End If

        If D02Systems.ObligatoryReceiver AndAlso tdbcEmployeeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Nguoi_tiep_nhan"))
            tdbcEmployeeID.Focus()
            Return False
        End If
        If c1dateVoucherDate.Value.ToString = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ngay_chung_tu"))
            tabMain.SelectedTab = tabPage1
            c1dateVoucherDate.Focus()
            Return False
        End If
        If tdbcVoucherTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Loai_phieu"))
            tabMain.SelectedTab = tabPage1
            tdbcVoucherTypeID.Focus()
            Return False
        End If
        If txtVoucherNo.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("So_phieu"))
            tabMain.SelectedTab = tabPage1
            txtVoucherNo.Focus()
            Return False
        End If

        Dim dr() As DataRow = dtGrid1.Select("IsNotAllocate = 0")
        If dr.Length > 0 Then
            If txtServiceLife.Text.Trim = "" Then
                D99C0008.MsgNotYetEnter(rL3("So_ky_khau_hao"))
                txtServiceLife.Focus()
                Return False
            End If
            If txtServiceLife.Text.Trim <> "" And Number(txtServiceLife.Text.Trim) = 0 Then
                D99C0008.MsgL3(rL3("So_ky_khau_hao_phai_lon_hon_0"))
                txtServiceLife.Focus()
                Return False
            End If
            If txtServiceLife.Text.Trim.Length > MaxInt Then
                D99C0008.MsgL3(rL3("So_vuot_qua_gioi_han"))
                txtServiceLife.Focus()
                Return False
            End If
        End If

        If txtDepreciatedPeriod.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("So_ky_da_khau_hao"))
            txtDepreciatedPeriod.Focus()
            Return False
        End If

        If txtDepreciatedPeriod.Text.Trim.Length > MaxInt Then
            D99C0008.MsgL3(rL3("So_vuot_qua_gioi_han"))
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

        If Number(txtConvertedAmount.Text) < Number(txtAmountDepreciation.Text) Then
            'Mức khấu hao không được lớn hơn nguyên giá
            D99C0008.MsgL3(rL3("Muc_khau_hao_khong_duoc_lon_hon_nguyen_gia"))
            tabMain.SelectedTab = tabPage1
            tdbg1.SplitIndex = SPLIT0
            tdbg1.Col = COL1_ConvertedAmount
            tdbg1.Bookmark = 0
            Return False
        End If

        If tdbg1.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tabMain.SelectedTab = tabPage1
            tdbg1.Focus()
            Return False
            ' Else
            'For i As Integer = 0 To tdbg1.RowCount - 1
            '    If tdbg1(i, COL1_DataID).ToString = "0" Then
            '        bIsOriginalAmount = True
            '        Exit For
            '    End If
            'Next
        End If

        'If Not bIsOriginalAmount Then
        '    D99C0008.MsgNotYetEnter(rl3("Nguyen_gia"))
        '    tabMain.SelectedTab = tabPage1
        '    tdbg1.SplitIndex = SPLIT0
        '    tdbg1.Col = COL1_DataName
        '    Return False
        'End If
        For i As Integer = 0 To tdbg1.RowCount - 1

            If tdbg1(i, COL1_DataName).ToString = "" Then
                D99C0008.MsgNotYetEnter(rL3("Du_lieu"))
                tdbg1.SplitIndex = SPLIT0
                tdbg1.Col = COL1_DataName
                tdbg1.Bookmark = i
                Return False
            End If

            If tdbg1(i, COL1_ObjectTypeID).ToString <> "" Then
                If tdbg1(i, COL1_ObjectID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rL3("Doi_tuong"))
                    tabMain.SelectedTab = tabPage1
                    tdbg1.SplitIndex = SPLIT0
                    tdbg1.Col = COL1_ObjectID
                    tdbg1.Bookmark = i
                    tdbg1.Focus()
                    Return False
                End If
            End If

            If tdbg1(i, COL1_ConvertedAmount).ToString = "" Then
                D99C0008.MsgNotYetEnter(rL3("So_tien"))
                tabMain.SelectedTab = tabPage1
                tdbg1.SplitIndex = SPLIT0
                tdbg1.Col = COL1_ConvertedAmount
                tdbg1.Bookmark = i
                tdbg1.Focus()
                Return False
            End If
            If tdbg1(i, COL1_ConvertedAmount).ToString <> "" Then
                If Number(tdbg1(i, COL1_ConvertedAmount).ToString) > MaxMoney Then
                    D99C0008.MsgL3(rL3("So_vuot_qua_gioi_han"))
                    tabMain.SelectedTab = tabPage1
                    tdbg1.SplitIndex = SPLIT0
                    tdbg1.Col = COL1_ConvertedAmount
                    tdbg1.Bookmark = i
                    tdbg1.Focus()
                    Return False
                End If
            End If

            If tdbg1(i, COL1_Num01).ToString <> "" Then
                If Number(tdbg1(i, COL1_Num01).ToString) > MaxMoney Then
                    D99C0008.MsgL3(rL3("So_vuot_qua_gioi_han"))
                    tabMain.SelectedTab = tabPage1
                    tdbg1.SplitIndex = SPLIT1
                    tdbg1.Col = COL1_Num01
                    tdbg1.Bookmark = i
                    tdbg1.Focus()
                    Return False
                End If
            End If

            If tdbg1(i, COL1_Num02).ToString <> "" Then
                If Number(tdbg1(i, COL1_Num02).ToString) > MaxMoney Then
                    D99C0008.MsgL3(rL3("So_vuot_qua_gioi_han"))
                    tabMain.SelectedTab = tabPage1
                    tdbg1.SplitIndex = SPLIT1
                    tdbg1.Col = COL1_Num02
                    tdbg1.Bookmark = i
                    tdbg1.Focus()
                    Return False
                End If
            End If
            If tdbg1(i, COL1_Num03).ToString <> "" Then
                If Number(tdbg1(i, COL1_Num03).ToString) > MaxMoney Then
                    D99C0008.MsgL3(rL3("So_vuot_qua_gioi_han"))
                    tabMain.SelectedTab = tabPage1
                    tdbg1.SplitIndex = SPLIT1
                    tdbg1.Col = COL1_Num03
                    tdbg1.Bookmark = i
                    tdbg1.Focus()
                    Return False
                End If
            End If
            If tdbg1(i, COL1_Num04).ToString <> "" Then
                If Number(tdbg1(i, COL1_Num04).ToString) > MaxMoney Then
                    D99C0008.MsgL3(rL3("So_vuot_qua_gioi_han"))
                    tabMain.SelectedTab = tabPage1
                    tdbg1.SplitIndex = SPLIT1
                    tdbg1.Col = COL1_Num04
                    tdbg1.Bookmark = i
                    tdbg1.Focus()
                    Return False
                End If
            End If
            If tdbg1(i, COL1_Num05).ToString <> "" Then
                If Number(tdbg1(i, COL1_Num05).ToString) > MaxMoney Then
                    D99C0008.MsgL3(rL3("So_vuot_qua_gioi_han"))
                    tabMain.SelectedTab = tabPage1
                    tdbg1.SplitIndex = SPLIT1
                    tdbg1.Col = COL1_Num05
                    tdbg1.Bookmark = i
                    tdbg1.Focus()
                    Return False
                End If
            End If
        Next

        Dim dTtalPercent As Double = 0
        If tdbg2.RowCount = 0 Then
            D99C0008.MsgNotYetEnter(rL3("Ma_phan_bo"))
            tabMain.SelectedTab = tabPage2
            tdbg2.SplitIndex = SPLIT0
            tdbg2.Col = COL2_AssignmentID
            tdbg2.Bookmark = 0
            Return False
        End If
        For i As Integer = 0 To tdbg2.RowCount - 1

            If tdbg2(i, COL2_AssignmentID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rL3("Ma_phan_bo"))
                tabMain.SelectedTab = tabPage2
                tdbg2.SplitIndex = SPLIT0
                tdbg2.Col = COL2_AssignmentID
                tdbg2.Bookmark = i
                Return False
            End If

            If tdbg2(i, COL2_PercentAmount).ToString <> "" Then
                dTtalPercent = dTtalPercent + CDbl(tdbg2(i, COL2_PercentAmount).ToString)
            End If
        Next
        If tdbcAssetID.Columns("AssignmentTypeID").Text = "0" Then
            If dTtalPercent <> 100 Then
                D99C0008.MsgL3(rL3("Tong_ty_le_phai_bang_100U"))
                tabMain.SelectedTab = tabPage2
                tdbg2.SplitIndex = SPLIT0
                tdbg2.Col = COL2_PercentAmount
                tdbg2.Bookmark = 0
                Return False
            End If
        ElseIf tdbcAssetID.Columns("AssignmentTypeID").Text = "2" Then
            Dim bCond1 As Boolean = False
            Dim bCond2 As Boolean = False
            Dim iCond2 As Integer = 0
            For i As Integer = 0 To tdbg2.RowCount - 1
                If tdbg2(i, COL2_Extend).ToString = "1" Then bCond1 = True
                If tdbg2(i, COL2_Extend).ToString = "2" Then
                    iCond2 = iCond2 + 1
                End If
            Next
            If iCond2 = 1 Then bCond2 = True
            If bCond1 = False Then
                D99C0008.MsgL3(rL3("Ban_phai_chon_tai_khoan_phan_bo"))
                tabMain.SelectedTab = tabPage2
                tdbg2.SplitIndex = SPLIT0
                tdbg2.Col = COL2_AssignmentID
                tdbg2.Bookmark = 0
                Return False
            End If

            If bCond2 = False Then
                D99C0008.MsgL3(rL3("Ban_phai_chon_tai_khoan_phan_bo_chenh_lech"))
                tabMain.SelectedTab = tabPage2
                tdbg2.SplitIndex = SPLIT0
                tdbg2.Col = COL2_AssignmentID
                tdbg2.Bookmark = 0
                Return False
            End If
        End If
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

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0001
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 09/10/2006 11:46:40
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0001_1() As StringBuilder
        Dim sUseMonth, sUseYear, sDepMonth, sDepYear As String
        If c1dateBeginUsing.Text = "" Then
            sUseMonth = ""
        Else
            sUseMonth = c1dateBeginUsing.Text.Substring(0, 2)
        End If

        If c1dateBeginUsing.Text = "" Then
            sUseYear = ""
        Else
            sUseYear = c1dateBeginUsing.Text.Substring(3, 4)
        End If

        If c1dateBeginDep.Text = "" Then
            sDepMonth = ""
        Else
            sDepMonth = c1dateBeginDep.Text.Substring(0, 2)
        End If

        If c1dateBeginDep.Text = "" Then
            sDepYear = ""
        Else
            sDepYear = c1dateBeginDep.Text.Substring(3, 4)
        End If

        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0001 Set ")
        sSQL.Append("ObjectTypeID = " & SQLString(tdbcObjectTypeID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ObjectID = " & SQLString(tdbcObjectID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("EmployeeID = " & SQLString(tdbcEmployeeID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("FullNameU = " & SQLStringUnicode(txtEmployeeName.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        'sSQL.Append("ConvertedAmount = " & SQLMoney(txtConvertedAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("DepreciatedAmount = " & SQLMoney(txtDepreciateAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        'sSQL.Append("RemainAmount = " & SQLMoney(txtRemainAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("DepreciatedPeriod = " & SQLNumber(txtDepreciatedPeriod.Text) & COMMA) 'int, NULL
        sSQL.Append("Percentage = " & SQLMoney(txtPercentage.Text, DxxFormat.D08_RatioDecimals) & COMMA) 'money, NULL
        sSQL.Append("AmountDepreciation = " & SQLMoney(txtAmountDepreciation.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("ServiceLife = " & SQLNumber(txtServiceLife.Text) & COMMA) 'int, NULL
        sSQL.Append("IsCompleted = " & SQLNumber(1) & COMMA) 'bit, NOT NULL
        sSQL.Append("IsRevalued = " & SQLNumber(0) & COMMA) 'bit, NOT NULL
        sSQL.Append("IsDisposed = " & SQLNumber(0) & COMMA) 'bit, NOT NULL
        sSQL.Append("SetUpFrom = 'BAL'" & COMMA) 'varchar[20], NULL
        sSQL.Append("UseDate = " & SQLDateTimeSave(c1dateUseDate.Value) & COMMA) 'datetime, NULL
        If sUseMonth = "" Then
            sSQL.Append("UseMonth = 0" & COMMA) 'int, NULL
        Else
            sSQL.Append("UseMonth = " & SQLNumber(sUseMonth) & COMMA) 'int, NULL
        End If
        If sUseYear = "" Then
            sSQL.Append("UseYear = 0" & COMMA) 'int, NULL
        Else
            sSQL.Append("UseYear = " & SQLNumber(sUseYear) & COMMA) 'int, NULL
        End If
        If sDepMonth = "" Then
            sSQL.Append("DepMonth = 0" & COMMA) 'datetime, NULL
        Else
            sSQL.Append("DepMonth = " & SQLNumber(sDepMonth) & COMMA) 'datetime, NULL
        End If
        If sDepYear = "" Then
            sSQL.Append("DepYear = 0" & COMMA) 'datetime, NULL
        Else
            sSQL.Append("DepYear = " & SQLNumber(sDepYear) & COMMA) 'datetime, NULL
        End If
        sSQL.Append("TranMonth = " & SQLNumber(giTranMonth) & COMMA) 'int, NULL
        sSQL.Append("TranYear = " & SQLNumber(giTranYear) & COMMA) 'int, NULL
        sSQL.Append("AssetDate = " & SQLDateTimeSave(c1dateAssetDate.Value) & COMMA) 'datetime, NULL
        '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        sSQL.Append("DepDate = " & SQLDateSave(c1dateDepDate.Text) & vbCrLf) 'datetime, NULL
        '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        sSQL.Append(COMMA & "IsCalPeriodByRate = " & SQLNumber(D02Systems.IsCalPeriodByRate) & vbCrLf) 'int, NULL
        If ReturnValueC1Combo(tdbcManagementObjID) <> "" Then
            sSQL.Append(COMMA & "ManagementObjTypeID = " & SQLString(ReturnValueC1Combo(tdbcManagementObjTypeID)))
            sSQL.Append(COMMA & "ManagementObjID = " & SQLString(ReturnValueC1Combo(tdbcManagementObjID)))
        End If
        sSQL.Append(" Where ")
        sSQL.Append("AssetID = " & SQLString(tdbcAssetID.Text))
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5000
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 09/10/2009 02:08:00
    '# Modified User: 
    '# Modified Date: 
    '# Description: Lưu lịch sử quản lý
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T5000_ManageHistory() As StringBuilder
        Dim sSQL As New StringBuilder
        Dim sHistoryID As String = ""

        If _FormState = EnumFormState.FormAdd Then
            sCreateUserID = gsUserID
            sCreateDate = sGetDate
        End If
        sSQL.Append("--Luu lich su quan ly" & vbCrLf)
        sSQL.Append("Insert Into D02T5000(")
        sSQL.Append("HistoryID, DivisionID, AssetID, BatchID, ")
        sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, HistoryTypeID, Status, InstanceID,")
        sSQL.Append(" ObjectTypeID, ObjectID, EmployeeID, ")
        sSQL.Append("CreateUserID,CreateDate, LastModifyUserID,LastModifyDate,FullNameU, ManagementObjTypeID, ManagementObjID, IsManagement")
        sHistoryID = CreateIGE("D02T5000", "HistoryID", "02", "HC", gsStringKey)
        'If bCheckIsManagement = True Then
        '    sHistoryID = CreateIGE("D02T5000", "HistoryID", "02", "HC", gsStringKey)

        'End If
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(sHistoryID) & COMMA) 'HistoryID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'AssetID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'BatchID, varchar[20], NOT NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'BeginYear, smallint, NOT NULL
        sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, smallint, NOT NULL
        sSQL.Append("'OB'" & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
        sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NULL
        sSQL.Append(SQLNumber(0) & COMMA) 'InstanceID, tinyint, NOT NULL
        If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" AndAlso bCheckIsManagement = True Then
            sSQL.Append(SQLString("") & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString("") & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString("") & COMMA) 'EmployeeID, varchar[20], NULL
        Else
            sSQL.Append(SQLString(tdbcObjectTypeID.Text) & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbcObjectID.Text) & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString(tdbcEmployeeID.Text) & COMMA) 'EmployeeID, varchar[20], NULL
        End If
        sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        sSQL.Append(SQLDateTimeSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
        sSQL.Append(SQLDateTimeSave(sGetDate) & COMMA) 'LastModifyDate, datetime, NULL
        If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" AndAlso bCheckIsManagement = True Then
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcManagementObjTypeID)) & COMMA)
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcManagementObjID)) & COMMA)
            sSQL.Append(SQLNumber(1))
            bCheckIsManagement = False
        Else
            sSQL.Append(SQLStringUnicode(txtEmployeeName.Text, gbUnicode, True) & COMMA) 'FullNameU, varchar[250], NULL
            sSQL.Append(SQLString("") & COMMA)
            sSQL.Append(SQLString("") & COMMA)
            sSQL.Append(SQLNumber(0))
        End If
        sSQL.Append(")")
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5000s
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 09/10/2009 03:33:36
    '# Modified User: 
    '# Modified Date: 
    '# Description: Lịch sử phân bổ
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T5000s_AssignHistory() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        Dim sHistoryID As String = ""
        Dim iFirstIGE As Long
        If _FormState = EnumFormState.FormAdd Then
            sCreateUserID = gsUserID
            sCreateDate = sGetDate
        End If
        For i As Integer = 0 To tdbg2.RowCount - 1
            'sHistoryID = CreateIGEs("D02T5000", "HistoryID", "02", "HC", gsStringKey, sHistoryID, tdbg2.RowCount)
            sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HC", gsStringKey, sHistoryID, tdbg2.RowCount, iFirstIGE)

            sSQL.Append("--Luu lich su phan bo" & vbCrLf)
            sSQL.Append("Insert Into D02T5000(")
            sSQL.Append("HistoryID, DivisionID, AssetID, BatchID, ")
            sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, HistoryTypeID, ")
            sSQL.Append("Status, InstanceID,  PercentAmount,AssignmentID, ")
            sSQL.Append("CreateUserID, CreateDate, LastModifyUserID, ")
            sSQL.Append("LastModifyDate")
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(sHistoryID) & COMMA) 'HistoryID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'AssetID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'BatchID, varchar[20], NOT NULL
            sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
            sSQL.Append(SQLNumber(giTranYear) & COMMA) 'BeginYear, smallint, NOT NULL
            sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
            sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, smallint, NOT NULL
            sSQL.Append("'AS'" & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
            sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NULL
            sSQL.Append(SQLNumber(0) & COMMA) 'InstanceID, tinyint, NOT NULL
            sSQL.Append(SQLMoney(tdbg2(i, COL2_PercentAmount), DxxFormat.DefaultNumber2) & COMMA) 'PercentAmount, money, NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_AssignmentID)) & COMMA) 'AssignmentID, varchar[20], NULL
            sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
            sSQL.Append(SQLDateTimeSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
            sSQL.Append(SQLDateTimeSave(sGetDate)) 'LastModifyDate, datetime, NULL
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5000
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 09/10/2009 03:47:17
    '# Modified User: 
    '# Modified Date: 
    '# Description: Lưu lịch sử tác động
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T5000_EffectHistory() As StringBuilder
        Dim sSQL As New StringBuilder
        Dim sRet As New StringBuilder
        Dim sBeginDepMonth As String = ""
        Dim sBeginDepYear As String = ""
        Dim sBeginUseMonth As String = ""
        Dim sBeginUseYear As String = ""
        Dim iFirstIGE As Long
        If _FormState = EnumFormState.FormAdd Then
            sCreateUserID = gsUserID
            sCreateDate = sGetDate
        End If

        If c1dateBeginDep.Text = "" Then
            sBeginDepMonth = ""
            sBeginDepYear = ""
        Else
            sBeginDepMonth = c1dateBeginDep.Text.Substring(0, 2)
            sBeginDepYear = c1dateBeginDep.Text.Substring(3, 4)
        End If

        If c1dateBeginUsing.Text = "" Then
            sBeginUseMonth = ""
            sBeginUseYear = ""
        Else
            sBeginUseMonth = c1dateBeginUsing.Text.Substring(0, 2)
            sBeginUseYear = c1dateBeginUsing.Text.Substring(3, 4)
        End If
        Dim sHistoryID As String = ""
        For i As Integer = 0 To 2
            'sHistoryID = CreateIGEs("D02T5000", "HistoryID", "02", "HC", gsStringKey, sHistoryID, 3)
            sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HC", gsStringKey, sHistoryID, 3, iFirstIGE)
            sSQL.Append("--Luu lich su tac dong" & vbCrLf)
            sSQL.Append("Insert Into D02T5000(")
            sSQL.Append("HistoryID, DivisionID, AssetID, BatchID, ")
            sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, Status, InstanceID,HistoryTypeID, ")
            sSQL.Append("IsStopDepreciation, IsStopUse, ServiceLife, ")
            sSQL.Append("CreateUserID, CreateDate, LastModifyUserID,LastModifyDate")
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(sHistoryID) & COMMA) 'HistoryID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'AssetID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'BatchID, varchar[20], NOT NULL
            If i = 0 Then
                If sBeginDepMonth = "" Then
                    sSQL.Append(0 & COMMA) 'BeginMonth, tinyint, NOT NULL
                    sSQL.Append(0 & COMMA) 'BeginYear, smallint, NOT NULL
                Else
                    sSQL.Append(SQLNumber(sBeginDepMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
                    sSQL.Append(SQLNumber(sBeginDepYear) & COMMA) 'BeginYear, smallint, NOT NULL
                End If
            Else
                If sBeginUseMonth = "" Then
                    sSQL.Append(0 & COMMA) 'BeginMonth, tinyint, NOT NULL
                    sSQL.Append(0 & COMMA) 'BeginYear, smallint, NOT NULL
                Else
                    sSQL.Append(SQLNumber(sBeginUseMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
                    sSQL.Append(SQLNumber(sBeginUseYear) & COMMA) 'BeginYear, smallint, NOT NULL
                End If
            End If
            sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
            sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, smallint, NOT NULL
            sSQL.Append(0 & COMMA) 'InstanceID, tinyint, NOT NULL
            sSQL.Append(0 & COMMA) 'Status, tinyint, NULL
            If i = 0 Then
                sSQL.Append("'SD'" & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
                sSQL.Append(0 & COMMA) 'IsStopDepreciation, tinyint, NULL
                sSQL.Append("NULL" & COMMA) 'IsStopUse, tinyint, NULL
                sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL
            ElseIf i = 1 Then
                sSQL.Append("'SU'" & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
                sSQL.Append("NULL" & COMMA) 'IsStopDepreciation, tinyint, NULL
                sSQL.Append(0 & COMMA) 'IsStopUse, tinyint, NULL
                sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL
            Else
                sSQL.Append("'SL'" & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
                sSQL.Append("NULL" & COMMA) 'IsStopDepreciation, tinyint, NULL
                sSQL.Append("NULL" & COMMA) 'IsStopUse, tinyint, NULL
                If txtServiceLife.Text = "" Then
                    sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL
                Else
                    sSQL.Append(SQLNumber(txtServiceLife.Text) & COMMA) 'ServiceLife, int, NULL
                End If
            End If
            sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
            sSQL.Append(SQLDateTimeSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
            sSQL.Append(SQLDateTimeSave(sGetDate)) 'LastModifyDate, datetime, NULL
            sSQL.Append(")")
            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet

    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5000
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 12/10/2009 11:01:35
    '# Modified User: 
    '# Modified Date: 
    '# Description: Lịch sử thanh lý
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T5000_LiquiDateHistory() As StringBuilder
        Dim sSQL As New StringBuilder
        Dim sHistoryID As String = ""
        If _FormState = EnumFormState.FormAdd Then
            sCreateUserID = gsUserID
            sCreateDate = sGetDate
        End If
        sHistoryID = CreateIGE("D02T5000", "HistoryID", "02", "HC", gsStringKey)
        sSQL.Append("--Luu lich su thanh ly" & vbCrLf)
        sSQL.Append("Insert Into D02T5000(")
        sSQL.Append("HistoryID, DivisionID, AssetID, BatchID, ")
        sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, HistoryTypeID, ")
        sSQL.Append("Status, ObjectTypeID, ObjectID, EmployeeID, ")
        sSQL.Append("CreateUserID,CreateDate, LastModifyUserID, ")
        sSQL.Append("LastModifyDate,IsLiquidated,FullNameU")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(sHistoryID) & COMMA) 'HistoryID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'AssetID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'BatchID, varchar[20], NOT NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'BeginYear, smallint, NOT NULL
        sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, smallint, NOT NULL   16/12/213 id 62153     '   sSQL.Append(SQLNumber(2009) & COMMA) 'EndYear, smallint, NOT NULL

        sSQL.Append("'IL'" & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
        sSQL.Append(0 & COMMA) 'Status, tinyint, NULL
        sSQL.Append(SQLString("") & COMMA) 'ObjectTypeID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'ObjectID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'EmployeeID, varchar[20], NULL
        sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        sSQL.Append(SQLDateTimeSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
        sSQL.Append(SQLDateTimeSave(sGetDate) & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(0 & COMMA)  'IsLiquidated, tinyint, NOT NULL
        sSQL.Append(SQLStringUnicode("")) 'FullNameU, varchar[250], NULL
        sSQL.Append(")")
        Return sSQL
    End Function

    Private Function SQLInsertD02T5000_AssetAccount(sHistoryID As String) As StringBuilder
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        Dim sSQL As New StringBuilder

        If _FormState = EnumFormState.FormAdd Then
            sCreateUserID = gsUserID
            sCreateDate = sGetDate
        End If

        sSQL.Append("--Luu dong lich su TK tai san " & vbCrLf)
        sSQL.Append("Insert Into D02T5000(")
        sSQL.Append("HistoryID, DivisionID, AssetID, BatchID, ")
        sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, HistoryTypeID, ")
        sSQL.Append("Status, ObjectTypeID, ObjectID, EmployeeID, ")
        sSQL.Append("CreateUserID,CreateDate, LastModifyUserID, ")
        sSQL.Append("LastModifyDate,IsLiquidated,FullNameU, AssetAccountID, DepAccountID")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(sHistoryID) & COMMA) 'HistoryID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'AssetID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'BatchID, varchar[20], NOT NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'BeginYear, smallint, NOT NULL
        sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, smallint, NOT NULL   16/12/213 id 62153     '   sSQL.Append(SQLNumber(2009) & COMMA) 'EndYear, smallint, NOT NULL

        sSQL.Append("'AAC'" & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
        sSQL.Append(0 & COMMA) 'Status, tinyint, NULL
        sSQL.Append(SQLString("") & COMMA) 'ObjectTypeID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'ObjectID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'EmployeeID, varchar[20], NULL
        sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        sSQL.Append(SQLDateTimeSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
        sSQL.Append(SQLDateTimeSave(sGetDate) & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(0 & COMMA)  'IsLiquidated, tinyint, NOT NULL
        sSQL.Append(SQLStringUnicode("") & COMMA) 'FullNameU, varchar[250], NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID, "AssetAccountID")) & COMMA) 'AssetAccountID, varchar[20], NULL
        sSQL.Append(SQLString("")) 'DepAccountID, varchar[20], NULL
        sSQL.Append(") " & vbCrLf)
        Return sSQL
    End Function

    Private Function SQLInsertD02T5000_DepAccount(sHistoryID As String) As StringBuilder
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        Dim sSQL As New StringBuilder

        If _FormState = EnumFormState.FormAdd Then
            sCreateUserID = gsUserID
            sCreateDate = sGetDate
        End If

        sSQL.Append("--Luu dong lich su TK khau hao " & vbCrLf)
        sSQL.Append("Insert Into D02T5000(")
        sSQL.Append("HistoryID, DivisionID, AssetID, BatchID, ")
        sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, HistoryTypeID, ")
        sSQL.Append("Status, ObjectTypeID, ObjectID, EmployeeID, ")
        sSQL.Append("CreateUserID,CreateDate, LastModifyUserID, ")
        sSQL.Append("LastModifyDate,IsLiquidated,FullNameU, AssetAccountID, DepAccountID")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(sHistoryID) & COMMA) 'HistoryID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'AssetID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'BatchID, varchar[20], NOT NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'BeginYear, smallint, NOT NULL
        sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, smallint, NOT NULL   16/12/213 id 62153     '   sSQL.Append(SQLNumber(2009) & COMMA) 'EndYear, smallint, NOT NULL

        sSQL.Append("'DAC'" & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
        sSQL.Append(0 & COMMA) 'Status, tinyint, NULL
        sSQL.Append(SQLString("") & COMMA) 'ObjectTypeID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'ObjectID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'EmployeeID, varchar[20], NULL
        sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        sSQL.Append(SQLDateTimeSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
        sSQL.Append(SQLDateTimeSave(sGetDate) & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(0 & COMMA)  'IsLiquidated, tinyint, NOT NULL
        sSQL.Append(SQLStringUnicode("") & COMMA) 'FullNameU, varchar[250], NULL
        sSQL.Append(SQLString("") & COMMA) 'AssetAccountID, varchar[20], NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID, "DepAccountID"))) 'DepAccountID, varchar[20], NULL
        sSQL.Append(") " & vbCrLf)
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5010
    '# Created User: 
    '# Created Date: 17/11/2021 05:22:24
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
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'BatchID, varchar[50], NOT NULL
        sSQL.Append(SQLString(sHistoryTypeID) & COMMA) 'HistoryTypeID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA & vbCrLf) 'AssetID, varchar[50], NOT NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'BeginMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'BeginYear, int, NOT NULL
        sSQL.Append(SQLNumber(12) & COMMA) 'EndMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(9999) & COMMA) 'EndYear, int, NOT NULL
        sSQL.Append(SQLString("") & COMMA & vbCrLf) 'GroupID, varchar[50], NOT NULL
        sSQL.Append(SQLString(sAccountID)) 'AccountID, varchar[50], NOT NULL
        sSQL.Append(") " & vbCrLf)

        Return sSQL
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0012s
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 12/10/2009 11:36:21
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0012s(Optional ByVal sTransID As String = "", Optional ByVal iFirstTrans As Long = 0) As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder


        Dim iCountIGETransID As Long 'Số dòng cần sinh IGE
        If _FormState = EnumFormState.FormAdd Then
            sCreateUserID = gsUserID
            sCreateDate = sGetDate
            iCountIGETransID = tdbg1.RowCount
        Else
            iCountIGETransID = dtGrid1.Select("TransactionID is Null or TransactionID=''").Length
        End If


        For i As Integer = 0 To tdbg1.RowCount - 1
            sSQL.Append("Insert Into D02T0012(")
            sSQL.Append("TransactionID, DivisionID, ModuleID, AssetID, VoucherTypeID, ")
            sSQL.Append("VoucherNo, VoucherDate, TranMonth, TranYear, ")
            sSQL.Append("DescriptionU, CurrencyID, ExchangeRate, DebitAccountID, ")
            sSQL.Append("CreditAccountID, OriginalAmount, ConvertedAmount, Status, TransactionTypeID, ")
            sSQL.Append(" Disabled, CreateUserID, CreateDate, ")
            sSQL.Append("LastModifyUserID, LastModifyDate, ObjectTypeID, ObjectID, BatchID, ")
            sSQL.Append("Ana01ID, Ana02ID, Ana03ID, ")
            sSQL.Append("Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, ")
            sSQL.Append("Ana09ID, Ana10ID, NotesU, Posted, SourceID,")
            sSQL.Append("Str01U, Str02U, Str03U, Str04U, Str05U, ")
            sSQL.Append("Num01, Num02, Num03, Num04, Num05, ")
            sSQL.Append("Date01, Date02, Date03, Date04, Date05,IsNotAllocate ")
            sSQL.Append(") Values(")


            If tdbg1(i, COL1_TransactionID).ToString = "" Then
                sTransID = CreateIGENewS("D02T0012", "TransactionID", "02", "TJ", gsStringKey, sTransID, iCountIGETransID, iFirstTrans)
                tdbg1(i, COL1_TransactionID) = sTransID
            End If


            sSQL.Append(SQLString(tdbg1(i, COL1_TransactionID)) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
            sSQL.Append("'02'" & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL

            sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'AssetID, varchar[20], NULL
            sSQL.Append(SQLString(tdbcVoucherTypeID.Text) & COMMA) 'VoucherTypeID, varchar[20], NULL
            sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[20], NULL
            sSQL.Append(SQLDateSave(c1dateVoucherDate.Text) & COMMA) 'VoucherDate, datetime, NULL
            sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
            sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL

            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Description), gbUnicode, True) & COMMA)  'DescriptionU, varchar[500], NULL
            sSQL.Append(SQLString(txtCurrenyID.Text) & COMMA) 'CurrencyID, varchar[20], NOT NULL
            sSQL.Append(SQLMoney(1, DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL
            If tdbg1(i, COL1_DataID).ToString = "0" Then
                sSQL.Append(SQLString(tdbg1(i, COL1_AccountID)) & COMMA) 'DebitAccountID, varchar[20], NULL
                sSQL.Append("null" & COMMA) 'CreditAccountID, varchar[20], NULL
            Else
                sSQL.Append("null" & COMMA) 'DebitAccountID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg1(i, COL1_AccountID)) & COMMA) 'CreditAccountID, varchar[20], NULL
            End If
            sSQL.Append(SQLMoney(tdbg1(i, COL1_ConvertedAmount), DxxFormat.DecimalPlaces) & COMMA) 'OriginalAmount, money, NULL
            sSQL.Append(SQLMoney(tdbg1(i, COL1_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL
            sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NOT NULL
            If tdbg1(i, COL1_DataID).ToString = "0" Then
                sSQL.Append("'SD'" & COMMA) 'TransactionTypeID, varchar[20], NULL
            Else
                sSQL.Append("'SDKH'" & COMMA) 'TransactionTypeID, varchar[20], NULL
            End If
            sSQL.Append(0 & COMMA) 'Disabled, bit, NOT NULL
            sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
            sSQL.Append(SQLDateTimeSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
            sSQL.Append(SQLDateTimeSave(sGetDate) & COMMA) 'LastModifyDate, datetime, NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_ObjectTypeID)) & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_ObjectID)) & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString(tdbcAssetID.Text) & COMMA) 'BatchID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana01ID)) & COMMA) 'Ana01ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana02ID)) & COMMA) 'Ana02ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana03ID)) & COMMA) 'Ana03ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana04ID)) & COMMA) 'Ana04ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana05ID)) & COMMA) 'Ana05ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana06ID)) & COMMA) 'Ana06ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana07ID)) & COMMA) 'Ana07ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana08ID)) & COMMA) 'Ana08ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana09ID)) & COMMA) 'Ana09ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana10ID)) & COMMA) 'Ana10ID, varchar[20], NULL
            sSQL.Append(SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'Notes, varchar[250], NULL
            sSQL.Append(SQLNumber(chkPosted.Checked) & COMMA) 'Posted, tinyint, NOT NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_SourceID)) & COMMA) 'SourceID, varchar[20], NULL

            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str01), gbUnicode, True) & COMMA) 'Str01U, varchar[250], NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str02), gbUnicode, True) & COMMA) 'Str02U, varchar[250], NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str03), gbUnicode, True) & COMMA) 'Str03U, varchar[250], NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str04), gbUnicode, True) & COMMA) 'Str04U, varchar[250], NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str05), gbUnicode, True) & COMMA) 'Str05U, varchar[250], NULL
            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num01), DxxFormat.DefaultNumber2) & COMMA) 'Num01, money, NULL
            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num02), DxxFormat.DefaultNumber2) & COMMA) 'Num02, money, NULL
            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num03), DxxFormat.DefaultNumber2) & COMMA) 'Num03, money, NULL
            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num04), DxxFormat.DefaultNumber2) & COMMA) 'Num04, money, NULL
            sSQL.Append(SQLMoney(tdbg1(i, COL1_Num05), DxxFormat.DefaultNumber2) & COMMA) 'Num05, money, NULL
            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date01)) & COMMA) 'Date01, datetime, NULL
            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date02)) & COMMA) 'Date02, datetime, NULL
            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date03)) & COMMA) 'Date03, datetime, NULL
            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date04)) & COMMA) 'Date04, datetime, NULL
            sSQL.Append(SQLDateSave(tdbg1(i, COL1_Date05)) & COMMA) 'Date05, datetime, NULL
            sSQL.Append(SQLNumber(tdbg1(i, COL1_IsNotAllocate))) 'IsNotAllocate, datetime, NULL
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0001
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 12/10/2009 02:18:02
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0001_2() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0001 Set ")
        sSQL.Append("ObjectTypeID = " & SQLString(tdbcObjectTypeID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ConvertedAmount = " & SQLMoney(txtConvertedAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("AmountDepreciation = " & SQLMoney(txtAmountDepreciation.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("RemainAmount = " & SQLMoney(txtRemainAmount.Text, DxxFormat.D90_ConvertedDecimals)) 'money, NULL
        sSQL.Append(" Where ")
        sSQL.Append("AssetID = " & SQLString(tdbcAssetID.Text))
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T5000
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 12/10/2009 02:29:18
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T5000() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T5000"
        sSQL &= " Where "
        sSQL &= "AssetID = " & SQLString(_assetID) & " And DivisionID=" & SQLString(gsDivisionID)
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0012
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 12/10/2009 02:31:24
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0012() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T0012"
        sSQL &= " Where "
        sSQL &= " AssetID= " & SQLString(_assetID) & " And DivisionID = " & SQLString(gsDivisionID) & " And " & vbCrLf
        sSQL &= "IsNull(TransactionTypeID,'') IN ('SD', 'SDKH') And ModuleID = '02' " & vbCrLf

        sSQL &= "Delete From D02T0012"
        sSQL &= " Where "
        sSQL &= " BatchID= " & SQLString(_assetID) & " And DivisionID = " & SQLString(gsDivisionID) & " And " & vbCrLf
        sSQL &= "IsNull(TransactionTypeID,'') IN ('SD', 'SDKH') And ModuleID = '02'"
        Return sSQL
    End Function

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        LoadTDBCAssetID()
        LoadAddNew()
        btnNext.Enabled = False
        btnSave.Enabled = True
        tdbcAssetID.Focus()
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Cap_nhat_thong_tin_so_du_tai_san_co_dinh_-_D02F1002") & UnicodeCaption(gbUnicode) 'CËp nhËt th¤ng tin sç d§ tªi s¶n cç ¢Ünh - D02F1002
        '================================================================ 
        lblAssetID.Text = rL3("Ma_tai_san") 'Mã tài sản

        lblEmployeeID.Text = rL3("Nguoi_tiep_nhan") 'Người tiếp nhận
        lblteVoucherDate.Text = rL3("Ngay_chung_tu") 'Ngày chứng từ
        lblVoucherTypeID.Text = rL3("Loai_phieu") 'Loại phiếu
        lblVoucherNo.Text = rL3("So_phieu") 'Số phiếu
        lblCurrenyID.Text = rL3("Loai_tien") 'Loại tiền
        lblDescription.Text = rL3("Dien_giai") 'Diễn giải
        lblMethodID.Text = rL3("Phuong_phap_KH") 'Phương pháp KH
        lblMethodEndID.Text = rL3("Khau_hao_ky_cuoi") 'Khấu hao kỳ cuối
        lblDeprTableName.Text = rL3("Bang_khau_hao") 'Bảng khấu hao
        lblConvertedAmount.Text = rL3("Nguyen_gia") 'Nguyên giá
        lblAmountDepreciation.Text = rL3("Hao_mon_luy_ke") 'Hao mòn luỹ kế
        lblRemainAmount.Text = rL3("Gia_tri_con_lai") 'Giá trị còn lại
        lblServiceLife.Text = rL3("So_ky_khau_hao") 'Số kỳ khấu hao
        lblDepreciatedPeriod.Text = rL3("So_ky_da_khau_hao") 'Số kỳ đã khấu hao
        lblPercentage.Text = rL3("Ty_le_khau_hao_%") 'Tỷ lệ khấu hao %
        lblteUseDate.Text = rL3("Ngay_su_dung") 'Ngày sử dụng
        lblteAssetDate.Text = rL3("Ngay_tiep_nhan") 'Ngày tiếp nhận
        lblDepreciateAmount.Text = rL3("Muc_khau_hao") 'Mức khấu hao
        lblteBeginUsing.Text = rL3("Ky_bat_dau_su_dung") 'Kỳ bắt đầu sử dụng
        lblteBeginDep.Text = rL3("Ky_bat_dau_KH") 'Kỳ bắt đầu KH
        lblObjectTypeID.Text = rL3("Bo_phan_tiep_nhan") 'Bộ phận tiếp nhận
        lblDepDate.Text = rL3("Ngay_bat_dau_khau_hao") 'Ngày bắt đầu khấu hao
        '================================================================ 
        btnHotKeys.Text = rL3("Phim_nong") 'Phím nóng
        btnSave.Text = rL3("_Luu") '&Lưu
        btnNext.Text = rL3("_Nhap_tiep") 'Nhập &tiếp
        btnClose.Text = rL3("Do_ng") 'Đó&ng
        '================================================================ 
        chkPosted.Text = rL3("Chuyen_but_toan_sang_tong_hop") 'Chuyển bút toán sang tổng hợp
        '================================================================ 
        grpAssetID.Text = rL3("Ma_tai_san") 'Mã tài sản
        grpFinancialInfo.Text = rL3("Thong_tin_tai_chinh") 'Thông tin tài chính
        '================================================================ 
        tabPage1.Text = "1. " & rL3("Dinh_khoan") '1. Định khoản
        tabPage2.Text = "2. " & rL3("Phan_bo_khau_hao") '2. Phân bổ khấu hao
        '================================================================ 
        tdbcEmployeeID.Columns("EmployeeID").Caption = rL3("Ma") 'Mã
        tdbcEmployeeID.Columns("EmployeeName").Caption = rL3("Ten") 'Tên

        tdbcObjectID.Columns("ObjectTypeID").Caption = rL3("Loai_DT") 'Loại ĐT
        tdbcObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên

        tdbcObjectTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcObjectTypeID.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcAssetID.Columns("AssetID").Caption = rL3("Ma") 'Mã
        tdbcAssetID.Columns("AssetName").Caption = rL3("Ten") 'Tên

        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rL3("Ma") 'Mã
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rL3("Dien_giai") 'Diễn giải

        '================================================================ 
        tdbdAna10ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna10ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdSourceID.Columns("SourceID").Caption = rL3("Ma") 'Mã
        tdbdSourceID.Columns("SourceName").Caption = rL3("Ten") 'Tên
        tdbdAna09ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna09ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna08ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna08ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbdObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        tdbdAna07ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna07ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdObjectTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbdObjectTypeID.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbdAna06ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna06ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna05ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna05ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna04ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna04ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna03ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna03ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna02ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna02ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna01ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna01ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAssignmentID.Columns("AssignmentID").Caption = rL3("Ma") 'Mã
        tdbdAssignmentID.Columns("AssignmentName").Caption = rL3("Ten") 'Tên

        '================================================================ 
        tdbg1.Columns(COL1_DataName).Caption = rL3("Du_lieu") 'Dữ liệu
        tdbg1.Columns(COL1_IsNotAllocate).Caption = rL3("Khong_tinh_KH") 'Không tính KH
        tdbg1.Columns(COL1_Description).Caption = rL3("Dien_giai") 'Diễn giải
        tdbg1.Columns(COL1_ObjectTypeID).Caption = rL3("Loai_doi_tuong1") 'Loại đối tượng
        tdbg1.Columns(COL1_ObjectID).Caption = rL3("Doi_tuong") 'Đối tượng
        tdbg1.Columns(COL1_AccountID).Caption = rL3("Ma_tai_khoan") 'Mã tài khoản
        tdbg1.Columns(COL1_ConvertedAmount).Caption = rL3("So_tien") 'Số tiền
        tdbg1.Columns(COL1_SourceID).Caption = rL3("Nguon_hinh_thanh") 'Nguồn hình thành

        tdbg2.Columns("AssignmentID").Caption = rL3("Ma_phan_bo") 'Mã phân bổ
        tdbg2.Columns("AssignmentName").Caption = rL3("Ten_phan_bo") 'Tên phân bổ
        tdbg2.Columns("DebitAccountID").Caption = rL3("TK_no") 'TK Nợ
        tdbg2.Columns("PercentAmount").Caption = rL3("Ty_le") 'Tỷ lệ
        '================================================================ 
        tdbcManagementObjTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcManagementObjTypeID.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcManagementObjID.Columns("ObjectTypeID").Caption = rL3("Loai_DT") 'Loại ĐT
        tdbcManagementObjID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcManagementObjID.Columns("ObjectName").Caption = rL3("Ten") 'Tên
    End Sub

    Private Sub tdbg1_OnAddNew(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbg1.OnAddNew
        tdbg1.Columns(COL1_IsNotAllocate).Value = 0
    End Sub

    'Incident 	73638
#Region "Chuẩn hóa sinh số phiếu"
    Dim sOldVoucherNo As String = "" 'Lưu lại số phiếu cũ
    Dim bEditVoucherNo As Boolean = False '= True: có nhấn F2; = False: không
    Dim bFirstF2 As Boolean = False 'Nhấn F2 lần đầu tiên
    Dim iPer_F5558 As Integer = 0 'Phân quyền cho Sửa số phiếu


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
        If _FormState = EnumFormState.FormAdd Then
            If tdbcVoucherTypeID.Columns("Auto").Text = "1" Then 'Sinh tự động
                txtVoucherNo.Text = CreateIGEVoucherNo(tdbcVoucherTypeID, False)
                ReadOnlyControl(txtVoucherNo)
            Else 'Không sinh tự động
                txtVoucherNo.Text = ""
                UnReadOnlyControl(txtVoucherNo, True)
            End If

        End If
    End Sub

    Private Sub txtVoucherNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVoucherNo.KeyDown
        If e.KeyCode = Keys.F2 Then
            'Loại phiếu hay Số phiếu = "" thì thoát
            If tdbcVoucherTypeID.Text = "" Or txtVoucherNo.Text = "" Then Exit Sub
            'Update 21/09/2010: Trường hợp Thêm mới phiếu và đã lưu Thành công thì không cho sửa Số phiếu
            If _FormState = EnumFormState.FormAdd And btnSave.Enabled = False Then Exit Sub
            'Kiểm tra quyền cho trường hợp Sửa
            If _FormState = EnumFormState.FormEdit And iPer_F5558 <= 2 Then Exit Sub
            'Cho sửa Số phiếu ở trạng thái Thêm mới hay Sửa
            If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormEdit Then
                'Trước khi gọi exe con thì nhớ lại Số phiếu cũ
                If bFirstF2 = False Then
                    sOldVoucherNo = txtVoucherNo.Text
                    bFirstF2 = True
                End If
                'Gọi exe con D91E0640
                'Dim frm As New D91F5558
                'With frm
                '    .FormName = "D91F5558"
                '    .FormPermission = "D02F5558" 'Màn hình phân quyền
                '    .ModuleID = D02 'Mã module hiện tại, VD: D22
                '    .TableName = "D02T0012" 'Tên bảng chứa số phiếu
                '    'Update 21/09/2010
                '    If _FormState = EnumFormState.FormAdd Then
                '        .VoucherID = "" 'Khóa sinh IGE là rỗng
                '    ElseIf _FormState = EnumFormState.FormEdit Then
                '        .VoucherID = tdbg1(0, COL1_TransactionID).ToString 'Khóa sinh IGE
                '    End If
                '    .VoucherNo = txtVoucherNo.Text 'Số phiếu cần sửa
                '    .Mode = "0" ' Tùy theo Module, mặc định là 0
                '    .KeyID01 = ""
                '    .KeyID02 = ""
                '    .KeyID03 = ""
                '    .KeyID04 = ""
                '    .KeyID05 = ""
                '    .ShowDialog()
                '    Dim sVoucherNo As String
                '    sVoucherNo = .Output02
                '    .Dispose()
                '    If sVoucherNo <> "" Then
                '        txtVoucherNo.Text = sVoucherNo 'Giá trị trả về Số phiếu mới
                '        ReadOnlyControl(txtVoucherNo) 'Lock text Số phiếu
                '        bEditVoucherNo = True 'Đã nhấn F2
                '        gbSavedOK = True
                '    End If
                'End With

                Dim arrPro() As StructureProperties = Nothing
                SetProperties(arrPro, "FormIDPermission", "D02F5558")
                SetProperties(arrPro, "VoucherTypeID", ReturnValueC1Combo(tdbcVoucherTypeID))
                If _FormState = EnumFormState.FormAdd Then
                    SetProperties(arrPro, "VoucherID", "")
                ElseIf _FormState = EnumFormState.FormEdit Then
                    SetProperties(arrPro, "VoucherID", tdbg1(0, COL1_TransactionID).ToString)
                End If
                SetProperties(arrPro, "Mode", 0)
                SetProperties(arrPro, "KeyID01", "")
                SetProperties(arrPro, "TableName", "D02T0012")
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
                    gbSavedOK = True
                End If
            End If
        End If
    End Sub
#End Region

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

End Class
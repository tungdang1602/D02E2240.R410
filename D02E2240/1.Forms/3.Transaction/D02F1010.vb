Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text
Imports System.IO
Imports System

Public Class D02F1010

#Region "Const of tdbg1"
    Private Const COL1_BatchID As String = "BatchID"                             ' BatchID
    Private Const COL1_TransactionID As String = "TransactionID"                 ' TransactionID
    Private Const COL1_CopyAmount As String = "CopyAmount"                       ' CopyAmount
    Private Const COL1_CipID As String = "CipID"                                 ' Mã XDCB
    Private Const COL1_IsNotAllocate As String = "IsNotAllocate"                 ' Không tính KH
    Private Const COL1_RefDate As String = "RefDate"                             ' Ngày hóa đơn
    Private Const COL1_SerialNo As String = "SerialNo"                           ' Số Sêri
    Private Const COL1_RefNo As String = "RefNo"                                 ' Số hóa đơn
    Private Const COL1_ObjectTypeID As String = "ObjectTypeID"                   ' Loại đối tượng
    Private Const COL1_ObjectID As String = "ObjectID"                           ' Đối tượng
    Private Const COL1_DebitAccountID As String = "DebitAccountID"               ' TK nợ
    Private Const COL1_CreditAccountID As String = "CreditAccountID"             ' TK có
    Private Const COL1_OriginalAmount As String = "OriginalAmount"               ' Số tiền nguyên tệ
    Private Const COL1_ConvertedAmount As String = "ConvertedAmount"             ' Số tiền quy đổi
    Private Const COL1_Description As String = "Description"                     ' Diễn giải
    Private Const COL1_SourceID As String = "SourceID"                           ' Nguồn hình thành
    Private Const COL1_InitialCovertedAmount As String = "InitialCovertedAmount" ' InitialCovertedAmount
    Private Const COL1_Str01 As String = "Str01"                                 ' Văn bản 01
    Private Const COL1_Str02 As String = "Str02"                                 ' Văn bản 02
    Private Const COL1_Str03 As String = "Str03"                                 ' Văn bản 03
    Private Const COL1_Str04 As String = "Str04"                                 ' Văn bản 04
    Private Const COL1_Str05 As String = "Str05"                                 ' Văn bản 05
    Private Const COL1_Num01 As String = "Num01"                                 ' Số 01
    Private Const COL1_Num02 As String = "Num02"                                 ' Số 02
    Private Const COL1_Num03 As String = "Num03"                                 ' Số 03
    Private Const COL1_Num04 As String = "Num04"                                 ' Số 04
    Private Const COL1_Num05 As String = "Num05"                                 ' Số 05
    Private Const COL1_Date01 As String = "Date01"                               ' Ngày 01
    Private Const COL1_Date02 As String = "Date02"                               ' Ngày 02
    Private Const COL1_Date03 As String = "Date03"                               ' Ngày 03
    Private Const COL1_Date04 As String = "Date04"                               ' Ngày 04
    Private Const COL1_Date05 As String = "Date05"                               ' Ngày 05
    Private Const COL1_Ana01ID As String = "Ana01ID"                             ' Khoản mục 01
    Private Const COL1_Ana02ID As String = "Ana02ID"                             ' Khoản mục 02
    Private Const COL1_Ana03ID As String = "Ana03ID"                             ' Khoản mục 03
    Private Const COL1_Ana04ID As String = "Ana04ID"                             ' Khoản mục 04
    Private Const COL1_Ana05ID As String = "Ana05ID"                             ' Khoản mục 05
    Private Const COL1_Ana06ID As String = "Ana06ID"                             ' Khoản mục 06
    Private Const COL1_Ana07ID As String = "Ana07ID"                             ' Khoản mục 07
    Private Const COL1_Ana08ID As String = "Ana08ID"                             ' Khoản mục 08
    Private Const COL1_Ana09ID As String = "Ana09ID"                             ' Khoản mục 09
    Private Const COL1_Ana10ID As String = "Ana10ID"                             ' Khoản mục 10
    Private Const COL1_IsEdit As String = "IsEdit"                               ' IsEdit
#End Region


#Region "Const of tdbg2"
    Private Const COL2_AssignmentID As String = "AssignmentID"     ' Mã phân bổ
    Private Const COL2_AssignmentName As String = "AssignmentName" ' Tên phân bổ
    Private Const COL2_DebitAccountID As String = "DebitAccountID" ' TK Nợ
    Private Const COL2_PercentAmount As String = "PercentAmount"   ' Tỷ lệ
    Private Const COL2_Extend As String = "Extend"                 ' Extend
    Private Const COL2_HistoryID As String = "HistoryID"           ' HistoryID
#End Region

    '---Kiểm tra khoản mục theo chuẩn gồm 7 bước
    '--- Chuẩn Khoản mục b1: Khai báo biến
    '-------Biến khai báo cho khoản mục
    Private Const SplitAna As Int16 = 2 ' Ghi nhận Khoản mục chứa ở Split nào
    Dim bUseAna As Boolean 'Kiểm tra có sử dụng Khoản mục không, để set thuộc tính Enabled nút Khoản mục 
    'Dim iDisplayAnaCol As Integer = 0 ' Cột Khoản mục đầu tiên được hiển thị, khi nhấn nút Khoản mục thì Focus đến cột đó
    Dim xCheckAna(9) As Boolean 'Khởi động tại Form_load: Ghi lại việc kiểm tra lần đầu Lưu, khi nhấn Lưu lần thứ 2 thì không cần kiểm tra nữa
    '------------------------------------------------------------------------------------------------
    Dim sAuditCode As String = "CIPToAsset"
    Dim sCreateUserID As String
    Dim sCreateDate As String
    'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b1:
    Dim sEditVoucherTypeID As String = ""
    Dim _setupFrom As String = "CIP"
    Dim iPerD02F0100 As Integer = ReturnPermission("D02F0100")
    Dim iPerD02F0087 As Integer = ReturnPermission("D02F0087") 'ID : 224617 - BỔ SUNG Cho phép gọi màn hình THIẾT LẬP DANH MỤC TÀI SẢN CỐ ĐỊNH tại bước hình thành TS

    Dim clsFilterCombo As Lemon3.Controls.FilterCombo
    Dim clsFilterDropdown As Lemon3.Controls.FilterDropdown
    Dim dtAnaCaption, dtObjectID As DataTable
    Dim dtManagementID As DataTable
    Private _assetID As String = ""

    Public Property AssetID() As String
        Get
            Return _assetID
        End Get
        Set(ByVal Value As String)
            _assetID = Value
        End Set
    End Property
    Private _fromCall As String = ""
    Public WriteOnly Property FromCall() As String 
        Set(ByVal Value As String )
            _fromCall = Value
        End Set
    End Property

    Private _d54ProjectID As String = ""
    Public WriteOnly Property D54ProjectID() As String 
        Set(ByVal Value As String )
            _d54ProjectID = Value
        End Set
    End Property

    Private _d27PropertyProductID As String = ""
    Public WriteOnly Property D27PropertyProductID() As String 
        Set(ByVal Value As String )
            _d27PropertyProductID = Value
        End Set
    End Property
    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
           
            _FormState = value
            '--- Chuẩn Khoản mục b2: Lấy caption cho 10 khoản mục
            bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg1, IndexOfColumn(tdbg1, COL1_Ana01ID), SplitAna, True, gbUnicode, dtAnaCaption)
            SetNewXaCheckAna()
            '--- Chuẩn Khoản mục b21: D91 có sử dụng Khoản mục
            If bUseAna Then
                '--- Chuẩn Khoản mục b3: Load 10 khoản mục
                LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg1, IndexOfColumn(tdbg1, COL1_Ana01ID), gbUnicode)
            Else
                tdbg1.RemoveHorizontalSplit(SplitAna)
            End If
            '------------------------------------
            Dim bUseSubInfo As Boolean = LoadCaptionSubInfo() 'load caption các thông tin phụ
            If Not bUseSubInfo Then tdbg1.RemoveHorizontalSplit(SPLIT1)
            LoadTDBCombo()
            LoadTDBDropDown()
            LoadtdbdAssignmentID()
            'ID 78489 07/08/2015
            'clsFilterCombo = New Lemon3.Controls.FilterCombo()
            'clsFilterCombo.CheckD91 = True
            'clsFilterCombo.UseFilterCombo(tdbcPropertyProductID)

            clsFilterCombo = New Lemon3.Controls.FilterCombo
            clsFilterCombo.CheckD91 = True 'Giá trị mặc định True: kiểm tra theo DxxFormat.LoadFormNotINV. Ngược lại luôn luôn Filter dạng mới (dùng cho Novaland)
            clsFilterCombo.AddPairObject(tdbcObjectTypeID, tdbcObjectID) 'Đã bổ sung cột Loại ĐT
            clsFilterCombo.AddPairObject(tdbcManagementObjTypeID, tdbcManagementObjID)
            clsFilterCombo.UseFilterComboObjectID()
            clsFilterCombo.UseFilterCombo(tdbcAssetID, tdbcEmployeeID, tdbcPropertyProductID)
            clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
            clsFilterCombo.LoadtdbcObjectID(tdbcManagementObjID, dtManagementID, ReturnValueC1Combo(tdbcManagementObjTypeID))
            ' Dim dic As New Dictionary(Of String, String)
            ' dic.Add(tdbdBudgetID.Name, "Note")'Ví dụ Cần lấy cột Note trong tdbdBudgetID. Nếu lấy Name, hoặc Description hoặc cột 1 thì không cần truyền
            clsFilterDropdown = New Lemon3.Controls.FilterDropdown()
            'clsFilterDropdown.SingleLine = True'Mặc đinh False. Chọn nhiều dòng gắn lại dữ liệu cho 1 dòng và cách nhau bằng ; (sử dụng Tài khoản D90F1110)
            clsFilterDropdown.CheckD91 = True 'Giá trị mặc định True: kiểm tra theo DxxFormat.LoadFormNotINV. Ngược lại luôn luôn Filter dạng mới (dùng cho Novaland)
            ' clsFilterDropdown.DicDDName = dic
            clsFilterDropdown.UseFilterDropdown(tdbg1, COL1_Ana01ID, COL1_Ana02ID, COL1_Ana03ID, COL1_Ana04ID, COL1_Ana05ID, COL1_Ana06ID, COL1_Ana07ID, COL1_Ana08ID, COL1_Ana09ID, COL1_Ana10ID, COL1_ObjectID)
            clsFilterDropdown.UseFilterDropdown(tdbg2, COL2_AssignmentID) 'Nếu dùng nhiều lưới

            txtPercentage.Tag = ""

            Select Case _FormState
                Case EnumFormState.FormAdd
                    LoadVoucherTypeID(tdbcVoucherTypeID, D02, sEditVoucherTypeID, gbUnicode)
                    LoadAddNew()
                    LoadTabpage1()
                    LoadTDBGrid2()
                    tdbg1.Enabled = False
                Case EnumFormState.FormCopy
                    LoadVoucherTypeID(tdbcVoucherTypeID, D02, sEditVoucherTypeID, gbUnicode)
                    LoadInherit()
                    LoadAddNew()
                    LoadTDBGrid2()
                    tdbg1.Enabled = False
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

    Private Sub SetNewXaCheckAna()
        ReDim xCheckAna(9)
        'Dim i As Integer
        'For i = 0 To 9
        '    xCheckAna(i) = False
        'Next i
    End Sub

    Private Sub D02F1010_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '12/10/2020, id 144622-Tài sản cố định_Lỗi chưa cảnh báo khi lưu
        If _FormState = EnumFormState.FormEdit Then
            If Not gbSavedOK Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            If (tdbcAssetID.Text <> "" Or tdbcObjectTypeID.Text <> "" Or tdbcObjectID.Text <> "" Or tdbcEmployeeID.Text <> "" Or c1dateVoucherDate.Text <> "" Or tdbcVoucherTypeID.Text <> "" Or txtVoucherNo.Text <> "" Or txtServiceLife.Text <> "" Or txtDepreciatedPeriod.Text <> "" Or c1dateBeginUsing.Text <> "" Or c1dateBeginDep.Text <> "" Or c1dateDepDate.Text <> "") Then
                If Not gbSavedOK Then
                    If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub D02F1002_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                'If DxxFormat.LoadFormNotINV = 1 Then
                '    If Me.ActiveControl.Parent IsNot Nothing AndAlso Me.ActiveControl.Parent.Name = tdbcPropertyProductID.Name Then Exit Sub
                'End If
                UseEnterAsTab(Me)
                Exit Sub
            Case Keys.F11
                HotKeyF11(Me, tdbg1)
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

    Dim sOldPropertyProductID As String
    Dim sOldAssetID As String
    Private Sub D02F1002_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        gbSavedOK = False
        CheckIdTextBox(txtVoucherNo, txtVoucherNo.MaxLength)
        InputbyUnicode(Me, gbUnicode)
        Loadlanguage()
        SetBackColorObligatory()
        txtCurrenyID.Text = DxxFormat.BaseCurrencyID
        iUseInvoiceCodeD02 = GetUseInvoiceCodeD02()
        'Nhập Ngày
        InputDateInTrueDBGrid(tdbg1, COL1_RefDate, COL1_Date01, COL1_Date02, COL1_Date03, COL1_Date04, COL1_Date05)
        '*************
        tdbg2_LockedColumns()
        'Nhập số
        Dim arr() As FormatColumn = Nothing
        AddNumberColumns(arr, SqlDbType.Money, COL1_OriginalAmount, "N" & DxxFormat.iD90_ConvertedDecimals)
        AddNumberColumns(arr, SqlDbType.Money, COL1_ConvertedAmount, "N" & DxxFormat.iD90_ConvertedDecimals)
        InputNumber(tdbg1, arr)
        '------------------------
        Dim arrPercent() As FormatColumn = Nothing
        AddPercentColumns(arrPercent, COL2_PercentAmount)
        InputNumber(tdbg2, arrPercent)
        '*****************
        If D02Options.LockConvertedAmount Then
            tdbg1.Splits(SPLIT0).DisplayColumns(COL1_ConvertedAmount).Locked = True
            tdbg1.Splits(SPLIT0).DisplayColumns(COL1_ConvertedAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        End If
        If Not D02Systems.CIPforPropertyProduct Then
            txtAssetName.Width = txtAssetName.Width + pnlProperty.Width
        End If

        pnlProperty.Visible = D02Systems.CIPforPropertyProduct
        tdbg1.Splits(0).DisplayColumns(COL1_IsNotAllocate).Visible = D02Systems.UseProperty
        sOldPropertyProductID = ReturnValueC1Combo(tdbcPropertyProductID)
        sOldAssetID = ReturnValueC1Combo(tdbcAssetID)

        '*****************

        ResetFooterGrid(tdbg1, 0, tdbg1.Splits.Count - 1)
        ResetFooterGrid(tdbg2, 0, tdbg2.Splits.Count - 1)
        LockServiceLife() '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

#Region "Events of tdbg1"

    Private Sub HeadClick(ByVal iCol As Integer)
        Select Case tdbg1.Columns(iCol).DataField
            Case COL1_OriginalAmount, COL1_ConvertedAmount
                Exit Sub
            Case Else
                CopyColumns(tdbg1, tdbg1.Col, tdbg1(tdbg1.Row, tdbg1.Col).ToString, tdbg1.Row)
        End Select

    End Sub

    Public Sub CopyColumns(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ColCopy As Integer, ByVal sValue As String, ByVal RowCopy As Int32)
        Try
            'If sValue = "" Or c1Grid.RowCount < 2 Then Exit Sub
            c1Grid.UpdateData()
            If c1Grid.RowCount < 2 Then Exit Sub

            sValue = c1Grid(RowCopy, ColCopy).ToString

            Dim Flag As DialogResult
            Flag = MessageBox.Show(rl3("Copy_cot_du_lieu_cho") & vbCrLf & rl3("____-_Tat_ca_cac_cot_(nhan_Yes)") & vbCrLf & rl3("____-_Nhung_dong_con_trong_(nhan_No)"), MsgAnnouncement, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

            If Flag = Windows.Forms.DialogResult.No Then ' Copy nhung dong con trong

                For i As Integer = RowCopy + 1 To c1Grid.RowCount - 1
                    If L3Int(c1Grid(i, COL1_IsEdit)) = 1 Then Continue For
                    If c1Grid(i, ColCopy).ToString = "" OrElse c1Grid(i, ColCopy).ToString = MaskFormatDateShort OrElse c1Grid(i, ColCopy).ToString = MaskFormatDate OrElse (L3IsNumeric(c1Grid(i, ColCopy).ToString) And Val(c1Grid(i, ColCopy).ToString) = 0) Then c1Grid(i, ColCopy) = sValue
                Next
                'c1Grid(RowCopy, ColCopy) = sValue

            ElseIf Flag = Windows.Forms.DialogResult.Yes Then ' Copy het
                For i As Integer = RowCopy + 1 To c1Grid.RowCount - 1
                    If L3Int(c1Grid(i, COL1_IsEdit)) = 1 Then Continue For
                    c1Grid(i, ColCopy) = sValue
                Next
                'c1Grid(0, ColCopy) = sValue
            Else
                Exit Sub
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub tdbg1_AfterDelete(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg1.AfterDelete
        If dtGrid1 IsNot Nothing And dtGrid1.Rows.Count > 0 Then txtConvertedAmount.Text = FormatNumber(dtGrid1.Compute("SUM(" & COL1_ConvertedAmount & ")", ""), DxxFormat.iD90_ConvertedDecimals)
    End Sub

    Private Sub tdbg1_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.HeadClick
        HeadClick(e.ColIndex)
    End Sub

    'Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
    '    Select Case e.KeyCode
    '        Case Keys.F5 'Xóa dữ liệu dòng hiện tại
    '            If D99C0008.MsgAskDelete = Windows.Forms.DialogResult.No Then Exit Sub
    '            For i As Integer = 1 To tdbg1.Columns.Count - 1
    '                tdbg1.Columns(i).Text = ""
    '            Next
    '            tdbg1.UpdateData()
    '        Case Keys.F7 'Copy cell
    '            Select Case tdbg1.Columns(tdbg1.Col).DataField
    '                Case COL1_ObjectTypeID, COL1_ObjectID
    '                    HotKeyF7(tdbg1, New Integer() {IndexOfColumn(tdbg1, COL1_ObjectTypeID), IndexOfColumn(tdbg1, COL1_ObjectID)})
    '                Case COL1_OriginalAmount, COL1_ConvertedAmount
    '                    HotKeyF7(tdbg1, New Integer() {IndexOfColumn(tdbg1, COL1_OriginalAmount), IndexOfColumn(tdbg1, COL1_ConvertedAmount)})
    '                Case Else
    '                    HotKeyF7(tdbg1)
    '            End Select
    '        Case Keys.F8 'Copy dòng
    '            HotKeyF8(tdbg1)
    '    End Select
    '    If e.Shift And e.KeyCode = Keys.Insert Then
    '        HotKeyShiftInsert(tdbg1, 0, IndexOfColumn(tdbg1, COL1_CipID), tdbg1.Columns.Count)
    '        Exit Sub
    '    End If
    '    If e.Control And e.KeyCode = Keys.S Then
    '        HeadClick(tdbg1.Col)
    '        Exit Sub
    '    End If
    '    If e.Control And e.KeyCode = Keys.Delete Then
    '        If L3Int(tdbg1.Columns(COL1_IsEdit).Text) = 1 Then Exit Sub
    '    End If
    '    HotKeyDownGrid(e, tdbg1, IndexOfColumn(tdbg1, COL1_CipID), SPLIT0, tdbg1.Splits.Count - 1, , , , IndexOfColumn(tdbg1, COL1_Description), txtDescription.Text)
    'End Sub

    Private Sub tdbg1_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.ComboSelect
        tdbg1.UpdateData()
    End Sub

    'Private Sub tdbg1_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg1.RowColChange
    'If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
    '    '--- Đổ nguồn cho các Dropdown phụ thuộc
    '    Select Case tdbg1.Columns(tdbg1.Col).DataField
    '        Case COL1_ObjectID
    '            LoadDataSource(tdbdObjectID, LoadObjectID(tdbg1.Columns(IndexOfColumn(tdbg1, COL1_ObjectTypeID)).Text), gbUnicode)
    '            'Case COL1_CipID
    '            '    tdbg1.Splits(0).DisplayColumns(tdbg1.Col).Button = L3Int(tdbg1(tdbg1.Row, COL1_IsEdit)) <> 1
    '    End Select
    'End Sub

    Private Sub tdbg1_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg1.BeforeColUpdate

        If tdbg1.Columns(COL1_TransactionID).Text = "" Then
            Select Case e.Column.DataColumn.DataField
                Case COL1_CipID
                    If tdbg1.Columns(e.ColIndex).Value.ToString <> tdbdCipID.Columns(COL1_CipID).Text OrElse ((_FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy) And tdbcAssetID.Text = "") Then tdbg1.Columns(e.ColIndex).Text = ""
                Case COL1_RefDate
                    If tdbg1.Columns(e.ColIndex).Text = "" Then tdbg1.Columns(e.ColIndex).Value = Now.Date
            End Select
        End If

        Select Case e.ColIndex
            Case IndexOfColumn(tdbg1, COL1_SerialNo), IndexOfColumn(tdbg1, COL1_RefNo)
                e.Cancel = L3IsID(tdbg1, e.ColIndex)
            Case IndexOfColumn(tdbg1, COL1_ObjectTypeID)
                If tdbg1.Columns(COL1_ObjectTypeID).Text <> tdbdObjectTypeID.Columns("ObjectTypeID").Text Then
                    tdbg1.Columns(COL1_ObjectTypeID).Text = ""
                    tdbg1.Columns(COL1_ObjectID).Text = ""
                End If
            Case IndexOfColumn(tdbg1, COL1_ObjectID)
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg1.Columns(e.ColIndex).Text <> tdbg1.Columns(e.ColIndex).DropDown.Columns("ObjectID").Text Then
                    tdbg1.Columns(e.ColIndex).Text = ""
                End If
            Case IndexOfColumn(tdbg1, COL1_DebitAccountID), IndexOfColumn(tdbg1, COL1_CreditAccountID), IndexOfColumn(tdbg1, COL1_SourceID)
                If tdbg1.Columns(e.ColIndex).Text <> tdbg1.Columns(e.ColIndex).DropDown.Columns(tdbg1.Columns(e.ColIndex).DropDown.DisplayMember).Text Then
                    tdbg1.Columns(e.ColIndex).Text = ""
                End If
            Case IndexOfColumn(tdbg1, COL1_OriginalAmount)
                If Number(txtExchangRate.Text) > 0 Then tdbg1.Columns(COL1_ConvertedAmount).Text = (Number(tdbg1.Columns(e.ColIndex).Text) * Number(txtExchangRate.Text)).ToString
                'Case IndexOfColumn(tdbg1, COL1_Ana01ID) To IndexOfColumn(tdbg1, COL1_Ana10ID)
                '    If tdbg1.Columns(e.ColIndex).Text = tdbg1.Columns(e.ColIndex).DropDown.Columns("AnaID").Text Then Exit Select
                '    Dim index As Integer = e.ColIndex - IndexOfColumn(tdbg1, COL1_Ana01ID) 'index của mảng giá trị
                '    If gbArrAnaValidate(index) Then 'Kiểm tra nhập trong danh sách
                '        tdbg1.Columns(e.ColIndex).Text = ""
                '    Else
                '        ' Kiểm tra chiều dài nhập vào
                '        If tdbg1.Columns(e.ColIndex).Text.Length > giArrAnaLength(index) Then tdbg1.Columns(e.ColIndex).Text = ""
                '    End If
            Case IndexOfColumn(tdbg1, COL1_Ana01ID) To IndexOfColumn(tdbg1, COL1_Ana10ID) 'Có nhập ngoài danh sách không bỏ
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg1.Columns(e.ColIndex).Text <> tdbg1.Columns(e.ColIndex).DropDown.Columns("AnaID").Text Then
                    CheckAfterColUpdateAna(tdbg1, IndexOfColumn(tdbg1, COL1_Ana01ID), e.ColIndex, dtAnaCaption) 'tham khảo hàm viết phía dưới
                End If

        End Select

    End Sub

    Private Sub tdbg1_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.AfterColUpdate
        Select Case e.Column.DataColumn.DataField
            Case COL1_CipID
                If tdbg1.Columns(e.ColIndex).Text = "" Then Exit Select
                If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
                    tdbg1.Columns(COL1_OriginalAmount).Text = tdbdCipID.Columns("SumConvertedAmo").Text
                Else
                    tdbg1.Columns(COL1_OriginalAmount).Text = (Number(tdbg1.Columns(COL1_CopyAmount).Text) + Number(tdbdCipID.Columns("SumConvertedAmo").Text)).ToString
                End If
                If tdbg1.Row = 0 Then 'Chỉ xử lý dòng đầu
                    If D02Options.DepAndEmpID Then
                        tdbcEmployeeID.SelectedValue = tdbdCipID.Columns("EmployeeID").Text
                        tdbcObjectTypeID.SelectedValue = tdbdCipID.Columns("ObjectTypeID").Text
                        tdbcObjectID.SelectedValue = tdbdCipID.Columns("ObjectID").Text
                    End If
                    If D02Options.CipDescription Then txtNotes.Text = tdbdCipID.Columns("Description").Text
                    If D02Options.CipName Then txtAssetName.Text = tdbdCipID.Columns("CipName").Text
                End If
                tdbg1.Columns(COL1_DebitAccountID).Text = ReturnValueC1Combo(tdbcAssetID, "AssetAccountID").ToString  'AssetAccountID
                tdbg1.Columns(COL1_CreditAccountID).Text = tdbdCipID.Columns("AccountID").Text
                tdbg1.Columns(COL1_InitialCovertedAmount).Text = tdbdCipID.Columns("SumConvertedAmo").Text
                tdbg1.Columns(COL1_ConvertedAmount).Text = tdbg1.Columns(COL1_OriginalAmount).Text
                tdbg1.UpdateData()
                GoTo 1
            Case COL1_ObjectTypeID
                tdbg1.Columns(COL1_ObjectID).Value = ""
                LoadtdbdObjectID(tdbg1.Columns(COL1_ObjectTypeID).Text)
            Case COL1_ObjectID
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
            Case COL1_ConvertedAmount, COL1_OriginalAmount
1:
                If dtGrid1 IsNot Nothing And dtGrid1.Rows.Count > 0 Then txtConvertedAmount.Text = FormatNumber(dtGrid1.Compute("SUM(" & COL1_ConvertedAmount & ")", ""), DxxFormat.iD90_ConvertedDecimals)
        End Select
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
            Case IndexOfColumn(tdbg1, COL1_ObjectID)
                If dr Is Nothing OrElse dr.Item("ObjectID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    tdbg1.Columns(COL1_ObjectID).Text = ""
                    Exit Sub
                End If
                'Chỉ dùng cho cột là Đối tượng
                If tdbg1.Columns(COL1_ObjectTypeID).Text = "" Then
                    tdbg1(tdbg1.Row, COL1_ObjectTypeID) = dr.Item("ObjectTypeID").ToString
                    LoadTDBDObjectID(tdbg1.Columns(COL1_ObjectTypeID).Text)
                End If
                tdbg1.Columns(COL1_ObjectID).Text = dr.Item("ObjectID").ToString

            Case IndexOfColumn(tdbg1, COL1_Ana01ID) To IndexOfColumn(tdbg1, COL1_Ana10ID)
                If dr Is Nothing OrElse dr.Item("AnaID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    CheckAfterColUpdateAna(tdbg1, IndexOfColumn(tdbg1, COL1_Ana01ID), iCol, dtAnaCaption) 'tham khảo hàm viết phía dưới
                    'tdbg1.Columns(iCol).Text = ""
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
            Case IndexOfColumn(tdbg1, COL1_ObjectID), IndexOfColumn(tdbg1, COL1_Ana01ID) To IndexOfColumn(tdbg1, COL1_Ana10ID)
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg1, tdbg1.Columns(tdbg1.Col).DataField)
                If tdbd Is Nothing Then Exit Select
                If tdbg1.Columns(COL1_ObjectTypeID).Text = "" Then
                    tdbdObjectID.DisplayColumns("ObjectTypeID").Visible = True
                End If
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg1, e, tdbd)
                If dr Is Nothing Then Exit Sub
                AfterColUpdate_tdbg1(tdbg1.Col, dr)
        End Select
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        Select Case e.KeyCode
            Case Keys.F5 'Xóa dữ liệu dòng hiện tại
                If D99C0008.MsgAskDelete = Windows.Forms.DialogResult.No Then Exit Sub
                For i As Integer = 1 To tdbg1.Columns.Count - 1
                    tdbg1.Columns(i).Text = ""
                Next
                tdbg1.UpdateData()
            Case Keys.F7 'Copy cell
                Select Case tdbg1.Columns(tdbg1.Col).DataField
                    Case COL1_ObjectTypeID, COL1_ObjectID
                        HotKeyF7(tdbg1, New Integer() {IndexOfColumn(tdbg1, COL1_ObjectTypeID), IndexOfColumn(tdbg1, COL1_ObjectID)})
                    Case COL1_OriginalAmount, COL1_ConvertedAmount
                        HotKeyF7(tdbg1, New Integer() {IndexOfColumn(tdbg1, COL1_OriginalAmount), IndexOfColumn(tdbg1, COL1_ConvertedAmount)})
                    Case Else
                        HotKeyF7(tdbg1)
                End Select
            Case Keys.F8 'Copy dòng
                HotKeyF8(tdbg1)
        End Select
        If e.Shift And e.KeyCode = Keys.Insert Then
            HotKeyShiftInsert(tdbg1, 0, IndexOfColumn(tdbg1, COL1_CipID), tdbg1.Columns.Count)
            Exit Sub
        End If
        If e.Control And e.KeyCode = Keys.S Then
            HeadClick(tdbg1.Col)
            Exit Sub
        End If
        If e.Control And e.KeyCode = Keys.Delete Then
            If L3Int(tdbg1.Columns(COL1_IsEdit).Text) = 1 Then Exit Sub
        End If
        HotKeyDownGrid(e, tdbg1, IndexOfColumn(tdbg1, COL1_CipID), SPLIT0, tdbg1.Splits.Count - 1, , , , IndexOfColumn(tdbg1, COL1_Description), txtDescription.Text)

        If clsFilterDropdown.CheckKeydownFilterDropdown(tdbg1, e) Then
            Select Case tdbg1.Col
                Case IndexOfColumn(tdbg1, COL1_ObjectID), IndexOfColumn(tdbg1, COL1_Ana01ID) To IndexOfColumn(tdbg1, COL1_Ana10ID)
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg1, tdbg1.Columns(tdbg1.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    If tdbg1.Columns(COL1_ObjectTypeID).Text = "" Then
                        tdbdObjectID.DisplayColumns("ObjectTypeID").Visible = True
                    End If
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg1, e, tdbd)
                    If dr Is Nothing Then Exit Sub
                    AfterColUpdate_tdbg1(tdbg1.Col, dr)
                    Exit Sub
            End Select

        End If
    End Sub

    Private Sub tdbg1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg1.KeyPress
        Select Case tdbg1.Columns(tdbg1.Col).DataField
            Case COL1_SerialNo
                e.KeyChar = UCase(e.KeyChar) 'Nhập các ký tự hoa
        End Select
    End Sub
#End Region

    Private Sub SetBackColorObligatory()
        tdbcAssetID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcObjectID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcEmployeeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateVoucherDate.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtVoucherNo.BackColor = COLOR_BACKCOLOROBLIGATORY
        'txtServiceLife.BackColor = COLOR_BACKCOLOROBLIGATORY
        txtDepreciatedPeriod.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateBeginUsing.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateBeginDep.BackColor = COLOR_BACKCOLOROBLIGATORY
        If D02Systems.IsCalDepByDate = True Then c1dateDepDate.BackColor = COLOR_BACKCOLOROBLIGATORY '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        tdbg1.Splits(0).DisplayColumns(COL1_CipID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbg2.Splits(0).DisplayColumns(COL2_AssignmentID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        If D02Systems.IsObligatoryManagement Then
            tdbcManagementObjTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
            tdbcManagementObjID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        End If
    End Sub

    Private Sub LoadtdbcAssetID()
        Dim sUnicode As String = ""
        If gbUnicode Then sUnicode = "U"
        Dim sSQL As String = ""

        sSQL &= "SELECT '+' As AssetID, N'<Thêm Mới>'  As AssetName,'' As AssetAccountID, '' As DepAccountID, " & vbCrLf 'ID : 224617 - BỔ SUNG Cho phép gọi màn hình THIẾT LẬP DANH MỤC TÀI SẢN CỐ ĐỊNH tại bước hình thành TS
        sSQL &= "'' As ObjectTypeID,'' As ObjectID,'' As EmployeeID,'' As  FullName,0 As ConvertedAmount,0 As Percentage, " & vbCrLf
        sSQL &= " 0 As ServiceLife,0 As AmountDepreciation ,Null As AssetDate,'' As  MethodID,'' As MethodEndID,  " & vbCrLf
        sSQL &= "  null as BeginUsing,  " & vbCrLf
        sSQL &= " null as BeginDep,  " & vbCrLf
        sSQL &= "0 As  DepreciatedPeriod,0 As DepreciatedAmount,0 As IsCompleted,0 As IsRevalued,0 As IsDisposed,  " & vbCrLf
        sSQL &= "0 As AssignmentTypeID,null As DeprTableName, null As UseDate,'' As Notes,null As DepDate, " & vbCrLf
        sSQL &= "'' As ManagementObjTypeID, '' As ManagementObjID" & vbCrLf
        sSQL &= "UNION" & vbCrLf
        sSQL &= "SELECT 	T1.AssetID, T1.AssetName" & UnicodeJoin(gbUnicode) & " as AssetName, T1.AssetAccountID, T1.DepAccountID," & vbCrLf
        If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
            sSQL &= "T1.ObjectTypeID, T1.ObjectID, T1.EmployeeID, T1.FullName" & sUnicode & " as FullName,"
        Else
            sSQL &= "N19.ObjectTypeID, N19.ObjectID, N19.EmployeeID, N19.FullName" & sUnicode & " as FullName,"
        End If
        sSQL &= "N19.CurrentCost as ConvertedAmount, T1.Percentage, N19.ServiceLife, N19.CurrentLTDDepreciation As AmountDepreciation, T1.AssetDate, A.Description" & UnicodeJoin(gbUnicode) & " AS MethodID," & vbCrLf
        sSQL &= " B.Description" & UnicodeJoin(gbUnicode) & " AS MethodEndID, " & vbCrLf
        sSQL &= "Case when T1.UseMonth <10 then '0' else '' end + ltrim(str(T1.UseMonth)) + '/' + ltrim(str(T1.UseYear)) as BeginUsing, "
        sSQL &= "Case when T1.DepMonth <10 then '0' else '' end + ltrim(str(T1.DepMonth)) + '/' + ltrim(str(T1.DepYear)) as BeginDep, "
        sSQL &= "T1.DepreciatedPeriod, T1.DepreciatedAmount," & vbCrLf
        sSQL &= "  T1.IsCompleted,T1.IsRevalued, T1.IsDisposed, T1.AssignmentTypeID, C.DeprTableName" & sUnicode & " as  DeprTableName, T1.UseDate" & vbCrLf
        sSQL &= ", T1.Notes" & UnicodeJoin(gbUnicode) & " as Notes, T1.DepDate, " & vbCrLf
        sSQL &= " T1.ManagementObjTypeID, T1.ManagementObjID" & vbCrLf
        sSQL &= "FROM D02T0001 T1 WITH(NOLOCK) " & vbCrLf
        sSQL &= " Inner Join D02T8000 A WITH(NOLOCK) On T1.MethodID = A.IntCode " & vbCrLf
        sSQL &= "Inner Join D02T8000 B WITH(NOLOCK) On T1.MethodEndID = B.IntCode" & vbCrLf
        sSQL &= " Left Join D02T0070 C WITH(NOLOCK) On T1.DeprTableID = C.DeprTableID " & vbCrLf
        sSQL &= " INNER Join D02N0019(" & giTranMonth & "," & giTranYear & ") N19 On T1.AssetID = N19.AssetID AND T1.DivisionID = N19.DivisionID" & vbCrLf
        sSQL &= " WHERE A.Language = " & SQLString(gsLanguage) & "  AND A.ModuleID = '02' And A.Type = 0 " & _
                        " AND B.Language = " & SQLString(gsLanguage) & "  AND B.ModuleID = '02' And B.Type = 1 " & _
                        " AND  T1.DivisionID = " & SQLString(gsDivisionID)
        If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then sSQL &= "AND  T1.IsCompleted = 0"
        sSQL &= vbCrLf & " ORDER BY AssetID, AssetName"

        LoadDataSource(tdbcAssetID, sSQL, gbUnicode)
        tdbcAssetID.Splits(0).DisplayColumns(0).Visible = True
        tdbcAssetID.Splits(0).DisplayColumns("AssetID").Width = 140
        tdbcAssetID.DropDownWidth = 400
        tdbcAssetID.Splits(0).DisplayColumns(2).Visible = False
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcAssetID
        LoadTDBCAssetID()
        'Load tdbcObjectTypeID/ tdbdObjectTypeID
        Dim dtObjectTypeID As DataTable = ReturnTableObjectTypeID(gbUnicode)
        LoadDataSource(tdbcObjectTypeID, dtObjectTypeID, gbUnicode)
        LoadDataSource(tdbdObjectTypeID, dtObjectTypeID.DefaultView.ToTable, gbUnicode)
        'Load tdbcEmployeeID
        sSQL = "Select ObjectID as EmployeeID, ObjectName" & UnicodeJoin(gbUnicode) & " as EmployeeName From Object WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where ObjectTypeID='NV' Order By ObjectID"
        LoadDataSource(tdbcEmployeeID, sSQL, gbUnicode)
        LoadtdbcPropertyProductID()

        sSQL = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) " & vbCrLf
        sSQL &= " Order By ObjectID "
        dtObjectID = ReturnDataTable(sSQL)

        Dim dtManagementObjTypeID As DataTable = ReturnTableObjectTypeID(gbUnicode)

        LoadDataSource(tdbcManagementObjTypeID, dtManagementObjTypeID, gbUnicode)
        dtManagementID = ReturnDataTable(LoadManagementObjID)

        LoadtdbdObjectID(tdbg1.Columns(COL1_ObjectTypeID).Text)
    End Sub

    Private Function LoadManagementObjID() As String
        Dim sSQL As String = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) " & vbCrLf
        sSQL &= " Order By ObjectID "
        Return sSQL
    End Function
    Private Function LoadObjectID(ByVal ID As String) As String
        Dim sSQL As String = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where ObjectTypeID=" & SQLString(ID) & vbCrLf
        sSQL &= " Order By ObjectID "
        Return sSQL
    End Function

    Private Sub LoadtdbdObjectID(ByVal ID As String)
        If ID = "" Then
            LoadDataSource(tdbdObjectID, dtObjectID, gbUnicode)
            tdbdObjectID.DisplayColumns("ObjectTypeID").Visible = True
        Else
            tdbdObjectID.DisplayColumns("ObjectTypeID").Visible = False
            LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObjectID, " ObjectTypeID = " & SQLString(ID), True), gbUnicode)
        End If

    End Sub


    Dim dtCipID As DataTable
    Private Sub LoadtdbdCipID()
        Dim sSQL As String = "--Do nguon cho ma XDCB " & vbCrLf
        sSQL &= SQLStoreD02P0012(True)
        'sSQL = "SELECT * FROM"
        'sSQL = sSQL & "(SELECT A.CipID, B.CipNo, B.AccountID, B.CipName" & UnicodeJoin(gbUnicode) & " as CipName, Sum(Case When A.DebitAccountID = B.AccountID then ConvertedAmount When a.CreditAccountID = b.AccountID then - ConvertedAmount else 0 end) as SumConvertedAmo, "
        'sSQL = sSQL & "Sum(Case When A.DebitAccountID = B.AccountID then OriginalAmount When a.CreditAccountID = B.AccountID then - OriginalAmount else 0 end) as SumOriginalAmo "
        'sSQL = sSQL & " FROM D02T0012 A INNER JOIN D02T0100 B ON A.CipID = B.CipID AND B.[Status] <> 2  "
        'sSQL = sSQL & " WHERE B.DivisionID = '" & gsDivisionID & "'  And Isnull (A.SplitbatchID,'') = ''"
        'sSQL = sSQL & " GROUP BY A.CipID, B.CipNo, B.AccountID, B.CipName" & UnicodeJoin(gbUnicode)
        'sSQL = sSQL & " HAVING MAX(A.TranMonth + A.TranYear * 100) <= " & giTranMonth & " + " & giTranYear & " * 100 )   A"
        'sSQL = sSQL & " WHERE A.SumConvertedAmo <> 0 "
        dtCipID = ReturnDataTable(sSQL)
        LoadDataSource(tdbdCipID, sSQL, gbUnicode)
    End Sub

    Dim dtPropertyProductID As DataTable
    Private Sub LoadtdbcPropertyProductID()
        Dim sSQL As String = "-- Do nguon cho ma BDS " & vbCrLf
        sSQL &= SQLStoreD02P0012() '"Exec D02P0012 'DIGINET', 2, 2012,1,0,'', 0, ''"
        dtPropertyProductID = ReturnDataTable(sSQL)
        LoadDataSource(tdbcPropertyProductID, dtPropertyProductID, gbUnicode)
    End Sub


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0012
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 21/02/2012 01:26:24
    '# Modified User: 
    '# Modified Date: 
    '# Description: Load tdbdCipID
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0012(Optional ByVal bForCipID As Boolean = False) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0012 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(IIf(_FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy, 0, 1)) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA 'AssetID, varchar[20], NOT NULL
        If bForCipID Then
            sSQL &= SQLNumber(0) & COMMA 'IsPropertyProduct, tinyint, NOT NULL
            sSQL &= SQLString(ReturnValueC1Combo(tdbcPropertyProductID)) 'PropertyProductID, varchar[50], NOT NULL
        Else
            sSQL &= SQLNumber(1) & COMMA 'IsPropertyProduct, tinyint, NOT NULL
            sSQL &= SQLString("") 'PropertyProductID, varchar[50], NOT NULL
        End If

        Return sSQL
    End Function



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0091
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 21/02/2012 01:27:17
    '# Modified User: 
    '# Modified Date: 
    '# Description: Load lưới XDCB. Nơi gọi : Truy vấn -> Chứng từ => menu Hình thành tài sản
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0091() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0091 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function

    Private Sub LoadTDBDropDown()
        LoadtdbdCipID()
        Dim sSQL As String = ""
        'Load AccountID
        Dim dtAccountID As DataTable = ReturnTableAccountID("AccountStatus = 0 ", gbUnicode)
        LoadDataSource(tdbdDebitAccountID, dtAccountID, gbUnicode)
        LoadDataSource(tdbdCreditAccountID, dtAccountID.DefaultView.ToTable, gbUnicode)
        'Load tdbdSourceID 
        sSQL = "SELECT 		SourceID, SourceName" & UnicodeJoin(gbUnicode) & " as SourceName " & vbCrLf
        sSQL &= "FROM 		D02T0013 WITH(NOLOCK)  " & vbCrLf
        sSQL &= "WHERE 		Disabled = 0 " & vbCrLf
        sSQL &= "ORDER BY 	SourceID" & vbCrLf
        LoadDataSource(tdbdSourceID, sSQL, gbUnicode)
        ''Load tdbdAssignmentID
        'sSQL = "Select '+' As AssignmentID, N'<Thêm Mới>'  As AssignmentName,'' As DebitAccountID, '' As Extend " & vbCrLf
        'sSQL &= "UNION" & vbCrLf
        'sSQL &= "SELECT 	AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " as AssignmentName, DebitAccountID, Extend " & vbCrLf
        'sSQL &= "FROM 		D02T0002 WITH(NOLOCK)  " & vbCrLf
        'sSQL &= "WHERE 		Disabled = 0 " & vbCrLf
        'sSQL &= "ORDER BY 	AssignmentID" & vbCrLf
        'LoadDataSource(tdbdAssignmentID, sSQL, gbUnicode)
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

#Region "Events tdbcAssetID with txtAssetName"

    Private Sub tdbcAssetID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetID.KeyDown
        If e.KeyCode <> Keys.F2 OrElse tdbcAssetID.ReadOnly Then Exit Sub
        Try
            If clsFilterCombo.IsNewFilter Then Exit Sub ' TH filter dạng mới thì F2 gọi D99F5555 đã có sẵn
            Dim arrPro() As StructureProperties = Nothing
            SetProperties(arrPro, "InListID", "25")
            SetProperties(arrPro, "WhereValue", tdbcAssetID.Text)
            SetProperties(arrPro, "InWhere", "( AssetName Like '%%' ) And IsCompleted = 0  And DivisionID = " & SQLString(gsDivisionID))
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

    Private Sub tdbcAssetID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetID.SelectedValueChanged
        If tdbcAssetID.SelectedValue Is Nothing Then
            txtAssetName.Text = ""
            tdbcPropertyProductID.Text = ""
        Else
            
            'Thêm ngày 20/7/2012 theo incident 49472 ThiHuan
            If _fromCall = "D02F2005" Then
                If Not D02Options.DepAndEmpID Then
                    'Bỏ ngày 23/7/2012 theo THIHUAN
                    'If tdbcAssetID.Columns("ObjectTypeID").Text <> "" Then
                    tdbcObjectTypeID.SelectedValue = tdbcAssetID.Columns("ObjectTypeID").Value
                    ' LoadDataSource(tdbcObjectID, LoadObjectID(ReturnValueC1Combo(tdbcObjectTypeID, "").ToString), gbUnicode)
                    tdbcObjectID.SelectedValue = tdbcAssetID.Columns("ObjectID").Value.ToString
                    'End If
                    ' If tdbcAssetID.Columns("EmployeeID").Value.ToString <> "" Then

                    'End If
                    tdbcEmployeeID.SelectedValue = tdbcAssetID.Columns("EmployeeID").Value.ToString
                End If
                If Not D02Options.CipName Then
                    txtAssetName.Text = tdbcAssetID.Columns("AssetName").Value.ToString
                End If
                If Not D02Options.CipDescription Then
                    txtNotes.Text = tdbcAssetID.Columns("Notes").Text

                End If
            Else
                txtAssetName.Text = tdbcAssetID.Columns("AssetName").Value.ToString
                If tdbcAssetID.Columns("ObjectTypeID").Text <> "" Then
                    tdbcObjectTypeID.SelectedValue = tdbcAssetID.Columns("ObjectTypeID").Value
                    ' LoadDataSource(tdbcObjectID, LoadObjectID(ReturnValueC1Combo(tdbcObjectTypeID, "").ToString), gbUnicode)
                    tdbcObjectID.SelectedValue = tdbcAssetID.Columns("ObjectID").Value.ToString
                End If
                txtNotes.Text = tdbcAssetID.Columns("Notes").Text
                '
                If tdbcAssetID.Columns("EmployeeID").Value.ToString <> "" Then tdbcEmployeeID.SelectedValue = tdbcAssetID.Columns("EmployeeID").Value.ToString
            End If
            '''''
            'txtEmployeeName.Text = tdbcAssetID.Columns("FullName").Value.ToString
            'txtConvertedAmount.Text = SQLNumber(tdbcAssetID.Columns("ConvertedAmount").Value.ToString, DxxFormat.D90_ConvertedDecimals)
            txtPercentage.Text = SQLNumber(tdbcAssetID.Columns("Percentage").Value.ToString, DxxFormat.D08_RatioDecimals)
            txtPercentage.Tag = txtPercentage.Text
            If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
                txtServiceLife.Text = SQLNumber(tdbcAssetID.Columns("ServiceLife").Value.ToString, DxxFormat.DefaultNumber0)
                txtAmountDepreciation.Text = SQLNumber(tdbcAssetID.Columns("AmountDepreciation").Value.ToString, DxxFormat.D90_ConvertedDecimals)
                txtRemainAmount.Text = SQLNumber(Number(txtConvertedAmount.Text) - Number(txtAmountDepreciation.Text), DxxFormat.D90_ConvertedDecimals)
                txtDepreciateAmount.Text = "0"
                txtDepreciatedPeriod.Text = "0"
            End If
            'txtDepreciateAmount.Text = SQLNumber(tdbcAssetID.Columns("DepreciatedAmount").Value.ToString, DxxFormat.D90_ConvertedDecimals)
            txtMethodEndID.Text = tdbcAssetID.Columns("MethodEndID").Value.ToString
            txtDeprTableName.Text = tdbcAssetID.Columns("DeprTableName").Value.ToString
            txtMethodID.Text = tdbcAssetID.Columns("MethodID").Text.ToString
            c1dateBeginDep.Value = tdbcAssetID.Columns("BeginDep").Value.ToString
            c1dateBeginUsing.Value = tdbcAssetID.Columns("BeginUsing").Value.ToString
            c1dateAssetDate.Value = tdbcAssetID.Columns("AssetDate").Text
            c1dateUseDate.Value = tdbcAssetID.Columns("UseDate").Text
            c1dateDepDate.Value = tdbcAssetID.Columns("DepDate").Text '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
            tdbcManagementObjTypeID.SelectedValue = tdbcAssetID.Columns("ManagementObjTypeID").Value
            tdbcManagementObjID.SelectedValue = tdbcAssetID.Columns("ManagementObjID").Value
            tdbg1.Enabled = True
            tdbg1.Rebind(True)
            For i As Integer = 0 To tdbg1.RowCount - 1
                tdbg1(i, COL1_DebitAccountID) = tdbcAssetID.Columns("AssetAccountID").Text
            Next
            LoadtdbcPropertyProductID()
            LoadtdbdCipID()
            'Xoá lưới
            '  LoadTDBGrid1(True)
            '***********
        End If
    End Sub

    Private Sub tdbcAssetID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetID.LostFocus
        'If tdbcAssetID.FindStringExact(tdbcAssetID.Text) = -1 Or tdbcAssetID.Text = "" Then
        '    tdbcAssetID.SelectedValue = ""
        '    tdbcAssetID.Text = ""
        '    tdbcPropertyProductID.Text = ""
        '    tdbcAssetID.Focus()
        'End If
        ' ShowD02F0087()

    End Sub

    Private Sub tdbcAssetID_Validated1(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetID.Validated
        clsFilterCombo.FilterCombo(tdbcAssetID, e)
        If _FormState <> EnumFormState.FormView AndAlso tdbg1.RowCount > 0 AndAlso sOldAssetID <> ReturnValueC1Combo(tdbcAssetID) Then
            If D99C0008.MsgAsk(rL3("Du_lieu_tren_luoi_se_bi_xoa_Ban_co_muon_thuc_hien_ko")) = Windows.Forms.DialogResult.Yes Then
                If dtGrid1 IsNot Nothing Then dtGrid1.Clear()
                LoadDataSource(tdbg1, dtGrid1, gbUnicode)
            Else
                tdbcAssetID.SelectedValue = sOldAssetID
            End If
        End If
        sOldAssetID = ReturnValueC1Combo(tdbcAssetID)

        If tdbcAssetID.FindStringExact(tdbcAssetID.Text) = -1 Or tdbcAssetID.Text = "" Then
            tdbcAssetID.SelectedValue = ""
            tdbcAssetID.Text = ""
            tdbcPropertyProductID.Text = ""
            tdbcAssetID.Focus()
        End If
        LoadtdbdCipID()
    End Sub

    Private Sub tdbcAssetID_Validated(sender As Object, e As EventArgs) Handles tdbcAssetID.Validated
        clsFilterCombo.FilterCombo(tdbcAssetID, e)
        If tdbcAssetID.FindStringExact(tdbcAssetID.Text) = -1 Then
            tdbcAssetID.Text = ""
        End If
        ShowD02F0087()

    End Sub
#End Region

#Region "Events tdbcObjectTypeID load tdbcObjectID with txtObjectName"

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.SelectedValueChanged
        'If tdbcObjectTypeID.SelectedValue Is Nothing OrElse tdbcObjectTypeID.Text = "" Then
        '    LoadDataSource(tdbcObjectID, LoadObjectID(""), gbUnicode)
        '    Exit Sub
        'End If
        'LoadDataSource(tdbcObjectID, LoadObjectID(tdbcObjectTypeID.SelectedValue.ToString()), gbUnicode)
        'tdbcObjectID.Text = ""
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

#Region "Events tdbcVoucherTypeID with txtVoucherNo"

    Private Sub tdbcVoucherTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.SelectedValueChanged
        bEditVoucherNo = False
        bFirstF2 = False

        If tdbcVoucherTypeID.SelectedValue Is Nothing OrElse tdbcVoucherTypeID.Text = "" Then
            txtVoucherNo.Text = ""
            ReadOnlyControl(txtVoucherNo)
            Exit Sub
        End If
        If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
            If tdbcVoucherTypeID.Columns("Auto").Text = "1" Then 'Sinh tu dong
                txtVoucherNo.Text = CreateIGEVoucherNo(tdbcVoucherTypeID, False)
                ReadOnlyControl(txtVoucherNo)
                'c1dateVoucherDate.Focus()
            Else
                txtVoucherNo.Text = ""
                UnReadOnlyControl(txtVoucherNo)
                'txtVoucherNo.Focus()
            End If
        End If
    End Sub

    Private Sub tdbcVoucherTypeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.LostFocus
        If tdbcVoucherTypeID.Text = "" Then Exit Sub
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then
            tdbcVoucherTypeID.Text = ""
        End If
        'GetVoucherNo(tdbcVoucherTypeID, txtVoucherNo, btnSetNewKey)
        If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
            If tdbcVoucherTypeID.Columns("Auto").Text = "1" Then 'Sinh tu dong
                c1dateVoucherDate.Focus()
            Else
                txtVoucherNo.Focus()
            End If
        End If
    End Sub

    '************************
    Dim sOldVoucherNo As String = "" 'Lưu lại số phiếu cũ
    Dim bEditVoucherNo As Boolean = False '= True: có nhấn F2; = False: không 
    Dim bFirstF2 As Boolean = False 'Nhấn F2 lần đầu tiên 
    Dim iPer_F5558 As Integer = 0 'Phân quyền cho Sửa số phiếu
    '************************

    Private Sub txtVoucherNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVoucherNo.KeyDown
        If e.KeyCode = Keys.F2 Then
            'Loại phiếu hay Số phiếu = "" thì thoát
            If tdbcVoucherTypeID.Text = "" Or txtVoucherNo.Text = "" Then Exit Sub

            'Update 21/09/2010: Trường hợp Thêm mới phiếu và đã lưu Thành công thì không cho sửa Số phiếu
            If (_FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy) And btnSave.Enabled = False Then Exit Sub
            'Kiểm tra quyền cho trường hợp Sửa
            If _FormState = EnumFormState.FormEdit And iPer_F5558 <= 2 Then Exit Sub

            'Cho sửa Số phiếu ở trạng thái Thêm mới hay Sửa
            If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormEdit OrElse _FormState = EnumFormState.FormCopy Then
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
                '    .TableName = "D07T1111" 'Tên bảng chứa số phiếu
                '    'Update 21/09/2010
                '    If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
                '        .VoucherID = "" 'Khóa sinh IGE là rỗng
                '    ElseIf _FormState = EnumFormState.FormEdit Then
                '        .VoucherID = _historyIDMaster   'Khóa sinh IGE
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
                If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
                    SetProperties(arrPro, "VoucherID", "")
                ElseIf _FormState = EnumFormState.FormEdit Then
                    SetProperties(arrPro, "VoucherID", _historyIDMaster)
                End If
                SetProperties(arrPro, "Mode", 0)
                SetProperties(arrPro, "KeyID01", "")
                SetProperties(arrPro, "TableName", "D07T1111")
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

#Region "Events tdbcPropertyProductID load tdbdCipID"

    Private Sub tdbcPropertyProductID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcPropertyProductID.SelectedValueChanged
        If tdbcPropertyProductID.SelectedValue Is Nothing OrElse tdbcPropertyProductID.Text = "" Then
            tdbcPropertyProductID.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub tdbcPropertyProductID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcPropertyProductID.Validated
        clsFilterCombo.FilterCombo(tdbcPropertyProductID, e)
        If _FormState <> EnumFormState.FormView AndAlso tdbg1.RowCount > 0 AndAlso sOldPropertyProductID <> ReturnValueC1Combo(tdbcPropertyProductID) Then
            If D99C0008.MsgAsk(rL3("Du_lieu_tren_luoi_se_bi_xoa_Ban_co_muon_thuc_hien_ko")) = Windows.Forms.DialogResult.Yes Then
                If dtGrid1 IsNot Nothing Then dtGrid1.Clear()
                LoadDataSource(tdbg1, dtGrid1, gbUnicode)
            Else
                tdbcPropertyProductID.SelectedValue = sOldPropertyProductID
            End If
        End If
        sOldPropertyProductID = ReturnValueC1Combo(tdbcPropertyProductID)

        If tdbcPropertyProductID.FindStringExact(tdbcPropertyProductID.Text) = -1 Then
            tdbcPropertyProductID.Text = ""
        End If
        LoadtdbdCipID()
    End Sub

#End Region

    Public Sub GetVoucherNo(ByVal tdbcVoucherTypeID As C1.Win.C1List.C1Combo, ByVal txtVoucherNo As TextBox, ByVal btnSetNewKey As Windows.Forms.Button)
        If tdbcVoucherTypeID.Text <> "" Then
            If tdbcVoucherTypeID.Columns("Auto").Text = "0" Then 'Không tạo mã tự động
                txtVoucherNo.ReadOnly = False
                txtVoucherNo.TabStop = True
                btnSetNewKey.Enabled = False
                txtVoucherNo.Text = ""
            Else
                gnNewLastKey = 0
                txtVoucherNo.ReadOnly = True
                txtVoucherNo.TabStop = False
                btnSetNewKey.Enabled = True

                txtVoucherNo.Text = CreateIGEVoucherNo(tdbcVoucherTypeID, False)
            End If
        End If
    End Sub

    Private Sub btnSetNewKey_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        GetNewVoucherNo(tdbcVoucherTypeID, txtVoucherNo)
    End Sub

    Private Sub LoadMasterAssetID()
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0500(_assetID, _setupFrom)
        'sSQL = "Exec D02P0500 'TUONGAN', 'ALL', 8, 2012, 'D02.AssetID =''03/OO/05''', '01', 1"
        Dim dtMaster As DataTable = ReturnDataTable(sSQL)
        If dtMaster.Rows.Count > 0 Then
            With dtMaster.Rows(0)
                'Không thấy load
                'tdbcObjectTypeID.SelectedValue = .Item("ObjectTypeID")
                'tdbcObjectID.SelectedValue = .Item("ObjectID")
                'tdbcEmployeeID.Text = .Item("EmployeeID").ToString
                'tdbcAssetID.Text = .Item("AssetID").ToString
                'txtAssetName.Text = .Item("AssetName").ToString
                '************
                txtConvertedAmount.Text = SQLNumber(.Item("ConvertedAmount").ToString, DxxFormat.D90_ConvertedDecimals)
                txtAmountDepreciation.Text = SQLNumber(.Item("AmountDepreciation").ToString, DxxFormat.D90_ConvertedDecimals)
                txtRemainAmount.Text = SQLNumber(.Item("RemainAmount").ToString, DxxFormat.D90_ConvertedDecimals)
                'txtMethodID.Text = tdbcAssetID.Columns("MethodID").Text
                'txtMethodEndID.Text = tdbcAssetID.Columns("MethodEndID").Value.ToString
                'txtDeprTableName.Text = tdbcAssetID.Columns("DeprTableName").Value.ToString
                txtServiceLife.Text = SQLNumber(.Item("ServiceLife").ToString, DxxFormat.DefaultNumber0)
                txtDepreciatedPeriod.Text = SQLNumber(.Item("DepreciatedPeriod").ToString, DxxFormat.DefaultNumber0)
                txtDepreciateAmount.Text = SQLNumber(.Item("DepreciatedAmount").ToString, DxxFormat.D90_ConvertedDecimals) ' 17/12/2013 id 62172 
                'txtPercentage.Text = SQLNumber(.Item("Percentage").ToString, DxxFormat.D08_RatioDecimals)
                'c1dateBeginUsing.Value = .Item("BeginUse").ToString
                'c1dateBeginDep.Value = .Item("BeginDep").ToString
                c1dateUseDate.Value = .Item("UseDate").ToString
                c1dateAssetDate.Value = .Item("AssetDate").ToString
                c1dateDepDate.Value = .Item("DepDate").ToString '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
                sCreateUserID = .Item("CreateUserID").ToString
                sCreateDate = .Item("CreateDate").ToString
            End With
        End If
    End Sub

    Dim dtGrid1 As DataTable

    Private Sub LoadTabpage1()
        Dim strLoad As New StringBuilder
        Dim sUnicode As String = ""
        If gbUnicode Then sUnicode = "U"
        strLoad.Append(" Select BatchID, A.AssetID, VoucherTypeID, VoucherNo, VoucherDate, ")
        If _FormState = EnumFormState.FormEdit Then
            strLoad.Append("1 as IsEdit,")
        Else
            strLoad.Append("0 as IsEdit,")
        End If
        strLoad.Append(" Notes" & sUnicode & " AS Notes, CurrencyID, ExchangeRate, OriginalAmount, ConvertedAmount,")
        strLoad.Append(" TransactionID, DebitAccountID, CreditAccountID, A.Description" & sUnicode & " AS Description,")
        strLoad.Append(" RefNo, RefDate, SeriNo as SerialNo, A.ObjectTypeID, A.ObjectID, ObjectName" & sUnicode & " AS ObjectName,")
        strLoad.Append(" Ana01ID, Ana02ID, Ana03ID, Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID,")
        strLoad.Append(" IsNull(A.CipID, '') CipID, SourceID, IsNull(B.CipNo, '') CipNo, ConvertedAmount As CopyAmount, C.SumConvertedAmo, C.SumConvertedAmo as InitialCovertedAmount ")
        strLoad.Append(", Str01" & sUnicode & " AS Str01, Str02" & sUnicode & " AS Str02, Str03" & sUnicode & " AS Str03, Str04" & sUnicode & " AS Str04, Str05" & sUnicode & " AS Str05, Num01, Num02, Num03, Num04, Num05 , Date01, Date02, Date03, Date04, Date05, Convert(bit, IsNotAllocate) As IsNotAllocate" & vbCrLf)
        strLoad.Append(" From D02T0012 A WITH(NOLOCK) Left Join D02T0100 B WITH(NOLOCK) On IsNull(A.CipID, '') = B.CipID ")
        strLoad.Append(" Left Join (SELECT * FROM(SELECT A.CipID,  Sum(Case When A.DebitAccountID = B.AccountID then ConvertedAmount When a.CreditAccountID = b.AccountID then - ConvertedAmount else 0 end) as SumConvertedAmo FROM D02T0012 A WITH(NOLOCK) " & _
                                   " INNER JOIN D02T0100 B WITH(NOLOCK) ON A.CipID = B.CipID WHERE A.TranMonth + A.TranYear * 100 <= " & giTranMonth & " + " & giTranYear & " * 100 " & " GROUP BY A.CipID) A " & _
                                           " WHERE A.SumConvertedAmo <> 0) C On IsNull(A.CipID, '') = C.CipID ")
        strLoad.Append(" Where ISNULL(A.AssetID,'') = " & SQLString(_assetID) & " And ModuleID = '02' And TransactionTypeID = 'XDCB'")
        If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then strLoad.Append(" And 1=0")
        dtGrid1 = ReturnDataTable(strLoad.ToString)
        LoadTDBGrid1()
        tdbcAssetID.SelectedValue = _assetID
        tdbcPropertyProductID.SelectedValue = _d27PropertyProductID
        If dtGrid1.Rows.Count = 0 OrElse _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then Exit Sub

        With dtGrid1.Rows(0)
            sEditVoucherTypeID = .Item("VoucherTypeID").ToString
            LoadVoucherTypeID(tdbcVoucherTypeID, D02, sEditVoucherTypeID, gbUnicode)
            tdbcVoucherTypeID.SelectedValue = sEditVoucherTypeID
            txtVoucherNo.Text = .Item("VoucherNo").ToString
            c1dateVoucherDate.Value = .Item("VoucherDate")
            txtDescription.Text = .Item("Notes").ToString
            'Không thấy load Loại tiền và Tỷ giá
        End With
    End Sub

    Private Sub LoadTDBGrid1(Optional ByVal bClear As Boolean = False)
        If bClear And dtGrid1 IsNot Nothing Then dtGrid1.Clear()
        LoadDataSource(tdbg1, dtGrid1, gbUnicode)
        FooterTotalGrid(tdbg1, COL1_CipID)
    End Sub

    Dim dtGrid2 As DataTable = Nothing

    Private Sub LoadTDBGrid2(Optional ByVal bLoadSQL As Boolean = True)
        If bLoadSQL Then
            Dim strSQL As String = ""
            strSQL = " Select D02T5000.HistoryID, D02T5000.AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " as AssignmentName, DebitAccountID, PercentAmount/100 AS PercentAmount"
            strSQL &= " From D02T5000 WITH(NOLOCK) Inner join D02T0002 WITH(NOLOCK) On D02T5000.AssignmentID = D02T0002.AssignmentID"
            strSQL &= " Where HistoryTypeID = 'AS'" 'Bo BeginYear,BeginMonth theo Incident 79529
            strSQL &= " And DivisionID = " & SQLString(gsDivisionID)
            strSQL &= " And ISNULL(AssetID,'') = " & SQLString(_assetID) & " Order by HistoryID"
            dtGrid2 = ReturnDataTable(strSQL)
        Else
            dtGrid2.Clear()
        End If
        LoadDataSource(tdbg2, dtGrid2, gbUnicode)
        FooterTotalGrid(tdbg2, COL2_AssignmentID)
    End Sub

    Private Sub LoadAddNew()
        btnNext.Enabled = False
        btnSave.Enabled = True
        c1dateVoucherDate.Value = Now.Date.Date
        c1dateBeginUsing.Value = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        c1dateBeginDep.Value = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        c1dateUseDate.Value = Date.Today
        c1dateAssetDate.Value = Date.Today
        c1dateDepDate.Value = Date.Today '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
    End Sub

    Private Sub LoadEdit()
        btnNext.Visible = False
        btnSave.Left = btnNext.Left
        ReadOnlyControl(tdbcAssetID, tdbcVoucherTypeID, txtVoucherNo)
        LoadMasterAssetID()
        LoadTabpage1()
        LoadTDBGrid2()
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

    Private Function LoadCaptionSubInfo() As Boolean
        Dim bUseSubInfo As Boolean = False
        Dim sUnicode As String = ""
        If gbUnicode Then sUnicode = "U"
        'Dim sSQL As String = ""
        'sSQL = "SELECT Data" & gsLanguage & sUnicode & " as Data, Description" & sUnicode & " as Description, Disabled, DataID, DataType, DecimalNum From D02T0003 WHERE DataID LIKE 'AnaStr%'" & vbCrLf
        'sSQL &= "Union All" & vbLf
        'sSQL &= "SELECT Data" & gsLanguage & sUnicode & " as Data, Description" & sUnicode & " as Description, Disabled, DataID, DataType, DecimalNum From D02T0003 WHERE DataID LIKE 'AnaNum%'" & vbCrLf
        'sSQL &= "Union All" & vbLf
        'sSQL &= "SELECT Data" & gsLanguage & sUnicode & " as Data, Description" & sUnicode & " as Description, Disabled, DataID, DataType, DecimalNum From D02T0003 WHERE DataID LIKE 'AnaDate%'" & vbCrLf
        Dim dtCaption As DataTable = ReturnDataTable(SQLStoreD02P0015)
        If dtCaption.Rows.Count = 0 Then Return False
        Dim arr() As FormatColumn = Nothing
        For i As Integer = 0 To dtCaption.Rows.Count - 1
            Dim sField As String = dtCaption.Rows(i).Item("DataID").ToString.Replace("Ana", "")
            If tdbg1.Columns.IndexOf(tdbg1.Columns(sField)) <= -1 Then Continue For
            tdbg1.Columns(sField).Caption = dtCaption.Rows(i).Item("Data" & gsLanguage).ToString
            tdbg1.Splits(SPLIT1).DisplayColumns(sField).HeadingStyle.Font = FontUnicode(gbUnicode)
            tdbg1.Splits(SPLIT1).DisplayColumns(sField).Visible = L3Bool(dtCaption.Rows(i).Item("Disabled"))
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

    Private Sub tdbg2_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.ComboSelect
        tdbg2.UpdateData()
    End Sub

    Private Sub tdbg2_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.AfterColUpdate
        Select Case e.Column.DataColumn.DataField
            Case COL2_AssignmentID 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, e.Column.DataColumn.DataField)
                If tdbd Is Nothing Then Exit Select
                If clsFilterDropdown.IsNewFilter Then
                    If tdbg2.Columns(e.ColIndex).Text = "+" Then 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                        tdbg2.Columns(e.ColIndex).Text = ""
                        ShowD02F0101(COL2_AssignmentID)
                    End If
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdownMulti(tdbg2, e, tdbd)
                    AfterColUpdate_tdbg2(e.ColIndex, dr)
                Else ' Nhập liệu dạng cũ (xổ dropdown)
                    '    Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg2.Columns(e.ColIndex).Text))
                    '    AfterColUpdate_tdbg2(e.ColIndex, row)
                    'End If
                    If tdbg2.Columns(e.ColIndex).Text = "+" Then 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                        tdbg2.Columns(e.ColIndex).Text = ""
                        ShowD02F0101(COL2_AssignmentID)
                    End If
                    Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg2.Columns(e.ColIndex).Text))
                    AfterColUpdate_tdbg2(e.ColIndex, row)
                End If
            Case COL1_ObjectID
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
            Case IndexOfColumn(tdbg2, COL2_AssignmentID)  'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, tdbg2.Columns(tdbg2.Col).DataField)
                'If tdbd Is Nothing Then Exit Select
                'Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                'If dr Is Nothing Then Exit Sub
                'AfterColUpdate_tdbg2(tdbg2.Col, dr)
                If tdbd Is Nothing Then Exit Select
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdownMulti(tdbg2, e, tdbd)
                If dr Is Nothing Then Exit Sub
                If dr(0).Item("AssignmentID").ToString = "+" Then 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                    tdbg2.Columns(COL2_AssignmentID).Text = ""
                    ShowD02F0101(COL2_AssignmentID)
                    Dim row As DataRow = ReturnDataRow(tdbdAssignmentID, "AssignmentID=" & SQLString(tdbg2.Columns(COL2_AssignmentID).Text))
                    AfterColUpdate_tdbg2(tdbg2.Col, row)
                Else
                    AfterColUpdate_tdbg2(tdbg2.Col, dr)
                End If
        End Select
    End Sub

    Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
        If clsFilterDropdown.CheckKeydownFilterDropdown(tdbg2, e) Then
            Select Case tdbg2.Col
                Case IndexOfColumn(tdbg2, COL2_AssignmentID)  'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, tdbg2.Columns(tdbg2.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdownMulti(tdbg2, e, tdbd)
                    If dr Is Nothing Then Exit Sub
                    If dr(0).Item("Assignment").ToString = "+" Then 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                        tdbg2.Columns(COL2_AssignmentID).Text = ""
                        ShowD02F0101(COL2_AssignmentID)
                        Dim row As DataRow = ReturnDataRow(COL2_AssignmentID, "AssignmentID=" & SQLString(tdbg2.Columns(COL2_AssignmentID).Text))
                        AfterColUpdate_tdbg2(tdbg2.Col, row)
                    Else
                        AfterColUpdate_tdbg2(tdbg2.Col, dr)
                    End If
                    Exit Sub
            End Select
        End If
    End Sub

    Private Sub AfterColUpdate_tdbg2(ByVal iCol As Integer, ByVal dr As DataRow)
        'Gán lại các giá trị phụ thuộc vào Dropdown
        Select Case iCol
            Case IndexOfColumn(tdbg2, COL2_AssignmentID)
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
                FooterTotalGrid(tdbg2, COL2_AssignmentID)
        End Select
    End Sub

    Private Sub tdbg2_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg2.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        Select Case e.Column.DataColumn.DataField
            Case COL2_AssignmentID
                If clsFilterDropdown.IsNewFilter Then Exit Sub 'ID : 214917 - BỔ SUNG CHỨC NĂNG THÊM MỚI PHÂN BỔ KHẤU HAO TRÊN DROPDOWN TÌM KIẾM PHÂN BỔ KHẤU HAO
                If tdbg2.Columns(COL2_AssignmentID).Text <> tdbdAssignmentID.Columns("AssignmentID").Text Then
                    tdbg2.Columns(COL2_AssignmentID).Text = ""
                    tdbg2.Columns(COL2_AssignmentName).Text = ""
                    tdbg2.Columns(COL2_DebitAccountID).Text = ""
                    tdbg2.Columns(COL2_Extend).Text = ""
                End If
        End Select
    End Sub

    Private Sub tdbg2_LockedColumns()
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_AssignmentName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    Dim _historyIDMaster As String = "" 'Sinh IGE cho master D02T5000
    Dim bCheckIsManagement As Boolean = False

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        tdbg1.UpdateData()
        tdbg2.UpdateData()
        If Not AllowSave() Then Exit Sub

        'sGetDate = SetGetDateSQL()
        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)
        If CheckVoucherDateInPeriod(c1dateVoucherDate.Text) = False Then
            tabMain.SelectedTab = tabPage1
            c1dateVoucherDate.Focus()
            Exit Sub
        End If
        Dim _isManagent As String = ""
        _isManagent = ReturnScalar("SELECT TOP 1 1 FROM D02T0001 WITH(NOLOCK) Where AssetID =" & SQLString(ReturnValueC1Combo(tdbcAssetID)) & "AND ManagementObjTypeID='' AND ManagementObjID=''")
        btnSave.Enabled = False
        btnClose.Enabled = False
        gbSavedOK = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        _historyIDMaster = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
        Select Case _FormState
            Case EnumFormState.FormAdd, EnumFormState.FormCopy
                ''Kiểm tra phiếu 
                '06/11/2012 thay đổi bảng D02T5000 thành bảng D02T0012 theo incident 52233
                If tdbcVoucherTypeID.Columns("Auto").Text = "1" And bEditVoucherNo = False Then 'Tự động
                    'txtVoucherNo.Text = CreateIGEVoucherNoNew(tdbcVoucherTypeID, "D02T5000", _historyIDMaster)
                    txtVoucherNo.Text = CreateIGEVoucherNoNew(tdbcVoucherTypeID, "D02T0012", _historyIDMaster)
                Else 'Không sinh tự động hay có nhấn F2
                    If bEditVoucherNo = False Then
                        If CheckDuplicateVoucherNoNew("D02", "D02T0012", _historyIDMaster, txtVoucherNo.Text) Then
                            Me.Cursor = Cursors.Default
                            btnSave.Enabled = True
                            btnClose.Enabled = True
                            txtVoucherNo.Focus()
                            Exit Sub
                        End If
                    Else 'Có nhấn F2 để sửa số phiếu
                        InsertD02T5558(_historyIDMaster, sOldVoucherNo, txtVoucherNo.Text)
                    End If
                    InsertVoucherNoD91T9111(txtVoucherNo.Text, "D02T0012", _historyIDMaster)
                End If
                bEditVoucherNo = False
                sOldVoucherNo = ""
                bFirstF2 = False
                ''****************************************
                sSQL.Append(SQLUpdateD02T0012.ToString & vbCrLf)
                'sSQL.Append(SQLUpdateD02T0100.ToString & vbCrLf)
                'sSQL.Append(SQLUpdateD02T0001.ToString & vbCrLf)'Chuyển về update cuối
                sSQL.Append(SQLInsertD02T5000(_historyIDMaster, giTranMonth, giTranYear, "OB", tdbcObjectTypeID.Text, tdbcObjectID.Text, tdbcEmployeeID.Text, txtEmployeeName.Text).ToString & vbCrLf)
                Dim _historyIDMasterNew As String = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
                If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" Then
                    bCheckIsManagement = True
                    sSQL.Append(SQLInsertD02T5000(_historyIDMasterNew, giTranMonth, giTranYear, "OB", tdbcObjectTypeID.Text, tdbcObjectID.Text, tdbcEmployeeID.Text, txtEmployeeName.Text).ToString & vbCrLf)
                End If
                sSQL.Append(SQLInsertD02T5000s().ToString & vbCrLf)

                Dim sHistory As String = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistory, giTranMonth, giTranYear, "IL").ToString & vbCrLf)

                sSQL.Append(SQLInsertD02T0012s().ToString & vbCrLf)

                txtConvertedAmount.Text = Format(dTotalConvertedAmount, DxxFormat.D90_ConvertedDecimals)
                sSQL.Append(SQLUpdateD02T0001.ToString & vbCrLf)
                sSQL.Append(SQLStoreD02P0101().ToString & vbCrLf)

                '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
                Dim sHistoryAAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistoryAAC, giTranMonth, giTranYear, "AAC", , , , , , , , , , ReturnValueC1Combo(tdbcAssetID, "AssetAccountID")).ToString & vbCrLf)
                Dim sHistoryDAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistoryDAC, giTranMonth, giTranYear, "DAC", , , , , , , , , , , ReturnValueC1Combo(tdbcAssetID, "DepAccountID")).ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5010(sHistoryAAC, "AAC", ReturnValueC1Combo(tdbcAssetID, "AssetAccountID")))
                sSQL.Append(SQLInsertD02T5010(sHistoryDAC, "DAC", ReturnValueC1Combo(tdbcAssetID, "DepAccountID")))
            Case EnumFormState.FormEdit, EnumFormState.FormEditOther
                sSQL.Append(SQLDeleteD02T5000() & vbCrLf)
                sSQL.Append(SQLDeleteD02T0012() & vbCrLf)
                'Bổ sung theo Nam 08/03/2012 : Trong VB6 không có tính năng sửa phiếu HT XDCB
                sSQL.Append(SQLUpdateD02T0012.ToString & vbCrLf)
                'sSQL.Append(SQLUpdateD02T0100.ToString & vbCrLf)
                'Update những CipID không chọn trên lưới Update =1
                Dim dtCipID As DataTable = CType(tdbdCipID.DataSource, DataTable)
                Dim dr() As DataRow = dtCipID.Select("CipID not In (" & arrCipID.ToString & ")")
                Dim arrNotCipID As String = ""
                For i As Integer = 0 To dr.Length - 1
                    If arrNotCipID <> "" Then arrNotCipID &= COMMA
                    arrNotCipID &= SQLString(dr(i).Item(COL1_CipID))
                Next
                If arrNotCipID <> "" Then sSQL.Append("Update D02T0100 Set Status = 1 " & _
                                                        "Where ISNULL(CipID,'') in (" & arrNotCipID.ToString & ") ")
                '***************************
                sSQL.Append(SQLUpdateD02T0001().ToString & vbCrLf)
                sSQL.Append(SQLStoreD02P0101().ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000(_historyIDMaster, giTranMonth, giTranYear, "OB", tdbcObjectTypeID.Text, tdbcObjectID.Text, tdbcEmployeeID.Text, txtEmployeeName.Text).ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5000s().ToString & vbCrLf)
                Dim sHistory As String = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistory, giTranMonth, giTranYear, "IL").ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T0012s().ToString & vbCrLf)

                '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
                Dim sHistoryAAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistoryAAC, giTranMonth, giTranYear, "AAC", , , , , , , , , , ReturnValueC1Combo(tdbcAssetID, "AssetAccountID")).ToString & vbCrLf)
                Dim sHistoryDAC As String = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
                sSQL.Append(SQLInsertD02T5000(sHistoryDAC, giTranMonth, giTranYear, "DAC", , , , , , , , , , , ReturnValueC1Combo(tdbcAssetID, "DepAccountID")).ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T5010(sHistoryAAC, "AAC", ReturnValueC1Combo(tdbcAssetID, "AssetAccountID")))
                sSQL.Append(SQLInsertD02T5010(sHistoryDAC, "DAC", ReturnValueC1Combo(tdbcAssetID, "DepAccountID")))
                Dim _historyIDMasterNew As String = CreateIGE("D02T5000", "HistoryID", "02", "HD", gsStringKey)
                If ReturnValueC1Combo(tdbcManagementObjTypeID) <> "" Then
                    bCheckIsManagement = True
                    sSQL.Append(SQLInsertD02T5000(_historyIDMasterNew, giTranMonth, giTranYear, "OB", tdbcObjectTypeID.Text, tdbcObjectID.Text, tdbcEmployeeID.Text, txtEmployeeName.Text).ToString & vbCrLf)
                End If
        End Select
        If tdbg1.RowCount = 1 Then 'Nếu lưới Định khoản có 1 dòng
            'Thực hiện Bước 6
            'Bạn có muốn cập nhật Nhà cung cấp và thông tin phụ của Mã XDCB sang thông tin phụ của Mã TSCD không?
            If D99C0008.Msg(rl3("Ban_co_muon_cap_nhat_Nha_cung_cap_va_thong_tin_phu_cua_Ma_XDCB_sang_thong_tin_phu_cua_Ma_TSCD_khong"), rl3("Thong_bao"), L3MessageBoxButtons.YesNo, L3MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                sSQL.Append(Level6())
            End If
        End If
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            gbSavedOK = True
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd, EnumFormState.FormCopy
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
            If ReturnValueC1Combo(tdbcPropertyProductID) <> "" Then
                ExecuteSQL(SQLStoreD02P1010)
            End If

        Else
            If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormCopy Then
                DeleteVoucherNoD91T9111_Transaction(txtVoucherNo.Text, "D02T0012", "VoucherNo", tdbcVoucherTypeID, bEditVoucherNo)
            End If
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0101
    '# Created User: HUỲNH KHANH
    '# Created Date: 25/01/2016 05:25:18
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0101() As String
        Dim sSQL As String = ""
        sSQL &= ("-- --kiem tra va cap nhat trang thai XDCB" & vbCrlf)
        sSQL &= "Exec D02P0101 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString("D02F1010") & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= SQLString(txtVoucherNo.Text) & COMMA 'VoucherNo, varchar[20], NOT NULL
        sSQL &= SQLString(arrCipID.ToString) 'CipID, varchar[20], NOT NULL
        Return sSQL
    End Function



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1010
    '# Created User: HUỲNH KHANH
    '# Created Date: 25/01/2015 12:07:09
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1010() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Update du lieu cho bang danh muc tai san co dinh" & vbCrlf)
        sSQL &= "Exec D02P1010 "
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA 'AssetID, varchar[50], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcPropertyProductID, "D54ProjectID")) & COMMA 'ProjectID, varchar[50], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcPropertyProductID)) & COMMA 'PropertyProductID, varchar[50], NOT NULL
        sSQL &= SQLString(gsDivisionID) 'DivisionID, varchar[50], NOT NULL
        Return sSQL
    End Function



    Dim arrCipID As New StringBuilder 'Danh sách các CipID trên lưới 1

    Private Function AllowSave() As Boolean
        If tdbcAssetID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Ma_tai_san"))
            tdbcAssetID.Focus()
            Return False
        End If
        If tdbcObjectTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Bo_phan_quan_ly"))
            tdbcObjectTypeID.Focus()
            Return False
        End If
        If tdbcObjectID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Bo_phan_quan_ly"))
            tdbcObjectID.Focus()
            Return False
        End If
        If tdbcEmployeeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Nguoi_tiep_nhan"))
            tdbcEmployeeID.Focus()
            Return False
        End If
        If Number(txtConvertedAmount.Text) < Number(txtAmountDepreciation.Text) Then
            'Mức khấu hao không được lớn hơn nguyên giá
            D99C0008.MsgL3(rl3("Muc_khau_hao_khong_duoc_lon_hon_nguyen_gia"))
            tabMain.SelectedTab = tabPage1
            If txtConvertedAmount.Enabled Then txtConvertedAmount.Focus()
            Return False
        End If
        If tdbcVoucherTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Loai_phieu"))
            tabMain.SelectedTab = tabPage1
            tdbcVoucherTypeID.Focus()
            Return False
        End If
        If txtVoucherNo.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rl3("So_phieu"))
            tabMain.SelectedTab = tabPage1
            txtVoucherNo.Focus()
            Return False
        End If
        If c1dateVoucherDate.Value.ToString = "" Then
            D99C0008.MsgNotYetEnter(rl3("Ngay_phieu"))
            tabMain.SelectedTab = tabPage1
            c1dateVoucherDate.Focus()
            Return False
        End If

        Dim dr() As DataRow = dtGrid1.Select("IsNotAllocate = 0")
        If dr.Length > 0 Then
            If Number(txtServiceLife.Text.Trim) = 0 Then
                D99C0008.MsgNotYetEnter(rL3("So_ky_khau_hao"))
                txtServiceLife.Focus()
                Return False
            End If
        End If

        If txtDepreciatedPeriod.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rl3("So_ky_da_khau_hao"))
            txtDepreciatedPeriod.Focus()
            Return False
        End If
        If Number(txtServiceLife.Text) < Number(txtDepreciatedPeriod.Text) Then
            D99C0008.MsgL3(rl3("So_ky_da_khau_hao_khong_duoc_lon_hon_So_ky_khau_hao"))
            txtDepreciatedPeriod.Focus()
            Return False
        End If
        'If txtDepreciatedPeriod.Text.Trim.Length > MaxInt Then
        '    D99C0008.MsgL3(rl3("So_vuot_qua_gioi_han"))
        '    txtDepreciatedPeriod.Focus()
        '    Return False
        'End If

        '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
        If D02Systems.IsCalPeriodByRate = True Then
            If Number(txtPercentage.Text) = 0 Then
                D99C0008.MsgNotYetEnter(rL3("Ty_le_khau_hao_%"))
                txtPercentage.Focus()
                Return False
            End If
        End If

        If c1dateBeginUsing.Value.ToString = "" Then
            D99C0008.MsgNotYetEnter(rl3("Ky_bat_dau_su_dung"))
            c1dateBeginUsing.Focus()
            Return False
        End If
        If c1dateBeginDep.Value.ToString = "" Then
            D99C0008.MsgNotYetEnter(rl3("Ky_bat_dau_KH"))
            c1dateBeginDep.Focus()
            Return False
        End If

        Dim bPeriod As Double = Year(CDate(c1dateBeginDep.Value)) * 100 + Month(CDate(c1dateBeginDep.Value))
        Dim uPeriod As Double = Year(CDate(c1dateBeginUsing.Value)) * 100 + Month(CDate(c1dateBeginUsing.Value))
        If bPeriod < uPeriod Then
            D99C0008.MsgL3(rl3("Ky_bat_dau_tinh_khau_hao_phai_lon_hon_hoac_bang_ky_su_dung"))
            c1dateBeginDep.Focus()
            Return False
        End If
        Dim curPeriod As Double = giTranYear * 100 + giTranMonth
        If bPeriod < curPeriod Then
            D99C0008.MsgL3(rl3("Ky_bat_dau_tinh_khau_hao_phai_lon_hon_hoac_bang_ky_hinh_thanh"))
            c1dateBeginDep.Focus()
            Return False
        End If
        If uPeriod < curPeriod Then
            D99C0008.MsgL3(rl3("Ky_su_dung_phai_lon_hon_hoac_bang_ky_hinh_thanh"))
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

        If tdbg1.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tabMain.SelectedTab = tabPage1
            tdbg1.Focus()
            Return False
        End If
        arrCipID = New StringBuilder
        For i As Integer = 0 To tdbg1.RowCount - 1
            If tdbg1(i, COL1_CipID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Ma_XDCB"))
                tabMain.SelectedTab = tabPage1
                tdbg1.Focus()
                tdbg1.SplitIndex = SPLIT0
                tdbg1.Col = IndexOfColumn(tdbg1, COL1_CipID)
                tdbg1.Bookmark = i
                Return False
            End If

            'Bỏ đoạn kiểm tra trùng đi theo incident 71702
            'For j As Integer = i + 1 To tdbg1.RowCount - 1
            '    If tdbg1(i, COL1_CipID).ToString = tdbg1(j, COL1_CipID).ToString Then
            '        D99C0008.MsgDuplicatePKey()
            '        tabMain.SelectedTab = tabPage1
            '        tdbg1.Focus()
            '        tdbg1.SplitIndex = SPLIT0
            '        tdbg1.Col = IndexOfColumn(tdbg1, COL1_CipID)
            '        tdbg1.Bookmark = j
            '        Return False
            '    End If
            'Next
            'Kiểm tra số tiền quy đổi
            'Dim iInitConvert As Double = Number(tdbg1(i, COL1_InitialCovertedAmount))
            'Dim iConvertedAmount As Double = Number(dtGrid1.Compute("SUM(ConvertedAmount)", "CipID = " & SQLString(tdbg1(i, COL1_CipID))))
            'If _FormState = EnumFormState.FormEdit Then iInitConvert += Number(tdbg1(i, COL1_CopyAmount))
            'If Number(tdbg1(i, COL1_ConvertedAmount)) > iInitConvert Then
            '    D99C0008.MsgNotYetEnter(rl3("So_tien_quy_doi_khong_duoc_lon_hon_so_du"))
            '    tabMain.SelectedTab = tabPage1
            '    tdbg1.Focus()
            '    tdbg1.SplitIndex = SPLIT0
            '    tdbg1.Col = IndexOfColumn(tdbg1, COL1_ConvertedAmount)
            '    tdbg1.Bookmark = i
            '    Return False
            'End If
            Dim iInitConvert As Double = Number(tdbg1(i, COL1_InitialCovertedAmount))
            Dim iConvertedAmount As Double = Number(dtGrid1.Compute("SUM(ConvertedAmount)", "CipID = " & SQLString(tdbg1(i, COL1_CipID))))
            If _FormState = EnumFormState.FormEdit Or _FormState = EnumFormState.FormEditOther Then iInitConvert += Number(tdbg1(i, COL1_CopyAmount))
            If iConvertedAmount > iInitConvert Then
                'D99C0008.MsgNotYetEnter(rL3("So_tien_quy_doi_khong_duoc_lon_hon_so_du"))
                D99C0008.Msg(rL3("So_tien_quy_doi_khong_duoc_lon_hon") & " " & iInitConvert)
                tabMain.SelectedTab = tabPage1
                tdbg1.Focus()
                tdbg1.SplitIndex = SPLIT0
                tdbg1.Col = IndexOfColumn(tdbg1, COL1_ConvertedAmount)
                tdbg1.Bookmark = i
                Return False
            End If
            If arrCipID.ToString <> "" Then arrCipID.Append(COMMA)
            arrCipID.Append(SQLString(tdbg1(i, COL1_CipID).ToString))
        Next
        'Tabpage 2
        If tdbg2.RowCount = 0 Then
            D99C0008.MsgNotYetEnter(rl3("Ma_phan_bo"))
            tabMain.SelectedTab = tabPage2
            tdbg2.Focus()
            tdbg2.SplitIndex = SPLIT0
            tdbg2.Col = IndexOfColumn(tdbg2, COL2_AssignmentID)
            tdbg2.Bookmark = 0
            Return False
        End If
        Dim dt As DataTable = dtGrid2.DefaultView.ToTable
        Dim dr1() As DataRow = dt.Select(COL2_AssignmentID & "=''")
        If dr1.Length > 0 Then
            If dr(0).Item(COL2_AssignmentID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rL3("Ma_phan_bo"))
                tabMain.SelectedTab = tabPage2
                tdbg2.Focus()
                tdbg2.SplitIndex = SPLIT0
                tdbg2.Col = IndexOfColumn(tdbg2, COL2_AssignmentID)
                tdbg2.Bookmark = dt.Rows.IndexOf(dr(0))
                Return False
            End If
        End If
        Dim dTotalPercent As Double = Number(dtGrid2.Compute("SUM(" & COL2_PercentAmount & ")", ""))
        If dTotalPercent <> 1 Then 'Kiểm tra không được nhỏ hơn hoặc lớn hơn 100%
            D99C0008.MsgL3(rl3("Tong_ty_le_phai_bang_100U"))
            tabMain.SelectedTab = tabPage2
            tdbg2.Focus()
            tdbg2.SplitIndex = SPLIT0
            tdbg2.Col = IndexOfColumn(tdbg2, COL2_PercentAmount)
            tdbg2.Bookmark = 0
            Return False
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


    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        LoadtdbcAssetID()
        LoadtdbdCipID()
        _assetID = ""
        tdbcAssetID.SelectedValue = ""
        tdbcPropertyProductID.SelectedValue = ""
        tdbcObjectTypeID.SelectedValue = ""
        tdbcEmployeeID.SelectedValue = ""
        tdbcVoucherTypeID.SelectedValue = ""
        tdbcObjectID.SelectedValue = ""
        txtConvertedAmount.Text = ""
        txtAmountDepreciation.Text = ""
        txtRemainAmount.Text = ""
        txtServiceLife.Text = ""
        txtDepreciatedPeriod.Text = ""
        txtPercentage.Text = ""
        txtDepreciateAmount.Text = ""
        txtVoucherNo.Text = ""
        txtDescription.Text = ""
        txtObjectName.Text = ""
        txtEmployeeName.Text = ""
        txtAssetName.Text = ""

        LoadAddNew()
        tabMain.SelectedTab = tabPage1
        LoadTDBGrid1(True)
        tdbg1.Enabled = False
        LoadTDBGrid2(False)
        tdbcAssetID.Focus()
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Hinh_thanh_tai_san_tu_XDCB_-_D02F1010") & UnicodeCaption(gbUnicode) 'HØnh thªnh tªi s¶n tô XDCB - D02F1010
        '================================================================ 
        lblAssetID.Text = rl3("Ma_tai_san") 'Mã tài sản

        lblEmployeeID.Text = rl3("Nguoi_tiep_nhan") 'Người tiếp nhận
        lblExchangRate.Text = rl3("Ty_gia") 'Tỷ giá
        lblteVoucherDate.Text = rl3("Ngay_phieu") 'Ngày phiếu
        lblVoucherTypeID.Text = rl3("Loai_phieu") 'Loại phiếu
        lblVoucherNo.Text = rl3("So_phieu") 'Số phiếu
        lblCurrenyID.Text = rl3("Loai_tien") 'Loại tiền
        lblDescription.Text = rl3("Dien_giai") 'Diễn giải
        lblMethodID.Text = rl3("Phuong_phap_KH") 'Phương pháp KH
        lblMethodEndID.Text = rl3("Khau_hao_ky_cuoi") 'Khấu hao kỳ cuối
        lblDeprTableName.Text = rl3("Bang_khau_hao") 'Bảng khấu hao
        lblConvertedAmount.Text = rl3("Nguyen_gia") 'Nguyên giá
        lblAmountDepreciation.Text = rl3("Hao_mon_luy_ke") 'Hao mòn luỹ kế
        lblRemainAmount.Text = rl3("Gia_tri_con_lai") 'Giá trị còn lại
        lblServiceLife.Text = rl3("So_ky_khau_hao") 'Số kỳ khấu hao
        lblDepreciatedPeriod.Text = rl3("So_ky_da_khau_hao") 'Số kỳ đã khấu hao
        lblPercentage.Text = rl3("Ty_le_khau_hao") & " %" 'Tỷ lệ khấu hao %
        lblUseDate.Text = rl3("Ngay_su_dung") 'Ngày sử dụng
        lblAssetDate.Text = rl3("Ngay_tiep_nhan") 'Ngày tiếp nhận
        lblDepreciateAmount.Text = rl3("Muc_khau_hao") 'Mức khấu hao
        lblBeginUsing.Text = rl3("Ky_bat_dau_su_dung") 'Kỳ bắt đầu sử dụng
        lblBeginDep.Text = rl3("Ky_bat_dau_KH") 'Kỳ bắt đầu KH
        lblNotes.Text = rL3("Ghi_chu")
        lblObjectTypeID.Text = rL3("Bo_phan_tiep_nhan") 'Bộ phận tiếp nhận
        lblPropertyProductID.Text = rL3("Ma_BDS") 'Mã BĐS
        lblDepDate.Text = rL3("Ngay_bat_dau_khau_hao") 'Ngày bắt đầu khấu hao

        '================================================================ 
        btnHotKeys.Text = rl3("Phim_nong") 'Phím nóng
        btnSave.Text = rl3("_Luu") '&Lưu
        btnNext.Text = rl3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        grpAssetID.Text = rl3("Ma_tai_san") 'Mã tài sản
        grpFinancialInfo.Text = rl3("Thong_tin_tai_chinh") 'Thông tin tài chính
        '================================================================ 
        tabPage1.Text = "1. " & rl3("Dinh_khoan") '1. Định khoản
        tabPage2.Text = "2. " & rl3("Phan_bo_khau_hao") '2. Phân bổ khấu hao
        '================================================================ 
        tdbcEmployeeID.Columns("EmployeeID").Caption = rl3("Ma") 'Mã
        tdbcEmployeeID.Columns("EmployeeName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbcObjectID.Columns("ObjectTypeID").Caption = rL3("Loai_DT") 'Loại ĐT
        tdbcObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên

        tdbcObjectTypeID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã
        tdbcObjectTypeID.Columns("ObjectTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAssetID.Columns("AssetID").Caption = rl3("Ma") 'Mã
        tdbcAssetID.Columns("AssetName").Caption = rl3("Ten") 'Tên
        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rl3("Ma") 'Mã
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcPropertyProductID.Columns("D54ProjectID").Caption = rL3("Ma_cong_trinh") 'Mã Dự án
        tdbcPropertyProductID.Columns("D27PropertyProductID").Caption = rL3("Ma_BDS") 'Mã BĐS
        tdbcPropertyProductID.Columns("CipNo").Caption = rL3("Ma_XDCB") 'Mã XDCB

        '================================================================ 
        tdbdCreditAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbdCreditAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        tdbdDebitAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbdDebitAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        tdbdCipID.Columns("CipNo").Caption = rl3("Ma") 'Mã
        tdbdCipID.Columns("CipName").Caption = rl3("Ten") 'Tên
        tdbdAna10ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna10ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdSourceID.Columns("SourceID").Caption = rl3("Ma") 'Mã
        tdbdSourceID.Columns("SourceName").Caption = rl3("Ten") 'Tên
        tdbdAna09ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna09ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna08ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna08ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        '================================================================ 
        tdbdObjectID.Columns("ObjectTypeID").Caption = rL3("Loai_DT") 'Loại ĐT
        tdbdObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbdObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên


        tdbdAna07ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna07ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdObjectTypeID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã
        tdbdObjectTypeID.Columns("ObjectTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        tdbdAna06ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna06ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna05ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna05ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna04ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna04ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna03ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna03ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna02ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna02ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna01ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna01ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAssignmentID.Columns("AssignmentID").Caption = rl3("Ma") 'Mã
        tdbdAssignmentID.Columns("AssignmentName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbg1.Columns("CipID").Caption = rL3("Ma_XDCB") 'Mã XDCB
        tdbg1.Columns("IsNotAllocate").Caption = rL3("Khong_tinh_KH") 'Không tính KH
        tdbg1.Columns("RefDate").Caption = rl3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg1.Columns("SerialNo").Caption = rl3("So_Seri") 'Số Sêri
        tdbg1.Columns("RefNo").Caption = rl3("So_hoa_don") 'Số hóa đơn
        tdbg1.Columns("ObjectTypeID").Caption = rl3("Loai_doi_tuong") 'Loại đối tượng
        tdbg1.Columns("ObjectID").Caption = rl3("Doi_tuong") 'Đối tượng
        tdbg1.Columns("DebitAccountID").Caption = rl3("TK_no") 'TK nợ
        tdbg1.Columns("CreditAccountID").Caption = rl3("TK_co") 'TK có
        tdbg1.Columns("OriginalAmount").Caption = rl3("So_tien_nguyen_te") 'Số tiền nguyên tệ
        tdbg1.Columns("ConvertedAmount").Caption = rl3("So_tien_quy_doi") 'Số tiền quy đổi
        tdbg1.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg1.Columns("SourceID").Caption = rl3("Nguon_hinh_thanh") 'Nguồn hình thành

        tdbg2.Columns("AssignmentID").Caption = rl3("Ma_phan_bo") 'Mã phân bổ
        tdbg2.Columns("AssignmentName").Caption = rl3("Ten_phan_bo") 'Tên phân bổ
        tdbg2.Columns("DebitAccountID").Caption = rl3("TK_no") 'TK Nợ
        tdbg2.Columns("PercentAmount").Caption = rL3("Ty_le") 'Tỷ lệ
        '================================================================ 
        tdbcManagementObjTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcManagementObjTypeID.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcManagementObjID.Columns("ObjectTypeID").Caption = rL3("Loai_DT") 'Loại ĐT
        tdbcManagementObjID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcManagementObjID.Columns("ObjectName").Caption = rL3("Ten") 'Tên
    End Sub

    Private Sub tdbg1_OnAddNew(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg1.OnAddNew
        FooterTotalGrid(tdbg1, COL1_CipID)
        tdbg1.Columns(COL1_IsNotAllocate).Value = 0
    End Sub

    Private Sub tdbg2_OnAddNew(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg2.OnAddNew
        FooterTotalGrid(tdbg1, COL1_CipID)
    End Sub



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0012
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 28/11/2011 03:28:32
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0012() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0012 Set Status = 1 ")
        sSQL.Append(" From D02T0012 A Inner join D02T0100 B on A.CipID = B.CipID ")
        sSQL.Append(" Where ISNULL(A.CipID,'') in (" & arrCipID.ToString & ") And ISNULL(A.AssetID,'')=''")
        Return sSQL
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
        sSQL.Append("Where ISNULL(CipID,'') in (" & arrCipID.ToString & ")")
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
        sSQL.Append("AssetDate = " & SQLDateSave(c1dateAssetDate.Text) & COMMA) 'datetime, NULL
        '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
        sSQL.Append("DepDate = " & SQLDateSave(c1dateDepDate.Text) & COMMA) 'datetime, NULL

        sSQL.Append("UseMonth = " & SQLNumber(Strings.Left(c1dateBeginUsing.Text, 2)) & COMMA) 'tinyint, NULL
        sSQL.Append("UseYear = " & SQLNumber(Strings.Right(c1dateBeginUsing.Text, 4)) & COMMA) 'smallint, NULL
        sSQL.Append("TranMonth = " & SQLNumber(giTranMonth) & COMMA) 'tinyint, NULL
        sSQL.Append("TranYear = " & SQLNumber(giTranYear) & COMMA) 'smallint, NULL
        sSQL.Append("DepMonth = " & SQLNumber(Strings.Left(c1dateBeginDep.Text, 2)) & COMMA) 'tinyint, NULL
        sSQL.Append("DepYear = " & SQLNumber(Strings.Right(c1dateBeginDep.Text, 4)) & COMMA) 'smallint, NULL
        sSQL.Append("ObjectTypeID = " & SQLString(ReturnValueC1Combo(tdbcObjectTypeID)) & COMMA) 'varchar[20], NULL
        sSQL.Append("ObjectID = " & SQLString(ReturnValueC1Combo(tdbcObjectID)) & COMMA) 'varchar[20], NULL
        sSQL.Append("EmployeeID = " & SQLString(ReturnValueC1Combo(tdbcEmployeeID)) & COMMA) 'varchar[20], NULL
        sSQL.Append("NotesU = " & SQLStringUnicode(txtNotes.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("AssetNameU = " & SQLStringUnicode(txtAssetName.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL

        sSQL.Append("FullNameU = " & SQLStringUnicode(txtEmployeeName.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("DepreciatedAmount = " & SQLMoney(txtDepreciateAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("RemainAmount = " & SQLMoney(txtRemainAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("DepreciatedPeriod = " & SQLNumber(txtDepreciatedPeriod.Text) & COMMA) 'int, NULL
        sSQL.Append("ConvertedAmount = " & SQLNumber(txtConvertedAmount.Text) & COMMA) 'int, NULL
        sSQL.Append("Percentage = " & SQLMoney(txtPercentage.Text, DxxFormat.DefaultNumber2) & COMMA) 'money, NULL
        sSQL.Append("AmountDepreciation = " & SQLMoney(txtAmountDepreciation.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
        sSQL.Append("ServiceLife = " & SQLNumber(txtServiceLife.Text) & COMMA) 'int, NULL
        sSQL.Append("IsCompleted = 1, IsRevalued =0, IsDisposed = 0, ")
        sSQL.Append("SetUpFrom = " & SQLString(_setupFrom) & COMMA) 'varchar[20], NULL
        sSQL.Append("UseDate = " & SQLDateSave(c1dateUseDate.Text)) 'datetime, NULL
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
        Optional ByVal PercentAmount As Object = "", Optional ByVal AssignmentID As Object = "", Optional AssetAccountID As String = "",
        Optional DepAccountID As String = "") As StringBuilder
    
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T5000(")
        sSQL.Append("HistoryID, DivisionID, AssetID, BatchID,  ")
        sSQL.Append("BeginMonth, BeginYear, EndMonth, EndYear, HistoryTypeID, Status, InstanceID,IsLiquidated,")
        sSQL.Append("ObjectTypeID, ObjectID, EmployeeID, FullNameU,  ")
        sSQL.Append(" IsStopDepreciation, IsStopUse, ServiceLife,PercentAmount,AssignmentID, ")
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
            sSQL.Append(SQLStringUnicode("") & COMMA) 'FullNameU, nvarchar, NOT NULL

        Else
            sSQL.Append(IIf(ObjectTypeID.ToString = "", "NULL", SQLString(ObjectTypeID)).ToString & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(IIf(ObjectID.ToString = "", "NULL", SQLString(ObjectID)).ToString & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(IIf(EmployeeID.ToString = "", "NULL", SQLString(EmployeeID)).ToString & COMMA) 'EmployeeID, varchar[20], NULL
            sSQL.Append(SQLStringUnicode(FullName, gbUnicode, True) & COMMA) 'FullNameU, nvarchar, NOT NULL

        End If
        sSQL.Append(IIf(IsStopDepreciation.ToString = "", "NULL", SQLNumber(IsStopDepreciation)).ToString & COMMA) 'IsStopDepreciation, tinyint, NULL
        sSQL.Append(IIf(IsStopUse.ToString = "", "NULL", SQLNumber(IsStopUse)).ToString & COMMA) 'IsStopUse, tinyint, NULL
        sSQL.Append(IIf(ServiceLife.ToString = "", "NULL", SQLNumber(ServiceLife)).ToString & COMMA) 'ServiceLife, int, NULL
        sSQL.Append(IIf(PercentAmount.ToString = "", "NULL", SQLMoney(Number(PercentAmount) * 100, DxxFormat.DefaultNumber2)).ToString & COMMA) 'PercentAmount, money, NULL
        sSQL.Append(IIf(AssignmentID.ToString = "", "NULL", SQLString(AssignmentID)).ToString & COMMA) 'AssignmentID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NOT NULL
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        sSQL.Append(SQLString(AssetAccountID) & COMMA) 'AssetAccountID, varchar[20], NULL
        sSQL.Append(SQLString(DepAccountID) & COMMA) 'DepAccountID, varchar[20], NULL
        'sSQL.Append(SQLMoney(?????, DxxFormat.?????) & COMMA) 'ConvertedAmount, money, NULL
        'sSQL.Append(SQLString(?????) & COMMA) 'GroupID, varchar[20], NOT NULL
        'sSQL.Append(SQLString(?????) & COMMA) 'AssetWHID, varchar[50], NOT NUL
       
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

    Private Function SQLInsertD02T5000s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim iCount As Long = tdbg2.RowCount + 3 'thêm 3 dòng cố định
        Dim iFirstHis As Long = 0
        Dim sHistoryID As String = ""
        Dim sSQL As New StringBuilder

        For i As Integer = 0 To tdbg2.RowCount - 1
            sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HD", gsStringKey, sHistoryID, iCount, iFirstHis)
            tdbg2(i, COL2_HistoryID) = sHistoryID

            sRet.Append(SQLInsertD02T5000(sHistoryID, giTranMonth, giTranYear, "AS", , , , , , , , tdbg2(i, COL2_PercentAmount), tdbg2(i, COL2_AssignmentID)).ToString & vbCrLf)
            sSQL = New StringBuilder
        Next
        sRet.Append(SQLInsertD02T5000_3().ToString & vbCrLf)
        Return sRet
    End Function

    Private Function SQLInsertD02T5000_3() As StringBuilder
        Dim sRet As New StringBuilder
        Dim iCount As Long = 3
        Dim iFirstHis As Long = 0
        Dim sHistoryID As String = ""
        Dim sSQL As New StringBuilder

        sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HD", gsStringKey, sHistoryID, iCount, iFirstHis)
        sRet.Append(SQLInsertD02T5000(sHistoryID, Strings.Left(c1dateBeginDep.Text, 2), Strings.Right(c1dateBeginDep.Text, 4), "SD", , , , , 0).ToString & vbCrLf)

        sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HD", gsStringKey, sHistoryID, iCount, iFirstHis)
        sRet.Append(SQLInsertD02T5000(sHistoryID, Strings.Left(c1dateBeginUsing.Text, 2), Strings.Right(c1dateBeginUsing.Text, 4), "SU", , , , , , 0).ToString & vbCrLf)

        sHistoryID = CreateIGENewS("D02T5000", "HistoryID", "02", "HD", gsStringKey, sHistoryID, iCount, iFirstHis)
        sRet.Append(SQLInsertD02T5000(sHistoryID, Strings.Left(c1dateBeginUsing.Text, 2), Strings.Right(c1dateBeginUsing.Text, 4), "SL", , , , , , , txtServiceLife.Text).ToString & vbCrLf)
        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5010
    '# Created User: 
    '# Created Date: 18/11/2021 10:36:43
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
        sSQL.Append(SQLNumber(9999) & COMMA & vbCrLf) 'EndYear, int, NOT NULL
        sSQL.Append(SQLString("") & COMMA) 'GroupID, varchar[50], NOT NULL
        sSQL.Append(SQLString(sAccountID)) 'AccountID, varchar[50], NOT NULL
        sSQL.Append(") ")

        Return sSQL
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

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0012s
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 29/11/2011 08:01:46
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Dim dTotalConvertedAmount As Double = 0
    Private Function SQLInsertD02T0012s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim iCount As Long = tdbg1.RowCount
        Dim iFirstTran As Long = 0
        Dim sTransactionID As String = ""
        Dim sSQL As New StringBuilder

        For i As Integer = 0 To tdbg1.RowCount - 1
            If iUseInvoiceCodeD02 = 0 Then GoTo 1
            If tdbg1(i, COL1_SerialNo).ToString <> "" And tdbg1(i, COL1_RefNo).ToString <> "" And IsNumeric(tdbg1(i, COL1_SerialNo)) Then
                Dim strSQL As String = SQLStoreD95P0105(tdbg1(i, COL1_SerialNo), tdbg1(i, COL1_RefNo))
                Dim dtTemp As DataTable = ReturnDataTable(strSQL)
                If dtTemp.Rows.Count = 0 Then GoTo 1
                If dtTemp.Rows(0).Item("Status").ToString <> "0" Then GoTo 1

                'Dim f As New D95M0240
                'f.ID01 = tdbg1(i, COL1_SerialNo).ToString
                'f.ID02 = tdbg1(i, COL1_RefNo).ToString
                'f.FormActive = "D95F0131"
                'f.ShowDialog()
                'Dim bClose As Boolean = f.bClose
                'f.Dispose()

                '==============
                '18/4/2017, id 96337-Đóng gói V4.1
                Dim bClose As Boolean = False
                Dim sField() As String = {"ID01", "ID02"}
                Dim sValue() As Object = {tdbg1(i, COL1_SerialNo).ToString, tdbg1(i, COL1_RefNo).ToString}
                Dim sOutput() As String = Lemon3.CallDxxMxx40("D95E0240", "D95F0131", "D95F0131", sField, sValue, New String() {"Close"})
                If sOutput.Length > 0 Then bClose = L3Bool(sOutput(0))
                '==============

                If Not bClose Then GoTo 1
                dtTemp = ReturnDataTable(strSQL)
                If dtTemp.Rows.Count = 0 Then GoTo 1 '"BÁn ch§a ¢Ünh nghÚa mÉu hâa ¢¥n cho Sç hâa ¢¥n nªy!"
                If dtTemp.Rows(0).Item("Status").ToString = "0" Then D99C0008.MsgL3("Bạn chưa định nghĩa mẫu hóa đơn cho Số hóa đơn này!")
            End If
1:
            sTransactionID = CreateIGENewS("D02T0012", "TransactionID", "02", "TK", gsStringKey, sTransactionID, iCount, iFirstTran)
            tdbg1(i, COL1_TransactionID) = sTransactionID

            sSQL.Append("Insert Into D02T0012(")
            sSQL.Append("TransactionID, DivisionID, ModuleID,  AssetID, ") 'SplitNo,
            sSQL.Append("VoucherTypeID, VoucherNo, VoucherDate, TranMonth, TranYear, ") 'TransactionDate,
            sSQL.Append(" CurrencyID, ExchangeRate, DebitAccountID, ")
            sSQL.Append("CreditAccountID, OriginalAmount, ConvertedAmount, Status, TransactionTypeID, ")
            sSQL.Append("RefNo, RefDate, Disabled, CreateUserID, CreateDate, ")
            sSQL.Append("LastModifyUserID, LastModifyDate, SeriNo, ObjectTypeID, ObjectID, BatchID,")
            sSQL.Append("Ana01ID, Ana02ID, Ana03ID,Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID,")
            sSQL.Append(" CipID, SourceID,")
            sSQL.Append("Num01, Num02, Num03, Num04, Num05, Date01, Date02, Date03, Date04, Date05,")
            sSQL.Append("DescriptionU,  NotesU, Str01U, Str02U, Str03U, Str04U, Str05U,IsNotAllocate") 'ObjectNameU,ItemNameU
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(tdbg1(i, COL1_TransactionID)) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString("02") & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA) 'AssetID, varchar[20], NULL
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcVoucherTypeID)) & COMMA) 'VoucherTypeID, varchar[20], NULL
            sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[50], NULL
            sSQL.Append(SQLDateSave(c1dateVoucherDate.Text) & COMMA) 'VoucherDate, datetime, NULL
            sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
            sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
            sSQL.Append(SQLString(txtCurrenyID.Text) & COMMA) 'CurrencyID, varchar[20], NOT NULL
            sSQL.Append(SQLMoney(txtExchangRate.Text, DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_DebitAccountID)) & COMMA) 'DebitAccountID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_CreditAccountID)) & COMMA) 'CreditAccountID, varchar[20], NULL
            sSQL.Append(SQLMoney(tdbg1(i, COL1_OriginalAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'OriginalAmount, money, NULL
            sSQL.Append(SQLMoney(tdbg1(i, COL1_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL
            dTotalConvertedAmount += Number(tdbg1(i, COL1_ConvertedAmount))
            sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NOT NULL
            sSQL.Append(SQLString("XDCB") & COMMA) 'TransactionTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_RefNo)) & COMMA) 'RefNo, varchar[50], NULL
            sSQL.Append(SQLDateSave(tdbg1(i, COL1_RefDate)) & COMMA) 'RefDate, datetime, NULL
            sSQL.Append(SQLNumber(0) & COMMA) 'Disabled, bit, NOT NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
            sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
            sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_SerialNo)) & COMMA) 'SeriNo, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_ObjectTypeID)) & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_ObjectID)) & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetID)) & COMMA) 'BatchID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana01ID)) & COMMA) 'Ana01ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana02ID)) & COMMA) 'Ana02ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana03ID)) & COMMA) 'Ana03ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana04ID)) & COMMA) 'Ana04ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana05ID)) & COMMA) 'Ana05ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana06ID)) & COMMA) 'Ana06ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana07ID)) & COMMA) 'Ana07ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana08ID)) & COMMA) 'Ana08ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana09ID)) & COMMA) 'Ana09ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_Ana10ID)) & COMMA) 'Ana10ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_CipID)) & COMMA) 'CipID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg1(i, COL1_SourceID)) & COMMA) 'SourceID, varchar[20], NULL
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
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Description), gbUnicode, True) & COMMA) 'DescriptionU, nvarchar, NOT NULL
            sSQL.Append(SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'NotesU, nvarchar, NOT NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str01), gbUnicode, True) & COMMA) 'Str01U, nvarchar, NOT NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str02), gbUnicode, True) & COMMA) 'Str02U, nvarchar, NOT NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str03), gbUnicode, True) & COMMA) 'Str03U, nvarchar, NOT NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str04), gbUnicode, True) & COMMA) 'Str04U, nvarchar, NOT NULL
            sSQL.Append(SQLStringUnicode(tdbg1(i, COL1_Str05), gbUnicode, True) & COMMA) 'Str05U, nvarchar, NOT NULL
            sSQL.Append(SQLNumber(tdbg1(i, COL1_IsNotAllocate))) 'IsNotAllocate, nvarchar, NOT NULL
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

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
        Return sSQL
    End Function


    Dim iUseInvoiceCodeD02 As Integer = 0
    Private Function GetUseInvoiceCodeD02() As Integer
        Dim strSQL As String = "SELECT UseInvoiceCodeD02 From D95T0000 WITH(NOLOCK) "
        Dim dtTemp As DataTable = ReturnDataTable(strSQL)
        If dtTemp.Rows.Count = 0 Then Exit Function
        Return L3Int(dtTemp.Rows(0).Item("UseInvoiceCodeD02"))
    End Function

    Private Function Level6() As String
        Dim sSQL As String = ""
        'Thêm ngày 18/7/2012 Theo incident 49472 THIHUAN
        sSQL = " UPDATE T01 " & _
                "SET T01.FAString01 = T10.CIPString01,	T01.FAString02 = T10.CIPString02,	" & _
                 "T01.FAString03 = T10.CIPString03,	T01.FAString04 = T10.CIPString04,	" & _
                 "T01.FAString05 = T10.CIPString05,	T01.FAString06 = T10.CIPString06,	" & _
                 "T01.FAString07 = T10.CIPString07,	T01.FAString08= T10.CIPString08,	" & _
                 "T01.FAString09 = T10.CIPString09,	T01.FAString10 = T10.CIPString10," & _
                "	T01.FADate01 = T10.CIPDate01,		T01.FADate02 = T10.CIPDate02,	" & _
                "	T01.FADate03 = T10.CIPDate03,		T01.FADate04 = T10.CIPDate04,	" & _
                "	T01.FADate05 = T10.CIPDate05,		T01.FADate06 = T10.CIPDate06,	" & _
                "	T01.FADate07 = T10.CIPDate07,		T01.FADate08 = T10.CIPDate08,	" & _
                "	T01.FADate09 = T10.CIPDate09,		T01.FADate10 = T10.CIPDate10," & _
                "	T01.FANum01 = T10.CIPNum01,		T01.FANum02 = T10.CIPNum02," & _
                "	T01.FANum03 = T10.CIPNum03,		T01.FANum04 = T10.CIPNum04," & _
                "	T01.FANum05 = T10.CIPNum05,		T01.FANum06 = T10.CIPNum06," & _
                "	T01.FANum07 = T10.CIPNum07,		T01.FANum08 = T10.CIPNum08," & _
                "T01.SupplierOTID = T10.SupplierOTID, T01.SupplierID = T10.SupplierID, " & _
                "T01.FAString01U = T10.CIPString01U, T01.FAString02U = T10.CIPString02U, " & _
                "T01.FAString03U = T10.CIPString03U, T01.FAString04U = T10.CIPString04U, " & _
                "T01.FAString05U = T10.CIPString05U, T01.FAString06U = T10.CIPString06U," & _
                "T01.FAString07U = T10.CIPString07U, T01.FAString08U= T10.CIPString08U, " & _
                "T01.FAString09U = T10.CIPString09U, T01.FAString10U = T10.CIPString10U, " & _
                "	T01.FANum09 = T10.CIPNum09,		T01.FANum10 = T10.CIPNum10" & vbCrLf
        sSQL &= "FROM 		D02T0001  T01 WITH(NOLOCK)" & vbCrLf
        sSQL &= "LEFT JOIN 	D02T0012 T12 WITH(NOLOCK) ON T01.AssetID = T12.AssetID" & vbCrLf
        sSQL &= "INNER JOIN 	D02T0100 T10 WITH(NOLOCK) ON T10.CipID = T12.CipID AND 	T10.Status = 2" & vbCrLf
        sSQL &= "WHERE 		ISNULL(T10.CipID,'') IN (" & SQLString(tdbg1.Columns(COL1_CipID).Value) & ") AND ISNULL(T12.AssetID,'')  = " & SQLString(tdbcAssetID.Text) & " AND 	T12.TransactionTypeID = 'XDCB'"
        Return sSQL
    End Function

#Region "Events of textbox"

    Private Sub txtAmountDepreciation_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmountDepreciation.TextChanged
        txtAmountDepreciation.Text = Format(Number(txtAmountDepreciation.Text), DxxFormat.D90_ConvertedDecimals)
        'RemainAmount = ConvertedAmount - AmountDepreciation
        txtRemainAmount.Text = Format(Number(txtConvertedAmount.Text) - Number(txtAmountDepreciation.Text), DxxFormat.D90_ConvertedDecimals)
    End Sub

    Private Sub txtConvertedAmount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtConvertedAmount.TextChanged
        txtConvertedAmount.Text = FormatNumber(Number(txtConvertedAmount.Text), DxxFormat.iD90_ConvertedDecimals)
        txtRemainAmount.Text = Format(Number(txtConvertedAmount.Text), DxxFormat.D90_ConvertedDecimals)
    End Sub

    Private Sub txtExchangRate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExchangRate.TextChanged
        txtExchangRate.Text = FormatNumber(Number(txtExchangRate.Text), DxxFormat.iExchangeRateDecimals)
    End Sub

    Private Sub txtNotes_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.TextChanged
        If txtDescription.Text = "" Then Exit Sub
        Dim valueright As String = ""
        If txtDescription.Text.Length > 1 Then valueright = txtDescription.Text.Substring(1)
        txtDescription.Text = Strings.Left(txtDescription.Text, 1).ToUpper & valueright
    End Sub

    Private Sub txtRemainAmount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemainAmount.TextChanged, txtDepreciateAmount.TextChanged
        Dim txtAmount As TextBox = CType(sender, TextBox)
        txtAmount.Text = Format(Number(txtAmount.Text), DxxFormat.D90_ConvertedDecimals)
    End Sub

    Private Sub txtServiceLife_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMethodID.KeyPress, txtMethodEndID.KeyPress, txtServiceLife.KeyPress, txtDepreciatedPeriod.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtServiceLife_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtServiceLife.TextChanged
        If D02Systems.IsCalPeriodByRate = True Then Exit Sub '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ

        'Lấy ServiceLife
        Dim dServiceLife As Double = Number(IIf(Number(txtServiceLife.Text) = 0, 1, Number(txtServiceLife.Text)))
        'DepreciatedAmount = ConvertedAmount / ServiceLife
        txtDepreciateAmount.Text = SQLNumber((Number(txtConvertedAmount.Text)) / dServiceLife, DxxFormat.D90_ConvertedDecimals)
        'Percentage = 100 / ServiceLife
        txtPercentage.Text = SQLNumber(100 / dServiceLife, DxxFormat.DefaultNumber2)
        txtServiceLife.Text = Format(Number(txtServiceLife.Text), DxxFormat.DefaultNumber0)
        txtPercentage.Tag = txtPercentage.Text
    End Sub

    Private Sub txtDepreciatedPeriod_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDepreciatedPeriod.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtDepreciatedPeriod_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDepreciatedPeriod.Validated
        txtDepreciatedPeriod.Text = SQLNumber(txtDepreciatedPeriod.Text, DxxFormat.DefaultNumber0)
    End Sub
#End Region

#Region "Nhập mã trên lưới"
    ''' <summary>
    ''' Kiểm tra Mã nhập vào lưới
    ''' </summary>
    ''' <param name="tdbg">grid name</param>
    ''' <param name="iCol">col id</param>
    ''' <returns>True/False</returns>
    ''' <remarks>Put tdbg_BeforeColUpdate events (Exp: e.Cancel = L3IsID(tdbg, e.ColIndex, ???))</remarks>
    Public Function L3IsID(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iCol As Integer) As Boolean
        Return L3IsIDFormula(tdbg, iCol, False)
    End Function

    'Kiểm tra Button Đóng có đặt Tên "Close"
    Private Function CheckContinue(ByVal ctrl As Control) As Boolean
        Try
            Dim form As Form = CType(ctrl.TopLevelControl, Form)
            If form.Controls.ContainsKey("btnClose") Then
                Dim btnClose As Control = CType(form.Controls("btnClose"), System.Windows.Forms.Button)
                If btnClose Is Nothing Then Return True 'không có nút đóng
                If btnClose.Focused Then Return False
                Dim arr() As String = ctrl.Tag.ToString.Split(";"c)
                If arr.Length > 1 Then Return False
                '************
            End If
        Catch ex As Exception

        End Try
        Return True
    End Function

    ''' <summary>
    ''' Thay đổi vị trí Select của chuỗi Vni
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="posFrom">vị trí bắt đầu</param>
    ''' <param name="posTo">Số ký tự được Select</param>
    ''' <remarks>Không cần kiểm tra khi Unicode</remarks>
    Private Sub ChangePositionIndexVNI(ByVal str As String, ByRef posFrom As Integer, ByRef posTo As Integer)
        If str = "" OrElse posFrom < 0 OrElse posFrom >= str.Length - 1 Then Exit Sub

        Dim arrChar() As String = {"Â", "Á", "À", "Å", "Ä", "Ã", "Ù", "Ø", "Û", "Õ", "Ï", "É", "È", "Ú", "Ü", "Ë", "Ê"}
        Dim c As String = (str.Substring(posFrom, 1)).ToUpper
        '"Ö", "Ô"
        Select Case c
            Case "Ö", "Ô" 'Ö: Ư; Ô: Ơ - không tăng vị trí, ngược lại thì tăng thêm 1 vị trí
                If L3FindArrString(arrChar, (str.Substring(posFrom + 1, 1)).ToUpper) Then posTo = 2
            Case Else 'kiểm tra trong danh sách arrChar
                If L3FindArrString(arrChar, c) Then
                    If posFrom > 0 Then posFrom -= 1
                    posTo = 2
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Kiểm tra Mã hợp lệ 
    ''' </summary>
    ''' <param name="str">Chuỗi kiểm tra</param>
    ''' <returns>Vị trí ký tự vi phạm</returns>
    ''' <remarks></remarks>
    Private Function IndexIdCharactor(ByVal str As String) As Integer
        '  If str.Length > iLength Then Return -2 'vượt chiều dài
        'BackSpace: 8
        For Each c As Char In str
            Select Case AscW(c)
                Case 13, 10 'Mutiline của textbox và phím Enter
                    Continue For
                Case Is < 33, Is > 127, 37, 39, 91, 93, 94 'Các ký tự đặc biệt: 37(%) 39(') 91([) 93(]) 94(^)
                    Return str.IndexOf(c)
            End Select
        Next
        Return -1
    End Function

    '''' Kiểm tra công thức hợp lệ
    '''' </summary>
    '''' <param name="str">Chuỗi kiểm tra</param>
    '''' <returns>Vị trí ký tự vi phạm</returns>
    '''' <remarks></remarks>
    Private Function IndexFormulaCharactor(ByVal str As String) As Integer
        '  If str.Length > iLength Then Return -2 'vượt chiều dài
        'BackSpace: 8
        For Each c As Char In str
            Select Case AscW(c)
                Case 13, 10 'Mutiline của textbox và phím Enter
                    Continue For
                Case Is < 33, Is > 127, 94 ''Các ký tự đặc biệt: 94(^)
                    Return str.IndexOf(c)
            End Select
        Next
        Return -1
    End Function

    Private Function L3IsIDFormula(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iCol As Integer, ByVal bFormula As Boolean) As Boolean
        'Nếu nhấn đóng thì không cần hiện thông báo
        If CheckContinue(tdbg) = False Then Return False

        Dim posFrom As Integer = -1
        If bFormula Then
            posFrom = IndexFormulaCharactor(tdbg.Columns(iCol).Text)
        Else
            posFrom = IndexIdCharactor(tdbg.Columns(iCol).Text)
        End If

        If posFrom >= 0 Then
            Dim posTo As Integer = 1
            If tdbg.Font.Name.Contains("Lemon3") Then ChangePositionIndexVNI(tdbg.Columns(iCol).Text, posFrom, posTo)
            Dim sValid As String = tdbg.Columns(iCol).Text.Substring(posFrom, posTo)
            If tdbg.Font.Name.Contains("Lemon3") And sValid <> " " Then sValid = ConvertVniToUnicode(sValid)
            D99C0008.MsgL3(rl3("Ma_co_ky_tu_khong_hop_le") & Space(1) & "[" & sValid & "]")
            Return True
        End If

        Return False
    End Function
#End Region

    Private Sub LoadInherit()
        LoadTabpage1()
        'Dim s As String = "Exec D02P0091 'LEMONADMIN', 'DRD32', 'TUONGAN', 8, 2012, 1"
        Dim dtTemp As DataTable = ReturnDataTable(SQLStoreD02P0091)
        If dtTemp.Rows.Count > 0 Then
            dtGrid1.Merge(dtTemp)
            'Thêm ngày 20/7/2012 theo incident 49472
            If _fromCall = "D02F2005" Then
                txtAssetName.Text = dtGrid1.Rows(0).Item("CipName").ToString
                txtNotes.Text = dtGrid1.Rows(0).Item("Desciption").ToString
            End If

            LoadTDBGrid1()
            If dtGrid1.Rows.Count = 0 Then Exit Sub
            'Gán ObjectTypeID nếu các dòng trùng nhau
            Dim dtObjectID As DataTable = dtGrid1.DefaultView.ToTable(True, COL1_ObjectTypeID, COL1_ObjectID)
            If dtObjectID.Rows.Count = dtGrid1.Rows.Count Then
                tdbcObjectTypeID.SelectedValue = dtGrid1.Rows(0).Item(COL1_ObjectTypeID)
                tdbcObjectID.SelectedValue = dtGrid1.Rows(0).Item(COL1_ObjectID)
            End If
            dtObjectID.Dispose()
            Dim dtEmployeeID As DataTable = dtGrid1.DefaultView.ToTable(True, "EmployeeID")
            If dtEmployeeID.Rows.Count = dtGrid1.Rows.Count Then tdbcEmployeeID.SelectedValue = dtGrid1.Rows(0).Item("EmployeeID")
            dtEmployeeID.Dispose()
            If dtGrid1 IsNot Nothing And dtGrid1.Rows.Count > 0 Then txtConvertedAmount.Text = FormatNumber(dtGrid1.Compute("SUM(" & COL1_ConvertedAmount & ")", ""), DxxFormat.iD90_ConvertedDecimals)
        End If
    End Sub

    Private Sub tdbg1_FetchRowStyle(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.FetchRowStyleEventArgs) Handles tdbg1.FetchRowStyle
        If tdbg1.RowCount = 0 Then Exit Sub
        If L3Int(tdbg1(e.Row, COL1_IsEdit)) = 1 Then e.CellStyle.Locked = True 'Nếu Mã XDCB đang TH sửa thì lock lại
        'tdbg1.Splits(0).DisplayColumns(tdbg1.Col).Button = L3Int(tdbg1(tdbg1.Row, COL1_IsEdit)) <> 1
    End Sub

    Private Sub tdbg1_BeforeDelete(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.CancelEventArgs) Handles tdbg1.BeforeDelete
        If L3Int(tdbg1.Columns(COL1_IsEdit).Text) = 1 Then e.Cancel = True
    End Sub



    Private Sub tdbg1_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg1.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        '--- Đổ nguồn cho các Dropdown phụ thuộc
        Select Case tdbg1.Columns(tdbg1.Col).DataField

            Case COL1_ObjectID
                LoadtdbdObjectID(tdbg1(tdbg1.Row, COL1_ObjectTypeID).ToString)

        End Select
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
            tdbg2.Columns(iCol).Text = L3String(GetProperties(frm, "AssignmentID"))
        End If
    End Sub
#End Region

    'ID : 224617 - BỔ SUNG Cho phép gọi màn hình THIẾT LẬP DANH MỤC TÀI SẢN CỐ ĐỊNH tại bước hình thành TS
#Region "Gọi đến D02F0087"
    Dim bFormD02F0087 As Boolean = False
    Private Sub ShowD02F0087()
        Dim sMethodID As String = ""
        Dim sSQLD91T1001_SaveLastKey As String = ""
        'tdbcAssetID.Text = ""

        If tdbcAssetID.Text = "+" Then
            If iPerD02F0087 < 2 Then
                D99C0008.MsgL3(rL3("Ban_khong_co_quyen_tao_ma_tu_dong")) 'BÁn kh¤ng câ quyÒn th£m mìi
                Exit Sub
            End If

            'If D02Systems.AssetAuto = 1 Then

            '    Dim arrPro() As StructureProperties = Nothing
            '    SetProperties(arrPro, "IndexTab", 0)
            '    SetProperties(arrPro, "FormIDPermission", "D02F3000")
            '    Dim frm As Form = CallFormShowDialog("D02D1240", "D02F3001", arrPro)
            '    If frm Is Nothing Then Exit Sub 'TH form đã gọi rồi thì không gọi nữa
            '    If L3Bool(GetProperties(frm, "SavedOk")) Then
            '        AssetID = GetProperties(frm, "AssetID").ToString
            '    End If

            'ElseIf D02Systems.AssetAuto = 2 Then
            '    If D02Systems.IsShowFormAutoCreate = True Then
            '        Dim arrPro() As StructureProperties = Nothing
            '        SetProperties(arrPro, "FormIDPermission", "D02F0087")
            '        Dim frm As Form = CallFormShowDialog("D02D1040", "D02F0087", arrPro)
            '        AssetID = GetProperties(frm, "sAssetID").ToString ' Tạo sAssetID để nhận AssetID được trả về từ Form gọi
            '        sMethodID = GetProperties(frm, "sMethodID").ToString
            '        sSQLD91T1001_SaveLastKey = GetProperties(frm, "sSQLD91T1001_SaveLastKey").ToString
            '        bFormD02F0087 = True
            '    Else
            '        LoadIGEMethodID()
            '        sMethodID = sDefaultIGEMethodID
            '    End If

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
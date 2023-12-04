'#-------------------------------------------------------------------------------------
'# Created Date: 11/09/2007 9:19:02 AM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 11/09/2007 9:19:02 AM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text

Public Class D02F5558
    Private _batchID As String
    Private _voucherNo As String

    Public Property BatchID() As String
        Get
            Return _batchID
        End Get
        Set(ByVal value As String)
            If BatchID = value Then
                _batchID = ""
                Return
            End If
            _batchID = value
        End Set
    End Property

    Public Property VoucherNo() As String
        Get
            Return _voucherNo
        End Get
        Set(ByVal value As String)
            If VoucherNo = value Then
                _voucherNo = ""
                Return
            End If
            _voucherNo = value
        End Set
    End Property

    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            _FormState = value
            Select Case _FormState
                Case EnumFormState.FormEditVoucher
                    txtVoucherNoOld.Enabled = False
                    txtVoucherNoOld.Text = _voucherNo
                    txtVoucherNoNew.Focus()
            End Select
        End Set
    End Property


    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F5558_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F5558_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Loadlanguage()
        btnSave.Enabled = ReturnPermission(PARA_FormIDPermission) > EnumPermission.View

    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)
        btnSave.Enabled = False
        btnClose.Enabled = False
        gbSavedOK = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder

        'Lưu LastKey của Số phiếu xuống Database (gọi hàm CreateIGEVoucherNo bật cờ True)
        'Kiểm tra trùng Số phiếu (gọi hàm CheckDuplicateVoucherNo)
        If CheckDuplicateVoucherNo(D02, "D02T0012", _batchID, txtVoucherNoNew.Text) Then
            Me.Cursor = Cursors.Default
            btnSave.Enabled = True
            btnClose.Enabled = True
            Exit Sub
        End If
        sSQL.Append(SQLStoreD02P5558)
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            gbSavedOK = True
            btnClose.Enabled = True
            btnSave.Enabled = True
            btnClose.Focus()
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9102
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 11/09/2007 09:30:35
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD91P9102() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D91P9102 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(D02) & COMMA 'ModuleID, varchar[20], NOT NULL
        sSQL &= SQLString("D02T0012") & COMMA 'TableName, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'VoucherID, varchar[20], NOT NULL
        sSQL &= SQLString(txtVoucherNoNew.Text) & COMMA 'VoucherNo, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) 'Language, varchar[20], NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P5558
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 11/09/2007 10:00:23
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P5558() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P5558 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(_batchID) & COMMA 'BatchID, varchar[20], NOT NULL
        sSQL &= SQLString(_voucherNo) & COMMA 'OldVoucherNo, varchar[20], NOT NULL
        sSQL &= SQLString(txtVoucherNoNew.Text) 'NewVoucherNo, varchar[20], NOT NULL
        Return sSQL
    End Function

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Sua_so_phieu_-_D02F5558") 'Sõa sç phiÕu - D02F5558
        '================================================================ 
        lblVoucherNoOld.Text = rl3("So_phieu_goc") 'Số phiếu gốc
        lblVoucherNoNew.Text = rl3("So_phieu_moi") 'rl3("So_phieu_moi") 'Số phiếu mới
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 

    End Sub


End Class
Public Class D02F0333

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F0333_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        Loadlanguage()
        optBAL.Enabled = IsMinPeriod(giTranMonth, giTranYear)
        btnNext.Enabled = ReturnPermission("D02F0300") > 0
        
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Hinh_thanh_TSCD_-_D02F0333") & UnicodeCaption(gbUnicode) 'HØnh thªnh TSC˜ - D02F0333
        '================================================================ 
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnNext.Text = "&" & rl3("Tiep_tuc") 'Tiếp tục
        '================================================================ 
        optBAL.Text = rl3("Nhap_so_du") 'Nhập số dư
        optCip.Text = rL3("Tu_xay_dung_co_ban") 'Từ xây dựng cơ bản
        optNew.Text = rL3("Mua_moi") 'Mua mới
        optAll.Text = rL3("Tat_ca1") 'Tất cả
        optCAP.Text = rL3("Dieu_dong_von")
        '================================================================ 
        grpOri.Text = rL3("Nguon_goc_hinh_thanh_TSCD") 'Nguồn gốc hình thành TSCĐ
    End Sub



    Public Function IsMinPeriod(ByVal Month As Integer, _
                            ByVal Year As Integer) As Boolean
        Dim sSQL As String = "Select top 1 TranMonth, TranYear From D02T9999" & vbCrLf
        sSQL &= "Where DivisionID=" & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "Order by Tranmonth + TranYear*100"
        Dim dtTemp As DataTable = ReturnDataTable(sSQL)
        If dtTemp.Rows.Count = 0 Then Return False
        If Month = L3Int(dtTemp.Rows(0).Item("TranMonth")) And Year = L3Int(dtTemp.Rows(0).Item("TranYear")) Then Return True
        Return False
    End Function

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Dim sSetupFrom As String = "ALL"
        If optAll.Checked Then
            sSetupFrom = "ALL"
        ElseIf optNew.Checked Then
            sSetupFrom = "NEW"
        ElseIf optCip.Checked Then
            sSetupFrom = "CIP"
        ElseIf optBAL.Checked Then
            sSetupFrom = "BAL"
        ElseIf optCAP.Checked Then
            sSetupFrom = "CAP"
        End If
        Dim f As New D02F0300
        f.SetupFrom = sSetupFrom
        f.ShowInTaskbar = True
        f.ShowDialog()
        f.Dispose()

    End Sub
End Class
Imports System.Text
''' <summary>
''' Module này dùng để khai báo các Sub và Function toàn cục
''' </summary>
''' <remarks>Các khai báo Sub và Function ở đây không được trùng với các khai báo
''' ở các module D99Xxxxx
''' </remarks>
Public Enum enumSetUpFrom
    ALL = 0 'Tất cả
    [NEW] = 1 'Mua mới
    CIP = 2 'XDCB
    BAL = 3 'Nhập số dư
    CAP = 4 'điều động vốn
End Enum

Module D02X0002
    ''' <summary>
    ''' Cập nhật số thứ tự cho lưới
    ''' </summary>
    Public Sub UpdateOrderNum(ByVal TDBGrid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iCol As Integer)
        For i As Integer = 0 To TDBGrid.RowCount - 1
            TDBGrid(i, iCol) = i + 1
        Next
    End Sub

    ''' <summary>
    ''' Kiểm tra sự tồn tại của 1 giá trị trong 1 cột trên lưới với nguồn dữ liệu trong TDBDropdown
    ''' </summary>
    Public Function CheckExist(ByVal pTDBD As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal piCol As Integer, ByVal sText As String) As Boolean
        For i As Integer = 0 To pTDBD.RowCount - 1
            pTDBD.Row = i
            If pTDBD.Columns(piCol).Text = sText Then Return True
        Next
        Return False
    End Function

    Private Function FindSxType(ByVal nType As String, ByVal s As String) As String
        Select Case nType.Trim
            Case "1" ' Theo tháng
                Return giTranMonth.ToString("00")
            Case "2" ' Theo năm
                Return giTranYear.ToString
            Case "3" ' Theo loại chứng từ
                Return s
            Case "4" ' Theo đơn vị
                Return gsDivisionID
            Case "5" ' Theo hằng
                Return s
            Case Else
                Return ""
        End Select
    End Function
  
    ''' <summary>
    ''' Xác định ví trí hiện hành của lưới
    ''' </summary>
    Public Sub SetCurrentRow(ByVal TDBGrid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iCol As Integer, ByVal sText As String)
        If TDBGrid.RowCount > 0 Then
            For i As Integer = 0 To TDBGrid.RowCount - 1
                If TDBGrid(i, iCol).ToString() = sText Then
                    TDBGrid.Row = i
                    Exit Sub
                End If
            Next
            TDBGrid.Row = 0
        End If
    End Sub

    '''' <summary>
    '''' Tính tổng cho 1 cột tương ứng trên lưới
    '''' </summary>
    '''' <param name="ipCol"></param>
    '''' <param name="C1Grid"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>

    'Public Function Sum(ByVal ipCol As Integer, ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid) As Double
    '    Dim lSum As Double = 0
    '    For i As Integer = 0 To C1Grid.RowCount - 1
    '        If C1Grid(i, ipCol) Is Nothing OrElse TypeOf (C1Grid(i, ipcol)) Is DBNull Then Continue For
    '        lSum += Convert.ToDouble(C1Grid(i, ipCol))
    '    Next
    '    Return lSum
    'End Function

   
    '#--------------------------------------------------------------------------
    '#CreateUser: Trần Thị Ái Trâm
    '#CreateDate: 04/09/2007
    '#ModifiedUser:
    '#ModifiedDate:
    '#Description: Hàm kiểm tra Audit log
    '#--------------------------------------------------------------------------
    Public Function PermissionAudit(ByVal sAuditCode As String) As Byte
        Dim sSQL As String
        Dim dt As DataTable

        sSQL = "Select Audit From D91T9200 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where AuditCode=" & SQLString(sAuditCode)

        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            If CByte(dt.Rows(0).Item("Audit")) = 1 Then
                Return 1
            Else
                Return 0
            End If
        Else
            Return 0
        End If
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9106
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 04/09/2007 11:30:16
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    'Private Function SQLStoreD91P9106(ByVal sAuditCode As String, ByVal sEventID As String, ByVal sDesc1 As String, ByVal sDesc2 As String, ByVal sDesc3 As String, ByVal sDesc4 As String, ByVal sDesc5 As String, ByVal nIsAuditDetail As Integer, ByVal sAuditItemID As String) As String
    '    Dim sSQL As String = ""
    '    sSQL &= "Exec D91P9106 "
    '    sSQL &= SQLDateTimeSave(Now) & COMMA 'AuditDate, datetime, NOT NULL
    '    sSQL &= SQLString(sAuditCode) & COMMA 'AuditCode, varchar[20], NOT NULL
    '    sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
    '    sSQL &= SQLString("02") & COMMA 'ModuleID, varchar[2], NOT NULL
    '    sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
    '    sSQL &= SQLString(sEventID) & COMMA 'EventID, varchar[20], NOT NULL
    '    sSQL &= SQLString(sDesc1) & COMMA 'Desc1, varchar[250], NOT NULL
    '    sSQL &= SQLString(sDesc2) & COMMA 'Desc2, varchar[250], NOT NULL
    '    sSQL &= SQLString(sDesc3) & COMMA 'Desc3, varchar[250], NOT NULL
    '    sSQL &= SQLString(sDesc4) & COMMA 'Desc4, varchar[250], NOT NULL
    '    sSQL &= SQLString(sDesc5) & COMMA 'Desc5, varchar[250], NOT NULL
    '    sSQL &= SQLNumber(nIsAuditDetail) & COMMA 'IsAuditDetail,tinyint
    '    sSQL &= SQLString(sAuditItemID)  'AuditItemID, varchar[50], NOT NULL

    '    Return sSQL
    'End Function

    '#--------------------------------------------------------------------------
    '#CreateUser: Trần Thị ÁiTrâm
    '#CreateDate: 04/09/2007
    '#ModifiedUser:
    '#ModifiedDate:
    '#Description: Thực thi store Audit Log
    '#--------------------------------------------------------------------------
    'Public Sub ExecuteAuditLog(ByVal sAuditCode As String, ByVal sEventID As String, Optional ByVal sDesc1 As String = "", Optional ByVal sDesc2 As String = "", Optional ByVal sDesc3 As String = "", Optional ByVal sDesc4 As String = "", Optional ByVal sDesc5 As String = "", Optional ByVal nIsAuditDetail As Integer = 0, Optional ByVal sAuditItemID As String = "")
    '    Dim sSQL As String
    '    sSQL = SQLStoreD91P9106(sAuditCode, sEventID, sDesc1, sDesc2, sDesc3, sDesc4, sDesc5, nIsAuditDetail, sAuditItemID)
    '    ExecuteSQL(sSQL)
    'End Sub

    'Public Sub LoadFormatsNew()
    '    '#------------------------------------------------------
    '    '#CreateUser: Trần Thị Ái Trâm
    '    '#CreateDate: 06/10/2009
    '    '#Description: Format so theo D91

    '    Const Number2 As String = "#,##0.00"
    '    Const Number4 As String = "#,##0.0000" 'dung Format ty le thue
    '    Const Number0 As String = "#,##0"
    '    Dim sSQL As String = "Exec D91P9300 "
    '    Dim dt As DataTable
    '    dt = ReturnDataTable(sSQL)
    '    With D02Format
    '        If dt.Rows.Count > 0 Then
    '            .ExchangeRate = InsertFormat(dt.Rows(0).Item("ExchangeRateDecimals"))
    '            .iExchangeRate = L3Int(dt.Rows(0).Item("ExchangeRateDecimals"))
    '            .DecimalPlaces = InsertFormat(dt.Rows(0).Item("DecimalPlaces"))
    '            .MyOriginal = .DecimalPlaces
    '            .iD90_Converted = L3Int(dt.Rows(0).Item("D90_ConvertedDecimals"))
    '            .D90_Converted = InsertFormat(dt.Rows(0).Item("D90_ConvertedDecimals"))
    '            .D07_Quantity = InsertFormat(dt.Rows(0).Item("D07_QuantityDecimals"))
    '            .D07_UnitCost = InsertFormat(dt.Rows(0).Item("D07_UnitCostDecimals"))
    '            .D08_Quantity = InsertFormat(dt.Rows(0).Item("D08_QuantityDecimals"))
    '            .D08_UnitCost = InsertFormat(dt.Rows(0).Item("D08_UnitCostDecimals"))
    '            .D08_Ratio = InsertFormat(dt.Rows(0).Item("D08_RatioDecimals"))
    '            .D90_ConvertedDecimals = CInt(dt.Rows(0).Item("D90_ConvertedDecimals"))
    '            .BaseCurrencyID = (IIf(IsDBNull(dt.Rows(0).Item("BaseCurrencyID").ToString), "", dt.Rows(0).Item("BaseCurrencyID").ToString)).ToString

    '            '.BOMQty = InsertFormat(dt.Rows(0).Item("BOMQtyDecimals"))
    '            '.BOMPrice = InsertFormat(dt.Rows(0).Item("BOMPriceDecimals"))
    '            '.BOMAmt = InsertFormat(dt.Rows(0).Item("BOMAmtDecimals"))
    '        Else
    '            .ExchangeRate = Number2
    '            .D90_Converted = Number2
    '            .D07_Quantity = Number2
    '            .D07_UnitCost = Number2
    '            .D08_Quantity = Number2
    '            .D08_UnitCost = Number2
    '            .D08_Ratio = Number2
    '            .D90_ConvertedDecimals = 0
    '            .DecimalSeparator = ","
    '            .ThousandSeparator = "."
    '            .BaseCurrencyID = ""
    '            '.BOMQty = Number2
    '            '.BOMPrice = Number2
    '            '.BOMAmt = Number2
    '        End If
    '        .DefaultNumber2 = Number2
    '        .DefaultNumber4 = Number4
    '        .DefaultNumber0 = Number0
    '    End With
    'End Sub

    'Public Function InsertFormat(ByVal ONumber As Object) As String
    '    Dim iNumber As Int16 = Convert.ToInt16(ONumber)
    '    Dim sRet As String = "#,##0"
    '    If iNumber = 0 Then
    '    Else
    '        sRet &= "." & Strings.StrDup(iNumber, "0")
    '    End If
    '    Return sRet
    'End Function

    'Public Function GetOriginalDecimal(ByVal sCurrencyID As String) As String

    '    Dim sSQL As String
    '    sSQL = "Select OriginalDecimal From D91V0010 Where CurrencyID = " & SQLString(sCurrencyID)
    '    Dim dt As DataTable
    '    dt = ReturnDataTable(sSQL)
    '    If dt.Rows.Count > 0 Then
    '        Return InsertFormat(dt.Rows(0).Item("OriginalDecimal"))
    '    Else
    '        Return DxxFormat.DecimalPlaces
    '    End If
    'End Function

    Public Function InserZero(ByVal NumZero As Byte) As String
        '#------------------------------------------------------
        '#CreateUser: Nguyen Thi Minh Hoa
        '#CreateDate: 04/04/2006
        '#ModifiedUser:  Nguyen Thi Minh Hoa
        '#ModifiedDate:  04/04/2006
        '#Description: Format so theo D91
        '#------------------------------------------------------
        If NumZero = 0 Then
            InserZero = ""
        Else
            InserZero = "."
            InserZero &= StrDup(NumZero, "0")
        End If
    End Function

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

    'Hàm ReturnTableFilter cải tiến
    Public Function ReturnTableFilter1(ByVal dt As DataTable, ByVal sWhereClause As String) As DataTable
        Dim dt1 As DataTable
        dt.DefaultView.RowFilter = sWhereClause
        dt1 = dt.DefaultView.ToTable
        Return dt1
    End Function

    Public Function SetGetDateSQL() As String
        Dim sSQL As String
        sSQL = "Select Getdate() as CreateDate "
        Return ReturnScalar(sSQL)
    End Function

    Public Sub Run1(ByVal sEXECHILD As String)
        If Not ExistFile(gsApplicationSetup & "\" & EXECHILD & ".exe") Then Exit Sub
        Dim pInfo As New System.Diagnostics.ProcessStartInfo(gsApplicationSetup & "\" & EXECHILD & ".exe")
        pInfo.Arguments = "/DigiNet Corporation"
        pInfo.WindowStyle = ProcessWindowStyle.Normal
        Process.Start(pInfo)
    End Sub

    ''' <summary>
    ''' Kiểm tra tồn tại exe con không ?
    ''' </summary>
    Private Function ExistFile(ByVal Path As String) As Boolean
        If System.IO.File.Exists(Path) Then Return True
        If geLanguage = EnumLanguage.Vietnamese Then
            D99C0008.MsgL3("Không tồn tại file " & EXECHILD & ".exe")
        Else
            D99C0008.MsgL3("Not exist file " & EXECHILD & ".exe")
        End If
        Return False
    End Function

    'Câu đổ nguồn chung cho SubReport
    Public Function SQLSubReport(ByVal sDivisionID As String) As String
        Dim sSQL As String = ""
        sSQL = "Select * From D91V0016" & vbCrLf
        sSQL &= "Where DivisionID = " & SQLString(sDivisionID)
        Return sSQL
    End Function



#Region "Màn hình chọn đường dẫn báo cáo"

    'Public Function GetReportPath(ByVal ReportTypeID As String, ByVal ReportName As String, ByVal CustomReport As String, ByRef ReportPath As String, Optional ByRef ReportTitle As String = "", Optional ByVal ModuleID As String = "02") As String
    '    'Dim bShowReportPath As Boolean
    '    'Dim iReportLanguage As Byte
    '    ''Lấy giá trị PARA_ModuleID từ module gọi đến
    '    ''Nếu là exe chính (không có biến PARA_ModuleID) thì lấy Dxx 
    '    'bShowReportPath = CType(D99C0007.GetModulesSetting("D" & PARA_ModuleID, ModuleOption.lmOptions, "ShowReportPath", "True"), Boolean)
    '    'iReportLanguage = CType(D99C0007.GetModulesSetting("D" & PARA_ModuleID, ModuleOption.lmOptions, "ReportLanguage", "0"), Byte)
    '    ''Lấy đường dẫn báo cáo từ module D99X0004
    '    'ReportPath = UnicodeGetReportPath(gbUnicode, iReportLanguage, "")
    '    'If bShowReportPath Then 'Hiển thị màn hình chọn đường dẫn báo cáo
    '    '    Dim frm As New D99F6666
    '    '    With frm
    '    '        .ModuleID = ModuleID '2 ký tự, tùy theo từng module có thể lấy theo module gốc chứa exe con hoặc module gọi đến.
    '    '        .ReportTypeID = ReportTypeID
    '    '        .ReportName = ReportName
    '    '        .CustomReport = CustomReport
    '    '        .ReportPath = ReportPath
    '    '        .ReportTitle = ReportTitle
    '    '        .ShowDialog()
    '    '        ReportName = .ReportName
    '    '        ReportPath = .ReportPath
    '    '        gsReportPath = ReportPath 'biến toàn cục đang dùng 
    '    '        ReportTitle = .ReportTitle
    '    '        SaveOptionReport(.ShowReportPath)
    '    '        .Dispose()
    '    '    End With
    '    'Else 'Không hiển thị thì lấy theo Loại nghiệp vụ (nếu có)
    '    '    If CustomReport <> "" Then
    '    '        ReportPath = gsApplicationSetup & "\XCustom\"
    '    '        ReportName = CustomReport
    '    '    End If
    '    'End If
    '    'ReportPath = ReportPath & ReportName & ".rpt"
    '    'Return ReportName
    '    Return Lemon3.Reports.GetReportPath(ReportTypeID, ReportName, CustomReport, ReportPath, ReportTitle, ModuleID, D02Options.ShowReportPath, D02Options.ReportLanguage)
    'End Function
    'Tùy thuộc từng module có biến lưu dưới Registry
    'Public Sub SaveOptionReport(ByVal bShowReportPath As Boolean)
    '    'D99C0007.SaveModulesSetting("D" & PARA_ModuleID, ModuleOption.lmOptions, "ShowReportPath", bShowReportPath)
    '    If "D" & PARA_ModuleID = D02 Then 'Module gốc
    '        'Nếu module nào có thêm code VB6 thì lưu thêm nhánh VB6
    '        'SaveSetting("Lemon3_D05", "Options", "NotShowDirectory", (Not bShowReportPath).ToString) 'Nhánh VB6
    '        D02Options.ShowReportPath = bShowReportPath 'Biến Tùy chọn
    '    End If
    'End Sub
#End Region

    Public Function convertStringToEnum(ByVal sValue As String) As enumSetUpFrom
        Dim eResult As enumSetUpFrom = enumSetUpFrom.ALL
        Select Case sValue.ToUpper.Trim
            Case "BAL"
                eResult = enumSetUpFrom.BAL
            Case "CIP"
                eResult = enumSetUpFrom.CIP
            Case "NEW"
                eResult = enumSetUpFrom.[NEW]
            Case "CAP"
                eResult = enumSetUpFrom.CAP
        End Select
        Return eResult
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0500
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 28/11/2011 08:21:47
    '# Modified User: 
    '# Modified Date: 
    '# Description: Load edit master
    '#---------------------------------------------------------------------------------------------------
    Public Function SQLStoreD02P0500(ByVal _assetID As String, ByVal _setupForm As String) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0500 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(_setupForm) & COMMA 'SetUpFrom, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString("D02.AssetID =" & SQLString(_assetID)) & COMMA 'strFind, varchar[8000], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode)
        Return sSQL
    End Function

    'Lấy kỳ đầu tiên dùng để phân quyền : màn hình D02F0300
    Public Sub GetFirstPeriod(ByVal gsDivisionID As String)
        Dim dt As DataTable
        Dim sSQL As String
        sSQL = "Select  TranMonth , TranYear From D02T9999 WITH(NOLOCK)  " & vbCrLf
        sSQL = sSQL & "Where DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL = sSQL & "Order By TranYear , TranMonth "
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            giFirstTranMonth = CInt(dt.Rows(0)("TranMonth").ToString())
            giFirstTranYear = CInt(dt.Rows(0)("TranYear").ToString())
        End If
    End Sub
    Public Sub SetImageButton(ByVal btnSave As System.Windows.Forms.Button, ByVal btnNotSave As System.Windows.Forms.Button, ByVal imgButton As ImageList)
        btnSave.Size = New System.Drawing.Size(76, 27)
        btnNotSave.Size = New System.Drawing.Size(100, 27)

        btnSave.ImageList = imgButton
        btnSave.ImageIndex = 0
        btnSave.ImageAlign = ContentAlignment.MiddleLeft

        btnNotSave.ImageList = imgButton
        btnNotSave.ImageIndex = 2
        btnNotSave.ImageAlign = ContentAlignment.MiddleLeft

        btnNotSave.Text = rL3("_Khong_luu")
        btnSave.Text = rL3("_Luu") '&Lưu
    End Sub


#Region "Import/Export Excel trực tiếp"

    'Public Const FileNameExcel As String = "Data.xls"
    'Public giVersion2007 As Integer = -1 ' =1 là Office2007, =0 là Office2003
    'Public Function ExportToExcelFromGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sFileName As String, Optional ByVal sFilter As String = "", Optional ByVal arrColAlwaysShow() As String = Nothing, Optional ByVal arrColAlwaysHide() As String = Nothing) As Boolean
    '    Try
    '        'Chỉ xuất những dòng phù hợp với sFilter truyền vào
    '        If sFileName = "" Then sFileName = FileNameExcel
    '        Dim dtG As DataTable = CType(tdbg.DataSource, DataTable)
    '        dtG.DefaultView.RowFilter = sFilter
    '        '=====================================================
    '        Dim oExcel As Object = CreateObject("Excel.Application") 'Microsoft.Office.Interop.Excel.Application 'Class() ' Create the Excel Application object
    '        'Update 04/07/2013 kiểm tra máy đang cài phiên bản Office nào

    '        If giVersion2007 = -1 Then
    '            CheckVersionExcel(oExcel)
    '        End If

    '        'D99C0008.Msg(1)
    '        If giVersion2007 = -1 Then
    '            ' Kiểm tra nếu dữ liệu > 65530 dong hoặc >256 cột thì chỉ chạy trên Office 2007
    '            If tdbg.RowCount > 65530 Then
    '                MessageBox.Show(ConvertUnicodeToVietwareF(rL3("So_dong_vuot_qua_gioi_han_cho_phep_cua_Excel") & " (" & tdbg.RowCount & " > 65530)"), MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                oExcel.Quit()
    '                Return False
    '            ElseIf tdbg.Columns.Count > 256 Then
    '                MessageBox.Show(ConvertUnicodeToVietwareF(rL3("So_cot_vuot_qua_gioi_han_cho_phep_cua_Excel") & " (" & tdbg.Columns.Count & "> 256)"), MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                oExcel.Quit()
    '                Return False
    '            End If
    '        End If

    '        '******************************************************
    '        'Kiểm tra tồn tại file Excel đang xuất
    '        If CloseProcessWindowMax(sFileName) = False Then oExcel.Quit() : Return False
    '        '******************************************************
    '        ' D99C0008.Msg(2)
    '        Dim oPathFile As String = System.Windows.Forms.Application.StartupPath + "\" + sFileName ' "Data.xlsx"
    '        Dim oWorkbook As Object = oExcel.Workbooks.Add(Type.Missing) ' Create a new Excel Workbook
    '        Dim oWorkSheet As Object ' Create a new Excel Worksheet
    '        Dim sFirstCol As String = "A"
    '        Dim iFirstCol As Integer = GetIntColumnExcel(sFirstCol) 'Đổi cột Chuỗi sang Số (VD: cột A đổi thành cột 0)

    '        Dim StartValue As Integer = 0 ' Vị trí bắt đầu của Rang excel
    '        Dim EndValue As Integer = 0 ' Vị trí cuối cùng của Rang excel
    '        Dim PrevPos As Integer = 0 ' Vị trí kế tiếp của Rang excel

    '        Dim iMaxRow As Long = tdbg.RowCount
    '        Dim iPackage As Integer = 0 'Số gói cần chạy 
    '        Dim iLimitRow As Integer 'Số dòng chạy cho 1 gói iPackage
    '        Dim iRowCount As Integer = 0 ' Tổng số dòng của 1 rang cần khởi tạo (rawData)
    '        Dim iStartRow As Integer = 0 ' Dòng bắt đầu chạy cho dtData trong 1 gói
    '        Dim iEndRow As Integer = 0 ' Dòng cuối cùng chạy cho dtData trong 1 gói

    '        Try
    '            ' Tạo Sheet mới
    '            oWorkSheet = oWorkbook.Worksheets(1) ' CType(oWorkbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
    '            oWorkSheet.Name = "Data"

    '            'Xác định số dòng dữ liệu cần xuất cho 1 gói (iPackage)
    '            If iMaxRow > 10000 Then
    '                iLimitRow = 1000
    '            Else
    '                iLimitRow = 100
    '            End If

    '            ' Tìm ký tự "A...Z" của cột
    '            Dim finalColLetter As String = String.Empty
    '            Dim colCharset As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    '            Dim colCharsetLen As Integer = colCharset.Length
    '            If tdbg.Columns.Count > colCharsetLen Then
    '                finalColLetter = colCharset.Substring((tdbg.Columns.Count - 1 + iFirstCol) \ colCharsetLen - 1, 1)
    '            End If
    '            finalColLetter += colCharset.Substring((tdbg.Columns.Count - 1 + iFirstCol) Mod colCharsetLen, 1)

    '            'Lấy số gói cần chạy để đưa vào rang excel
    '            If iMaxRow <= iLimitRow Then ' Nếu Số dòng xuất <= iLimitRow thì 1 Package
    '                iPackage = 1
    '            Else
    '                Dim iDecimal As Integer = CInt(iMaxRow Mod iLimitRow)
    '                If iDecimal = 0 Then ' Chia hết
    '                    iPackage = CInt(iMaxRow / iLimitRow)
    '                Else ' Chia dư -> Package = Package + 1
    '                    iPackage = CInt(Int(iMaxRow / iLimitRow))
    '                    iPackage += 1
    '                End If
    '            End If

    '            'Tạo bảng và Add Data
    '            For iX As Long = 1 To iPackage
    '                If iX < iPackage Then
    '                    If iX = 1 Then 'Lần đầu tiên
    '                        iStartRow = iEndRow
    '                    Else
    '                        iStartRow = iEndRow + 1
    '                    End If
    '                    iEndRow = CInt(iLimitRow * iX) - 1
    '                    iRowCount = iLimitRow
    '                Else
    '                    If iPackage = 1 Then
    '                        iStartRow = 0
    '                        iEndRow = CInt(iMaxRow - 1)
    '                        iRowCount = CInt(iMaxRow)
    '                    Else
    '                        iStartRow = iEndRow + 1
    '                        iEndRow = CInt(iMaxRow - 1)
    '                        iRowCount = CInt(iMaxRow - (iLimitRow * (iX - 1)))
    '                    End If
    '                End If
    '                ' D99C0008.Msg(3)
    '                StartValue = PrevPos + 1
    '                If iX = 1 Then ' Table đầu tiên dành vị trí A1 cho header
    '                    EndValue = iRowCount + PrevPos + 3
    '                Else ' Các Table sau nối tiếp theo không tạo header
    '                    EndValue = iRowCount + PrevPos
    '                End If

    '                'Tạo mảng dữ liệu để đưa vào file excel
    '                Dim arrData(iRowCount + 2, tdbg.Columns.Count - 1) As Object
    '                Dim sColumnFieldName As String = ""
    '                Dim sColumnCaption As String = ""
    '                Dim iRow_Data As Integer = 0
    '                'Phải lấy dữ liệu của bảng dtCaptionCols để kiểm tra, vì bảng dtData có thứ tự cột không đúng như trên lưới
    '                For col As Integer = 0 To tdbg.Columns.Count - 1
    '                    sColumnFieldName = tdbg.Columns(col).DataField
    '                    sColumnCaption = tdbg.Columns(col).Caption
    '                    'If Not gbUnicode Then sColumnCaption = ConvertVniToUnicode(sColumnCaption)

    '                    iRow_Data = 0
    '                    If iX = 1 Then ' Đưa dữ liệu vào dòng tiêu đề (Header)
    '                        arrData(0, col) = "'" & sColumnFieldName
    '                        arrData(1, col) = tdbg.Columns(col).NumberFormat
    '                        arrData(2, col) = "'" & sColumnCaption
    '                        iRow_Data += 3
    '                    End If

    '                    ' Đưa dữ liệu vào các dòng kế tiếp
    '                    For row As Integer = iStartRow To iEndRow
    '                        'Dim dr As DataRow = dtData.Rows(row)
    '                        ''Nếu cột là chuỗi thì thêm dấu ' phía trước để khi xuất Excel thì hiểu giá trị là chuỗi
    '                        If tdbg.Columns(col).DataType.Name = "String" Then
    '                            If gbUnicode Then
    '                                arrData(iRow_Data, col) = "'" & tdbg(row, col).ToString
    '                            Else 'Nếu nhập liệu VNI thì ConvertVniToUnicode dữ liệu dạng chuỗi sang Unicode
    '                                arrData(iRow_Data, col) = "'" & ConvertVniToUnicode(tdbg(row, col).ToString)
    '                            End If
    '                        ElseIf tdbg.Columns(col).DataType.Name = "Boolean" Then
    '                            arrData(iRow_Data, col) = SQLNumber(tdbg(row, col))
    '                        Else
    '                            arrData(iRow_Data, col) = tdbg(row, col)
    '                        End If
    '                        iRow_Data += 1
    '                    Next
    '                Next
    '                ' D99C0008.Msg(4)
    '                ' Fast data export to Excel
    '                Dim excelRange As String = String.Format(sFirstCol & "{0}:{1}{2}", StartValue, finalColLetter, EndValue)
    '                oWorkSheet.Range(excelRange, Type.Missing).Value2 = arrData
    '                'Khung
    '                oWorkSheet.Range(excelRange, Type.Missing).Borders.LineStyle = 1 ' Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                oWorkSheet.Range(excelRange, Type.Missing).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)

    '                'Định dạng các cột Excel
    '                Dim range As Object 'Microsoft.Office.Interop.Excel.Range
    '                'oWorkSheet.Columns.AutoFit()
    '                Dim colIndex As Integer = iFirstCol '0

    '                For i As Integer = 0 To tdbg.Columns.Count - 1
    '                    'Xác định vị trí vùng Range
    '                    range = oWorkSheet.Range(GetStringColumnExcel(colIndex) & IIf(iX = 1, StartValue + 3, StartValue + 1).ToString, GetStringColumnExcel(colIndex) & EndValue)  'DirectCast(oWorkSheet.Range(GetStringColumnExcel(colIndex) & IIf(iX = 1, StartValue + 3, StartValue + 1).ToString, GetStringColumnExcel(colIndex) & EndValue), Microsoft.Office.Interop.Excel.Range)
    '                    '=======================================================
    '                    Dim bVisible As Boolean = False
    '                    Dim dWidth As Integer = 0
    '                    For iSplit As Integer = 0 To tdbg.Splits.Count - 1
    '                        If GetIndexofArray(arrColAlwaysShow, tdbg.Columns(i).DataField) <> -1 Then ' Cột được hiện khi xuất Excel
    '                            bVisible = True
    '                            dWidth = tdbg.Splits(iSplit).DisplayColumns(i).Width
    '                        ElseIf GetIndexofArray(arrColAlwaysHide, tdbg.Columns(i).DataField) <> -1 Then ' Cột được ẩn khi xuất Excel
    '                            bVisible = False
    '                        Else
    '                            bVisible = tdbg.Splits(iSplit).DisplayColumns(i).Visible
    '                            dWidth = tdbg.Splits(iSplit).DisplayColumns(i).Width
    '                        End If
    '                        If bVisible Then Exit For
    '                    Next
    '                    range.EntireColumn.ColumnWidth = dWidth * (1 / 6)
    '                    range.EntireColumn.Hidden = Not bVisible
    '                    '=======================================================
    '                    Select Case tdbg.Columns(i).DataType.Name
    '                        Case "Decimal" 'Số thập phân
    '                            If tdbg.Columns(i).NumberFormat = "Percent" Then
    '                                range.EntireColumn.NumberFormat = "0.00%"
    '                            Else
    '                                range.EntireColumn.NumberFormat = "#,##0" & InsertZero(L3Int(tdbg.Columns(i).NumberFormat.Replace("N", "")))
    '                            End If
    '                            range.EntireColumn.HorizontalAlignment = -4152 '= Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
    '                        Case "Boolean", "Byte" ' Boolean, Byte là cột checkbox
    '                            range.HorizontalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                        Case "Integer" 'Số nguyên
    '                            range.NumberFormat = "#,##0"
    '                            range.HorizontalAlignment = -4152 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
    '                        Case "DateTime" 'Ngày
    '                            range.EntireColumn.NumberFormat = tdbg.Columns(i).NumberFormat
    '                            range.EntireColumn.HorizontalAlignment = -4108  'Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                        Case Else
    '                            range.HorizontalAlignment = -4131 'Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
    '                    End Select
    '                    'Định dạng hiển thị dữ liệu cho cột
    '                    colIndex = colIndex + 1
    '                Next
    '                If iX = 1 Then 'Header
    '                    range = oWorkSheet.Rows(3, Type.Missing) 'TryCast(oWorkSheet.Rows(3, Type.Missing), Microsoft.Office.Interop.Excel.Range)
    '                    range.Font.Bold = True
    '                    range.EntireRow.VerticalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                    range.EntireRow.HorizontalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                    'Mau nen
    '                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
    '                End If
    '                PrevPos = EndValue 'Giữ vị trí cuối cùng của table trước đó
    '            Next
    '            ' D99C0008.Msg(5)
    '            'Hide row 1,2
    '            Dim rrow As Object = oWorkSheet.Range("A1", "A2")
    '            rrow.EntireRow.Hidden = True

    '            'Tắt cảnh báo hỏi có muốn Save As không?
    '            oExcel.DisplayAlerts = False

    '            If giVersion2007 = 1 Then
    '                oWorkbook.SaveAs(oPathFile, FileFormat:=56)
    '            Else
    '                oWorkbook.SaveAs(oPathFile)
    '            End If

    '            oExcel.Workbooks.Open(oPathFile)
    '            oExcel.Visible = True
    '            Return True
    '        Catch ex As Exception
    '            oWorkbook.Close(False, Type.Missing, Type.Missing)
    '            ' Release the Application object
    '            oExcel.Quit()
    '        Finally
    '            oWorkSheet = Nothing
    '            oWorkbook = Nothing
    '            If oExcel IsNot Nothing Then oExcel = Nothing
    '            System.GC.Collect()
    '            System.GC.WaitForPendingFinalizers()
    '        End Try
    '        dtG.DefaultView.RowFilter = ""
    '    Catch ex As Exception
    '        D99C0008.Msg(ex.Message)
    '    End Try
    'End Function

    'Private Function GetIndexofArray(ByVal arr() As String, ByVal sValue As String) As Integer
    '    For i As Integer = 0 To arr.Length - 1
    '        If L3String(arr(i)) = sValue Then Return i
    '    Next
    '    Return -1
    'End Function

    'Public Function CheckVersionExcel(ByVal appExcel As Object) As Boolean
    '    'Dim appExcel As New Microsoft.Office.Interop.Excel.Application
    '    ' appExcel = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
    '    If L3Int(appExcel.Version) >= 12 Then
    '        giVersion2007 = 1
    '        Return True
    '    End If
    '    Return False
    'End Function

    'Public Function CloseProcessWindowMax(Optional ByVal FileName As String = FileNameExcel, Optional ByVal bShowMessage As Boolean = True) As Boolean
    '    'Doan code dung de dong file Excel mo san (khong phai do Chuong trinh mo)
    '    Dim p As System.Diagnostics.Process = Nothing
    '    Dim sWindowName As String = "Microsoft Excel - " & FileName
    '    Try
    '        For Each pr As Process In Process.GetProcessesByName("EXCEL")
    '            If sWindowName = pr.MainWindowTitle OrElse pr.MainWindowTitle = sWindowName.Substring(0, sWindowName.LastIndexOf(".")) Then
    '                If p Is Nothing Then
    '                    p = pr
    '                ElseIf p.StartTime < pr.StartTime Then
    '                    p = pr
    '                End If
    '            End If
    '        Next
    '        If p IsNot Nothing Then
    '            If bShowMessage Then
    '                If (D99C0008.MsgAsk(rL3("Ban_phai_dong_File") & Space(1) & FileName & Space(1) & rL3("truoc_khi_xuat_Excel") & "." & vbCrLf & rL3("Ban_co_muon_dong_khong")) = Windows.Forms.DialogResult.Yes) Then
    '                    p.Kill()
    '                    Return True
    '                Else
    '                    Return False
    '                End If
    '            Else
    '                p.Kill()
    '                Return True
    '            End If

    '        End If
    '        Return True 'False
    '    Catch ex As Exception
    '        Return True
    '    End Try
    'End Function

    'Private Function GetIntColumnExcel(ByVal sColumn As String) As Integer
    '    'Update 8/1/2014: Cách mới cho Office từ 2007 về sau
    '    Dim charColumn() As Char = sColumn.ToCharArray()
    '    Dim sum As Integer = 0
    '    For i As Integer = 0 To charColumn.Length - 1
    '        sum *= 26
    '        sum += (Asc(sColumn(i)) - Asc("A") + 1)
    '    Next
    '    Return sum - 1
    'End Function

    'Private Function GetStringColumnExcel(ByVal sColumn As Integer) As String
    '    Dim divNumber As Integer = sColumn + 1
    '    Dim columnName As String = ""
    '    Dim modNumber As Integer
    '    While divNumber > 0
    '        modNumber = (divNumber - 1) Mod 26
    '        columnName = Convert.ToChar(65 + modNumber).ToString() & columnName
    '        divNumber = CInt(((divNumber - modNumber) / 26))
    '    End While
    '    Return columnName
    'End Function

    Public Function ImportExcelToGrid(ByRef dtExport As DataTable, Optional ByVal sFilter As String = "", Optional ByRef sOutputPath As String = "") As Boolean
        Dim file As New OpenFileDialog
        file.Filter = "Excel File (*.xls;*.xlsx)|*.xlsx;*.xls"
        If (file.ShowDialog() = System.Windows.Forms.DialogResult.OK) Then
            Try
                sOutputPath = file.FileName
                Dim stream As System.IO.Stream = System.IO.File.Open(file.FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                Dim excelReader As Excel.IExcelDataReader
                If file.FileName.Contains(".xlsx") Then
                    excelReader = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream)
                Else
                    excelReader = Excel.ExcelReaderFactory.CreateBinaryReader(stream)
                End If
                excelReader.IsFirstRowAsColumnNames = True
                Dim ds As DataSet = excelReader.AsDataSet
                excelReader.Close()
                Try
                    Dim dtTEm As DataTable = ds.Tables(0)
                    'Xóa những dòng tương ứng với sFilter
                    If sFilter <> "" Then
                        Dim dr() As DataRow = dtExport.Select(sFilter)
                        For i As Integer = dr.Length - 1 To 0 Step -1
                            dtExport.Rows.Remove(dr(i))
                        Next
                    End If
                    '=============================================
                    For k As Integer = 2 To dtTEm.Rows.Count - 1
                        dtExport.Rows.Add()
                        For i As Integer = 0 To dtExport.Columns.Count - 1
                            Try
                                If dtExport.Columns(dtExport.Columns(i).Caption).DataType.Name = "Boolean" Then
                                    dtExport.Rows(dtExport.Rows.Count - 1).Item(dtExport.Columns(i).Caption) = L3Bool(dtTEm.Rows(k).Item(dtExport.Columns(i).Caption))
                                ElseIf dtExport.Columns(dtExport.Columns(i).Caption).DataType.Name = "DateTime" Then
                                    If L3String(dtTEm.Rows(k).Item(dtExport.Columns(i).Caption)) = "" Then
                                        '   dtExport.Rows(dtExport.Rows.Count - 1).Item(dtExport.Columns(i).Caption) = Nothing
                                    Else
                                        dtExport.Rows(dtExport.Rows.Count - 1).Item(dtExport.Columns(i).Caption) = DateTime.FromOADate(Number(dtTEm.Rows(k).Item(dtExport.Columns(i).Caption)))
                                    End If

                                Else
                                    dtExport.Rows(dtExport.Rows.Count - 1).Item(dtExport.Columns(i).Caption) = dtTEm.Rows(k).Item(dtExport.Columns(i).Caption)
                                End If
                            Catch ex As Exception

                            End Try
                        Next
                    Next
                Catch ex As Exception
                    D99C0008.MsgL3(ex.Message)
                End Try
                Return True
            Catch ex As Exception
                D99C0008.MsgL3(ex.Message)
            End Try
        End If
        Return False
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T5558
    '# Created User: HUỲNH KHANH
    '# Created Date: 24/09/2014 09:36:11
    '#---------------------------------------------------------------------------------------------------
    'Private Function SQLInsertD02T5558(ByVal sVoucherIGE As String, ByVal sOldVoucherNo As String, ByVal sNewVoucherNo As String) As StringBuilder
    '    Dim sSQL As New StringBuilder
    '    sSQL.Append("-- Cap nhat VoucherID vào D02T5558" & vbCrLf)
    '    sSQL.Append("Insert Into D02T5558(")
    '    sSQL.Append("BatchID, OldVoucherNo, NewVoucherNo, CreateUserID, CreateDate, ")
    '    sSQL.Append("TranMonth, TranYear, DivisionID")
    '    sSQL.Append(") Values(" & vbCrLf)
    '    sSQL.Append(SQLString(sVoucherIGE) & COMMA) 'BatchID, varchar[20], NOT NULL
    '    sSQL.Append(SQLString(sOldVoucherNo) & COMMA) 'OldVoucherNo, varchar[20], NOT NULL
    '    sSQL.Append(SQLString(sNewVoucherNo) & COMMA) 'NewVoucherNo, varchar[20], NOT NULL
    '    sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
    '    sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
    '    sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NOT NULL
    '    sSQL.Append(SQLNumber(giTranYear)) 'TranYear, int, NOT NULL
    '    sSQL.Append(COMMA & SQLString(gsDivisionID))
    '    sSQL.Append(")")
    '    Return sSQL
    'End Function

    Public Sub InsertD02T5558(ByVal sVoucherID As String, ByVal sOldVoucherNo As String, ByVal sNewVoucherNo As String)
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Cap nhat VoucherID vao D02T5558" & vbCrLf)
        sSQL.Append("Insert Into D02T5558(")
        sSQL.Append("BatchID, OldVoucherNo, NewVoucherNo, CreateUserID, CreateDate, ")
        sSQL.Append("TranMonth, TranYear, DivisionID") '18/10/2018, id 1141596-Lỗi hình thành tài sản cố định
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(sVoucherID) & COMMA) 'VoucherID, varchar[20], NOT NULL
        sSQL.Append(SQLString(sOldVoucherNo) & COMMA) 'OldVoucherNo, varchar[20], NOT NULL
        sSQL.Append(SQLString(sNewVoucherNo) & COMMA) 'NewVoucherNo, varchar[20], NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NOT NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, int, NOT NULL
        sSQL.Append(SQLString(gsDivisionID))
        sSQL.Append(") ")

        ExecuteSQL(sSQL.ToString)
    End Sub

#End Region

End Module

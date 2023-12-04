Imports System
Public Module D99X6001

#Region "Import/Export Excel trực tiếp từ lưới"

    Private Const FileNameExcel As String = "Data.xls"
    Private giVersion2007 As Integer = -1 ' =1 là Office2007, =0 là Office2003
    Private MAX_ROW_EXCEL, iFirstCol As Integer
    Private sFirstCol As String = "A"

    Private Sub ExportExcelForWorksheet(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dtG As DataTable, ByVal oWorkbook As Object, ByVal iSheet As Integer, Optional ByVal arrColAlwaysShow() As String = Nothing, Optional ByVal arrColAlwaysHide() As String = Nothing)
        Dim StartValue As Integer = 0 ' Vị trí bắt đầu của Rang excel
        Dim EndValue As Integer = 0 ' Vị trí cuối cùng của Rang excel
        Dim PrevPos As Integer = 0 ' Vị trí kế tiếp của Rang excel

        Dim iMaxRow As Long = dtG.Rows.Count
        Dim iPackage As Integer = 0 'Số gói cần chạy 
        Dim iLimitRow As Integer 'Số dòng chạy cho 1 gói iPackage
        Dim iRowCount As Integer = 0 ' Tổng số dòng của 1 rang cần khởi tạo (rawData)
        Dim iStartRow As Integer = 0 ' Dòng bắt đầu chạy cho dtData trong 1 gói
        Dim iEndRow As Integer = 0 ' Dòng cuối cùng chạy cho dtData trong 1 gói

        Dim oWorkSheet As Object
        ' Tạo Sheet mới
        If iSheet < 4 Then
            oWorkSheet = oWorkbook.Worksheets(iSheet) ' CType(oWorkbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)        oWorkSheet.Name = "Data"
        Else
            oWorkSheet = oWorkbook.Worksheets.Add(Type.Missing, oWorkbook.Worksheets(iSheet - 1), 1, Type.Missing)
        End If
        If iSheet = 1 Then
            oWorkSheet.Name = "Data"
        Else
            oWorkSheet.Name = "Data " & (iSheet - 1).ToString("00")
        End If

        'Xác định số dòng dữ liệu cần xuất cho 1 gói (iPackage)
        If iMaxRow > 10000 Then
            iLimitRow = 1000
        Else
            iLimitRow = 100
        End If

        ' Tìm ký tự "A...Z" của cột
        Dim finalColLetter As String = String.Empty
        Dim colCharset As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim colCharsetLen As Integer = colCharset.Length
        If C1Grid.Columns.Count > colCharsetLen Then
            finalColLetter = colCharset.Substring((C1Grid.Columns.Count - 1 + iFirstCol) \ colCharsetLen - 1, 1)
        End If
        finalColLetter += colCharset.Substring((C1Grid.Columns.Count - 1 + iFirstCol) Mod colCharsetLen, 1)

        'Lấy số gói cần chạy để đưa vào rang excel
        If iMaxRow <= iLimitRow Then ' Nếu Số dòng xuất <= iLimitRow thì 1 Package
            iPackage = 1
        Else
            Dim iDecimal As Integer = CInt(iMaxRow Mod iLimitRow)
            If iDecimal = 0 Then ' Chia hết
                iPackage = CInt(iMaxRow / iLimitRow)
            Else ' Chia dư -> Package = Package + 1
                iPackage = CInt(Int(iMaxRow / iLimitRow))
                iPackage += 1
            End If
        End If

        'Tạo bảng và Add Data
        For iX As Long = 1 To iPackage
            If iX < iPackage Then
                If iX = 1 Then 'Lần đầu tiên
                    iStartRow = iEndRow
                Else
                    iStartRow = iEndRow + 1
                End If
                iEndRow = CInt(iLimitRow * iX) - 1
                iRowCount = iLimitRow
            Else
                If iPackage = 1 Then
                    iStartRow = 0
                    iEndRow = CInt(iMaxRow - 1)
                    iRowCount = CInt(iMaxRow)
                Else
                    iStartRow = iEndRow + 1
                    iEndRow = CInt(iMaxRow - 1)
                    iRowCount = CInt(iMaxRow - (iLimitRow * (iX - 1)))
                End If
            End If
            ' D99C0008.Msg(3)
            StartValue = PrevPos + 1
            If iX = 1 Then ' Table đầu tiên dành vị trí A1 cho header
                EndValue = iRowCount + PrevPos + 3
            Else ' Các Table sau nối tiếp theo không tạo header
                EndValue = iRowCount + PrevPos
            End If

            'Tạo mảng dữ liệu để đưa vào file excel
            Dim arrData(iRowCount + 2, C1Grid.Columns.Count - 1) As Object
            Dim sColumnFieldName As String = ""
            Dim sColumnCaption As String = ""
            Dim iRow_Data As Integer = 0
            'Phải lấy dữ liệu của bảng dtCaptionCols để kiểm tra, vì bảng dtData có thứ tự cột không đúng như trên lưới
            For col As Integer = 0 To C1Grid.Columns.Count - 1
                sColumnFieldName = C1Grid.Columns(col).DataField
                If sColumnFieldName = "" Then Continue For
                sColumnCaption = C1Grid.Columns(col).Caption
                'If Not gbUnicode Then sColumnCaption = ConvertVniToUnicode(sColumnCaption)

                iRow_Data = 0
                If iX = 1 Then ' Đưa dữ liệu vào dòng tiêu đề (Header)
                    arrData(0, col) = "'" & sColumnFieldName
                    arrData(1, col) = C1Grid.Columns(col).NumberFormat
                    arrData(2, col) = "'" & sColumnCaption
                    iRow_Data += 3
                End If

                ' Đưa dữ liệu vào các dòng kế tiếp
                For row As Integer = iStartRow To iEndRow
                    'Dim dr As DataRow = dtData.Rows(row)
                    ''Nếu cột là chuỗi thì thêm dấu ' phía trước để khi xuất Excel thì hiểu giá trị là chuỗi
                    If C1Grid.Columns(col).DataType.Name = "String" Then
                        If gbUnicode Then
                            arrData(iRow_Data, col) = "'" & L3String(dtG.Rows(row).Item(C1Grid.Columns(col).DataField))
                        Else 'Nếu nhập liệu VNI thì ConvertVniToUnicode dữ liệu dạng chuỗi sang Unicode
                            arrData(iRow_Data, col) = "'" & ConvertVniToUnicode(L3String(dtG.Rows(row).Item(C1Grid.Columns(col).DataField)))
                        End If
                    ElseIf C1Grid.Columns(col).DataType.Name = "Boolean" Then
                        arrData(iRow_Data, col) = SQLNumber(dtG.Rows(row).Item(C1Grid.Columns(col).DataField))
                    Else
                        arrData(iRow_Data, col) = dtG.Rows(row).Item(C1Grid.Columns(col).DataField)
                    End If
                    iRow_Data += 1
                Next
            Next
            ' D99C0008.Msg(4)
            ' Fast data export to Excel
            Dim excelRange As String = String.Format(sFirstCol & "{0}:{1}{2}", StartValue, finalColLetter, EndValue)
            oWorkSheet.Range(excelRange, Type.Missing).Value2 = arrData
            'Khung
            oWorkSheet.Range(excelRange, Type.Missing).Borders.LineStyle = 1 ' Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            oWorkSheet.Range(excelRange, Type.Missing).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)

            'Định dạng các cột Excel
            Dim range As Object 'Microsoft.Office.Interop.Excel.Range
            'oWorkSheet.Columns.AutoFit()
            Dim colIndex As Integer = iFirstCol '0

            For i As Integer = 0 To C1Grid.Columns.Count - 1
                'Xác định vị trí vùng Range
                range = oWorkSheet.Range(GetStringColumnExcel(colIndex) & IIf(iX = 1, StartValue + 3, StartValue + 1).ToString, GetStringColumnExcel(colIndex) & EndValue)  'DirectCast(oWorkSheet.Range(GetStringColumnExcel(colIndex) & IIf(iX = 1, StartValue + 3, StartValue + 1).ToString, GetStringColumnExcel(colIndex) & EndValue), Microsoft.Office.Interop.Excel.Range)
                '=======================================================
                Dim bVisible As Boolean = False
                Dim dWidth As Integer = 0
                For iSplit As Integer = 0 To C1Grid.Splits.Count - 1
                    If CheckContainsInArray(arrColAlwaysShow, C1Grid.Columns(i).DataField) Then ' Cột được hiện khi xuất Excel
                        bVisible = True
                        dWidth = C1Grid.Splits(iSplit).DisplayColumns(i).Width
                    ElseIf CheckContainsInArray(arrColAlwaysHide, C1Grid.Columns(i).DataField) Then ' Cột được ẩn khi xuất Excel
                        bVisible = False
                    Else
                        bVisible = C1Grid.Splits(iSplit).DisplayColumns(i).Visible
                        dWidth = C1Grid.Splits(iSplit).DisplayColumns(i).Width
                    End If
                    If bVisible Then Exit For
                Next
                range.EntireColumn.ColumnWidth = dWidth * (1 / 6)
                range.EntireColumn.Hidden = Not bVisible
                '=======================================================
                Select Case C1Grid.Columns(i).DataType.Name
                    Case "Decimal" 'Số thập phân
                        If C1Grid.Columns(i).NumberFormat = "Percent" Or C1Grid.Columns(i).NumberFormat.Contains("%") Then
                            range.EntireColumn.NumberFormat = "0.00%"
                        Else
                            If C1Grid.Columns(i).NumberFormat.Contains("#,##") Then
                                range.EntireColumn.NumberFormat = C1Grid.Columns(i).NumberFormat
                            Else
                                range.EntireColumn.NumberFormat = "#,##0" & InsertZero(L3Int(C1Grid.Columns(i).NumberFormat.Replace("N", "")))
                            End If
                        End If
                        range.EntireColumn.HorizontalAlignment = -4152 '= Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                    Case "Boolean", "Byte" ' Boolean, Byte là cột checkbox
                        range.HorizontalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Case "Integer", "Int32" 'Số nguyên
                        '  range.NumberFormat = "#,##0"
                        range.HorizontalAlignment = -4152 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                    Case "DateTime" 'Ngày
                        range.EntireColumn.NumberFormat = C1Grid.Columns(i).NumberFormat
                        range.EntireColumn.HorizontalAlignment = -4108  'Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    Case "String" 'Update 07/08/2015
                        range.EntireColumn.NumberFormat = "@" ' Dùng cho TH cột Ngày nhưng định dạng là dạng chuỗi
                        range.HorizontalAlignment = -4131
                    Case Else
                        range.HorizontalAlignment = -4131 'Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                End Select
                'Định dạng hiển thị dữ liệu cho cột
                colIndex = colIndex + 1
            Next
            If iX = 1 Then 'Header
                range = oWorkSheet.Rows(3, Type.Missing) 'TryCast(oWorkSheet.Rows(3, Type.Missing), Microsoft.Office.Interop.Excel.Range)
                range.Font.Bold = True
                range.EntireRow.VerticalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                range.EntireRow.HorizontalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                'Mau nen
                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
            End If
            PrevPos = EndValue 'Giữ vị trí cuối cùng của table trước đó
        Next
        ' D99C0008.Msg(5)
        'Hide row 1,2
        Dim rrow As Object = oWorkSheet.Range("A1", "A2")
        rrow.EntireRow.Hidden = True

    End Sub

    ''' <summary>
    ''' Xuất Excel từ lưới
    ''' </summary>
    ''' <param name="C1Grid">Lưới cần Xuất Excel</param>
    ''' <param name="sFileName"></param>
    ''' <param name="sFilter">Điều kiện filter do PSD quy định (VD: Chỉ xuất những dòng có Choose = 1)</param>
    ''' <param name="arrColAlwaysShow">Cột ẩn trên lưới nhưng cần hiện khi xuất ra file Excel</param>
    ''' <param name="arrColAlwaysHide">Cột hiện trên lưới nhưng cần ẩn khi xuất ra file Excel</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportToExcelFromGrid(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sFileName As String, Optional ByVal sFilter As String = "", Optional ByVal arrColAlwaysShow() As String = Nothing, Optional ByVal arrColAlwaysHide() As String = Nothing, Optional ByVal bExportAll As Boolean = False) As Boolean
        Try
            'Chỉ xuất những dòng phù hợp với sFilter truyền vào
            If sFileName = "" Then sFileName = FileNameExcel
            Dim dtG, dtSource As DataTable
            If bExportAll Then 'Chỉ xuất những dòng trên lưới (trong TH có Filter bar)
                dtSource = ReturnTableFilter(CType(C1Grid.DataSource, DataTable), sFilter, True)
            Else ' Xuất hết dữ liệu của dtGrid
                dtSource = ReturnTableFilter(CType(C1Grid.DataSource, DataTable).DefaultView.ToTable, sFilter)
            End If

            '    dtG.DefaultView.RowFilter = sFilter
            '=====================================================
            Dim oExcel As Object = CreateObject("Excel.Application") 'Microsoft.Office.Interop.Excel.Application 'Class() ' Create the Excel Application object
            'Update 04/07/2013 kiểm tra máy đang cài phiên bản Office nào


            If giVersion2007 = -1 Then
                CheckVersionExcel(oExcel)
            End If

            'Update 20/11/2015: Sửa theo Incident 82133: lưu file theo dạng Office đang tồn tại trên máy đó
            If giVersion2007 = 1 AndAlso sFileName.EndsWith(".xls") Then 'là Office2007: Chỉ Replate khi File truyền vào dạng .xls (TH bên ngoài truyền vào .xlsx thì không Replace
                sFileName = sFileName.ToLower.Replace(".xls", ".xlsx")
            ElseIf giVersion2007 <> 1 AndAlso sFileName.EndsWith(".xlsx") Then
                sFileName = sFileName.ToLower.Replace(".xlsx", ".xls") 'Truyền .xlsx thì file xuát là .xls
            End If


            'D99C0008.Msg(1)
            If giVersion2007 = -1 Then
                ' Kiểm tra nếu dữ liệu > 65530 dong hoặc >256 cột thì chỉ chạy trên Office 2007
                'If dtG.Rows.Count > MAX_ROW_EXCEL Then
                '    MessageBox.Show(ConvertUnicodeToVietwareF(rL3("So_dong_vuot_qua_gioi_han_cho_phep_cua_Excel") & " (" & dtG.Rows.Count & " > 65530)"), MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Error)
                '    oExcel.Quit()
                '    Return False
                'Else
                If C1Grid.Columns.Count > 256 Then
                    MessageBox.Show(ConvertUnicodeToVietwareF(rL3("So_cot_vuot_qua_gioi_han_cho_phep_cua_Excel") & " (" & C1Grid.Columns.Count & "> 256)"), MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    oExcel.Quit()
                    Return False
                End If
            End If

            '******************************************************
            'Kiểm tra tồn tại file Excel đang xuất
            If CloseProcessWindowMax(sFileName) = False Then oExcel.Quit() : Return False
            '******************************************************
            ' D99C0008.Msg(2)
            Dim oPathFile As String = gsApplicationPath + "\" + sFileName ' "Data.xlsx"
            Dim oWorkbook As Object = oExcel.Workbooks.Add(Type.Missing) ' Create a new Excel Workbook
            Dim oWorkSheet As Object ' Create a new Excel Worksheet
            iFirstCol = GetIntColumnExcel(sFirstCol) 'Đổi cột Chuỗi sang Số (VD: cột A đổi thành cột 0)

            'Dim StartValue As Integer = 0 ' Vị trí bắt đầu của Rang excel
            'Dim EndValue As Integer = 0 ' Vị trí cuối cùng của Rang excel
            'Dim PrevPos As Integer = 0 ' Vị trí kế tiếp của Rang excel

            'Dim iMaxRow As Long = dtG.Rows.Count
            'Dim iPackage As Integer = 0 'Số gói cần chạy 
            'Dim iLimitRow As Integer 'Số dòng chạy cho 1 gói iPackage
            'Dim iRowCount As Integer = 0 ' Tổng số dòng của 1 rang cần khởi tạo (rawData)
            'Dim iStartRow As Integer = 0 ' Dòng bắt đầu chạy cho dtData trong 1 gói
            'Dim iEndRow As Integer = 0 ' Dòng cuối cùng chạy cho dtData trong 1 gói
            Try
                MAX_ROW_EXCEL = oWorkbook.Worksheets(1).Rows.Count - 5 ' Tổng số dòng của Worksheet
                If dtSource.Rows.Count > MAX_ROW_EXCEL Then
                    dtG = dtSource.Clone()
                    Dim iSheetCount As Integer
                    iSheetCount = Math.Ceiling(dtSource.Rows.Count / MAX_ROW_EXCEL) 'Làm tròn lên
                    For i As Integer = 0 To iSheetCount - 1 ' Tổng số Sheets
                        dtG.Clear()
                        For j As Integer = 0 To MAX_ROW_EXCEL - 1
                            If i * MAX_ROW_EXCEL + j > dtSource.Rows.Count - 1 Then Exit For
                            dtG.ImportRow(dtSource.Rows(i * MAX_ROW_EXCEL + j))
                        Next
                        ExportExcelForWorksheet(C1Grid, dtG, oWorkbook, i + 1, arrColAlwaysShow, arrColAlwaysHide)
                    Next
                Else
                    dtG = dtSource.DefaultView.ToTable
                    ExportExcelForWorksheet(C1Grid, dtG, oWorkbook, 1, arrColAlwaysShow, arrColAlwaysHide)
                End If
                '' Tạo Sheet mới
                'oWorkSheet = oWorkbook.Worksheets(1) ' CType(oWorkbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                'oWorkSheet.Name = "Data"

                ''Xác định số dòng dữ liệu cần xuất cho 1 gói (iPackage)
                'If iMaxRow > 10000 Then
                '    iLimitRow = 1000
                'Else
                '    iLimitRow = 100
                'End If

                '' Tìm ký tự "A...Z" của cột
                'Dim finalColLetter As String = String.Empty
                'Dim colCharset As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                'Dim colCharsetLen As Integer = colCharset.Length
                'If C1Grid.Columns.Count > colCharsetLen Then
                '    finalColLetter = colCharset.Substring((C1Grid.Columns.Count - 1 + iFirstCol) \ colCharsetLen - 1, 1)
                'End If
                'finalColLetter += colCharset.Substring((C1Grid.Columns.Count - 1 + iFirstCol) Mod colCharsetLen, 1)

                ''Lấy số gói cần chạy để đưa vào rang excel
                'If iMaxRow <= iLimitRow Then ' Nếu Số dòng xuất <= iLimitRow thì 1 Package
                '    iPackage = 1
                'Else
                '    Dim iDecimal As Integer = CInt(iMaxRow Mod iLimitRow)
                '    If iDecimal = 0 Then ' Chia hết
                '        iPackage = CInt(iMaxRow / iLimitRow)
                '    Else ' Chia dư -> Package = Package + 1
                '        iPackage = CInt(Int(iMaxRow / iLimitRow))
                '        iPackage += 1
                '    End If
                'End If

                ''Tạo bảng và Add Data
                'For iX As Long = 1 To iPackage
                '    If iX < iPackage Then
                '        If iX = 1 Then 'Lần đầu tiên
                '            iStartRow = iEndRow
                '        Else
                '            iStartRow = iEndRow + 1
                '        End If
                '        iEndRow = CInt(iLimitRow * iX) - 1
                '        iRowCount = iLimitRow
                '    Else
                '        If iPackage = 1 Then
                '            iStartRow = 0
                '            iEndRow = CInt(iMaxRow - 1)
                '            iRowCount = CInt(iMaxRow)
                '        Else
                '            iStartRow = iEndRow + 1
                '            iEndRow = CInt(iMaxRow - 1)
                '            iRowCount = CInt(iMaxRow - (iLimitRow * (iX - 1)))
                '        End If
                '    End If
                '    ' D99C0008.Msg(3)
                '    StartValue = PrevPos + 1
                '    If iX = 1 Then ' Table đầu tiên dành vị trí A1 cho header
                '        EndValue = iRowCount + PrevPos + 3
                '    Else ' Các Table sau nối tiếp theo không tạo header
                '        EndValue = iRowCount + PrevPos
                '    End If

                '    'Tạo mảng dữ liệu để đưa vào file excel
                '    Dim arrData(iRowCount + 2, C1Grid.Columns.Count - 1) As Object
                '    Dim sColumnFieldName As String = ""
                '    Dim sColumnCaption As String = ""
                '    Dim iRow_Data As Integer = 0
                '    'Phải lấy dữ liệu của bảng dtCaptionCols để kiểm tra, vì bảng dtData có thứ tự cột không đúng như trên lưới
                '    For col As Integer = 0 To C1Grid.Columns.Count - 1
                '        sColumnFieldName = C1Grid.Columns(col).DataField
                '        If sColumnFieldName = "" Then Continue For
                '        sColumnCaption = C1Grid.Columns(col).Caption
                '        'If Not gbUnicode Then sColumnCaption = ConvertVniToUnicode(sColumnCaption)

                '        iRow_Data = 0
                '        If iX = 1 Then ' Đưa dữ liệu vào dòng tiêu đề (Header)
                '            arrData(0, col) = "'" & sColumnFieldName
                '            arrData(1, col) = C1Grid.Columns(col).NumberFormat
                '            arrData(2, col) = "'" & sColumnCaption
                '            iRow_Data += 3
                '        End If

                '        ' Đưa dữ liệu vào các dòng kế tiếp
                '        For row As Integer = iStartRow To iEndRow
                '            'Dim dr As DataRow = dtData.Rows(row)
                '            ''Nếu cột là chuỗi thì thêm dấu ' phía trước để khi xuất Excel thì hiểu giá trị là chuỗi
                '            If C1Grid.Columns(col).DataType.Name = "String" Then
                '                If gbUnicode Then
                '                    arrData(iRow_Data, col) = "'" & L3String(dtG.Rows(row).Item(C1Grid.Columns(col).DataField))
                '                Else 'Nếu nhập liệu VNI thì ConvertVniToUnicode dữ liệu dạng chuỗi sang Unicode
                '                    arrData(iRow_Data, col) = "'" & ConvertVniToUnicode(L3String(dtG.Rows(row).Item(C1Grid.Columns(col).DataField)))
                '                End If
                '            ElseIf C1Grid.Columns(col).DataType.Name = "Boolean" Then
                '                arrData(iRow_Data, col) = SQLNumber(dtG.Rows(row).Item(C1Grid.Columns(col).DataField))
                '            Else
                '                arrData(iRow_Data, col) = dtG.Rows(row).Item(C1Grid.Columns(col).DataField)
                '            End If
                '            iRow_Data += 1
                '        Next
                '    Next
                '    ' D99C0008.Msg(4)
                '    ' Fast data export to Excel
                '    Dim excelRange As String = String.Format(sFirstCol & "{0}:{1}{2}", StartValue, finalColLetter, EndValue)
                '    oWorkSheet.Range(excelRange, Type.Missing).Value2 = arrData
                '    'Khung
                '    oWorkSheet.Range(excelRange, Type.Missing).Borders.LineStyle = 1 ' Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                '    oWorkSheet.Range(excelRange, Type.Missing).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)

                '    'Định dạng các cột Excel
                '    Dim range As Object 'Microsoft.Office.Interop.Excel.Range
                '    'oWorkSheet.Columns.AutoFit()
                '    Dim colIndex As Integer = iFirstCol '0

                '    For i As Integer = 0 To C1Grid.Columns.Count - 1
                '        'Xác định vị trí vùng Range
                '        range = oWorkSheet.Range(GetStringColumnExcel(colIndex) & IIf(iX = 1, StartValue + 3, StartValue + 1).ToString, GetStringColumnExcel(colIndex) & EndValue)  'DirectCast(oWorkSheet.Range(GetStringColumnExcel(colIndex) & IIf(iX = 1, StartValue + 3, StartValue + 1).ToString, GetStringColumnExcel(colIndex) & EndValue), Microsoft.Office.Interop.Excel.Range)
                '        '=======================================================
                '        Dim bVisible As Boolean = False
                '        Dim dWidth As Integer = 0
                '        For iSplit As Integer = 0 To C1Grid.Splits.Count - 1
                '            If CheckContainsInArray(arrColAlwaysShow, C1Grid.Columns(i).DataField) Then ' Cột được hiện khi xuất Excel
                '                bVisible = True
                '                dWidth = C1Grid.Splits(iSplit).DisplayColumns(i).Width
                '            ElseIf CheckContainsInArray(arrColAlwaysHide, C1Grid.Columns(i).DataField) Then ' Cột được ẩn khi xuất Excel
                '                bVisible = False
                '            Else
                '                bVisible = C1Grid.Splits(iSplit).DisplayColumns(i).Visible
                '                dWidth = C1Grid.Splits(iSplit).DisplayColumns(i).Width
                '            End If
                '            If bVisible Then Exit For
                '        Next
                '        range.EntireColumn.ColumnWidth = dWidth * (1 / 6)
                '        range.EntireColumn.Hidden = Not bVisible
                '        '=======================================================
                '        Select Case C1Grid.Columns(i).DataType.Name
                '            Case "Decimal" 'Số thập phân
                '                If C1Grid.Columns(i).NumberFormat = "Percent" Or C1Grid.Columns(i).NumberFormat.Contains("%") Then
                '                    range.EntireColumn.NumberFormat = "0.00%"
                '                Else
                '                    If C1Grid.Columns(i).NumberFormat.Contains("#,##") Then
                '                        range.EntireColumn.NumberFormat = C1Grid.Columns(i).NumberFormat
                '                    Else
                '                        range.EntireColumn.NumberFormat = "#,##0" & InsertZero(L3Int(C1Grid.Columns(i).NumberFormat.Replace("N", "")))
                '                    End If
                '                End If
                '                range.EntireColumn.HorizontalAlignment = -4152 '= Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                '            Case "Boolean", "Byte" ' Boolean, Byte là cột checkbox
                '                range.HorizontalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                '            Case "Integer", "Int32" 'Số nguyên
                '                '  range.NumberFormat = "#,##0"
                '                range.HorizontalAlignment = -4152 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                '            Case "DateTime" 'Ngày
                '                range.EntireColumn.NumberFormat = C1Grid.Columns(i).NumberFormat
                '                range.EntireColumn.HorizontalAlignment = -4108  'Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                '            Case "String" 'Update 07/08/2015
                '                range.EntireColumn.NumberFormat = "@" ' Dùng cho TH cột Ngày nhưng định dạng là dạng chuỗi
                '                range.HorizontalAlignment = -4131
                '            Case Else
                '                range.HorizontalAlignment = -4131 'Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                '        End Select
                '        'Định dạng hiển thị dữ liệu cho cột
                '        colIndex = colIndex + 1
                '    Next
                '    If iX = 1 Then 'Header
                '        range = oWorkSheet.Rows(3, Type.Missing) 'TryCast(oWorkSheet.Rows(3, Type.Missing), Microsoft.Office.Interop.Excel.Range)
                '        range.Font.Bold = True
                '        range.EntireRow.VerticalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                '        range.EntireRow.HorizontalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                '        'Mau nen
                '        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
                '    End If
                '    PrevPos = EndValue 'Giữ vị trí cuối cùng của table trước đó
                'Next
                '' D99C0008.Msg(5)
                ''Hide row 1,2
                'Dim rrow As Object = oWorkSheet.Range("A1", "A2")
                'rrow.EntireRow.Hidden = True

                'Tắt cảnh báo hỏi có muốn Save As không?
                oExcel.DisplayAlerts = False
                'If dtSource.Rows.Count > MAX_ROW_EXCEL Then
                Try
                    oWorkSheet = oWorkbook.Worksheets(1)
                    oWorkSheet.Activate()
                    oWorkSheet.Select()
                Catch ex As Exception
                End Try

                ' End If

                '                If giVersion2007 = 1 Then
                '                    oWorkbook.SaveAs(oPathFile, FileFormat:=56)
                '                Else
                '                    oWorkbook.SaveAs(oPathFile)
                '                End If
                'Update 20/11/2015: Sửa theo Incident 82133: lưu file theo dạng Office đang tồn tại trên máy đó
                oWorkbook.SaveAs(oPathFile)

                oExcel.Workbooks.Open(oPathFile)
                oExcel.Visible = True
                Return True
            Catch ex As Exception
                D99C0008.MsgL3(ex.Message)
                oWorkbook.Close(False, Type.Missing, Type.Missing)
                ' Release the Application object
                oExcel.Quit()
            Finally
                oWorkSheet = Nothing
                oWorkbook = Nothing
                If oExcel IsNot Nothing Then oExcel = Nothing
                System.GC.Collect()
                System.GC.WaitForPendingFinalizers()
            End Try
            dtSource.DefaultView.RowFilter = ""
        Catch ex As Exception
            D99C0008.Msg(ex.Message)
        End Try
    End Function

    Public Function CheckContainsInArray(ByVal arr() As Object, ByVal sValue As Object) As Boolean
        If arr Is Nothing Then Return False

        Dim arrList As New List(Of Object)(arr)
        Return arrList.Contains(sValue)
    End Function

    Public Function CheckVersionExcel(ByVal appExcel As Object) As Boolean
        'Dim appExcel As New Microsoft.Office.Interop.Excel.Application
        ' appExcel = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        If L3Int(appExcel.Version) >= 12 Then
            giVersion2007 = 1
            Return True
        End If
        Return False
    End Function

    Public Function CloseProcessWindowMax(Optional ByVal FileName As String = FileNameExcel, Optional ByVal bShowMessage As Boolean = True) As Boolean
        'Doan code dung de dong file Excel mo san (khong phai do Chuong trinh mo)
        Dim p As System.Diagnostics.Process = Nothing
        Dim sWindowName As String = "Microsoft Excel - " & FileName
        Try
            For Each pr As Process In Process.GetProcessesByName("EXCEL")
                If sWindowName = pr.MainWindowTitle OrElse pr.MainWindowTitle = sWindowName.Substring(0, sWindowName.LastIndexOf(".")) Then
                    If p Is Nothing Then
                        p = pr
                    ElseIf p.StartTime < pr.StartTime Then
                        p = pr
                    End If
                End If
            Next
            If p IsNot Nothing Then
                If bShowMessage Then
                    If (D99C0008.MsgAsk(rL3("Ban_phai_dong_File") & Space(1) & FileName & Space(1) & rL3("truoc_khi_xuat_Excel") & "." & vbCrLf & rL3("Ban_co_muon_dong_khong")) = Windows.Forms.DialogResult.Yes) Then
                        p.Kill()
                        Return True
                    Else
                        Return False
                    End If
                Else
                    p.Kill()
                    Return True
                End If

            End If
            Return True 'False
        Catch ex As Exception
            Return True
        End Try
    End Function

    Private Function GetIntColumnExcel(ByVal sColumn As String) As Integer
        'Update 8/1/2014: Cách mới cho Office từ 2007 về sau
        Dim charColumn() As Char = sColumn.ToCharArray()
        Dim sum As Integer = 0
        For i As Integer = 0 To charColumn.Length - 1
            sum *= 26
            sum += (Asc(sColumn(i)) - Asc("A") + 1)
        Next
        Return sum - 1
    End Function

    Private Function GetStringColumnExcel(ByVal sColumn As Integer) As String
        Dim divNumber As Integer = sColumn + 1
        Dim columnName As String = ""
        Dim modNumber As Integer
        While divNumber > 0
            modNumber = (divNumber - 1) Mod 26
            columnName = Convert.ToChar(65 + modNumber).ToString() & columnName
            divNumber = CInt(((divNumber - modNumber) / 26))
        End While
        Return columnName
    End Function

    ''' <summary>
    ''' Import file Excel vào DataTable
    ''' </summary>
    ''' <param name="dtImport">DataTable trả ra</param>
    ''' <param name="sFilter">Điều kiện filter do PSD quy định (VD: Chỉ import những dòng có Choose = 1)</param>
    ''' <param name="sOutputPath">Trả ra đường dẫn file Excel khi import</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ImportExcelToGrid(ByRef dtImport As DataTable, Optional ByVal sFilter As String = "", Optional ByRef sOutputPath As String = "") As Boolean
        Return ImportExcelToGrid(Nothing, dtImport, sFilter, sOutputPath)
    End Function

    Public Function ImportExcelToGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dtImport As DataTable, Optional ByVal sFilter As String = "", Optional ByRef sOutputPath As String = "") As Boolean
        If Not System.IO.File.Exists(gsApplicationSetup & "\\Excel.dll") Then
            D99C0008.MsgL3("Không tồn tại DLL Excel.dll")
            Return False
        End If
        If Not System.IO.File.Exists(gsApplicationSetup & "\\ICSharpCode.SharpZipLib.dll") Then
            D99C0008.MsgL3("Không tồn tại DLL ICSharpCode.SharpZipLib.dll")
            Return False
        End If

        Return ExcuteImportExcelToGrid(tdbg, dtImport, sFilter, sOutputPath)
    End Function

    Private Function ExcuteImportExcelToGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dtImport As DataTable, ByVal sFilter As String, ByRef sOutputPath As String) As Boolean
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
                Dim ds As DataSet
                If file.FileName.Contains(".xlsx") Then
                    ds = excelReader.AsDataSet(True)
                Else
                    ds = excelReader.AsDataSet()
                End If
                excelReader.Close()
                Try
                    Dim dtTEm As DataTable = ds.Tables(0)
                    'Xóa những dòng tương ứng với sFilter
                    If sFilter <> "" Then
                        Dim dr() As DataRow = dtImport.Select(sFilter)
                        For i As Integer = dr.Length - 1 To 0 Step -1
                            dtImport.Rows.Remove(dr(i))
                        Next
                    Else
                        If dtImport.Rows.Count > 0 AndAlso D99C0008.MsgAsk(rL3("Ban_co_muon_xoa_tat_ca_du_lieu_hien_tai_tren_luoi_khong")) = DialogResult.Yes Then
                            If dtImport IsNot Nothing Then dtImport.Clear()
                        End If
                    End If
                    '=============================================
                    For k As Integer = 2 To dtTEm.Rows.Count - 1
                        dtImport.Rows.Add()
                        For i As Integer = 0 To dtImport.Columns.Count - 1
                            Try
                                If dtImport.Columns(dtImport.Columns(i).Caption).DataType.Name = "Boolean" Then
                                    dtImport.Rows(dtImport.Rows.Count - 1).Item(dtImport.Columns(i).Caption) = L3Bool(dtTEm.Rows(k).Item(dtImport.Columns(i).Caption))
                                ElseIf dtImport.Columns(dtImport.Columns(i).Caption).DataType.Name = "DateTime" Then
                                    If L3String(dtTEm.Rows(k).Item(dtImport.Columns(i).Caption)) = "" Then
                                        '   dtExport.Rows(dtExport.Rows.Count - 1).Item(dtExport.Columns(i).Caption) = Nothing
                                    Else
                                        If IsNumeric(dtTEm.Rows(k).Item(dtImport.Columns(i).Caption)) Then
                                            dtImport.Rows(dtImport.Rows.Count - 1).Item(dtImport.Columns(i).Caption) = DateTime.FromOADate(Number(dtTEm.Rows(k).Item(dtImport.Columns(i).Caption)))
                                        Else
                                            dtImport.Rows(dtImport.Rows.Count - 1).Item(dtImport.Columns(i).Caption) = dtTEm.Rows(k).Item(dtImport.Columns(i).Caption)
                                        End If
                                    End If
                                Else
                                    dtImport.Rows(dtImport.Rows.Count - 1).Item(dtImport.Columns(i).Caption) = dtTEm.Rows(k).Item(dtImport.Columns(i).Caption)
                                End If
                                If tdbg IsNot Nothing Then SetDefaultValue(tdbg, dtImport, dtImport.Rows.Count - 1, dtImport.Columns(i).Caption) 'Bổ sung set giá trị mặc định 10/08/2015

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

    'Những cột có DefaultValue thì gắn
    Private Sub SetDefaultValue(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dtImport As DataTable, ByVal rowImp As Integer, ByVal colImp As String)
        If L3String(tdbg.Columns(colImp).DefaultValue) = "" Then Exit Sub 'Không default giá trị

        Try
            If tdbg.Columns.IndexOf(tdbg.Columns(colImp)) >= 0 Then
                If IsDBNull(dtImport.Rows(rowImp).Item(colImp)) OrElse dtImport.Rows(rowImp).Item(colImp) Is Nothing Then 'chưa có dữ liệu
                    dtImport.Rows(rowImp).Item(colImp) = tdbg.Columns(colImp).DefaultValue
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "Export trực tiếp từ DataTable"
    Private Sub ExportExcelForWorksheet_FromDataTable(ByVal dtCaption As DataTable, ByVal dtG As DataTable, ByVal oWorkbook As Object, ByVal iSheet As Integer)
        Try
            Dim StartValue As Integer = 0 ' Vị trí bắt đầu của Rang excel
            Dim EndValue As Integer = 0 ' Vị trí cuối cùng của Rang excel
            Dim PrevPos As Integer = 0 ' Vị trí kế tiếp của Rang excel

            Dim iMaxRow As Long = dtG.Rows.Count
            Dim iPackage As Integer = 0 'Số gói cần chạy 
            Dim iLimitRow As Integer 'Số dòng chạy cho 1 gói iPackage
            Dim iRowCount As Integer = 0 ' Tổng số dòng của 1 rang cần khởi tạo (rawData)
            Dim iStartRow As Integer = 0 ' Dòng bắt đầu chạy cho dtData trong 1 gói
            Dim iEndRow As Integer = 0 ' Dòng cuối cùng chạy cho dtData trong 1 gói

            Dim oWorkSheet As Object
            ' Tạo Sheet mới
            If iSheet < 4 Then
                oWorkSheet = oWorkbook.Worksheets(iSheet) ' CType(oWorkbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)        oWorkSheet.Name = "Data"
            Else
                oWorkSheet = oWorkbook.Worksheets.Add(Type.Missing, oWorkbook.Worksheets(iSheet - 1), 1, Type.Missing)
            End If
            If iSheet = 1 Then
                oWorkSheet.Name = "Data"
            Else
                oWorkSheet.Name = "Data " & (iSheet - 1).ToString("00")
            End If

            'Xác định số dòng dữ liệu cần xuất cho 1 gói (iPackage)
            If iMaxRow > 10000 Then
                iLimitRow = 1000
            Else
                iLimitRow = 100
            End If

            ' Tìm ký tự "A...Z" của cột
            Dim finalColLetter As String = String.Empty
            Dim colCharset As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Dim colCharsetLen As Integer = colCharset.Length
            If dtG.Columns.Count > colCharsetLen Then
                finalColLetter = colCharset.Substring((dtG.Columns.Count - 1 + iFirstCol) \ colCharsetLen - 1, 1)
            End If
            finalColLetter += colCharset.Substring((dtG.Columns.Count - 1 + iFirstCol) Mod colCharsetLen, 1)

            'Lấy số gói cần chạy để đưa vào rang excel
            If iMaxRow <= iLimitRow Then ' Nếu Số dòng xuất <= iLimitRow thì 1 Package
                iPackage = 1
            Else
                Dim iDecimal As Integer = CInt(iMaxRow Mod iLimitRow)
                If iDecimal = 0 Then ' Chia hết
                    iPackage = CInt(iMaxRow / iLimitRow)
                Else ' Chia dư -> Package = Package + 1
                    iPackage = CInt(Int(iMaxRow / iLimitRow))
                    iPackage += 1
                End If
            End If
            '******************************
            Dim dr() As DataRow
            Dim sColumnFieldName As String = ""

            'Tạo bảng và Add Data
            For iX As Long = 1 To iPackage
                If iX < iPackage Then
                    If iX = 1 Then 'Lần đầu tiên
                        iStartRow = iEndRow
                    Else
                        iStartRow = iEndRow + 1
                    End If
                    iEndRow = CInt(iLimitRow * iX) - 1
                    iRowCount = iLimitRow
                Else
                    If iPackage = 1 Then
                        iStartRow = 0
                        iEndRow = CInt(iMaxRow - 1)
                        iRowCount = CInt(iMaxRow)
                    Else
                        iStartRow = iEndRow + 1
                        iEndRow = CInt(iMaxRow - 1)
                        iRowCount = CInt(iMaxRow - (iLimitRow * (iX - 1)))
                    End If
                End If
                StartValue = PrevPos + 1
                If iX = 1 Then ' Table đầu tiên dành vị trí A1 cho header
                    EndValue = iRowCount + PrevPos + 3
                Else ' Các Table sau nối tiếp theo không tạo header
                    EndValue = iRowCount + PrevPos
                End If

                'Tạo mảng dữ liệu để đưa vào file excel
                Dim arrData(iRowCount + 2, dtG.Columns.Count - 1) As Object
                Dim sColumnCaption As String = ""
                Dim iRow_Data As Integer = 0

                'Phải lấy dữ liệu của bảng dtCaptionCols để kiểm tra, vì bảng dtData có thứ tự cột không đúng như trên lưới
                For col As Integer = 0 To dtG.Columns.Count - 1
                    sColumnFieldName = dtG.Columns(col).ColumnName
                    If sColumnFieldName = "" Then Continue For
                    '**************************
                    dr = dtCaption.Select("FieldName = " & SQLString(sColumnFieldName))
                    If dr.Length <= 0 Then Continue For
                    sColumnCaption = dr(0).Item("Caption").ToString
                    '**************************
                    iRow_Data = 0
                    If iX = 1 Then ' Đưa dữ liệu vào dòng tiêu đề (Header)
                        arrData(0, col) = "'" & sColumnFieldName
                        arrData(1, col) = dr(0).Item("DataFormat").ToString 'C1Grid.Columns(col).NumberFormat
                        arrData(2, col) = "'" & sColumnCaption
                        iRow_Data += 3
                    End If

                    ' Đưa dữ liệu vào các dòng kế tiếp
                    For row As Integer = iStartRow To iEndRow
                        'Nếu cột là chuỗi thì thêm dấu ' phía trước để khi xuất Excel thì hiểu giá trị là chuỗi
                        If dtG.Columns(col).DataType.Name = "String" Then
                            If gbUnicode Then
                                arrData(iRow_Data, col) = "'" & L3String(dtG.Rows(row).Item(col))
                            Else 'Nếu nhập liệu VNI thì ConvertVniToUnicode dữ liệu dạng chuỗi sang Unicode
                                arrData(iRow_Data, col) = "'" & ConvertVniToUnicode(L3String(dtG.Rows(row).Item(col)))
                            End If
                        ElseIf dtG.Columns(col).DataType.Name = "Boolean" Then
                            arrData(iRow_Data, col) = SQLNumber(dtG.Rows(row).Item(col))
                        Else
                            arrData(iRow_Data, col) = dtG.Rows(row).Item(col)
                        End If
                        iRow_Data += 1
                    Next
                Next
                ' Fast data export to Excel
                Dim excelRange As String = String.Format(sFirstCol & "{0}:{1}{2}", StartValue, finalColLetter, EndValue)
                oWorkSheet.Range(excelRange, Type.Missing).Value2 = arrData
                'Khung
                oWorkSheet.Range(excelRange, Type.Missing).Borders.LineStyle = 1 ' Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                oWorkSheet.Range(excelRange, Type.Missing).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)

                'Định dạng các cột Excel
                Dim range As Object 'Microsoft.Office.Interop.Excel.Range
                Dim colIndex As Integer = iFirstCol '0

                For i As Integer = 0 To dtG.Columns.Count - 1
                    'Xác định vị trí vùng Range
                    range = oWorkSheet.Range(GetStringColumnExcel(colIndex) & IIf(iX = 1, StartValue + 3, StartValue + 1).ToString, GetStringColumnExcel(colIndex) & EndValue)  'DirectCast(oWorkSheet.Range(GetStringColumnExcel(colIndex) & IIf(iX = 1, StartValue + 3, StartValue + 1).ToString, GetStringColumnExcel(colIndex) & EndValue), Microsoft.Office.Interop.Excel.Range)
                    '=======================================================
                    Dim bVisible As Boolean = False
                    Dim dWidth As Integer = 0
                    '**************************
                    sColumnFieldName = dtG.Columns(i).ColumnName
                    If sColumnFieldName = "" Then Continue For
                    dr = dtCaption.Select("FieldName = " & SQLString(sColumnFieldName))
                    If dr.Length <= 0 Then Continue For
                    '**************************
                    bVisible = L3Bool(dr(0).Item("Visible"))
                    dWidth = L3Int(dr(0).Item("Width"))

                    range.EntireColumn.ColumnWidth = dWidth * (1 / 4)
                    range.EntireColumn.Hidden = Not bVisible
                    '=======================================================
                    Select Case dtG.Columns(i).DataType.Name
                        Case "Decimal" 'Số thập phân
                            If dr(0).Item("DataFormat").ToString = "Percent" Or dr(0).Item("DataFormat").ToString.Contains("%") Then
                                range.EntireColumn.NumberFormat = "0.00%"
                            Else
                                If dr(0).Item("DataFormat").ToString.Contains("#,##") Then
                                    range.EntireColumn.NumberFormat = dr(0).Item("DataFormat").ToString
                                Else
                                    range.EntireColumn.NumberFormat = "#,##0" & InsertZero(L3Int(dr(0).Item("DataFormat").ToString.Replace("N", "")))
                                End If
                            End If
                            range.EntireColumn.HorizontalAlignment = -4152
                        Case "Boolean", "Byte" ' Boolean, Byte là cột checkbox
                            range.HorizontalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        Case "Integer", "Int32" 'Số nguyên
                            range.HorizontalAlignment = -4152 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                        Case "DateTime" 'Ngày
                            range.EntireColumn.NumberFormat = dr(0).Item("DataFormat").ToString
                            range.EntireColumn.HorizontalAlignment = -4108  'Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        Case "String" 'Update 07/08/2015
                            range.EntireColumn.NumberFormat = "@" ' Dùng cho TH cột Ngày nhưng định dạng là dạng chuỗi
                            range.HorizontalAlignment = -4131
                        Case Else
                            range.HorizontalAlignment = -4131 'Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    End Select
                    'Định dạng hiển thị dữ liệu cho cột
                    colIndex = colIndex + 1
                Next
                If iX = 1 Then 'Header
                    range = oWorkSheet.Rows(3, Type.Missing) 'TryCast(oWorkSheet.Rows(3, Type.Missing), Microsoft.Office.Interop.Excel.Range)
                    range.Font.Bold = True
                    range.EntireRow.VerticalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    range.EntireRow.HorizontalAlignment = -4108 ' Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    'Mau nen
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
                End If
                PrevPos = EndValue 'Giữ vị trí cuối cùng của table trước đó
            Next

            'Hide row 1,2
            Dim rrow As Object = oWorkSheet.Range("A1", "A2")
            rrow.EntireRow.Hidden = True
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Xuất Excel từ Data Table
    ''' </summary>
    ''' <param name="dtCaption">Table chứa caption</param>
    ''' <param name="sFileName"></param>
    ''' <param name="sFilter">Điều kiện filter do PSD quy định (VD: Chỉ xuất những dòng có Choose = 1)</param>
    ''' <remarks></remarks>
    Public Function ExportToExcelFromDataTable(ByVal dtCaption As DataTable, ByVal dtSource As DataTable, ByVal sFileName As String, Optional ByVal sFilter As String = "") As Boolean
        Try
            'Chỉ xuất những dòng phù hợp với sFilter truyền vào
            If sFileName = "" Then sFileName = FileNameExcel
            If sFilter <> "" Then dtSource = ReturnTableFilter(dtSource, sFilter)
            '=====================================================
            Dim oExcel As Object = CreateObject("Excel.Application")  ' Create the Excel Application object
            'Update 04/07/2013 kiểm tra máy đang cài phiên bản Office nào
            If giVersion2007 = -1 Then CheckVersionExcel(oExcel)

            'Update 20/11/2015: Sửa theo Incident 82133: lưu file theo dạng Office đang tồn tại trên máy đó
            If giVersion2007 = 1 AndAlso sFileName.EndsWith(".xls") Then 'là Office2007: Chỉ Replate khi File truyền vào dạng .xls (TH bên ngoài truyền vào .xlsx thì không Replace
                sFileName = sFileName.ToLower.Replace(".xls", ".xlsx")
            ElseIf giVersion2007 <> 1 AndAlso sFileName.EndsWith(".xlsx") Then
                sFileName = sFileName.ToLower.Replace(".xlsx", ".xls") 'Truyền .xlsx thì file xuát là .xls
            End If

            If giVersion2007 = -1 Then
                ' Kiểm tra nếu dữ liệu > 65530 dong hoặc >256 cột thì chỉ chạy trên Office 2007
                If dtSource.Columns.Count > 256 Then
                    MessageBox.Show(ConvertUnicodeToVietwareF(rL3("So_cot_vuot_qua_gioi_han_cho_phep_cua_Excel") & " (" & dtSource.Columns.Count & "> 256)"), MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    oExcel.Quit()
                    Return False
                End If
            End If
            '******************************************************
            'Kiểm tra tồn tại file Excel đang xuất
            If CloseProcessWindowMax(sFileName) = False Then oExcel.Quit() : Return False
            '******************************************************
            Dim dtG As DataTable
            Dim oPathFile As String = gsApplicationPath + "\" + sFileName ' "Data.xlsx"
            Dim oWorkbook As Object = oExcel.Workbooks.Add(Type.Missing) ' Create a new Excel Workbook
            Dim oWorkSheet As Object ' Create a new Excel Worksheet
            iFirstCol = GetIntColumnExcel(sFirstCol) 'Đổi cột Chuỗi sang Số (VD: cột A đổi thành cột 0)

            Try
                MAX_ROW_EXCEL = oWorkbook.Worksheets(1).Rows.Count - 5 ' Tổng số dòng của Worksheet
                If dtSource.Rows.Count > MAX_ROW_EXCEL Then
                    dtG = dtSource.Clone()
                    Dim iSheetCount As Integer
                    iSheetCount = Math.Ceiling(dtSource.Rows.Count / MAX_ROW_EXCEL) 'Làm tròn lên
                    For i As Integer = 0 To iSheetCount - 1 ' Tổng số Sheets
                        dtG.Clear()
                        For j As Integer = 0 To MAX_ROW_EXCEL - 1
                            If i * MAX_ROW_EXCEL + j > dtSource.Rows.Count - 1 Then Exit For
                            dtG.ImportRow(dtSource.Rows(i * MAX_ROW_EXCEL + j))
                        Next
                        ExportExcelForWorksheet_FromDataTable(dtCaption, dtG, oWorkbook, i + 1)
                    Next
                Else
                    dtG = dtSource.DefaultView.ToTable
                    ExportExcelForWorksheet_FromDataTable(dtCaption, dtG, oWorkbook, 1)
                End If

                'Tắt cảnh báo hỏi có muốn Save As không?
                oExcel.DisplayAlerts = False
                'If dtSource.Rows.Count > MAX_ROW_EXCEL Then
                Try
                    oWorkSheet = oWorkbook.Worksheets(1)
                    oWorkSheet.Activate()
                    oWorkSheet.Select()
                Catch ex As Exception
                End Try

                'Update 20/11/2015: Sửa theo Incident 82133: lưu file theo dạng Office đang tồn tại trên máy đó
                oWorkbook.SaveAs(oPathFile)

                oExcel.Workbooks.Open(oPathFile)
                oExcel.Visible = True
                Return True
            Catch ex As Exception
                D99C0008.MsgL3(ex.Message)
                oWorkbook.Close(False, Type.Missing, Type.Missing)
                ' Release the Application object
                oExcel.Quit()
            Finally
                oWorkSheet = Nothing
                oWorkbook = Nothing
                If oExcel IsNot Nothing Then oExcel = Nothing
                System.GC.Collect()
                System.GC.WaitForPendingFinalizers()
            End Try
            dtSource.DefaultView.RowFilter = ""
        Catch ex As Exception
            D99C0008.Msg(ex.Message)
        End Try
    End Function
#End Region

End Module

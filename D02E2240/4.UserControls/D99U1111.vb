Imports System
Imports System.Text
Public Class D99U1111


#Region "Const of tdbg"
    Private Const COL_IsUsed As String = "IsUsed"               ' Chọn
    Private Const COL_Description As String = "Description"     ' Tên cột
    Private Const COL_FieldName As String = "FieldName"         ' FieldName
    Private Const COL_Obligatory As String = "Obligatory"       ' Obligatory
    Private Const COL_GridName As String = "GridName"           ' GridName
    Private Const COL_IsUpdate As String = "IsUpdate"           ' IsUpdate
    Private Const COL_SplitIndex As String = "SplitIndex"       ' SplitIndex
    Private Const COL_SplitSizeMode As String = "SplitSizeMode" ' SplitSizeMode
    Private Const COL_SplitSize As String = "SplitSize"         ' SplitSize
    Private Const COL_Mode As String = "Mode"                   ' Mode
    Private Const COL_Button As String = "Button"                   ' Button
#End Region

    Dim _moduleID As String
    Dim _mode As Object
    Dim _formID As Control
    Dim _dtGrid As DataTable 'Load lưới
    Dim _tdbgO As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Dim _Button As Object = -1


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Public Sub New(ByVal FormID As Control, ByVal tdbgO As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dtCaption As DataTable, _
      Optional ByVal iMode As Object = 0, Optional ByVal bStatusSave As Boolean = True, _
      Optional ByVal bUseTemplateInquiry As Boolean = False, Optional ByVal iButton As Object = -1)

        ' This call is required by the Windows Form Designer. 
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call. 
        _formID = FormID
        _moduleID = FormID.Name.Substring(1, 2)
        _mode = iMode
        _dtGrid = dtCaption
        _tdbgO = tdbgO
        _Button = iButton

        LoadTDBGrid(iMode, bUseTemplateInquiry)
        'Hiển thị nút Save
        btnSave.Enabled = bStatusSave
        '*************

        LoadLanguage()
        SetResolutionForm(Me)
    End Sub

    'Khi thay đổi Mode thì gọi hàm này
    Private Sub LoadTDBGrid(ByVal iMode As Object, ByVal bUseTemplateInquiry As Boolean)
        Dim dtD91T2022 As New DataTable
        'Nếu đã sử dụng Mẫu template thì không lấy dữ liệu trong bảng D91T2022
        If Not bUseTemplateInquiry Then
            Dim sSQL As String = ""
            sSQL = "Select Distinct FieldName" & UnicodeJoin(gbUnicode) & " as FieldName, cast(0 as bit) As IsUsed, Mode, 0 as IsUpdate" & vbCrLf ' 1 as IsUpdate,
            sSQL &= "From D91T2022 WITH(NOLOCK) " & vbCrLf
            sSQL &= "WHERE ModuleID = " & SQLString(_moduleID) & " And FormID = " & SQLString(_formID.Name)
            sSQL &= " And FieldName" & UnicodeJoin(gbUnicode) & " <>'' And (Mode = '-1' Or Mode = " & SQLString(_mode) & ")"
            sSQL &= " And UserID = " & SQLString(gsUserID)
            dtD91T2022 = ReturnDataTable(sSQL)
            'If dtD91T2022.Rows.Count > 0 Then
            '    dtD91T2022.PrimaryKey = New DataColumn() {dtD91T2022.Columns("FieldName"), dtD91T2022.Columns("Mode")}
            '    _dtGrid.Merge(dtD91T2022, False, MissingSchemaAction.AddWithKey)
            '    'Xóa dòng không tồn tại với hiện tại
            '    _dtGrid.DefaultView.RowFilter = COL_Description & " <>''"
            '    _dtGrid = _dtGrid.DefaultView.ToTable
            '    '***************
            'End If
        End If

        '*******************************************
        _dtGrid.DefaultView.RowFilter = COL_Mode & "= '-1' Or " & COL_Mode & "=" & SQLString(iMode.ToString)
        For i As Integer = 0 To _dtGrid.DefaultView.Count - 1 'Set lại IsUsed
            Dim dr() As DataRow = dtD91T2022.Select("FieldName=" & SQLString(_dtGrid.DefaultView(i).Item(COL_FieldName)))
            If dr.Length = 0 Then Continue For
            _dtGrid.DefaultView(i).Item(COL_IsUsed) = dr(0).Item(COL_IsUsed)
            _dtGrid.DefaultView(i).Item(COL_IsUpdate) = dr(0).Item(COL_IsUpdate)
        Next
        _dtGrid.AcceptChanges()
        '*******************
        _dtGrid.DefaultView.RowFilter = COL_Mode & "= '-1' Or " & COL_Mode & "=" & SQLString(iMode.ToString)
        LoadDataSource(tdbgD99U1111, _dtGrid, True)
        '*******************************************
        'Co dãn cột Description -> Lưới -> UserControl, set vị trí các nút 
        Dim iOldWidthDes As Integer = tdbgD99U1111.Splits(0).DisplayColumns(COL_Description).Width
        tdbgD99U1111.Splits(0).DisplayColumns(COL_Description).AutoSize()
        'Nếu cột Diễn giải có width > mặc định thì mới gắn lại Width
        If tdbgD99U1111.Splits(0).DisplayColumns(COL_Description).Width > iOldWidthDes Then
            Dim iDisFormGrid As Integer = Me.Size.Width - tdbgD99U1111.Width
            tdbgD99U1111.Width = tdbgD99U1111.Splits(0).DisplayColumns(COL_IsUsed).Width + tdbgD99U1111.Splits(0).DisplayColumns(COL_Description).Width + tdbgD99U1111.Splits(0).RecordSelectorWidth * 3 '60
            Me.Size = New Size(tdbgD99U1111.Width + iDisFormGrid, Me.Height)
        End If
        'If dtD91T2022.Rows.Count > 0 Then
        btnRefresh_Click(Nothing, Nothing)
        dtD91T2022.Dispose()
    End Sub

    Private Function CreatTableF12() As DataTable
        Dim dtCaption As New DataTable

        For i As Integer = 0 To tdbgD99U1111.Columns.Count - 1
            dtCaption.Columns.Add(tdbgD99U1111.Columns(i).DataField)
        Next
        dtCaption.Columns(COL_IsUsed).DataType = Type.GetType("System.Boolean")
        dtCaption.Columns(COL_IsUpdate).DataType = Type.GetType("System.Int32")
        Return dtCaption
    End Function

    Public Sub InsertColVisible(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iSplit As Integer, ByRef dtCaption As DataTable, ByVal pos As Integer, _
                     ByVal iMode As Object, ByVal iObl As Integer, ByVal iCol As Integer, Optional ByVal iButton As Integer = -1)
        Dim dr As DataRow
        dr = dtCaption.NewRow
        dr = CreateDataRow(c1Grid, iSplit, dr, iMode, iObl, iCol, iButton)
        dtCaption.Rows.InsertAt(dr, pos)
    End Sub

    Public Sub InsertColVisible(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iSplit As Integer, ByRef dtCaption As DataTable, ByVal pos As Integer, _
          Optional ByVal iMode As Object = 0, Optional ByVal ArrColObl() As Object = Nothing, _
          Optional ByVal oColBegin As Object = 0, Optional ByVal oColEnd As Object = -1, Optional ByVal bCheckVis As Boolean = True, Optional ByVal iButton As Integer = -1)
        Dim iColBegin As Integer
        Dim iColEnd As Integer
        'Lấy đúng kiểu cột
        CastObjToIntCol(c1Grid, oColBegin, iColBegin)
        CastObjToIntCol(c1Grid, oColEnd, iColEnd)
        '********************
        If iColEnd = -1 Then iColEnd = c1Grid.Columns.Count - 1
        For i As Integer = iColBegin To iColEnd
            If c1Grid.Splits(iSplit).DisplayColumns(i).Visible = False AndAlso bCheckVis Then Continue For
            InsertColVisible(c1Grid, iSplit, dtCaption, pos, iMode, GetObl(c1Grid, i, ArrColObl), i, iButton)
            pos += 1
        Next i
    End Sub

    Private Sub AddColVisible(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iSplit As Integer, ByRef dtCaption As DataTable, _
                        ByVal iMode As Object, ByVal iObl As Integer, ByVal iCol As Integer, Optional ByVal iButton As Integer = -1)
        Dim dr As DataRow
        dr = dtCaption.NewRow
        dr = CreateDataRow(c1Grid, iSplit, dr, iMode, iObl, iCol, iButton)
        dtCaption.Rows.Add(dr)
    End Sub

    Private Function CreateDataRow(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iSplit As Integer, ByVal dr As DataRow, _
                    ByVal iMode As Object, ByVal iObl As Integer, ByVal iCol As Integer, Optional ByVal iButton As Integer = -1) As DataRow
        Dim dc As C1.Win.C1TrueDBGrid.C1DataColumn = c1Grid.Columns(iCol)
        ' Dim dr As DataRow
        '  dr = dtCaption.NewRow
        dr.Item(COL_IsUsed) = 1
        dr.Item(COL_FieldName) = dc.DataField
        '**************
        Dim sDes As String = dc.Caption
        If c1Grid.Splits(iSplit).DisplayColumns(iCol).HeadingStyle.Font.Name.Contains("Lemon3") Then sDes = ConvertVniToUnicode(sDes)
        dr.Item(COL_Description) = sDes
        '**************
        dr.Item(COL_Obligatory) = iObl
        '**************
        dr.Item(COL_Mode) = iMode
        dr.Item(COL_GridName) = c1Grid.Name
        dr.Item(COL_IsUpdate) = 0
        dr.Item(COL_SplitIndex) = iSplit
        dr.Item(COL_SplitSize) = c1Grid.Splits(iSplit).SplitSize
        dr.Item(COL_SplitSizeMode) = CType(c1Grid.Splits(iSplit).SplitSizeMode, Integer)
        dr.Item(COL_Button) = iButton
        '   dtCaption.Rows.Add(dr)
        Return dr
    End Function


    Private Function CastObjToIntCol(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal oColumn As Object, ByRef iCol As Integer) As Boolean
        If oColumn Is Nothing Then Return False
        If IsNumeric(oColumn) Then
            iCol = L3Int(oColumn)
        Else
            iCol = tdbg.Columns.IndexOf(tdbg.Columns(oColumn.ToString))
        End If
        Return True
    End Function

    Private Function GetObl(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iCol As Integer, ByVal ArrColObl() As Object) As Integer
        Dim iObl As Integer = 0
        If ArrColObl IsNot Nothing AndAlso ArrColObl.Length > 0 Then
            If IsNumeric(ArrColObl(0)) Then 'Là Integer
                If L3FindArr(ArrColObl, iCol) Then iObl = 1
            Else 'Cột kiểu chuỗi
                If L3FindArr(ArrColObl, c1Grid.Columns(iCol).DataField) Then iObl = 1
            End If
        End If
        Return iObl
    End Function

    Public Sub EditModeColVisible(ByRef dtCaption As DataTable, _
                      ByVal iMode As Object, ByVal ParamArray sCol() As String)

        For i As Integer = 0 To sCol.Length - 1
            Dim dr() As DataRow = dtCaption.Select(COL_FieldName & "=" & SQLString(sCol(i)))
            If dr.Length = 0 Then Continue For
            dr(0).Item(COL_Mode) = iMode
        Next
    End Sub

    Public Sub EditModeColVisible(ByRef dtCaption As DataTable, _
                    ByVal iMode As Object, ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal oColBegin As Object, ByVal oColEnd As Object)
        Try
            Dim iColBegin As Integer
            Dim iColEnd As Integer
            'Lấy đúng kiểu cột
            CastObjToIntCol(c1Grid, oColBegin, iColBegin)
            CastObjToIntCol(c1Grid, oColEnd, iColEnd)
            '********************
            For i As Integer = iColBegin To iColEnd
                Dim dr() As DataRow = dtCaption.Select(COL_FieldName & "=" & SQLString(c1Grid.Columns(i).DataField))
                If dr.Length = 0 Then Continue For
                dr(0).Item(COL_Mode) = iMode
            Next
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try
    End Sub

    'ArrColObl: có thể String or Integer
    ''' <summary>
    ''' Thêm cột đang hiển thị trên lưới từ đến trong split
    ''' </summary>
    ''' <param name="c1Grid"></param>
    ''' <param name="iSplit">split</param>
    ''' <param name="dtCaption"></param>
    ''' <param name="iMode">Mode hiển thị. Mặc định chung là 0</param>
    ''' <param name="ArrColObl">DS bắt buộc hiển thị</param>
    ''' <param name="oColBegin">Cột bắt đầu duyệt trong split</param>
    ''' <param name="oColEnd">Cột cuối cùng duyệt trong split</param>
    ''' <param name="bCheckVis">Có kiểm tra Visible thì mới thêm vào F12 không? Mặc định là kiểm tra</param>
    ''' <param name="iSplitIndex">TH có RemoveSplit (có Check Chi tiết). Split thực sự hiển khi chkIsDetail = False</param>
    ''' <remarks></remarks>
    Public Sub AddColVisible(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iSplit As Integer, ByRef dtCaption As DataTable, _
             Optional ByVal iMode As Object = 0, Optional ByVal ArrColObl() As Object = Nothing, _
             Optional ByVal oColBegin As Object = 0, Optional ByVal oColEnd As Object = -1, Optional ByVal bCheckVis As Boolean = True, Optional ByVal iButton As Integer = -1)
        Try
            If dtCaption Is Nothing Then dtCaption = CreatTableF12() 'Tạo bảng
            Dim iColBegin As Integer
            Dim iColEnd As Integer
            'Lấy đúng kiểu cột
            CastObjToIntCol(c1Grid, oColBegin, iColBegin)
            CastObjToIntCol(c1Grid, oColEnd, iColEnd)
            '********************
            If iColEnd = -1 Then iColEnd = c1Grid.Columns.Count - 1
            For i As Integer = iColBegin To iColEnd
                If c1Grid.Splits(iSplit).DisplayColumns(i).Visible = False AndAlso bCheckVis Then Continue For
                AddColVisible(c1Grid, iSplit, dtCaption, iMode, GetObl(c1Grid, i, ArrColObl), i, iButton)
            Next i
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Thêm cột đang hiển thị trên lưới
    ''' </summary>
    ''' <param name="c1Grid"></param>
    ''' <param name="dtCaption">table hiển thị F12</param>
    ''' <param name="ArrColObl">DS cột bắt buộc hiển thị</param>
    ''' <param name="iMode">Mode hiển thị. Mặc định chung là 0</param>
    ''' <remarks></remarks>
    Public Sub AddColVisible(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dtCaption As DataTable, _
         Optional ByVal ArrColObl() As Object = Nothing, Optional ByVal iMode As Object = 0, Optional ByVal iButton As Integer = -1)
        For i As Integer = 0 To c1Grid.Splits.Count - 1
            AddColVisible(c1Grid, i, dtCaption, iMode, ArrColObl, , , , iButton)
        Next
    End Sub

    Private Sub D99U1111_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'cho Focus vào lưới (vì sự kiện VisibleChanged k activate ở lần đầu usrctrl show)
        'FOCUS VÀO LƯỚI MỖI KHI D09U1111 VISIBLE
        D99U1111_VisibleChanged(Nothing, Nothing)
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        If _dtGrid Is Nothing Then Exit Sub
        Dim i As Integer
        Try
            If Not AllowRefresh() Then Exit Sub
            Dim dr() As DataRow = _dtGrid.Select("(Mode ='" & _mode.ToString & "' Or Mode ='-1' ) And ( Button ='" & _Button.ToString & "' Or Button ='-1')")
            'Ẩn cột theo UserControl
            For i = 0 To dr.Length - 1
                Dim iSplit As Integer = L3Int(dr(i).Item(COL_SplitIndex))
                ' If _Button > -1 AndAlso L3Int(dr(i).Item(COL_Button)) <> _Button Then Continue For 'Nếu sử dụng nút
                If iSplit < _tdbgO.Splits.Count Then _tdbgO.Splits(iSplit).DisplayColumns(dr(i).Item(COL_FieldName).ToString).Visible = L3Bool(dr(i).Item(COL_IsUsed))
                dr(i).Item(COL_IsUpdate) = 0
            Next i
            'CHỈNH LẠI SPLITSIZE CỦA LƯỚI, XỬ LÝ CHO 2 TH

            For iSplitIndex As Integer = 0 To _tdbgO.Splits.Count - 1 'Split
                ResetSplitSize(iSplitIndex)
            Next
            '*********************
        Catch ex As Exception
            MessageBox.Show(i.ToString & " - FieldName: " & " - Source = " & ex.Source & " - Message = " & ex.Message & " - ToString = " & ex.ToString)
        End Try
    End Sub

    Private Sub ResetSplitSize(ByVal iSplitIndex As Integer)
        Dim iTotalSplitSize As Integer = 0
        Dim iFirstCol As Integer = -1
        With _tdbgO

            Dim bSplitDisable As Boolean = False

            For i As Integer = 0 To .Splits.ColCount - 1
                Dim bVisible As Boolean = False
                For j As Integer = 0 To .Columns.Count - 1
                    If .Splits(i).DisplayColumns(j).Visible Then
                        If iFirstCol = -1 Then iFirstCol = j
                        bVisible = True
                        Exit For
                    End If
                Next j

                If bVisible Then
                    Dim splitSize As Integer = 0
                    Dim splitSizeMode As Integer = 0
                    GetInfoSplit(i, splitSize, splitSizeMode)
                    If splitSize > 0 Then
                        .Splits(i).SplitSizeMode = CType(splitSizeMode, C1.Win.C1TrueDBGrid.SizeModeEnum)
                        .Splits(i).SplitSize = splitSize
                        '.Splits(i).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.Always
                    End If
                Else
                    .Splits(i).SplitSizeMode = C1.Win.C1TrueDBGrid.SizeModeEnum.Exact
                    .Splits(i).SplitSize = 0
                    .Splits(i).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.None
                    bSplitDisable = True
                End If

                iTotalSplitSize += .Splits(i).SplitSize

                'Tim split cuối có hiển thị
                If i = .Splits.ColCount - 1 Then
                    If .Splits(i).SplitSize > 0 And bSplitDisable Then
                        .Splits(i).SplitSizeMode = C1.Win.C1TrueDBGrid.SizeModeEnum.Scalable
                    End If
                End If
            Next i
            'Set lại cột để ScrollBar kéo về đầu
            If .Splits.Count = 1 Then .Col = iFirstCol

        End With
    End Sub


    Private Sub GetInfoSplit(ByVal index As Integer, ByRef splitSize As Integer, ByRef splitSizeMode As Integer)
        Dim dr() As DataRow = _dtGrid.Select(COL_SplitIndex & " = " & index.ToString & " And GridName=" & SQLString(_tdbgO.Name) & " And ( Mode ='-1' or Mode = " & SQLString(_mode) & ")")
        If dr.Length = 0 Then Exit Sub
        splitSize = L3Int(dr(0).Item("SplitSize"))
        splitSizeMode = L3Int(dr(0).Item("SplitSizeMode"))
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD91T2022
    '# Created User: Đỗ Minh Dũng
    '# Created Date: 28/07/2009 01:55:54
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD91T2022() As String
        Dim sSQL As String = ""
        sSQL &= "Delete D91T2022"
        sSQL &= " Where ModuleID = " & SQLString(_moduleID) & " And FormID = " & SQLString(_formID.Name) & vbCrLf
        sSQL &= " And ( Mode ='-1' or Mode = " & SQLString(_mode) & ") And FieldName" & UnicodeJoin(gbUnicode) & " <> ''"
        sSQL &= " And UserID = " & SQLString(gsUserID)
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T2022s
    '# Created User: Đỗ Minh Dũng
    '# Created Date: 28/07/2009 01:56:52
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T2022s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        Dim dr() As DataRow = CType(tdbgD99U1111.DataSource, DataTable).Select(COL_IsUsed & "=0")
        For i As Integer = 0 To dr.Length - 1
            sSQL.Append("Insert Into D91T2022(")
            sSQL.Append("ModuleID, FormID, FieldName, FieldNameU, UserID, Mode")
            sSQL.Append(") Values(" & vbCrLf)
            sSQL.Append(SQLString(_moduleID) & COMMA) 'ModuleID, int, NULL
            sSQL.Append(SQLString(_formID.Name) & COMMA) 'FormID, varchar[10], NULL
            If gbUnicode Then
                sSQL.Append("''" & COMMA) 'FieldName, varchar[50], NULL
                sSQL.Append("N" & SQLString(dr(i).Item(COL_FieldName)) & COMMA) 'FieldNameU, varchar[50], NULL
            Else
                sSQL.Append(SQLString(dr(i).Item(COL_FieldName)) & COMMA) 'FieldName, varchar[50], NULL
                sSQL.Append("N''" & COMMA) 'FieldNameU, varchar[50], NULL
            End If
            sSQL.Append(SQLString(gsUserID) & COMMA) 'UserID, varchar[10], NULL
            sSQL.Append(SQLString(dr(i).Item(COL_Mode))) 'Mode,  varchar[50], NULL
            'sSQL.Append(COMMA & SQLString(_TemplateID)) 'TemplateID,  varchar[50], NULL
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub

        If Not AllowRefresh() Then Exit Sub

        btnSave.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        sSQL.Append(SQLDeleteD91T2022() & vbCrLf)
        sSQL.Append(SQLInsertD91T2022s())

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            btnRefresh_Click(Nothing, Nothing)
        Else
            SaveNotOK()
        End If
        btnSave.Enabled = True
    End Sub

    Public Sub picClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picClose.Click
        If _dtGrid Is Nothing Then Exit Sub
        tdbgD99U1111.UpdateData()
        'Dim dr() As DataRow = _dtGrid.Select(COL_IsUpdate & "=1")
        Dim dr() As DataRow = _dtGrid.Select(COL_IsUpdate & "=1 and  (Mode ='" & _mode.ToString & "' Or Mode ='-1')")
        If dr.Length > 0 Then
            If D99C0008.MsgAsk(r("Thong_tin_tren_luoi_da_thay_doi_ban_co_muon_Refresh_lai_khong")) = Windows.Forms.DialogResult.Yes Then
                tdbgD99U1111.UpdateData()
                btnRefresh_Click(Nothing, Nothing)
            Else
                For i As Integer = 0 To dr.Length - 1
                    dr(i).Item(COL_IsUpdate) = 0
                    dr(i).Item(COL_IsUsed) = Not L3Bool(dr(i).Item(COL_IsUsed))
                Next
            End If
        End If
        _formID.Focus()
        Me.Hide()
    End Sub

    Private Sub LoadLanguage()
        lbl1.Text = r("Chon_cot_hien_thi")
        tdbgD99U1111.Columns("IsUsed").Caption = r("Chon")
        tdbgD99U1111.Columns("Description").Caption = r("Ten_cotU")
        btnSave.Text = r("_Luu")
    End Sub

#Region "Events of tdbg"
    Private Sub tdbg_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgD99U1111.AfterColUpdate
        tdbgD99U1111.Columns(COL_IsUpdate).Text = "1"
    End Sub

    Private Sub tdbg_BeforeColEdit(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColEditEventArgs) Handles tdbgD99U1111.BeforeColEdit
        If tdbgD99U1111.Columns(COL_Obligatory).Text = "1" Then e.Cancel = True
    End Sub

    Dim bHeadClick As Boolean = True

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgD99U1111.HeadClick
        If e.ColIndex = IndexOfColumn(tdbgD99U1111, COL_IsUsed) Then L3HeadClick()
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbgD99U1111.KeyDown
        If e.Control And e.KeyCode = Keys.S Then
            If tdbgD99U1111.Col = IndexOfColumn(tdbgD99U1111, COL_IsUsed) Then
                L3HeadClick()
            End If
        Else
            HotKeyDownGrid(e, tdbgD99U1111, IndexOfColumn(tdbgD99U1111, COL_IsUsed))
        End If
    End Sub

    Private Sub L3HeadClick()
        bHeadClick = Not bHeadClick
        For i As Integer = 0 To tdbgD99U1111.RowCount - 1
            If tdbgD99U1111(i, COL_Obligatory).ToString = "1" Then Continue For

            tdbgD99U1111(i, COL_IsUpdate) = 1
            tdbgD99U1111(i, COL_IsUsed) = bHeadClick
        Next i
    End Sub
#End Region

    Private Function AllowRefresh() As Boolean
        Dim dr() As DataRow = _dtGrid.Select(COL_IsUsed & "=1")

        If dr.Length = 0 Then
            D99C0008.MsgL3(r("MSG000010"))
            tdbgD99U1111.Col = IndexOfColumn(tdbgD99U1111, COL_IsUsed)
            tdbgD99U1111.Row = 0
            tdbgD99U1111.SplitIndex = SPLIT0
            If tdbgD99U1111.Splits(SPLIT0).MarqueeStyle <> C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor Then tdbgD99U1111.Focus()
            Return False
        End If

        Return True
    End Function

    Private Sub tdbg_OwnerDrawCell(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.OwnerDrawCellEventArgs) Handles tdbgD99U1111.OwnerDrawCell

        If tdbgD99U1111(e.Row, COL_Obligatory).ToString = "0" Then Exit Sub

        Dim objPen As New Pen(Color.FromName("Green"))
        Dim pt As New Point()
        Dim X As Integer = e.CellRect.X + L3Int(FormatNumber(((e.CellRect.Width - 11) / 2), 0)) - 2
        Dim rect As New Rectangle(X, e.CellRect.Y + 2, 12, e.CellRect.Height - 5) '(X,Y)(width, Height)

        'Update 31/05/2010
        If e.Row Mod 2 = 0 Then
            e.Graphics.FillRectangle(Brushes.White, e.CellRect)
        Else
            e.Graphics.FillRectangle(Brushes.Beige, e.CellRect)
        End If
        e.Graphics.FillRectangle(Brushes.DarkGray, rect)
        e.Graphics.DrawRectangle(objPen, rect)

        'The commented out line below causes a red checkmark to be drawn in the topmost cell only of the column
        pt.X = e.CellRect.X + L3Int(FormatNumber((e.CellRect.Width / 2), 0)) - 5 '3
        'No red checkmark is drawn in any cell in the column using the Y-cord below
        pt.Y = rect.Y + 5 'e.CellRect.Y + L3Int(((e.CellRect.Height - 11) / 2)) - 5

        'Lines moving downward left to right--left side of check (The checkmark is 3 lines thick)
        e.Graphics.DrawLine(objPen, pt.X, pt.Y + 0, pt.X + 2, pt.Y + 2)
        e.Graphics.DrawLine(objPen, pt.X, pt.Y + 1, pt.X + 2, pt.Y + 3)
        e.Graphics.DrawLine(objPen, pt.X, pt.Y + 2, pt.X + 2, pt.Y + 4)
        'Lines moving upward left to right--right side of check
        e.Graphics.DrawLine(objPen, pt.X + 2, pt.Y + 2, pt.X + 6, pt.Y - 2)
        e.Graphics.DrawLine(objPen, pt.X + 2, pt.Y + 3, pt.X + 6, pt.Y - 1)
        e.Graphics.DrawLine(objPen, pt.X + 2, pt.Y + 4, pt.X + 6, pt.Y - 0)

        e.Handled = True

    End Sub

    Private Sub D99U1111_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If Me.Visible = False Then Exit Sub
        tdbgD99U1111.SplitIndex = 0
        tdbgD99U1111.Col = 0
        tdbgD99U1111.Row = 0
        tdbgD99U1111.Focus()
    End Sub

    Friend Sub GetButtonPress(ByVal iButton As Integer)
        _Button = iButton
        btnRefresh_Click(Nothing, Nothing)
    End Sub


    ''Mẫu truy vấn
    'Friend Sub ShowForm(ByVal tdbgO As C1.Win.C1TrueDBGrid.C1TrueDBGrid, Optional ByVal iSplit As Integer = 0, Optional ByVal ArrColObl() As String = Nothing)
    '    If btnSave.Enabled = False Then btnRefresh.Left = btnSave.Left
    '    'set Group là Bắt buộc nhập
    '    For j As Integer = 0 To tdbgO.GroupedColumns.Count - 1
    '        Dim dr() As DataRow = _dtGrid.Select(COL_FieldName & "=" & tdbgO.GroupedColumns(j).DataField)
    '        For index As Integer = 0 To dr.Length - 1
    '            If tdbgO.GroupedColumns(j).Caption.Trim = dr(index).Item(COL_Description).ToString.Trim Then '' OrElse tdbgO.GroupedColumns(j).DataField = dr(COL_FieldName).ToString Then
    '                dr(index).Item("IsUsed") = 1
    '                dr(index).Item("Obligatory") = 1 ' Group là Bắt buộc nhập
    '                Exit For
    '            Else
    '                'Nếu trên lưới tdbgO bỏ Group thì Set lại giá trị bắt buộc nhập = 0
    '                If ArrColObl IsNot Nothing AndAlso ArrColObl.Length > 0 Then
    '                    If Not L3FindArrString(ArrColObl, tdbgO.GroupedColumns(j).DataField) Then dr(index).Item("Obligatory") = 0 ' Không Bắt buộc nhập
    '                End If
    '            End If
    '        Next
    '    Next
    'End Sub
End Class


Imports System
Public Class D91F6010

    Private WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker
    Private ChildName As String = "D91E0240"
    Dim exe As D91E0240
    Dim p As System.Diagnostics.Process

    Private _whereValue As String ' Giá trị cần tìm kiếm
    Public WriteOnly Property WhereValue() As String
        Set(ByVal Value As String)
            _whereValue = Value
        End Set
    End Property

    Private _inListID As String
    Public WriteOnly Property InListID() As String
        Set(ByVal Value As String)
            _inListID = Value
        End Set
    End Property

    Private _inWhere As String 'Điều kiện tìm kiếm, truyen vao mot chuoi tim kiem
    Public WriteOnly Property InWhere() As String
        Set(ByVal Value As String)
            _inWhere = Value
        End Set
    End Property

    Private _outPut01 As String ' Kết quả tìm kiếm trả về
    Public ReadOnly Property OutPut01() As String
        Get
            Return _outPut01
        End Get

    End Property

    Private _outPut02 As String ' Kết quả tìm kiếm trả về ObjectTypeID cho ListID = 2
    Public ReadOnly Property OutPut02() As String
        Get
            Return _outPut02
        End Get

    End Property

    Private _formPermision As String
    Public WriteOnly Property FormPermision() As String
        Set(ByVal Value As String)
            _formPermision = Value
        End Set
    End Property

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundWorker1.DoWork
        'Tạo một process gắn với exe con, process này sẽ quan sát exe con.
        Dim p As System.Diagnostics.Process
        Try
            p = Process.GetProcessesByName(ChildName)(0)
            If p Is Nothing Then
                Exit Sub
            End If
            p.EnableRaisingEvents = True
            p.WaitForExit()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub FormLock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Ẩn form trung gian
        Me.Size = New Size(0, 0)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        '----Truyền tham số exe con------
        exe = New D91E0240(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString, gsDivisionID, giTranMonth, giTranYear)
        exe.FormActive = D91E0240Form.D91F6010
        exe.InListID = _inListID
        exe.InWhere = _inWhere
        exe.InWhereValue = _whereValue
        exe.Run()

        'Bắt đầu chạy cơ chế background
        backgroundWorker1 = New System.ComponentModel.BackgroundWorker
        backgroundWorker1.RunWorkerAsync()
    End Sub

    'sự kiện hoàn thành và dừng của Background
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundWorker1.RunWorkerCompleted
        _outPut01 = exe.Output01
        _outPut02 = exe.Output02
        Me.Close()
    End Sub

End Class
Imports System

Public Class D95M0240

    Private WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker
    Private ChildName As String = "D95E0240"
    Dim exe As D95E0240

    Private _FormActive As String
    Public WriteOnly Property FormActive() As String
        Set(ByVal Value As String)
            _FormActive = Value
        End Set
    End Property

    Private _formPermission As String = ""
    Public WriteOnly Property FormPermission() As String
        Set(ByVal Value As String)
            _formPermission = Value
        End Set
    End Property

    Private _iD01 As String = ""
    Public WriteOnly Property ID01() As String
        Set(ByVal Value As String)
            _iD01 = Value
        End Set
    End Property

    Private _iD02 As String
    Public WriteOnly Property ID02() As String
        Set(ByVal Value As String)
            _iD02 = Value
        End Set
    End Property

    Private _bClose As Boolean = True
    Public ReadOnly Property bClose() As Boolean
        Get
            Return _bClose
        End Get
    End Property

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundWorker1.DoWork
        'Tạo một process gắn với exe con, process này sẽ quan sát exe con.
        Try
            Dim p As System.Diagnostics.Process
            p = Process.GetProcessesByName(ChildName)(0)
            If p Is Nothing Then
                Exit Sub
            End If
            p.EnableRaisingEvents = True
            'p.WaitForExit()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FormLock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Ẩn form trung gian
        Me.Size = New Size(0, 0)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        '----Truyền tham số exe con------
        exe = New D95E0240(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString, gsDivisionID, giTranMonth, giTranYear)
        exe.FormActive = _FormActive
        If _formPermission = "" Then _formPermission = _FormActive
        exe.FormPermission = _formPermission
        exe.ID01 = _iD01
        exe.ID02 = _iD02
        exe.Run()

        'Bắt đầu chạy cơ chế background
        backgroundWorker1 = New System.ComponentModel.BackgroundWorker
        backgroundWorker1.RunWorkerAsync()
    End Sub

    'sự kiện hoàn thành và dừng của Background
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundWorker1.RunWorkerCompleted
        _bClose = exe.Close
        Me.Close()
    End Sub

End Class

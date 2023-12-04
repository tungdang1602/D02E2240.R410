Imports System

Public Class D95E0240
    Private Const EXEMODULE As String = "D95"
    Private Const EXECHILD As String = "D95E0240"
    Private sLanguage As String

    ''' <summary>
    ''' Khởi tạo exe con D95E0140
    ''' </summary>
    ''' <param name="Server">Server kết nối đến hệ thống</param>
    ''' <param name="Database">Database kết nối đến hệ thống</param>
    ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
    ''' <param name="Password">Password kết nối đến hệ thống</param>
    ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
    ''' <param name="Language">Ngôn ngữ sử dụng</param>
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String)
        sLanguage = Language
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ServerName", Server, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DBName", Database, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ConnectionUserID", UserDatabaseID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Password", Password, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "UserID", UserID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Language", Language)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "CreateBy", gsCreateBy)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "CodeTable", gbUnicode.ToString)
    End Sub

    ''' <summary>
    ''' Khởi tạo exe con D95E0140
    ''' </summary>
    ''' <param name="Server">Server kết nối đến hệ thống</param>
    ''' <param name="Database">Database kết nối đến hệ thống</param>
    ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
    ''' <param name="Password">Password kết nối đến hệ thống</param>
    ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
    ''' <param name="Language">Ngôn ngữ sử dụng</param>
    ''' <param name="DivisionID">Đơn vị hiện tại</param>
    ''' <param name="TranMonth">Tháng kế toán hiện tại</param>
    ''' <param name="TranYear">Năm kế toán hiện tại</param>
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String, ByVal DivisionID As String, ByVal TranMonth As Integer, ByVal TranYear As Integer)
        Me.New(Server, Database, UserDatabaseID, Password, UserID, Language)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DivisionID", DivisionID)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranMonth", TranMonth.ToString)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranYear", TranYear.ToString)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Closed", gbClosed.ToString)
    End Sub

    ''' <summary>
    ''' Màn hình cần hiển thị cho exe con
    ''' </summary>
    Public WriteOnly Property FormActive() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", Value)
        End Set
    End Property

    Public WriteOnly Property Ctrl02() As EnumFormState
        Set(ByVal Value As EnumFormState)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl02", CType(Value, Integer).ToString)
        End Set
    End Property

    ''' <summary>
    ''' Màn hình phân quyền cho exe con
    ''' </summary>
    Public WriteOnly Property FormPermission() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl03", Value)
        End Set
    End Property

    ''' <summary>
    ''' FormState
    ''' </summary>
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal Value As EnumFormState)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "FormState", CType(Value, Integer).ToString)
        End Set
    End Property

    Public WriteOnly Property ID01() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID01", Value)
        End Set
    End Property


    Public WriteOnly Property ID02() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID02", Value)
        End Set
    End Property

    Public ReadOnly Property Close() As Boolean
        Get
            Return L3Bool(D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Close", "True"))
        End Get
    End Property

    ''' <summary>
    ''' Thực thi exe con
    ''' </summary>
    Public Sub Run()
        If Not ExistFile(gsApplicationSetup & "\" & EXECHILD & ".exe") Then Exit Sub
        Dim pInfo As New System.Diagnostics.ProcessStartInfo(gsApplicationSetup & "\" & EXECHILD & ".exe")
        pInfo.Arguments = "/DigiNet Corporation"
        pInfo.WindowStyle = ProcessWindowStyle.Normal
        SaveRunningExeSettings(MODULED02, EXECHILD)
        Process.Start(pInfo)
    End Sub

    ''' <summary>
    ''' Kiểm tra tồn tại exe con không ?
    ''' </summary>
    Private Function ExistFile(ByVal Path As String) As Boolean
        If System.IO.File.Exists(Path) Then Return True
        If sLanguage = "0" Then
            D99C0008.MsgL3("Không tồn tại file " & EXECHILD & ".exe")
        Else
            D99C0008.MsgL3("Not exist file" & EXECHILD & ".exe")
        End If
        Return False
    End Function

End Class

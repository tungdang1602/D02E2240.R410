''' <summary>
''' Module này liên qua đến các khai báo biến, enum, ... toàn cục
''' </summary>
''' <remarks>Các khai báo ở đây không được trùng với các khai báo ở các Module D99Xxxxx</remarks>
Module D02X0001

    ''' <summary>
    ''' Module đang coding D02E2240
    ''' </summary>
    Public Const MODULED02 As String = "D02E2240"
    ''' <summary>
    ''' Chuỗi D02
    ''' </summary>
    Public Const D02 As String = "D02"
    ''' <summary>
    ''' Dùng cho kiểm tra Security theo chuẩn của DIGINET
    ''' </summary>
    Public Const L3_APP_NAME As String = "Lemon3"
    ''' <summary>
    ''' Dùng cho kiểm tra Security theo chuẩn của DIGINET
    ''' </summary>
    Public Const L3_HS_SECTION As String = "HandshakeR360"
    ''' <summary>
    ''' Dùng cho kiểm tra Security theo chuẩn của DIGINET
    ''' </summary>
    Public Const L3_HS_MODULE As String = "D02"
    ''' <summary>
    ''' Dùng cho kiểm tra Security theo chuẩn của DIGINET
    ''' </summary>
    Public Const L3_HS_VALUE As String = "R3.60.00.Y2007"
    ''' <summary>
    ''' Dùng cho kiểm tra lưu thành công hay không
    ''' </summary>
    Public gbSavedOK As Boolean = False
    Public gbEnabledMenuFind As Boolean = False

    ''' <summary>
    ''' Khai báo structure cho phần định dạng format
    ''' </summary>
    'Public Structure StructureFormat

    '    ''' <summary>
    '    ''' format thành tiền
    '    ''' </summary>
    '    Public OriginalAmount As String
    '    ''' <summary>
    '    ''' Số làm tròn của thành tiền
    '    ''' </summary>
    '    Public OriginalAmountRound As Integer
    '    ''' <summary>
    '    ''' format thành tiền quy đổi
    '    ''' </summary>
    '    Public ConvertedAmount As String
    '    ''' <summary>
    '    ''' Số làm tròn của thành tiền quy đổi
    '    ''' </summary>
    '    Public ConvertedAmountRound As Integer
    '    ''' <summary>
    '    ''' format tỷ giá
    '    ''' </summary>
    '    Public ExchangeRate As String
    '    ''' <summary>
    '    ''' Số làm tròn của tỷ giá
    '    ''' </summary>
    '    Public ExchangeRateRound As Integer
    '    ''' <summary>
    '    ''' Nguyên tệ gốc
    '    ''' </summary>
    '    Public BaseCurrencyID As String
    '    ''' <summary>
    '    ''' Dấu phân cách thập phân
    '    ''' </summary>
    '    Public DecimalSeperator As String
    '    ''' <summary>
    '    ''' Dấu phân cách hàng ngàn
    '    ''' </summary>
    '    Public ThousandSeperator As String
    '    Public DefaultNumber2 As String
    '    Public DefaultNumber4 As String
    '    Public Percentage As String
    '    '------------------------------------------------------------------------
    '    '  D91 Format here
    '    '------------------------------------------------------------------------
    'End Structure

    ''' <summary>
    ''' Khai bao de chon cac button tren luoi vd: Khoan muc, doi tuong,mat hang...
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum Button
        ObjectButton = 0
        Inventory = 1
        Ana = 2
        Other = 3
    End Enum

    'Public D02Format As StructureFormat
    ''' <summary>
    ''' Lưu trữ các thiết lập tùy chọn
    ''' </summary>
    'Public D02Options As StructureOption

    ''' <summary>
    ''' Khai báo Structure cho phần Tùy chọn của Module
    ''' </summary>
    'Public Structure StructureOption
    '    ''' <summary>
    '    ''' Hỏi trước khi lưu
    '    ''' </summary>
    '    Public MessageAskBeforeSave As Boolean
    '    ''' <summary>
    '    ''' Thông báo khi lưu thành công
    '    ''' </summary>
    '    Public MessageWhenSaveOK As Boolean
    '    ''' <summary>
    '    ''' Hiển thị form chọn kỳ kế toán khi chạy chương trình
    '    ''' </summary>
    '    Public ViewFormPeriodWhenAppRun As Boolean
    '    ''' <summary>
    '    ''' Lưu giá trị gần nhất
    '    ''' </summary>
    '    Public SaveLastRecent As Boolean
    '    ''' <summary>
    '    ''' Lưu đơn vị mặc định
    '    ''' </summary>
    '    Public DefaultDivisionID As String
    '    ''' <summary>
    '    ''' Khóa thành tiền quy đổi
    '    ''' </summary>
    '    Public LockConvertedAmount As Boolean
    '    ''' <summary>
    '    ''' Làm tròn thành tiền quy đổi
    '    ''' </summary>
    '    Public RoundConvertedAmount As Boolean
    '    ''' <summary>
    '    ''' Hiển thị quy trình sơ đồ nghiệp vụ
    '    ''' </summary>
    '    Public ViewWorkflow As Boolean
    '    ''' <summary>
    '    ''' Ngôn ngữ báo cáo
    '    ''' </summary>
    '    Public ReportLanguage As Byte

    '    Public ShowReportPath As Boolean
    '    '------------------------------------------------------------------------
    '    '  D02 Options here
    '    '------------------------------------------------------------------------
    '    '------------------------------------------------------------------------
    '    Public DepAndEmpID As Boolean 'Bộ phận tiếp nhận và người tiếp nhận
    '    Public CipName As Boolean 'Tên XDCB
    '    Public CipDescription As Boolean 'Diễn giải
    'End Structure


    ''' <summary>
    ''' Lưu trữ các thiết lập Thông tin hệ thống
    ''' </summary>
    'Public D02Systems As StructureSystem

    ''' <summary>
    ''' Khai báo structure cho phần Thiết lập hệ thống
    ''' </summary>
    'Public Structure StructureSystem
    '    ''' <summary>
    '    ''' Đơn vị mặc định
    '    ''' </summary>
    '    Public DefaultDivisionID As String
    '    ''' <summary>
    '    ''' TK tài sản
    '    ''' </summary>
    '    Public DefAssetAccountID As String
    '    ''' <summary>
    '    ''' TK khấu hao
    '    ''' </summary>
    '    Public DefDepreciationAccountID As String
    '    ''' <summary>
    '    ''' Nguồn vốn
    '    ''' </summary>
    '    Public DefSourceID As String
    '    ''' <summary>
    '    ''' Phân bổ KH
    '    ''' </summary>
    '    Public DefAssignmentID As String
    '    ''' <summary>
    '    ''' Phương pháp KH
    '    ''' </summary>
    '    Public MethodID As String
    '    ''' <summary>
    '    ''' Xử lý KH kỳ cuối
    '    ''' </summary>
    '    Public MethodEndID As String
    '    ''' <summary>
    '    ''' Các bút toán giảm TS
    '    ''' </summary>
    '    Public DecreaseAsset As Boolean
    '    ''' <summary>
    '    ''' Tạo mã tự động cho tài sản cố định
    '    ''' </summary>
    '    Public AssetAuto As Integer
    '    ''' <summary>
    '    ''' Phân loại 1
    '    ''' </summary>
    '    Public AssetS1Enabled As Boolean
    '    ''' <summary>
    '    ''' Phân loại 2
    '    ''' </summary>
    '    Public AssetS2Enabled As Boolean
    '    ''' <summary>
    '    ''' Phân loại 3
    '    ''' </summary>
    '    Public AssetS3Enabled As Boolean
    '    ''' <summary>
    '    ''' AssetS1Default
    '    ''' </summary>
    '    Public AssetS1Default As String
    '    ''' <summary>
    '    ''' AssetS2Default
    '    ''' </summary>
    '    Public AssetS2Default As String
    '    ''' <summary>
    '    ''' AssetS3Default
    '    ''' </summary>
    '    Public AssetS3Default As String
    '    ''' <summary>
    '    ''' S1Length
    '    ''' </summary>
    '    Public S1Length As Integer
    '    ''' <summary>
    '    ''' S2Length
    '    ''' </summary>
    '    Public S2Length As Integer
    '    ''' <summary>
    '    ''' S3Length
    '    ''' </summary>
    '    Public S3Length As Integer
    '    ''' <summary>
    '    ''' Dấu phân cách
    '    ''' </summary>
    '    Public AssetSeperated As Boolean
    '    ''' <summary>
    '    ''' Dấu phân cách
    '    ''' </summary>
    '    Public AssetSeperator As String
    '    ''' <summary>
    '    ''' Dạng hiển thị
    '    ''' </summary>
    '    Public AssetOutputOrder As String
    '    ''' <summary>
    '    ''' Độ dài số
    '    ''' </summary>
    '    Public AutoNumberLength As Integer
    '    ''' <summary>
    '    ''' Độ dài mã
    '    ''' </summary>
    '    Public AssetOutputLength As Integer
    '    ''' <summary>
    '    ''' Độ dài mã Incident 69247
    '    ''' </summary>
    '    Public CIPAuto As Boolean
    '    ''' <summary>
    '    ''' Mã TSCĐ và mã CCDC tăng liên tục
    '    ''' </summary>
    '    Public IsAssetIDForD02D43 As Boolean
    '    ''' <summary>
    '    ''' BĐS đầu tư
    '    ''' </summary>
    '    Public UseProperty As Boolean
    '    ''' <summary>
    '    ''' Bắt buộc nhập người tiếp nhận
    '    ''' </summary>
    '    Public ObligatoryReceiver As Boolean
    '    ''' <summary>
    '    ''' Tập hợp XDCB cho mã BĐS
    '    ''' </summary>
    '    Public CIPforPropertyProduct As Boolean
    '    '------------------------------------------------------------------------
    '    '  D02 Systems here
    '    '------------------------------------------------------------------------
    'End Structure

    ''' <summary>
    ''' Khai báo cho phần định dạng chung lấy từ D91P9300
    ''' Createdate 20/12/2007
    ''' </summary>
    ''' <remarks></remarks>
    'Public Structure StructureFormatNew
    '    ''' <summary>
    '    ''' Format tỷ giá
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public ExchangeRate As String
    '    Public iExchangeRate As Integer
    '    ''' <summary>
    '    ''' Format nguyên tệ 
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public DecimalPlaces As String
    '    ''' <summary>
    '    ''' Format nguyên tệ ứng với mỗi loại tiền
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public MyOriginal As String
    '    ''' <summary>
    '    ''' Format tiền quy đổi
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public D90_Converted As String
    '    Public iD90_Converted As Integer
    '    ''' <summary>
    '    ''' Format số lượng, số lượng quy đổi theo nhóm sản xuất
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public D07_Quantity As String
    '    ''' <summary>
    '    ''' Format đơn giá theo nhóm sản xuất
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public D07_UnitCost As String
    '    Public D08_Quantity As String
    '    Public D08_UnitCost As String
    '    Public D08_Ratio As String
    '    Public DecimalSeparator As String
    '    Public ThousandSeparator As String
    '    Public D90_ConvertedDecimals As Integer
    '    Public BaseCurrencyID As String 'Loai tien hoach toan
    '    ''' <summary>
    '    ''' Format 2 số lẽ
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public DefaultNumber2 As String
    '    ''' <summary>
    '    ''' Format 4 số lẽ
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public DefaultNumber4 As String
    '    ''' <summary>
    '    ''' Format 0 số lẽ
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public DefaultNumber0 As String
    'End Structure

    'Public D02Format As StructureFormatNew

    'Quyền Sửa số phiếu
    'Public giPerF5558 As Integer

    ''' <summary>
    ''' Dung de luu các trạng thai cua AuditLog
    ''' </summary>
    ''' <remarks></remarks>
    Public Const PathReport9 As String = "\XReports\"
    ''' <summary>
    ''' Dùng cho form Chọn đường dẫn báo cáo: Custom Report
    ''' </summary>
    ''' <remarks></remarks>
    Public Const PathCustomizedReport9 As String = "\XCustom\"
    ''' <summary>
    '''  Dùng cho form Chọn đường dẫn báo cáo
    ''' </summary>
    ''' <remarks></remarks>
    Public gsReportPath As String
    Public giFirstTranMonth As Integer 'Dùng để phân cho D02F0300 :Số dư đầu kỳ
    Public giFirstTranYear As Integer 'Dùng để phân cho D02F0300 :Số dư đầu kỳ
End Module

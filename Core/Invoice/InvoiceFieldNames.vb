''' <summary>
''' 发票/行程单字段的标准字段名常量。
''' 统一 key 命名，供本地解析、百度 OCR 映射、扩展字段字典共用，避免各处手写字符串不一致。
''' </summary>
Public NotInheritable Class InvoiceFieldNames

    Private Sub New()
    End Sub

    ' —— 增值税电子发票通用字段 ——
    Public Const InvoiceCode As String = "发票代码"
    Public Const InvoiceNumber As String = "发票号码"
    Public Const InvoiceDate As String = "开票日期"
    Public Const BuyerName As String = "购买方名称"
    Public Const BuyerTaxId As String = "购买方纳税人识别号"
    Public Const SellerName As String = "销售方名称"
    Public Const SellerTaxId As String = "销售方纳税人识别号"
    Public Const ItemName As String = "项目名称"
    Public Const Specification As String = "规格型号"
    Public Const Unit As String = "单位"
    Public Const Quantity As String = "数量"
    Public Const UnitPrice As String = "单价"
    Public Const Amount As String = "金额"
    Public Const TaxRate As String = "税率"
    Public Const TaxAmount As String = "税额"
    Public Const TotalWithTax As String = "价税合计"
    Public Const Remark As String = "备注"
    Public Const CheckCode As String = "校验码"
    Public Const Payee As String = "收款人"
    Public Const Reviewer As String = "复核人"
    Public Const Drawer As String = "开票人"

    ' —— 网约车/滴滴行程单字段 ——
    Public Const Passenger As String = "乘车人"
    Public Const DepartureTime As String = "出发时间"
    Public Const ArrivalTime As String = "到达时间"
    Public Const StartLocation As String = "起点"
    Public Const EndLocation As String = "终点"
    Public Const ServiceType As String = "车型服务类型"
    Public Const OrderNumber As String = "订单号"
    Public Const TripAmount As String = "行程金额"
    Public Const City As String = "城市"
    Public Const DriverInfo As String = "司机车辆信息"

End Class

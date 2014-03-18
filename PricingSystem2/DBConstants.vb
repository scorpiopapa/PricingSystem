Module DBConstants
#If DEBUG Then
    Public ReadOnly DB_FILE = "..\..\..\price.mdb"
    'Private ReadOnly DB_FILE = "F:\Project\供应商系统\PricingSystem2\PricingSystem2\price.mdb"
#Else
    Public ReadOnly DB_FILE = Application.StartupPath + "\price.mdb"
#End If

    Public ReadOnly CONNECTION_STRING As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DB_FILE + ";Persist Security Info=true"

    'Public ReadOnly REPORT_MASTER_TABLE = "report_master"
    'Public ReadOnly REPORT_NAME_COLUMN = "report_name"
    'Public ReadOnly REPORT_TYPE_COLUMN = "类型"
    'Public ReadOnly REPORT_ID_COLUMN = "id"

    Public ReadOnly INT As System.Type = System.Type.GetType("System.Int32")
    Public ReadOnly STR As System.Type = System.Type.GetType("System.String")
    Public ReadOnly DT As System.Type = System.Type.GetType("System.DateTime")

    Public ReadOnly NUMBER_TYPE As String = "数值"
    Public ReadOnly TEXT_TYPE As String = "文本"
    Public ReadOnly DATE_TYPE As String = "日期"

    Public ReadOnly ITEM_ID_COLUMN As String = "item_id"
    Public ReadOnly MATAURE_DATE_COLUMN As String = "到期日"
    Public ReadOnly PROJECT_NAME_COLUMN As String = "所属项目"
    Public ReadOnly BID_COMPANY As String = "报价单位"
    Public ReadOnly ITEM_NAME_COLUMN As String = "品目名称"
    Public ReadOnly BID_PRICE_COLUMN As String = "报价"
    Public ReadOnly BID_COMMENT_COLUMN As String = "报价备注"
    Public ReadOnly OFFER_PRICE_COLUMN As String = "审价"
    Public ReadOnly OFFER_COMMENT_COLUMN As String = "审价备注"
End Module

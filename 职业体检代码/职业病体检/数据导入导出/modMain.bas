Attribute VB_Name = "modMain"
Option Explicit


Sub main()
    Dim lstrServer  As String
    Dim lstrData As String
    
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")

    '初始化数据访问对象。
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    

    
    FrmLogin.Show
End Sub

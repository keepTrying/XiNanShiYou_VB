Attribute VB_Name = "modMain"
Option Explicit


Sub main()
    Dim lstrServer  As String
    Dim lstrData As String
    
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    lstrData = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")

    '��ʼ�����ݷ��ʶ���
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    

    
    FrmLogin.Show
End Sub

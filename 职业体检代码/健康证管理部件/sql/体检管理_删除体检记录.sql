SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


ALTER   proc ������_ɾ������¼
	@pϵͳ��� varchar(20)='00110306260005'
as
set nocount on 
declare @l����������� varchar(20)

--�ж��Ƿ����½��ۡ�
if exists(select * from ������_��������Ϣ�� where ϵͳ���=@pϵͳ��� and ���״̬=3)
	select 1	
else begin
	--��ȡ������š�
	select @l�����������=����������� from ������_��������Ϣ�� where ϵͳ���=@pϵͳ���
	select @l�����������=isnull(@l�����������,'')

	--ɾ����������Ϣ��ͨ������������ɾ��������ر����ݡ�
	delete ������_��������Ϣ�� where ϵͳ���=@pϵͳ���
	delete ������_�����ʱ�־�� where ϵͳ���=@pϵͳ���

	--����ֻ����1������¼��ɾ������������
	if not exists(select * from (select count(*) as num from ������_��������Ϣ�� where �����������=@l�����������) a where a.num>0)
		delete ������_�����Ա������Ϣ�� where �����������=@l�����������
		
	--����©���ı�š�
	--insert into ������_©���ı�ű�(�������,���,��Ԥ��)
	--values('ϵͳ���',@pϵͳ���,0)

    select 0
end	


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


ALTER      view dbo.ְҵ�����_��ѯͳ����ͼ
as
select a.ϵͳ���,
		--b.�շ�����,
		a.����,a.�Ա�,a.����,f.��λ����,
		b.������ as ����,b.������� as �����Ա����,a.Σ������,
		b.������,
		--b.�������,
		convert(varchar(10),b.�������,120) as �������,
		a.�ֹ���,f.Ƭ��,f.��������,f.��ҵ���,
		--b.����������,
		--b.����ϵͳ���,
		--g.���� as ��������,
		--d.���ֽ���,d.��������,
		--e.��� as ҽʦ���,
		--e.���� as ҽʦ����,
		--�շѽ�� = case when d.���� = 16 then b.�շѽ�� 
		--else null
		--end,
		���״̬ = case
			--2012.12.10 ���� ��
			--���Ӷ������״̬
			when b.���״̬=0 then 'δУ��'
			when b.���״̬=1 then 'δ���嵥'
			when b.���״̬=2 then 'δ¼���ܼ��߸�����Ϣ'
			when b.���״̬=3 then '�����'
			when b.���״̬=4 then 'δ�½���'
			when b.���״̬=5 then '���½���'
			when b.���״̬=6 then '�Ѹ���'
			when b.���״̬=7 then '�ѷ�����'
			when b.���״̬=8 then '������'
			end
			--2012.12.10 ���� ��
      from ְҵ�����_�����Ա������Ϣ�� a left join ְҵ�����_��������Ϣ�� b on a.ϵͳ���=b.ϵͳ���
					 --left join ְҵ�����_ְҵʷ�� c on b.ϵͳ���=c.ϵͳ���
					 --left join ְҵ�����_���ҽ��۱� d on a.ϵͳ��� = d.ϵͳ���
					 --left join ϵͳ����_Ա��������Ϣ�� e on a.ҽ����� = e.���
					 left join ��λ����_��λ������Ϣ��ͼ f on a.��λ������=f.������
					 --left join ϵͳ����_�ֵ�_�ֵ����ݱ� g on d.���� = g.��� and g.ID = 84


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

ALTER    proc dbo.ְҵ�����_ɾ������¼
	@pϵͳ��� varchar(20)='00110306260005'
as
set nocount on 
--declare @l����������� varchar(20)

--�ж��Ƿ����½��ۡ�
if exists(select * from ְҵ�����_��������Ϣ�� where ϵͳ���=@pϵͳ��� and ���״̬=3)
	select 1	
else begin
	--��ȡ������š�
	--select @l�����������=����������� from ְҵ�����_��������Ϣ�� where ϵͳ���=@pϵͳ���
	--select @l�����������=isnull(@l�����������,'')

	--ɾ����������Ϣ��ͨ������������ɾ��������ر����ݡ�
	delete ְҵ�����_��������Ϣ�� where ϵͳ���=@pϵͳ���
	delete 	dbo.ְҵ�����_��������ʷ�� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_ְҵʷ�� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_���ҽ��۱� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_��� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_������ where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_������� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_�򳣹滯��� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_Ѫ���滯��� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_B��Ӱ��� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_���߿� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_X��Ӱ��� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_�ڿ� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_��ٿ� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_�ι���Ӱ��� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_�ĵ�� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_Ⱦɫ�廯��� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼��� where ϵͳ���=@pϵͳ���
	delete 	dbo.ְҵ�����_������ʷ�� where ϵͳ���=@pϵͳ���
	delete dbo.ְҵ�����_�Ծ�֢״�� where ϵͳ���=@pϵͳ���
	
	--delete ������_�����ʱ�־�� where ϵͳ���=@pϵͳ���

	--����ֻ����1������¼��ɾ������������
	if not exists(select * from (select count(*) as num from ְҵ�����_��������Ϣ�� where ϵͳ���=@pϵͳ���) a where a.num>0)
		delete ְҵ�����_�����Ա������Ϣ�� where ϵͳ���=@pϵͳ���
		
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




SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





---------------------------------------------
/*
  ��ѯ�����Դ�ӡ������Ѿ���ӡ�������word����
  ���ߣ������ 
  ʱ�䣺2013-1-14 
*/
---------------------------------------------
ALTER                 proc [dbo].[ְҵ�����_��ѯ��챨����Ϣ]
	@p��ʼ���� datetime ='2004-01-01',
	@p��ֹ���� datetime ='2005-12-01',
	@p�������� VARCHAR ( 50 )='',
	@pϵͳ��� varchar(40)='',
	@p��λ���� VARCHAR ( 100 )='',
        @p���� VARCHAR ( 20 )=''
as
set nocount on 


if @p��ֹ����<>'' 
   if charindex(' ',@p��ֹ����)=0 select @p��ֹ����=@p��ֹ����+' 23:59:59'

	select a.ϵͳ���,b.������,c.����,c.�Ա�,c.����,c.��λ����,a.������ as ��������,c.Σ������,c.�ֹ���,a.�������,a.������, Convert(varchar(100),a.�������,23)as �������,
		  case a.���״̬ when 6 then 'δ��ӡ' when 7 then '�Ѵ�ӡ'end as ����״̬
			from ְҵ�����_��������Ϣ�� a left join ְҵ�����_��챨����Ϣ�� b on a.ϵͳ���=b.ϵͳ���  left join ְҵ�����_�����Ա������Ϣ�� c  
				 on a.ϵͳ���=c.ϵͳ���
			where ((�������>=@p��ʼ���� or @p��ʼ����='')
		and  (�������<=@p��ֹ���� or @p��ֹ����='')
		and (a.ϵͳ���=@pϵͳ��� or @pϵͳ���='')
        	and (a.������=@p�������� or @p��������='')
		and ���״̬ in(6,7)
		and  (��λ���� like '%'+@p��λ����+'%' or @p��λ����='')
		and  (���� like '%'+@p����+'%' or @p����='')
			 )		
order by c.��λ����








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


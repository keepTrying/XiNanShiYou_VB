if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_���ҽʦ��Ŀ�������ݿ�]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ְҵ�����_���ҽʦ��Ŀ�������ݿ�]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create   VIEW dbo.ְҵ�����_���ҽʦ��Ŀ�������ݿ�
AS 
select a.ҽʦ���, ҽʦ���� = b.����, a.�����Ŀ,�����Ŀ���� = c.����,
	c.����,c.ö����Դ,c.ȱʡֵ,case when c.�ȽϷ�ʽ='����' or c.�ȽϷ�ʽ='=' or c.�ȽϷ�ʽ is null then '' else c.�ȽϷ�ʽ end +c.��׼ֵ as ��׼ֵ,c.��λ
    from ������_���ҽʦ��Ŀ���ñ� a,
        ϵͳ����_Ա��������Ϣ�� b,
        ְҵ�����_�����Ŀ���ñ� c
    where a.ҽʦ��� *= b.��� and
         a.�����Ŀ = c.����

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO








if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_���������ݿ�]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ְҵ�����_���������ݿ�]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--������쵥��
CREATE          VIEW dbo.ְҵ�����_���������ݿ�
AS 
    select a.ϵͳ���,b.������ݺ���,b.����, b.�Ա�,b.����,b.��������,b.Σ������,b.ְҵ����,
	b.����Դ,b.�ֹ���,b.ְ���ְ��,b.�������,b.����,b.ְҵΣ������,b.�绰����,b.סַ,
	b.�ʱ�,b.�Ļ��̶�,b.����,b.����,b.���,a.�������,
        b.��������,b.������,b.��λ������, b.��λ����, a.�Թܱ��, a.������, a.�������,
        a.������,a.�½�������,��Ϻʹ������, a.�½���ҽʦ,�½���ҽʦ����=c.����,
        a.������, a.���״̬, a.����������, a.����ϵͳ���, a.�շ�����, a.�������״̬
    FROM ְҵ�����_��������Ϣ�� a, 
        ְҵ�����_�����Ա������Ϣ�� b,
        ϵͳ����_Ա��������Ϣ�� c
    WHERE a.ϵͳ���=b.ϵͳ��� and
          a.�½���ҽʦ*= c.���

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO









if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�������ͼ]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ְҵ�����_�������ͼ]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE           view ְҵ�����_�������ͼ as 
--2012-07-04 �ڵ�� ��
--ְҵ�����_�������Ϣ�� ֮���ɾ
--select *
--from  ְҵ�����_�������Ϣ�� 
--union
--2012-07-04 �ڵ�� ��

select * from dbo.ְҵ�����_�����Ϣ_��ٿ�
union select * from dbo.ְҵ�����_�����Ϣ_�ڿ�
union select * from dbo.ְҵ�����_�����Ϣ_���
union select * from dbo.ְҵ�����_�����Ϣ_Ѫ���滯���
union select * from dbo.ְҵ�����_�����Ϣ_�ι��ܻ����
union select * from dbo.ְҵ�����_�����Ϣ_�򳣹滯���
union select * from dbo.ְҵ�����_�����Ϣ_Ⱦɫ�廯���
union select * from dbo.ְҵ�����_�����Ϣ_�������
union select * from dbo.ְҵ�����_�����Ϣ_X��Ӱ���
union select * from dbo.ְҵ�����_�����Ϣ_�ĵ��
union select * from dbo.ְҵ�����_�����Ϣ_B��Ӱ���
union select * from dbo.ְҵ�����_�����Ϣ_�ι���Ӱ���
union select * from dbo.ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼���

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO











if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����շ���ͼ]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ְҵ�����_����շ���ͼ]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE   view ְҵ�����_����շ���ͼ as 
select   a.ϵͳ���,a.�����Ŀ,b.��λ������,b.��λ����,c.����,b.�������,b.������
from  ְҵ�����_�������ͼ a, ְҵ�����_���������ݿ� b,ְҵ�����_�����Ŀ���ñ� c
where a.ϵͳ���=b.ϵͳ��� and c.����=a.�����Ŀ

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����������ѯ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ְҵ�����_����������ѯ]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE    proc dbo.ְҵ�����_����������ѯ 
	@p��ʼ���� VARCHAR ( 10 )='2004-01-01',
	@p��ֹ���� VARCHAR ( 10 )='2005-12-01',
	@p�������� VARCHAR ( 50 )='',
    @p��λ���� VARCHAR ( 100 )='',
    @p���� VARCHAR ( 20 )='',
	@p��쵥�� varchar(20)='',
    @p�Թܱ�� VARCHAR ( 20 )='',
	@pϵͳ��� varchar(40)=''
AS
set nocount on
if @p��ֹ����<>'' 
   if charindex(' ',@p��ֹ����)=0 select @p��ֹ����=@p��ֹ����+' 23:59:59'

select a.ϵͳ���,����,�Ա�,����,��λ����,�Թܱ��,������,convert(varchar(10),�������,120) as �������,������,isnull(����������,'') as ����������,isnull(����ϵͳ���,'') as ����ϵͳ���,���״̬=case when ���״̬=3 then '���½���' else 'δ�½���' end
from ְҵ�����_��������Ϣ�� a inner join  ְҵ�����_�����Ա������Ϣ�� b on a.ϵͳ���=b.ϵͳ���
where ((�������>=@p��ʼ���� or @p��ʼ����='')
		and  (�������<=@p��ֹ���� or @p��ֹ����='')
		and  (��λ���� like '%'+@p��λ����+'%' or @p��λ����='')
		and  (���� like '%'+@p����+'%' or @p����='')
        and (�Թܱ��=@p�Թܱ�� or @p�Թܱ��='')
	and (a.ϵͳ���=@pϵͳ��� or @pϵͳ���='')
)	
or (���״̬=3 and isnull(����������,'')<>'' and isnull(����ϵͳ���,'')='')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_ɾ������¼]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ְҵ�����_ɾ������¼]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE   proc ְҵ�����_ɾ������¼
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
	delete ������_�����ʱ�־�� where ϵͳ���=@pϵͳ���

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




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_ְҵ��ʷ��������ѯ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ְҵ�����_ְҵ��ʷ��������ѯ]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE      proc dbo.ְҵ�����_ְҵ��ʷ��������ѯ 
	--2012-07-06 �ڵ�� ��
	--�����и��Ŀ�ʼ�������ֹ���ھ�ȷ�����Ӻ��룬������������¸�ʽ 
	--@p��ʼ���� VARCHAR ( 10 )='2004-01-01',
	--@p��ֹ���� VARCHAR ( 10 )='2005-12-01',
	@p��ʼ���� datetime ='2004-01-01',
	@p��ֹ���� datetime ='2005-12-01',
	--2012-07-06 �ڵ�� ��
	@p�������� VARCHAR ( 50 )='',
    @p��λ���� VARCHAR ( 100 )='',
    @p���� VARCHAR ( 20 )='',
	@p��쵥�� varchar(20)='',
    @p�Թܱ�� VARCHAR ( 20 )='',
	@pϵͳ��� varchar(40)=''
AS
set nocount on
if @p��ֹ����<>'' 
   if charindex(' ',@p��ֹ����)=0 select @p��ֹ����=@p��ֹ����+' 23:59:59'

select a.ϵͳ���,����,�Ա�,����,��λ����,�Թܱ��,������,convert(varchar(10),�������,120) as �������,������,isnull(����������,'') as ����������,isnull(����ϵͳ���,'') as ����ϵͳ���,���״̬=
	case when ���״̬=3 then '���½���' when ���״̬=0 then '����ʷ¼��' when ���״̬=1 then '�����' else 'δ�½���' end
from ְҵ�����_��������Ϣ�� a inner join  ְҵ�����_�����Ա������Ϣ�� b on a.ϵͳ���=b.ϵͳ���
where ((�������>=@p��ʼ���� or @p��ʼ����='')
		and  (�������<=@p��ֹ���� or @p��ֹ����='')
		and  (��λ���� like '%'+@p��λ����+'%' or @p��λ����='')
		and  (���� like '%'+@p����+'%' or @p����='')
        --and (�Թܱ��=@p�Թܱ�� or @p�Թܱ��='')
	and (a.ϵͳ���=@pϵͳ��� or @pϵͳ���='')
)	
or (���״̬=3 and isnull(����������,'')<>'' and isnull(����ϵͳ���,'')='')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO











if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_��������ݿ�]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ְҵ�����_��������ݿ�]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE        VIEW dbo.ְҵ�����_��������ݿ�
AS 
    SELECT a.ϵͳ���,b.������ݺ���, 
        b.����, b.�Ա�, b.��������, a.�Թܱ��, 
        a.�������, a.���״̬, c.�����Ŀ, �����Ŀ����=d.����, 
        c.�����, c.���ҽʦ, ���ҽʦ����=e.����, c.��д����, d.����,
        d.ȱʡֵ,d.ö����Դ,d.������,f.����,d.��׼ֵ,d.��λ,d.����,c.�������
    FROM ְҵ�����_��������Ϣ�� a, 
        ְҵ�����_�����Ա������Ϣ�� b, 
        ְҵ�����_�������Ϣ�� c, 
        ְҵ�����_�����Ŀ���ñ� d,
        ϵͳ����_Ա��������Ϣ�� e,ϵͳ����_�ֵ�_�ֵ����ݱ� f
    WHERE a.ϵͳ���=c.ϵͳ��� and 
	a.ϵͳ���=b.ϵͳ��� and
        c.�����Ŀ= d.���� and 
        c.���ҽʦ*= e.��� and d.������=f.InnerID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����ģ�����������ݿ�]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ְҵ�����_����ģ�����������ݿ�]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-------------------------------------------------------------------------------------------------
create  VIEW dbo.ְҵ�����_����ģ�����������ݿ�
AS 
    Select a.��������, a.������, ���������� = b.����
    From ְҵ�����_����ģ�������۱� a, 
        ϵͳ����_�ֵ�_�ֵ����ݱ� b
    Where a.������=b.InnerID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO








if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����ģ�帽����Ŀ���ݿ�]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ְҵ�����_����ģ�帽����Ŀ���ݿ�]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

create  VIEW dbo.ְҵ�����_����ģ�帽����Ŀ���ݿ�
AS
SELECT dbo.ְҵ�����_����ģ�帽����Ŀ��Ϣ��.��������, 
      dbo.ְҵ�����_����ģ�帽����Ŀ��Ϣ��.������Ŀ,
      dbo.ְҵ�����_����ģ�帽����Ŀ��Ϣ��.���,
      dbo.ְҵ�����_����ģ�帽����Ŀ��Ϣ��.�Ƿ��¼, 
      dbo.ְҵ�����_�����Ա������Ŀ���ñ�.¼�����, 
      dbo.ְҵ�����_�����Ա������Ŀ���ñ�.��������, 
      dbo.ְҵ�����_�����Ա������Ŀ���ñ�.���ݳ���, 
      dbo.ְҵ�����_�����Ա������Ŀ���ñ�.ö��ֵ
FROM dbo.ְҵ�����_����ģ�帽����Ŀ��Ϣ�� INNER JOIN
      dbo.ְҵ�����_�����Ա������Ŀ���ñ� ON 
      dbo.ְҵ�����_����ģ�帽����Ŀ��Ϣ��.������Ŀ = dbo.ְҵ�����_�����Ա������Ŀ���ñ�.������Ŀ

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_��ѯͳ����ͼ]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ְҵ�����_��ѯͳ����ͼ]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


create view ְҵ�����_��ѯͳ����ͼ
as
select a.ϵͳ���,b.�շ�����,a.����,a.�Ա�,b.������ as ����,b.������� as �����Ա����,
		b.������,b.�������,c.����,f.��λ����,f.Ƭ��,f.��������,f.��ҵ���,b.����������,
		b.����ϵͳ���,g.���� as ��������,d.���ֽ���,d.��������,e.��� as ҽʦ���,e.���� as ҽʦ����,
		�շѽ�� = case
		when d.���� = 16 then b.�շѽ��
		else null
		end
	
      from ְҵ�����_�����Ա������Ϣ�� a left join ְҵ�����_��������Ϣ�� b on a.ϵͳ���=b.ϵͳ���
					 left join ְҵ�����_ְҵʷ�� c on b.ϵͳ���=c.ϵͳ���
					 left join ְҵ�����_���ҽ��۱� d on c.ϵͳ��� = d.ϵͳ���
					 left join ϵͳ����_Ա��������Ϣ�� e on d.ҽ����� = e.���
					 left join ��λ����_��λ������Ϣ��ͼ f on a.��λ������=f.������
					 left join ϵͳ����_�ֵ�_�ֵ����ݱ� g on d.���� = g.��� and g.id = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sel����ҽʦ��Ϣ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sel����ҽʦ��Ϣ]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure sel����ҽʦ��Ϣ(@para���� varchar(16),@para��� varchar(10))
as
begin
declare @str varchar(10)
set @str = rtrim(ltrim(@para���))

declare @intID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')

if len(@str) = 0
select a.����,a.���,c.���� from ϵͳ����_Ա��������Ϣ�� a,ְҵ�����_�û�����Ȩ�ޱ� b,ϵͳ����_�ֵ�_�ֵ����ݱ� c where a.���=b.�û���� and b.���ұ�� = c.��� and c.ID = @intID and a.���� = @para���� and ���� not in('���Ǽ�','ҵ������','ְҵ��ʷ¼��') order by c.���;
else
select a.����,a.���,c.���� from ϵͳ����_Ա��������Ϣ�� a,ְҵ�����_�û�����Ȩ�ޱ� b,ϵͳ����_�ֵ�_�ֵ����ݱ� c where a.���=b.�û���� and b.���ұ�� = c.��� and c.ID = @intID and a.��� = @para��� and ���� not in('���Ǽ�','ҵ������','ְҵ��ʷ¼��') order by c.���;
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


--�޸��Զ��½��۴洢����
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

ALTER     PROCEDURE autoConclusion
@paraSysNo varchar(16),		--ϵͳ���
@paraItem varchar(16),		--��������
@paraResult varchar(16),	--������
@paraDoctor varchar(16),	--ҽʦ����
@paraFillDate varchar(16),	--��дʱ��
@paraConclusion varchar(16),	--�������
@paraTableName varchar(40)	--��Ŀ���ڿ���
as
begin
  declare @paraItemID varchar(6)	--��Ŀ���
  declare @paraFlag int	--���ر�־
  declare @sqlstr nvarchar(4000)	--ִ�е�sql���
  declare @standard varchar(50)		--��׼ֵ
  
  select @paraItemID = ���� from ְҵ�����_�����Ŀ���ñ� where ���� = @paraItem	--�õ���Ŀ���
  select @standard = ��׼ֵ from ְҵ�����_�����Ŀ���ñ� where ���� = @paraItem and ���� = @paraItemID	--�õ���׼ֵ
  if '����'=@paraResult
  begin
    set @paraConclusion = '�ϸ�'
  end
  if '�쳣'=@paraResult
  begin
    set @paraConclusion = '���ϸ�'
  end

  if @standard<>''
  begin
	  if isnumeric(@paraResult)=1
	  begin
		if (convert(numeric,@paraResult)-convert(numeric,@standard))<>0
		begin
		  set @paraConclusion = '���ϸ�'
		end
		else
		begin
		  set @paraConclusion = '�ϸ�'
		end
	  end
  end
  set @sqlstr = N'select @paraFlag=count(*) from '+ @paraTableName +' where ϵͳ��� = '''+@paraSysNo+''' and �����Ŀ = '''+@paraItemID+''';'
  execute Sp_executeSql @sqlstr, N'@paraFlag int out',@paraFlag out
  if @paraFlag > 0
  begin
    set @sqlstr = 'update '+@paraTableName+' set �����='''+@paraResult+''',���ҽʦ='''+@paraDoctor+''',��дʱ��='''+@paraFillDate+''',�������='''+@paraConclusion+''' where ϵͳ���='''+@paraSysNo+''' and �����Ŀ='''+@paraItemID+''';'
    exec (@sqlstr)
  end
  else
  begin
    set @sqlstr = 'insert into '+@paraTableName+' values( '''+@paraSysNo+''' , '''+@paraItemID+''' , '''+@paraResult+''' , '''+@paraDoctor+''' , '+@paraFillDate+' , '''+@paraConclusion+''')'
    exec (@sqlstr)
  end
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


--2012-07-05 �ڵ�� 
--�����洢���̣�����ְҵ��ʷ(�ܼ��߸�����Ϣ)�����ѯʹ��
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_ְҵ��ʷ��������ѯ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ְҵ�����_ְҵ��ʷ��������ѯ]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE      proc dbo.ְҵ�����_ְҵ��ʷ��������ѯ 
	@p��ʼ���� VARCHAR ( 10 )='2004-01-01',
	@p��ֹ���� VARCHAR ( 10 )='2005-12-01',
	@p�������� VARCHAR ( 50 )='',
    @p��λ���� VARCHAR ( 100 )='',
    @p���� VARCHAR ( 20 )='',
	@p��쵥�� varchar(20)='',
    @p�Թܱ�� VARCHAR ( 20 )='',
	@pϵͳ��� varchar(40)=''
AS
set nocount on
if @p��ֹ����<>'' 
   if charindex(' ',@p��ֹ����)=0 select @p��ֹ����=@p��ֹ����+' 23:59:59'

select a.ϵͳ���,����,�Ա�,����,��λ����,�Թܱ��,������,convert(varchar(10),�������,120) as �������,������,isnull(����������,'') as ����������,isnull(����ϵͳ���,'') as ����ϵͳ���,���״̬=
	case when ���״̬=3 then '���½���' when ���״̬=0 then '����ʷ¼��' when ���״̬=1 then '�����' else 'δ�½���' end
from ְҵ�����_��������Ϣ�� a inner join  ְҵ�����_�����Ա������Ϣ�� b on a.ϵͳ���=b.ϵͳ���
where ((�������>=@p��ʼ���� or @p��ʼ����='')
		and  (�������<=@p��ֹ���� or @p��ֹ����='')
		and  (��λ���� like '%'+@p��λ����+'%' or @p��λ����='')
		and  (���� like '%'+@p����+'%' or @p����='')
        --and (�Թܱ��=@p�Թܱ�� or @p�Թܱ��='')
	and (a.ϵͳ���=@pϵͳ��� or @pϵͳ���='')
)	
or (���״̬=3 and isnull(����������,'')<>'' and isnull(����ϵͳ���,'')='')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


--2012-07-05 �ڵ��
--�����洢���̣�����ְҵ������ض�����ϵͳ���
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_���ɱ����ˮ��]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ְҵ�����_���ɱ����ˮ��]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



------------------------------------------------------------------------------------------
CREATE   PROCEDURE dbo.ְҵ�����_���ɱ����ˮ�� 
        	@p������� varchar(20)
AS
/****************************************************************************
    ���ݡ�������_��ŵ������ˮ�ű���¼��ָ��������Ƶĵ�ǰ��ʹ�õ������ˮ�ţ�
�����������ˮ�ŷ��أ���������ˮ�ż��������ˮ�ű�.
******************************************************************************/
    set nocount on 
    declare @l��� int
    declare @l���� int
    declare @l����ֵ varchar(10)
    declare @l�������� int

    --select @l�������� = count(*) from  [������_������ɹ����] where ������� = @p������� and  ��� in ('yy','mm','dd')  
    --set @l�������� =isnull(@l�������� ,1) 
    --if @l�������� = 1 
        begin    
        if exists(select ������Ŀ from [ְҵ�����_ҵ��������Ϣ��] 
                      where ������Ŀ = '��������ˮ��' and datepart(yyyy,˵��) = datepart(yyyy,getdate()))
            begin
            select @l���=����ֵ from [ְҵ�����_ҵ��������Ϣ��] 
                      where ������Ŀ = '��������ˮ��' and datepart(yyyy,˵��) = datepart(yyyy,getdate())
            set @l��� = @l��� + 1
            update [ְҵ�����_ҵ��������Ϣ��] set ����ֵ = @l��� 
                  where  ������Ŀ = '��������ˮ��' and datepart(yyyy, ˵��) = datepart(yyyy,getdate() )
            end
        else
            begin
            set @l��� = 1
            insert into [ְҵ�����_ҵ��������Ϣ��](������Ŀ,����ֵ,˵��) values('��������ˮ��','1',getdate())
            end
        end
    
    set @l����=7
    set @l����ֵ =  convert(varchar(10),@l���)
    set @l����ֵ = replicate('0',@l���� - len(@l����ֵ)) + @l����ֵ
    select @l����ֵ


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



--2012-07-05 �ڵ��
--�����洢���̣��˻�ְҵ�����ģ���ϵͳ���
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�˻ر����ˮ��]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ְҵ�����_�˻ر����ˮ��]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE  PROCEDURE dbo.ְҵ�����_�˻ر����ˮ��
	@p������� varchar(20)='ϵͳ���',
	@p��� varchar(20)='102000001'
AS
/****************************************************************************
���ܣ����ݡ�������_��ŵ������ˮ�ű���¼��ָ��������Ƶĵ�ǰ��ʹ�õ������ˮ�ţ�
     �������ˮ�ż�һ��
******************************************************************************/
set nocount on 
declare @l��� int          --��ǰ�����ˮ�š�
declare @l���� int          --��ˮ�ų��ȡ�
declare @l�������� int      --0 û�����ɹ���1 �����ֻ������ݣ�2 ����а������+�·ݣ�3 ������������ա�
declare @l��ˮ�� varchar(10)--����@p����е���ˮ�š�
declare @lintInsert as int

 
BEGIN
    select @lintInsert =0

    --��ȡ��ǰ��ŵ���ˮ�ų��ȡ�
    set @l����=isnull(@l����,7)

    --��ȡ����е���ˮ�š�
    set @l��ˮ��=right(@p���,@l����)
    begin
	--����а�����ݣ����ű���ÿ���ͷ��ʼ�ࡣ
        set @l��� =cast(@l��ˮ�� as int) - 1
		if not exists(select * from [ְҵ�����_ҵ��������Ϣ��]
			    	    where  ������Ŀ='��������ˮ��' and datepart(yyyy, ˵��) = datepart(yyyy,getdate()) 
						and ����ֵ=cast(@l��ˮ�� as int))
			select @lintInsert=1
		else
	        update [ְҵ�����_ҵ��������Ϣ��] set ����ֵ = @l��� 
    	    where  ������Ŀ='��������ˮ��' and datepart(yyyy, ˵��) = datepart(yyyy,getdate()) 
				and ����ֵ=cast(@l��ˮ�� as int)
    end		
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




--2012-07-05 �ڵ��
--�����洢���̣�����ְҵ������������ѯ����
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����������ѯ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ְҵ�����_����������ѯ]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE          proc dbo.ְҵ�����_����������ѯ 
	--2012-07-06 �ڵ�� ��
	--�����и��Ŀ�ʼ�������ֹ���ھ�ȷ�����Ӻ��룬������������¸�ʽ 
	--@p��ʼ���� VARCHAR ( 10 )='2004-01-01',
	--@p��ֹ���� VARCHAR ( 10 )='2005-12-01',
	@p��ʼ���� datetime ='2004-01-01',
	@p��ֹ���� datetime ='2005-12-01',
	--2012-07-06 �ڵ�� ��
	@p�������� VARCHAR ( 50 )='',
    @p��λ���� VARCHAR ( 100 )='',
    @p���� VARCHAR ( 20 )='',
	@p��쵥�� varchar(20)='',
    @p�Թܱ�� VARCHAR ( 20 )='',
	@pϵͳ��� varchar(40)=''
AS
set nocount on
if @p��ֹ����<>'' 
   if charindex(' ',@p��ֹ����)=0 select @p��ֹ����=@p��ֹ����+' 23:59:59'

select a.ϵͳ���,����,�Ա�,����,��λ����,�Թܱ��,������,
	convert(varchar(10),�������,120) as �������,
	������,isnull(����������,'') as ����������,
	isnull(����ϵͳ���,'') as ����ϵͳ���,
	���״̬=case 
		--2012-06-15 �ڵ�� ��
		--���Ӷ������״̬
		--when ���״̬=3 then '���½���' 
		--else 'δ�½���' 
		--end
		when ���״̬=0 then 'δУ��'
		when ���״̬=1 then 'δ���嵥'
		when ���״̬=2 then 'δ¼���ܼ��߸�����Ϣ'
		when ���״̬=3 then '�����'
		when ���״̬=4 then 'δ�½���'
		when ���״̬=5 then '���½���'
		when ���״̬=6 then '�Ѹ���'
		when ���״̬=7 then '�ѷ�����'
		when ���״̬=8 then '������'
		end
		--2012-06-15 �ڵ�� ��
from ְҵ�����_��������Ϣ�� a inner join  ְҵ�����_�����Ա������Ϣ�� b on a.ϵͳ���=b.ϵͳ���
where ((�������>=@p��ʼ���� or @p��ʼ����='')
		and  (�������<=@p��ֹ���� or @p��ֹ����='')
		and  (��λ���� like '%'+@p��λ����+'%' or @p��λ����='')
		and  (���� like '%'+@p����+'%' or @p����='')
        and (�Թܱ��=@p�Թܱ�� or @p�Թܱ��='')
	and (a.ϵͳ���=@pϵͳ��� or @pϵͳ���='')
        and (a.������=@p�������� or @p��������='')
)	
or (���״̬=3 and isnull(����������,'')<>'' and isnull(����ϵͳ���,'')='')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****************************************************************************
���ܣ����ݴ���ġ��������ơ�������š���ѯ�ÿ��ҵĽ����Ϣ��
******************************************************************************/

ALTER    proc [dbo].[ְҵ�����_�����Ϣ]
	@p���� varchar(20) = '��ٿ�',
	@p��� varchar(20) = '01'

as
set nocount on
declare @lsql varchar(1000)

	select @lsql='select distinct a.ϵͳ���,a.����,a.�Ա�,a.����,a.�������,a.��λ����,convert(varchar(10),b.��дʱ��,2) ��дʱ�� 
	  from ְҵ�����_���������ݿ� a, ְҵ�����_�����Ϣ_'+@p����+' b 
            where 1=1 and a.ϵͳ���=b.ϵͳ��� and (a.���״̬=''2'' or a.���״̬=''3'' or a.���״̬=''4'')
			   and  (substring(a.�������״̬,'+@p���+',1)=''1'' or substring(a.�������״̬,'+@p���+',1)=''2'')'

exec(@lsql)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  function z_fc(@s varchar(20))
returns varchar(1000)
as
begin
declare @str varchar(1000)
select @str = isnull(@str,' ') +����+',' from ְҵ�����_�������ͼ a,ְҵ�����_�����Ŀ���ñ� b 
	where a.�����Ŀ = b.���� and ϵͳ��� = @s and a.������� = '���ϸ�'
return stuff(@str,1,1,'')
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



--�޸�ְҵ�����_��������Ϣ��
ALTER TABLE ְҵ�����_��������Ϣ�� ADD ����״̬ varchar(2) null


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





ALTER               proc dbo.ְҵ�����_����������ѯ
	--2012-07-06 �ڵ�� ��
	--�����и��Ŀ�ʼ�������ֹ���ھ�ȷ�����Ӻ��룬������������¸�ʽ 
	--@p��ʼ���� VARCHAR ( 10 )='2004-01-01',
	--@p��ֹ���� VARCHAR ( 10 )='2005-12-01',
	@p��ʼ���� datetime ='2004-01-01',
	@p��ֹ���� datetime ='2012-11-01',
	--2012-07-06 �ڵ�� ��
	@p�������� VARCHAR ( 50 )='',
    @p��λ���� VARCHAR ( 100 )='',
    @p���� VARCHAR ( 20 )='',
	@p��쵥�� varchar(20)='',
    @p�Թܱ�� VARCHAR ( 20 )='',
	@pϵͳ��� varchar(40)=''
AS
set nocount on
if @p��ֹ����<>'' 
   if charindex(' ',@p��ֹ����)=0 select @p��ֹ����=@p��ֹ����+' 23:59:59'

select a.ϵͳ���,����,�Ա�,����,��λ����,�Թܱ��,������ as ������,
	convert(varchar(10),�������,120) as �������,
	������,isnull(����������,'') as ����������,
	isnull(����ϵͳ���,'') as ����ϵͳ���,
	���״̬=case 
		--2012-06-15 �ڵ�� ��
		--���Ӷ������״̬
		--when ���״̬=3 then '���½���' 
		--else 'δ�½���' 
		--end
		when ���״̬=0 then 'δУ��'
		when ���״̬=1 then 'δ���嵥'
		when ���״̬=2 then 'δ¼���ܼ��߸�����Ϣ'
		when ���״̬=3 then '�����'
		when ���״̬=6 then '������'
		when ���״̬=4 then 'δ�½���'
		--when ���״̬=5 then '���½���'
		when ���״̬=6 then '�Ѹ���'
		when ���״̬=7 then '�ѷ�����'
		--when a.����ϵͳ��� is not null then '������'
		when ���״̬=8 then '������'
		--when ���״̬=8 then '������'
		end
		--2012-06-15 �ڵ�� ��
from ְҵ�����_��������Ϣ�� a inner join  ְҵ�����_�����Ա������Ϣ�� b on a.ϵͳ���=b.ϵͳ���
where ((�������>=@p��ʼ���� or @p��ʼ����='')
		and  (�������<=@p��ֹ���� or @p��ֹ����='')
		and  (��λ���� like '%'+@p��λ����+'%' or @p��λ����='')
		and  (���� like '%'+@p����+'%' or @p����='')
        and (�Թܱ��=@p�Թܱ�� or @p�Թܱ��='')
	and (a.ϵͳ���=@pϵͳ��� or @pϵͳ���='')
        and (a.������=@p�������� or @p��������='')
)	
or (���״̬=3 and isnull(����������,'')<>'' and isnull(����ϵͳ���,'')='')



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






--2006-12-12������������С������������֤��
ALTER   PROCEDURE dbo.����֤_�Զ���ȡ������
		@pϵͳ��� varchar(20)='50110612120003'
AS
SET NOCOUNT ON
declare @�������� varchar(250)
declare @���������� varchar(250)
declare @l������ varchar(250)
declare @tmp������ varchar(250)
declare @l������ varchar(250)
declare @l��ѵ���� varchar(250),
	@l�������� varchar(50)
declare @i int
declare  @l int

select @��������=�������� from ����֤_ҵ�����ñ� 
select @��������=isnull(@��������,'')
if right(@��������,1)<>',' set @��������=@��������+','

--��ȡ���˵���ѵ���ۺ������ۡ�
select @l��ѵ����=b.���� from ����֤_��ҵ��Ա����֤������Ϣ�� a inner join ����֤_��ѵ�����ֵ���ͼ b 
 					on a.��ѵ����=b.InnerID where ϵͳ���=@pϵͳ���
select @l������=������ from ����֤_��ҵ��Ա����֤������Ϣ�� where ϵͳ���=@pϵͳ���
set @l������=rtrim(isnull(@l������,''))
if @l������<>'' set @l������=@l������+','

--��ȡ��������(--2006-12-12������������С������������֤����
select @l��������=b.���� from ����֤_��ҵ��Ա����֤������Ϣ�� a,ϵͳ����_���������ֵ�� b
where a.��������=b.InnerID 
and a.ϵͳ���=@pϵͳ���

set @i=1
if @l��ѵ����='�ϸ�'
begin
	set @����������='ͬ�֤ⷢ'
	set @l������='ͬ�֤ⷢ'

	--ֻҪ�������а�������һ����Ҫ����Ľ��ۣ����ж�Ϊ���롣
	while charindex(',',@l������,@i)>0 
	begin
		set @l=charindex(',',@l������,@i)
		set @tmp������=rtrim(substring(@l������,@i,@l-@i))
	 	--select @l,@tmp������

		if @tmp������<>'' begin
			set @tmp������=@tmp������+','

			--2006-12-12������������С������������֤��
			If charindex(@tmp������,@��������)>0  
				and not (@l�������� like '����%' and @tmp������ like '%�Ҹ�%')
		 	begin
	 			set @����������='��λ����'
		 		set @l������='����'
	 	  		break
		 	end 
     	end
		set @i=@l+1
	end 
end 
else
begin
  set @����������='��λ����'
  set @l������='����'
end 

--����������InnerID����������
select a.InnerID,@���������� from ϵͳ����_�ֵ�_�ֵ����ݱ� a inner join ϵͳ����_�ֵ�_�ֵ���б� b on a.ID=b.ID and b.����='����֤_��鴦������ֵ��' where a.����=@l������


GO



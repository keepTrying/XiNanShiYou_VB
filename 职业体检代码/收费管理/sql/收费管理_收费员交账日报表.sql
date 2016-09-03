SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



--exec �շѹ���_�շ�Ա�����ձ��� '2008-08-21','4002'
ALTER    PROCEDURE dbo.�շѹ���_�շ�Ա�����ձ���
	@p���� varchar(10) =null,
	@p�û���� varchar(10)=null
AS
SET NOCOUNT ON

if @p���� ='' or @p����='*' set @p����=null
if @p�û���� ='' or @p�û����='*' set @p�û����=null

create table #TMP_�շ���ʱ�� (
	col1 varchar(100) null,
	col2 varchar(100) null,
	col3 varchar(100) null,
	col4 varchar(100) null,
)
declare @date datetime
select @date=convert(datetime,@p����)

declare @projNo varchar(3),@projName varchar(50)

declare proj cursor for select �շ���Ŀ���,�շ���Ŀ���� from �շѹ���_�շ���Ŀ�ֵ�� where len(�շ���Ŀ���)=3 order by �շ���Ŀ���
open proj
fetch next from proj into @projNo,@projName
while @@fetch_status=0
    begin
	--��һ����Ŀͳ���ֽ��շ��ܶ�
	insert into #TMP_�շ���ʱ�� (col1,col2)
	select @projName,convert(varchar(10),sum(���))
	from �շѹ���_������Ϣ�� 
	where �շ���=@p�û���� and �շ�״̬=1 and left(�շ���Ŀ���,3)=@projNo and ���ѷ�ʽ=1
	and ��������=@date and ����>0
	--��һ����Ŀͳ��֧Ʊ�շ��ܶ�
	update #TMP_�շ���ʱ�� set col3=(
	select convert(varchar(10),sum(���)) from �շѹ���_������Ϣ�� 
	where �շ���=@p�û���� and �շ�״̬=1 and left(�շ���Ŀ���,3)=@projNo and ���ѷ�ʽ=2
	and ��������=@date and ����>0)
	where col1=@projName
	--��һ����Ŀͳ���˷��ܶ�
	update #TMP_�շ���ʱ�� set col4=(
	select convert(varchar(10),sum(���)) from �շѹ���_������Ϣ�� 
	where �շ���=@p�û���� and �շ�״̬=1 and left(�շ���Ŀ���,3)=@projNo 
	and ��������=@date and ����<0)
	where col1=@projName

	fetch next from proj into @projNo,@projName
    end
close proj
deallocate proj

--����Ʊ�ݵĺŶ�
declare @minNo int,@str1 varchar(100),@num int

select @minNo=0
Process1:
if exists(select * from �շѹ���_������Ϣ�� where �շ���=@p�û���� and ��������=@date and convert(int,�վݺ�)>@minNo)
    begin
	select @minNo=convert(int,min(�վݺ�)) from �շѹ���_������Ϣ��
	where �շ���=@p�û���� and ��������=@date and convert(int,�վݺ�)>@minNo
	select @str1=convert(varchar(10),@minNo),@num=1
	while exists(select * from �շѹ���_������Ϣ�� where convert(int,�վݺ�)=@minNo+1 and �շ���=@p�û���� and ��������=@date)
	    begin
		select @minNo=@minNo+1
		select @num=@num+1
	    end
	if @num=1
		insert into #TMP_�շ���ʱ�� values('�Ŷ�','1',@str1,'')
	else
		insert into #TMP_�շ���ʱ�� values('�Ŷ�',convert(varchar(4),@num),@str1+'~'+convert(varchar(10),@minNo),'')
	goto Process1
    end
select @num=count(distinct �շѱ��) from �շѹ���_������Ϣ�� where �շ���=@p�û���� and ��������=@date
insert into #TMP_�շ���ʱ�� values('����',convert(varchar(4),@num),'','')
--����Ʊ��
declare @no varchar(20)

declare proj1 cursor for select distinct �վݺ� from �շѹ���_������Ϣ�� where �շ���=@p�û���� and ��������=@date and �շ�״̬<>1
open proj1
fetch next from proj1 into @no
while @@fetch_status=0
    begin
	insert into #TMP_�շ���ʱ�� values('����Ʊ��',@no,'','')
	fetch next from proj1 into @no
    end
close proj1
deallocate proj1

--�˷�Ʊ��
declare proj1 cursor for select distinct �վݺ� from �շѹ���_������Ϣ�� where �շ���=@p�û���� and ��������=@date and �շ�״̬=1 and ����<0
open proj1
fetch next from proj1 into @no
while @@fetch_status=0
    begin
	insert into #TMP_�շ���ʱ�� values('�˷�Ʊ��',@no,'','')
	fetch next from proj1 into @no
    end
close proj1
deallocate proj1
select * from #TMP_�շ���ʱ��





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


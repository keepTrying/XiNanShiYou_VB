
--exec ������_��������� '2006-8-8'
--exec ������_��������� '2006-08-03'
if exists(select * from sysobjects where name='������_���������' and type='P')	
	drop proc ������_���������
go
create  proc dbo.������_���������
	@pDate as smalldatetime
as

set nocount on
    declare @l��� int
    declare @l������� varchar(20)
    declare @Weekday as int
    declare @lFirst as datetime

--��ȡָ���������ڵ�����һ��
select @Weekday=datepart(dw,@pDate)
if @Weekday=1 select @Weekday=7
select @Weekday=@Weekday-2
if @Weekday>0 
	select @lFirst=dateadd(day,-@Weekday,@pDate)
else 
	select @lFirst=@pDate

select @l�������='�����'

if exists(select �����ˮ�� from [������_��������ˮ�ű�] where ������� = @l������� and ���� = @lFirst)
    begin
    	update [������_��������ˮ�ű�] set �����ˮ�� =  isnull(�����ˮ��,0)+1
          where ������� = @l������� and ���� = @lFirst

    	select @l��� = �����ˮ�� from [������_��������ˮ�ű�] 
          where ������� = @l������� and ���� = @lFirst
    	
    end
else
    begin
    	set @l��� = 1
    	insert into [������_��������ˮ�ű�](�������,�����ˮ��,����) values(@l�������,1,@lFirst)
    end


select @l���

GO



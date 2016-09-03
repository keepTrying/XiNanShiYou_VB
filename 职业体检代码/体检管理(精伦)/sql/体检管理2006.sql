
--exec 体检管理_生成体检表号 '2006-8-8'
--exec 体检管理_生成体检表号 '2006-08-03'
if exists(select * from sysobjects where name='体检管理_生成体检表号' and type='P')	
	drop proc 体检管理_生成体检表号
go
create  proc dbo.体检管理_生成体检表号
	@pDate as smalldatetime
as

set nocount on
    declare @l编号 int
    declare @l编号名称 varchar(20)
    declare @Weekday as int
    declare @lFirst as datetime

--获取指定日期所在的星期一。
select @Weekday=datepart(dw,@pDate)
if @Weekday=1 select @Weekday=7
select @Weekday=@Weekday-2
if @Weekday>0 
	select @lFirst=dateadd(day,-@Weekday,@pDate)
else 
	select @lFirst=@pDate

select @l编号名称='体检表号'

if exists(select 最大流水号 from [体检管理_编号最大流水号表] where 编号名称 = @l编号名称 and 日期 = @lFirst)
    begin
    	update [体检管理_编号最大流水号表] set 最大流水号 =  isnull(最大流水号,0)+1
          where 编号名称 = @l编号名称 and 日期 = @lFirst

    	select @l编号 = 最大流水号 from [体检管理_编号最大流水号表] 
          where 编号名称 = @l编号名称 and 日期 = @lFirst
    	
    end
else
    begin
    	set @l编号 = 1
    	insert into [体检管理_编号最大流水号表](编号名称,最大流水号,日期) values(@l编号名称,1,@lFirst)
    end


select @l编号

GO



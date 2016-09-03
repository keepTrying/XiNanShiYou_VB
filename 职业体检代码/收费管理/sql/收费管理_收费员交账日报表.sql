SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



--exec 收费管理_收费员交账日报表 '2008-08-21','4002'
ALTER    PROCEDURE dbo.收费管理_收费员交账日报表
	@p日期 varchar(10) =null,
	@p用户编号 varchar(10)=null
AS
SET NOCOUNT ON

if @p日期 ='' or @p日期='*' set @p日期=null
if @p用户编号 ='' or @p用户编号='*' set @p用户编号=null

create table #TMP_收费临时表 (
	col1 varchar(100) null,
	col2 varchar(100) null,
	col3 varchar(100) null,
	col4 varchar(100) null,
)
declare @date datetime
select @date=convert(datetime,@p日期)

declare @projNo varchar(3),@projName varchar(50)

declare proj cursor for select 收费项目编号,收费项目名称 from 收费管理_收费项目字典表 where len(收费项目编号)=3 order by 收费项目编号
open proj
fetch next from proj into @projNo,@projName
while @@fetch_status=0
    begin
	--按一级项目统计现金收费总额
	insert into #TMP_收费临时表 (col1,col2)
	select @projName,convert(varchar(10),sum(金额))
	from 收费管理_费用信息表 
	where 收费人=@p用户编号 and 收费状态=1 and left(收费项目编号,3)=@projNo and 交费方式=1
	and 交费日期=@date and 数量>0
	--按一级项目统计支票收费总额
	update #TMP_收费临时表 set col3=(
	select convert(varchar(10),sum(金额)) from 收费管理_费用信息表 
	where 收费人=@p用户编号 and 收费状态=1 and left(收费项目编号,3)=@projNo and 交费方式=2
	and 交费日期=@date and 数量>0)
	where col1=@projName
	--按一级项目统计退费总额
	update #TMP_收费临时表 set col4=(
	select convert(varchar(10),sum(金额)) from 收费管理_费用信息表 
	where 收费人=@p用户编号 and 收费状态=1 and left(收费项目编号,3)=@projNo 
	and 交费日期=@date and 数量<0)
	where col1=@projName

	fetch next from proj into @projNo,@projName
    end
close proj
deallocate proj

--计算票据的号段
declare @minNo int,@str1 varchar(100),@num int

select @minNo=0
Process1:
if exists(select * from 收费管理_费用信息表 where 收费人=@p用户编号 and 交费日期=@date and convert(int,收据号)>@minNo)
    begin
	select @minNo=convert(int,min(收据号)) from 收费管理_费用信息表
	where 收费人=@p用户编号 and 交费日期=@date and convert(int,收据号)>@minNo
	select @str1=convert(varchar(10),@minNo),@num=1
	while exists(select * from 收费管理_费用信息表 where convert(int,收据号)=@minNo+1 and 收费人=@p用户编号 and 交费日期=@date)
	    begin
		select @minNo=@minNo+1
		select @num=@num+1
	    end
	if @num=1
		insert into #TMP_收费临时表 values('号段','1',@str1,'')
	else
		insert into #TMP_收费临时表 values('号段',convert(varchar(4),@num),@str1+'~'+convert(varchar(10),@minNo),'')
	goto Process1
    end
select @num=count(distinct 收费编号) from 收费管理_费用信息表 where 收费人=@p用户编号 and 交费日期=@date
insert into #TMP_收费临时表 values('张数',convert(varchar(4),@num),'','')
--作废票号
declare @no varchar(20)

declare proj1 cursor for select distinct 收据号 from 收费管理_费用信息表 where 收费人=@p用户编号 and 交费日期=@date and 收费状态<>1
open proj1
fetch next from proj1 into @no
while @@fetch_status=0
    begin
	insert into #TMP_收费临时表 values('作废票号',@no,'','')
	fetch next from proj1 into @no
    end
close proj1
deallocate proj1

--退费票号
declare proj1 cursor for select distinct 收据号 from 收费管理_费用信息表 where 收费人=@p用户编号 and 交费日期=@date and 收费状态=1 and 数量<0
open proj1
fetch next from proj1 into @no
while @@fetch_status=0
    begin
	insert into #TMP_收费临时表 values('退费票号',@no,'','')
	fetch next from proj1 into @no
    end
close proj1
deallocate proj1
select * from #TMP_收费临时表





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


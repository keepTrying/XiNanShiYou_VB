SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


ALTER   proc 体检管理_删除体检记录
	@p系统编号 varchar(20)='00110306260005'
as
set nocount on 
declare @l健康档案编号 varchar(20)

--判断是否已下结论。
if exists(select * from 体检管理_体检基本信息表 where 系统编号=@p系统编号 and 体检状态=3)
	select 1	
else begin
	--获取档案编号。
	select @l健康档案编号=健康档案编号 from 体检管理_体检基本信息表 where 系统编号=@p系统编号
	select @l健康档案编号=isnull(@l健康档案编号,'')

	--删除体检基本信息，通过触发器级联删除其他相关表内容。
	delete 体检管理_体检基本信息表 where 系统编号=@p系统编号
	delete 体检管理_体检访问标志表 where 系统编号=@p系统编号

	--该人只存在1条体检记录，删除健康档案。
	if not exists(select * from (select count(*) as num from 体检管理_体检基本信息表 where 健康档案编号=@l健康档案编号) a where a.num>0)
		delete 体检管理_体检人员基本信息表 where 健康档案编号=@l健康档案编号
		
	--插入漏掉的编号。
	--insert into 体检管理_漏掉的编号表(编号名称,编号,已预定)
	--values('系统编号',@p系统编号,0)

    select 0
end	


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


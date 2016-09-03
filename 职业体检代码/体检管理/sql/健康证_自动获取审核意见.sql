

--2006-12-12（公共场所的小三阳允许发健康证）
ALTER   PROCEDURE dbo.健康证_自动获取审核意见
		@p系统编号 varchar(20)='50110612120003'
AS
SET NOCOUNT ON
declare @调离设置 varchar(250)
declare @返回审核意见 varchar(250)
declare @l审核意见 varchar(250)
declare @tmp体检结论 varchar(250)
declare @l体检结论 varchar(250)
declare @l培训结论 varchar(250),
	@l卫生种类 varchar(50)
declare @i int
declare  @l int

select @调离设置=调离设置 from 健康证_业务配置表 
select @调离设置=isnull(@调离设置,'')
if right(@调离设置,1)<>',' set @调离设置=@调离设置+','

--获取该人的培训结论和体检结论。
select @l培训结论=b.名称 from 健康证_从业人员健康证申请信息表 a inner join 健康证_培训结论字典视图 b 
 					on a.培训结论=b.InnerID where 系统编号=@p系统编号
select @l体检结论=体检结论 from 健康证_从业人员健康证申请信息表 where 系统编号=@p系统编号
set @l体检结论=rtrim(isnull(@l体检结论,''))
if @l体检结论<>'' set @l体检结论=@l体检结论+','

--获取卫生种类(--2006-12-12（公共场所的小三阳允许发健康证）。
select @l卫生种类=b.名称 from 健康证_从业人员健康证申请信息表 a,系统管理_卫生种类字典表 b
where a.卫生种类=b.InnerID 
and a.系统编号=@p系统编号

set @i=1
if @l培训结论='合格'
begin
	set @返回审核意见='同意发证'
	set @l审核意见='同意发证'

	--只要体检结论中包括任意一个需要调离的结论，都判断为调离。
	while charindex(',',@l体检结论,@i)>0 
	begin
		set @l=charindex(',',@l体检结论,@i)
		set @tmp体检结论=rtrim(substring(@l体检结论,@i,@l-@i))
	 	--select @l,@tmp体检结论

		if @tmp体检结论<>'' begin
			set @tmp体检结论=@tmp体检结论+','

			--2006-12-12（公共场所的小三阳允许发健康证）
			If charindex(@tmp体检结论,@调离设置)>0  
				and not (@l卫生种类 like '公共%' and @tmp体检结论 like '%乙肝%')
		 	begin
	 			set @返回审核意见='岗位调离'
		 		set @l审核意见='调离'
	 	  		break
		 	end 
     	end
		set @i=@l+1
	end 
end 
else
begin
  set @返回审核意见='岗位调离'
  set @l审核意见='调离'
end 

--返回审核意见InnerID，审核意见。
select a.InnerID,@返回审核意见 from 系统管理_字典_字典内容表 a inner join 系统管理_字典_字典表列表 b on a.ID=b.ID and b.名称='健康证_审查处理意见字典表' where a.名称=@l审核意见


GO



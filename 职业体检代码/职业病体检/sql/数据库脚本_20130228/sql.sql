SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


ALTER      view dbo.职业病体检_查询统计视图
as
select a.系统编号,
		--b.收费批号,
		a.姓名,a.性别,a.年龄,f.单位名称,
		b.体检表编号 as 体检表,b.体检类型 as 体检人员类型,a.危害因素,
		b.体检类别,
		--b.体检日期,
		convert(varchar(10),b.体检日期,120) as 体检日期,
		a.现工种,f.片区,f.卫生种类,f.行业类别,
		--b.复查体检表编号,
		--b.复查系统编号,
		--g.名称 as 科室名称,
		--d.文字结论,d.结论日期,
		--e.编号 as 医师编号,
		--e.姓名 as 医师姓名,
		--收费金额 = case when d.科室 = 16 then b.收费金额 
		--else null
		--end,
		体检状态 = case
			--2012.12.10 张令 ↓
			--增加多种体检状态
			when b.体检状态=0 then '未校核'
			when b.体检状态=1 then '未打清单'
			when b.体检状态=2 then '未录入受检者个人信息'
			when b.体检状态=3 then '体检中'
			when b.体检状态=4 then '未下结论'
			when b.体检状态=5 then '已下结论'
			when b.体检状态=6 then '已复核'
			when b.体检状态=7 then '已发报告'
			when b.体检状态=8 then '待复查'
			end
			--2012.12.10 张令 ↑
      from 职业病体检_体检人员基本信息表 a left join 职业病体检_体检基本信息表 b on a.系统编号=b.系统编号
					 --left join 职业病体检_职业史表 c on b.系统编号=c.系统编号
					 --left join 职业病体检_科室结论表 d on a.系统编号 = d.系统编号
					 --left join 系统管理_员工基本信息表 e on a.医生编号 = e.编号
					 left join 单位档案_单位基本信息视图 f on a.单位申请编号=f.申请编号
					 --left join 系统管理_字典_字典内容表 g on d.科室 = g.编号 and g.ID = 84


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

ALTER    proc dbo.职业病体检_删除体检记录
	@p系统编号 varchar(20)='00110306260005'
as
set nocount on 
--declare @l健康档案编号 varchar(20)

--判断是否已下结论。
if exists(select * from 职业病体检_体检基本信息表 where 系统编号=@p系统编号 and 体检状态=3)
	select 1	
else begin
	--获取档案编号。
	--select @l健康档案编号=健康档案编号 from 职业病体检_体检基本信息表 where 系统编号=@p系统编号
	--select @l健康档案编号=isnull(@l健康档案编号,'')

	--删除体检基本信息，通过触发器级联删除其他相关表内容。
	delete 职业病体检_体检基本信息表 where 系统编号=@p系统编号
	delete 	dbo.职业病体检_个人生活史表 where 系统编号=@p系统编号
	delete dbo.职业病体检_职业史表 where 系统编号=@p系统编号
	delete dbo.职业病体检_科室结论表 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_外科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_生化科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_电测听科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_尿常规化验科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_血常规化验科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_B超影像科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_免疫科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_X光影像科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_内科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_五官科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_肺功能影像科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_心电科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_染色体化验科 where 系统编号=@p系统编号
	delete dbo.职业病体检_结果信息_受检者个人信息录入科 where 系统编号=@p系统编号
	delete 	dbo.职业病体检_既往病史表 where 系统编号=@p系统编号
	delete dbo.职业病体检_自觉症状表 where 系统编号=@p系统编号
	
	--delete 体检管理_体检访问标志表 where 系统编号=@p系统编号

	--该人只存在1条体检记录，删除健康档案。
	if not exists(select * from (select count(*) as num from 职业病体检_体检基本信息表 where 系统编号=@p系统编号) a where a.num>0)
		delete 职业病体检_体检人员基本信息表 where 系统编号=@p系统编号
		
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




SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





---------------------------------------------
/*
  查询出可以打印报告和已经打印并保存的word报告
  作者：罗李奎 
  时间：2013-1-14 
*/
---------------------------------------------
ALTER                 proc [dbo].[职业病体检_查询体检报告信息]
	@p开始日期 datetime ='2004-01-01',
	@p截止日期 datetime ='2005-12-01',
	@p体检表名称 VARCHAR ( 50 )='',
	@p系统编号 varchar(40)='',
	@p单位名称 VARCHAR ( 100 )='',
        @p姓名 VARCHAR ( 20 )=''
as
set nocount on 


if @p截止日期<>'' 
   if charindex(' ',@p截止日期)=0 select @p截止日期=@p截止日期+' 23:59:59'

	select a.系统编号,b.报告编号,c.姓名,c.性别,c.年龄,c.单位名称,a.体检表编号 as 体检表名称,c.危害因素,c.现工种,a.体检类型,a.体检类别, Convert(varchar(100),a.体检日期,23)as 体检日期,
		  case a.体检状态 when 6 then '未打印' when 7 then '已打印'end as 报告状态
			from 职业病体检_体检基本信息表 a left join 职业病体检_体检报告信息表 b on a.系统编号=b.系统编号  left join 职业病体检_体检人员基本信息表 c  
				 on a.系统编号=c.系统编号
			where ((体检日期>=@p开始日期 or @p开始日期='')
		and  (体检日期<=@p截止日期 or @p截止日期='')
		and (a.系统编号=@p系统编号 or @p系统编号='')
        	and (a.体检表编号=@p体检表名称 or @p体检表名称='')
		and 体检状态 in(6,7)
		and  (单位名称 like '%'+@p单位名称+'%' or @p单位名称='')
		and  (姓名 like '%'+@p姓名+'%' or @p姓名='')
			 )		
order by c.单位名称








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


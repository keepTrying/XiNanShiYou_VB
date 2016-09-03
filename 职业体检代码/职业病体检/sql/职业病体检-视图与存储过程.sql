if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检医师项目设置数据库]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[职业病体检_体检医师项目设置数据库]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create   VIEW dbo.职业病体检_体检医师项目设置数据库
AS 
select a.医师编号, 医师姓名 = b.姓名, a.体检项目,体检项目名称 = c.名称,
	c.属性,c.枚举来源,c.缺省值,case when c.比较方式='属于' or c.比较方式='=' or c.比较方式 is null then '' else c.比较方式 end +c.标准值 as 标准值,c.单位
    from 体检管理_体检医师项目设置表 a,
        系统管理_员工基本信息表 b,
        职业病体检_体检项目设置表 c
    where a.医师编号 *= b.编号 and
         a.体检项目 = c.编码

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO








if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检基本数据库]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[职业病体检_体检基本数据库]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--增加体检单号
CREATE          VIEW dbo.职业病体检_体检基本数据库
AS 
    select a.系统编号,b.公民身份号码,b.姓名, b.性别,b.年龄,b.建档日期,b.危害因素,b.职业分类,
	b.照射源,b.现工种,b.职务或职称,b.放射剂量,b.工龄,b.职业危害工龄,b.电话号码,b.住址,
	b.邮编,b.文化程度,b.籍贯,b.民族,b.婚否,a.体检类型,
        b.出生日期,b.出生地,b.单位申请编号, b.单位名称, a.试管编号, a.体检表编号, a.体检日期,
        a.体检结论,a.下结论日期,诊断和处理意见, a.下结论医师,下结论医师姓名=c.姓名,
        a.体检类别, a.体检状态, a.复查体检表编号, a.复查系统编号, a.收费批号, a.各科体检状态
    FROM 职业病体检_体检基本信息表 a, 
        职业病体检_体检人员基本信息表 b,
        系统管理_员工基本信息表 c
    WHERE a.系统编号=b.系统编号 and
          a.下结论医师*= c.编号

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO









if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检结果视图]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[职业病体检_体检结果视图]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE           view 职业病体检_体检结果视图 as 
--2012-07-04 于登淼 ↓
--职业病体检_体检结果信息表 之后会删
--select *
--from  职业病体检_体检结果信息表 
--union
--2012-07-04 于登淼 ↑

select * from dbo.职业病体检_结果信息_五官科
union select * from dbo.职业病体检_结果信息_内科
union select * from dbo.职业病体检_结果信息_外科
union select * from dbo.职业病体检_结果信息_血常规化验科
union select * from dbo.职业病体检_结果信息_肝功能化验科
union select * from dbo.职业病体检_结果信息_尿常规化验科
union select * from dbo.职业病体检_结果信息_染色体化验科
union select * from dbo.职业病体检_结果信息_电测听科
union select * from dbo.职业病体检_结果信息_X光影像科
union select * from dbo.职业病体检_结果信息_心电科
union select * from dbo.职业病体检_结果信息_B超影像科
union select * from dbo.职业病体检_结果信息_肺功能影像科
union select * from dbo.职业病体检_结果信息_受检者个人信息录入科

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO











if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检收费视图]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[职业病体检_体检收费视图]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE   view 职业病体检_体检收费视图 as 
select   a.系统编号,a.体检项目,b.单位申请编号,b.单位名称,c.单价,b.体检类型,b.体检类别
from  职业病体检_体检结果视图 a, 职业病体检_体检基本数据库 b,职业病体检_体检项目设置表 c
where a.系统编号=b.系统编号 and c.编码=a.体检项目

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检管理界面查询]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[职业病体检_体检管理界面查询]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE    proc dbo.职业病体检_体检管理界面查询 
	@p开始日期 VARCHAR ( 10 )='2004-01-01',
	@p截止日期 VARCHAR ( 10 )='2005-12-01',
	@p体检表名称 VARCHAR ( 50 )='',
    @p单位名称 VARCHAR ( 100 )='',
    @p姓名 VARCHAR ( 20 )='',
	@p体检单号 varchar(20)='',
    @p试管编号 VARCHAR ( 20 )='',
	@p系统编号 varchar(40)=''
AS
set nocount on
if @p截止日期<>'' 
   if charindex(' ',@p截止日期)=0 select @p截止日期=@p截止日期+' 23:59:59'

select a.系统编号,姓名,性别,年龄,单位名称,试管编号,体检表编号,convert(varchar(10),体检日期,120) as 体检日期,体检结论,isnull(复查体检表编号,'') as 复查体检表编号,isnull(复查系统编号,'') as 复查系统编号,体检状态=case when 体检状态=3 then '已下结论' else '未下结论' end
from 职业病体检_体检基本信息表 a inner join  职业病体检_体检人员基本信息表 b on a.系统编号=b.系统编号
where ((体检日期>=@p开始日期 or @p开始日期='')
		and  (体检日期<=@p截止日期 or @p截止日期='')
		and  (单位名称 like '%'+@p单位名称+'%' or @p单位名称='')
		and  (姓名 like '%'+@p姓名+'%' or @p姓名='')
        and (试管编号=@p试管编号 or @p试管编号='')
	and (a.系统编号=@p系统编号 or @p系统编号='')
)	
or (体检状态=3 and isnull(复查体检表编号,'')<>'' and isnull(复查系统编号,'')='')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_删除体检记录]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[职业病体检_删除体检记录]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE   proc 职业病体检_删除体检记录
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
	delete 体检管理_体检访问标志表 where 系统编号=@p系统编号

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




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_职业病史管理界面查询]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[职业病体检_职业病史管理界面查询]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE      proc dbo.职业病体检_职业病史管理界面查询 
	--2012-07-06 于登淼 ↓
	--代码中更改开始日期与截止日期精确到分钟和秒，所以这里更改下格式 
	--@p开始日期 VARCHAR ( 10 )='2004-01-01',
	--@p截止日期 VARCHAR ( 10 )='2005-12-01',
	@p开始日期 datetime ='2004-01-01',
	@p截止日期 datetime ='2005-12-01',
	--2012-07-06 于登淼 ↑
	@p体检表名称 VARCHAR ( 50 )='',
    @p单位名称 VARCHAR ( 100 )='',
    @p姓名 VARCHAR ( 20 )='',
	@p体检单号 varchar(20)='',
    @p试管编号 VARCHAR ( 20 )='',
	@p系统编号 varchar(40)=''
AS
set nocount on
if @p截止日期<>'' 
   if charindex(' ',@p截止日期)=0 select @p截止日期=@p截止日期+' 23:59:59'

select a.系统编号,姓名,性别,年龄,单位名称,试管编号,体检表编号,convert(varchar(10),体检日期,120) as 体检日期,体检结论,isnull(复查体检表编号,'') as 复查体检表编号,isnull(复查系统编号,'') as 复查系统编号,体检状态=
	case when 体检状态=3 then '已下结论' when 体检状态=0 then '待病史录入' when 体检状态=1 then '待体检' else '未下结论' end
from 职业病体检_体检基本信息表 a inner join  职业病体检_体检人员基本信息表 b on a.系统编号=b.系统编号
where ((体检日期>=@p开始日期 or @p开始日期='')
		and  (体检日期<=@p截止日期 or @p截止日期='')
		and  (单位名称 like '%'+@p单位名称+'%' or @p单位名称='')
		and  (姓名 like '%'+@p姓名+'%' or @p姓名='')
        --and (试管编号=@p试管编号 or @p试管编号='')
	and (a.系统编号=@p系统编号 or @p系统编号='')
)	
or (体检状态=3 and isnull(复查体检表编号,'')<>'' and isnull(复查系统编号,'')='')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO











if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检结果数据库]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[职业病体检_体检结果数据库]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE        VIEW dbo.职业病体检_体检结果数据库
AS 
    SELECT a.系统编号,b.公民身份号码, 
        b.姓名, b.性别, b.出生日期, a.试管编号, 
        a.体检日期, a.体检状态, c.体检项目, 体检项目名称=d.名称, 
        c.体检结果, c.体检医师, 体检医师姓名=e.姓名, c.填写日期, d.属性,
        d.缺省值,d.枚举来源,d.体检大类,f.名称,d.标准值,d.单位,d.单价,c.单项结论
    FROM 职业病体检_体检基本信息表 a, 
        职业病体检_体检人员基本信息表 b, 
        职业病体检_体检结果信息表 c, 
        职业病体检_体检项目设置表 d,
        系统管理_员工基本信息表 e,系统管理_字典_字典内容表 f
    WHERE a.系统编号=c.系统编号 and 
	a.系统编号=b.系统编号 and
        c.体检项目= d.编码 and 
        c.体检医师*= e.编号 and d.体检大类=f.InnerID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检表模板体检结论数据库]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[职业病体检_体检表模板体检结论数据库]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-------------------------------------------------------------------------------------------------
create  VIEW dbo.职业病体检_体检表模板体检结论数据库
AS 
    Select a.体检表名称, a.体检结论, 体检结论名称 = b.名称
    From 职业病体检_体检表模板体检结论表 a, 
        系统管理_字典_字典内容表 b
    Where a.体检结论=b.InnerID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO








if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检表模板附加项目数据库]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[职业病体检_体检表模板附加项目数据库]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

create  VIEW dbo.职业病体检_体检表模板附加项目数据库
AS
SELECT dbo.职业病体检_体检表模板附加项目信息表.体检表名称, 
      dbo.职业病体检_体检表模板附加项目信息表.附加项目,
      dbo.职业病体检_体检表模板附加项目信息表.序号,
      dbo.职业病体检_体检表模板附加项目信息表.是否必录, 
      dbo.职业病体检_体检人员附加项目设置表.录入标题, 
      dbo.职业病体检_体检人员附加项目设置表.数据类型, 
      dbo.职业病体检_体检人员附加项目设置表.数据长度, 
      dbo.职业病体检_体检人员附加项目设置表.枚举值
FROM dbo.职业病体检_体检表模板附加项目信息表 INNER JOIN
      dbo.职业病体检_体检人员附加项目设置表 ON 
      dbo.职业病体检_体检表模板附加项目信息表.附加项目 = dbo.职业病体检_体检人员附加项目设置表.附加项目

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_查询统计视图]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[职业病体检_查询统计视图]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


create view 职业病体检_查询统计视图
as
select a.系统编号,b.收费批号,a.姓名,a.性别,b.体检表编号 as 体检表,b.体检类型 as 体检人员类型,
		b.体检类别,b.体检日期,c.工种,f.单位名称,f.片区,f.卫生种类,f.行业类别,b.复查体检表编号,
		b.复查系统编号,g.名称 as 科室名称,d.文字结论,d.结论日期,e.编号 as 医师编号,e.姓名 as 医师姓名,
		收费金额 = case
		when d.科室 = 16 then b.收费金额
		else null
		end
	
      from 职业病体检_体检人员基本信息表 a left join 职业病体检_体检基本信息表 b on a.系统编号=b.系统编号
					 left join 职业病体检_职业史表 c on b.系统编号=c.系统编号
					 left join 职业病体检_科室结论表 d on c.系统编号 = d.系统编号
					 left join 系统管理_员工基本信息表 e on d.医生编号 = e.编号
					 left join 单位档案_单位基本信息视图 f on a.单位申请编号=f.申请编号
					 left join 系统管理_字典_字典内容表 g on d.科室 = g.编号 and g.id = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sel返回医师信息]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sel返回医师信息]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure sel返回医师信息(@para姓名 varchar(16),@para编号 varchar(10))
as
begin
declare @str varchar(10)
set @str = rtrim(ltrim(@para编号))

declare @intID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')

if len(@str) = 0
select a.姓名,a.编号,c.名称 from 系统管理_员工基本信息表 a,职业病体检_用户科室权限表 b,系统管理_字典_字典内容表 c where a.编号=b.用户编号 and b.科室编号 = c.编号 and c.ID = @intID and a.姓名 = @para姓名 and 名称 not in('体检登记','业务设置','职业病史录入') order by c.编号;
else
select a.姓名,a.编号,c.名称 from 系统管理_员工基本信息表 a,职业病体检_用户科室权限表 b,系统管理_字典_字典内容表 c where a.编号=b.用户编号 and b.科室编号 = c.编号 and c.ID = @intID and a.编号 = @para编号 and 名称 not in('体检登记','业务设置','职业病史录入') order by c.编号;
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


--修改自动下结论存储过程
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

ALTER     PROCEDURE autoConclusion
@paraSysNo varchar(16),		--系统编号
@paraItem varchar(16),		--单项名称
@paraResult varchar(16),	--单项结果
@paraDoctor varchar(16),	--医师名字
@paraFillDate varchar(16),	--填写时间
@paraConclusion varchar(16),	--单项结论
@paraTableName varchar(40)	--项目所在科室
as
begin
  declare @paraItemID varchar(6)	--项目编号
  declare @paraFlag int	--返回标志
  declare @sqlstr nvarchar(4000)	--执行的sql语句
  declare @standard varchar(50)		--标准值
  
  select @paraItemID = 编码 from 职业病体检_体检项目设置表 where 名称 = @paraItem	--得到项目编号
  select @standard = 标准值 from 职业病体检_体检项目设置表 where 名称 = @paraItem and 编码 = @paraItemID	--得到标准值
  if '正常'=@paraResult
  begin
    set @paraConclusion = '合格'
  end
  if '异常'=@paraResult
  begin
    set @paraConclusion = '不合格'
  end

  if @standard<>''
  begin
	  if isnumeric(@paraResult)=1
	  begin
		if (convert(numeric,@paraResult)-convert(numeric,@standard))<>0
		begin
		  set @paraConclusion = '不合格'
		end
		else
		begin
		  set @paraConclusion = '合格'
		end
	  end
  end
  set @sqlstr = N'select @paraFlag=count(*) from '+ @paraTableName +' where 系统编号 = '''+@paraSysNo+''' and 体检项目 = '''+@paraItemID+''';'
  execute Sp_executeSql @sqlstr, N'@paraFlag int out',@paraFlag out
  if @paraFlag > 0
  begin
    set @sqlstr = 'update '+@paraTableName+' set 体检结果='''+@paraResult+''',体检医师='''+@paraDoctor+''',填写时间='''+@paraFillDate+''',单项结论='''+@paraConclusion+''' where 系统编号='''+@paraSysNo+''' and 体检项目='''+@paraItemID+''';'
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


--2012-07-05 于登淼 
--创建存储过程，用于职业病史(受检者个人信息)界面查询使用
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_职业病史管理界面查询]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[职业病体检_职业病史管理界面查询]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE      proc dbo.职业病体检_职业病史管理界面查询 
	@p开始日期 VARCHAR ( 10 )='2004-01-01',
	@p截止日期 VARCHAR ( 10 )='2005-12-01',
	@p体检表名称 VARCHAR ( 50 )='',
    @p单位名称 VARCHAR ( 100 )='',
    @p姓名 VARCHAR ( 20 )='',
	@p体检单号 varchar(20)='',
    @p试管编号 VARCHAR ( 20 )='',
	@p系统编号 varchar(40)=''
AS
set nocount on
if @p截止日期<>'' 
   if charindex(' ',@p截止日期)=0 select @p截止日期=@p截止日期+' 23:59:59'

select a.系统编号,姓名,性别,年龄,单位名称,试管编号,体检表编号,convert(varchar(10),体检日期,120) as 体检日期,体检结论,isnull(复查体检表编号,'') as 复查体检表编号,isnull(复查系统编号,'') as 复查系统编号,体检状态=
	case when 体检状态=3 then '已下结论' when 体检状态=0 then '待病史录入' when 体检状态=1 then '待体检' else '未下结论' end
from 职业病体检_体检基本信息表 a inner join  职业病体检_体检人员基本信息表 b on a.系统编号=b.系统编号
where ((体检日期>=@p开始日期 or @p开始日期='')
		and  (体检日期<=@p截止日期 or @p截止日期='')
		and  (单位名称 like '%'+@p单位名称+'%' or @p单位名称='')
		and  (姓名 like '%'+@p姓名+'%' or @p姓名='')
        --and (试管编号=@p试管编号 or @p试管编号='')
	and (a.系统编号=@p系统编号 or @p系统编号='')
)	
or (体检状态=3 and isnull(复查体检表编号,'')<>'' and isnull(复查系统编号,'')='')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


--2012-07-05 于登淼
--创建存储过程，生成职业病体检特定规则系统编号
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_生成编号流水号]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[职业病体检_生成编号流水号]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



------------------------------------------------------------------------------------------
CREATE   PROCEDURE dbo.职业病体检_生成编号流水号 
        	@p编号名称 varchar(20)
AS
/****************************************************************************
    根据“体检管理_编号的最大流水号表”记录的指定编号名称的当前已使用的最大流水号，
生成新最大流水号返回，并把新流水号记入最大流水号表.
******************************************************************************/
    set nocount on 
    declare @l编号 int
    declare @l长度 int
    declare @l返回值 varchar(10)
    declare @l日期类型 int

    --select @l日期类型 = count(*) from  [体检管理_编号生成规则表] where 编号名称 = @p编号名称 and  组成 in ('yy','mm','dd')  
    --set @l日期类型 =isnull(@l日期类型 ,1) 
    --if @l日期类型 = 1 
        begin    
        if exists(select 设置项目 from [职业病体检_业务设置信息表] 
                      where 设置项目 = '编号最大流水号' and datepart(yyyy,说明) = datepart(yyyy,getdate()))
            begin
            select @l编号=设置值 from [职业病体检_业务设置信息表] 
                      where 设置项目 = '编号最大流水号' and datepart(yyyy,说明) = datepart(yyyy,getdate())
            set @l编号 = @l编号 + 1
            update [职业病体检_业务设置信息表] set 设置值 = @l编号 
                  where  设置项目 = '编号最大流水号' and datepart(yyyy, 说明) = datepart(yyyy,getdate() )
            end
        else
            begin
            set @l编号 = 1
            insert into [职业病体检_业务设置信息表](设置项目,设置值,说明) values('编号最大流水号','1',getdate())
            end
        end
    
    set @l长度=7
    set @l返回值 =  convert(varchar(10),@l编号)
    set @l返回值 = replicate('0',@l长度 - len(@l返回值)) + @l返回值
    select @l返回值


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



--2012-07-05 于登淼
--创建存储过程，退回职业病体检模块的系统编号
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_退回编号流水号]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[职业病体检_退回编号流水号]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE  PROCEDURE dbo.职业病体检_退回编号流水号
	@p编号名称 varchar(20)='系统编号',
	@p编号 varchar(20)='102000001'
AS
/****************************************************************************
功能：根据“体检管理_编号的最大流水号表”记录的指定编号名称的当前已使用的最大流水号，
     把最大流水号减一。
******************************************************************************/
set nocount on 
declare @l编号 int          --当前最大流水号。
declare @l长度 int          --流水号长度。
declare @l日期类型 int      --0 没有生成规则，1 编号中只包含年份，2 编号中包含年份+月份，3 编号中有年月日。
declare @l流水号 varchar(10)--参数@p编号中的流水号。
declare @lintInsert as int

 
BEGIN
    select @lintInsert =0

    --获取当前编号的流水号长度。
    set @l长度=isnull(@l长度,7)

    --获取编号中的流水号。
    set @l流水号=right(@p编号,@l长度)
    begin
	--编号中包含年份，则编号必须每年从头开始编。
        set @l编号 =cast(@l流水号 as int) - 1
		if not exists(select * from [职业病体检_业务设置信息表]
			    	    where  设置项目='编号最大流水号' and datepart(yyyy, 说明) = datepart(yyyy,getdate()) 
						and 设置值=cast(@l流水号 as int))
			select @lintInsert=1
		else
	        update [职业病体检_业务设置信息表] set 设置值 = @l编号 
    	    where  设置项目='编号最大流水号' and datepart(yyyy, 说明) = datepart(yyyy,getdate()) 
				and 设置值=cast(@l流水号 as int)
    end		
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




--2012-07-05 于登淼
--创建存储过程，用于职业病体检管理界面查询功能
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检管理界面查询]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[职业病体检_体检管理界面查询]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE          proc dbo.职业病体检_体检管理界面查询 
	--2012-07-06 于登淼 ↓
	--代码中更改开始日期与截止日期精确到分钟和秒，所以这里更改下格式 
	--@p开始日期 VARCHAR ( 10 )='2004-01-01',
	--@p截止日期 VARCHAR ( 10 )='2005-12-01',
	@p开始日期 datetime ='2004-01-01',
	@p截止日期 datetime ='2005-12-01',
	--2012-07-06 于登淼 ↑
	@p体检表名称 VARCHAR ( 50 )='',
    @p单位名称 VARCHAR ( 100 )='',
    @p姓名 VARCHAR ( 20 )='',
	@p体检单号 varchar(20)='',
    @p试管编号 VARCHAR ( 20 )='',
	@p系统编号 varchar(40)=''
AS
set nocount on
if @p截止日期<>'' 
   if charindex(' ',@p截止日期)=0 select @p截止日期=@p截止日期+' 23:59:59'

select a.系统编号,姓名,性别,年龄,单位名称,试管编号,体检表编号,
	convert(varchar(10),体检日期,120) as 体检日期,
	体检结论,isnull(复查体检表编号,'') as 复查体检表编号,
	isnull(复查系统编号,'') as 复查系统编号,
	体检状态=case 
		--2012-06-15 于登淼 ↓
		--增加多种体检状态
		--when 体检状态=3 then '已下结论' 
		--else '未下结论' 
		--end
		when 体检状态=0 then '未校核'
		when 体检状态=1 then '未打清单'
		when 体检状态=2 then '未录入受检者个人信息'
		when 体检状态=3 then '体检中'
		when 体检状态=4 then '未下结论'
		when 体检状态=5 then '已下结论'
		when 体检状态=6 then '已复核'
		when 体检状态=7 then '已发报告'
		when 体检状态=8 then '待复查'
		end
		--2012-06-15 于登淼 ↑
from 职业病体检_体检基本信息表 a inner join  职业病体检_体检人员基本信息表 b on a.系统编号=b.系统编号
where ((体检日期>=@p开始日期 or @p开始日期='')
		and  (体检日期<=@p截止日期 or @p截止日期='')
		and  (单位名称 like '%'+@p单位名称+'%' or @p单位名称='')
		and  (姓名 like '%'+@p姓名+'%' or @p姓名='')
        and (试管编号=@p试管编号 or @p试管编号='')
	and (a.系统编号=@p系统编号 or @p系统编号='')
        and (a.体检表编号=@p体检表名称 or @p体检表名称='')
)	
or (体检状态=3 and isnull(复查体检表编号,'')<>'' and isnull(复查系统编号,'')='')


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
功能：根据传入的“科室名称“、”编号”查询该科室的结果信息。
******************************************************************************/

ALTER    proc [dbo].[职业病体检_结果信息]
	@p科室 varchar(20) = '五官科',
	@p编号 varchar(20) = '01'

as
set nocount on
declare @lsql varchar(1000)

	select @lsql='select distinct a.系统编号,a.姓名,a.性别,a.年龄,a.体检类型,a.单位名称,convert(varchar(10),b.填写时间,2) 填写时间 
	  from 职业病体检_体检基本数据库 a, 职业病体检_结果信息_'+@p科室+' b 
            where 1=1 and a.系统编号=b.系统编号 and (a.体检状态=''2'' or a.体检状态=''3'' or a.体检状态=''4'')
			   and  (substring(a.各科体检状态,'+@p编号+',1)=''1'' or substring(a.各科体检状态,'+@p编号+',1)=''2'')'

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
select @str = isnull(@str,' ') +名称+',' from 职业病体检_体检结果视图 a,职业病体检_体检项目设置表 b 
	where a.体检项目 = b.编码 and 系统编号 = @s and a.单项结论 = '不合格'
return stuff(@str,1,1,'')
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



--修改职业病体检_体检基本信息表
ALTER TABLE 职业病体检_体检基本信息表 ADD 复查状态 varchar(2) null


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





ALTER               proc dbo.职业病体检_体检管理界面查询
	--2012-07-06 于登淼 ↓
	--代码中更改开始日期与截止日期精确到分钟和秒，所以这里更改下格式 
	--@p开始日期 VARCHAR ( 10 )='2004-01-01',
	--@p截止日期 VARCHAR ( 10 )='2005-12-01',
	@p开始日期 datetime ='2004-01-01',
	@p截止日期 datetime ='2012-11-01',
	--2012-07-06 于登淼 ↑
	@p体检表名称 VARCHAR ( 50 )='',
    @p单位名称 VARCHAR ( 100 )='',
    @p姓名 VARCHAR ( 20 )='',
	@p体检单号 varchar(20)='',
    @p试管编号 VARCHAR ( 20 )='',
	@p系统编号 varchar(40)=''
AS
set nocount on
if @p截止日期<>'' 
   if charindex(' ',@p截止日期)=0 select @p截止日期=@p截止日期+' 23:59:59'

select a.系统编号,姓名,性别,年龄,单位名称,试管编号,体检表编号 as 体检表名,
	convert(varchar(10),体检日期,120) as 体检日期,
	体检结论,isnull(复查体检表编号,'') as 复查体检表编号,
	isnull(复查系统编号,'') as 复查系统编号,
	体检状态=case 
		--2012-06-15 于登淼 ↓
		--增加多种体检状态
		--when 体检状态=3 then '已下结论' 
		--else '未下结论' 
		--end
		when 体检状态=0 then '未校核'
		when 体检状态=1 then '未打清单'
		when 体检状态=2 then '未录入受检者个人信息'
		when 体检状态=3 then '体检中'
		when 体检状态=6 then '待复核'
		when 体检状态=4 then '未下结论'
		--when 体检状态=5 then '已下结论'
		when 体检状态=6 then '已复核'
		when 体检状态=7 then '已发报告'
		--when a.复查系统编号 is not null then '待复查'
		when 体检状态=8 then '待复核'
		--when 体检状态=8 then '待复查'
		end
		--2012-06-15 于登淼 ↑
from 职业病体检_体检基本信息表 a inner join  职业病体检_体检人员基本信息表 b on a.系统编号=b.系统编号
where ((体检日期>=@p开始日期 or @p开始日期='')
		and  (体检日期<=@p截止日期 or @p截止日期='')
		and  (单位名称 like '%'+@p单位名称+'%' or @p单位名称='')
		and  (姓名 like '%'+@p姓名+'%' or @p姓名='')
        and (试管编号=@p试管编号 or @p试管编号='')
	and (a.系统编号=@p系统编号 or @p系统编号='')
        and (a.体检表编号=@p体检表名称 or @p体检表名称='')
)	
or (体检状态=3 and isnull(复查体检表编号,'')<>'' and isnull(复查系统编号,'')='')



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


--exec  健康证管理_导入已体检完毕人员信息
ALTER      PROCEDURE 健康证管理_导入已体检完毕人员信息 
AS
set nocount on

insert into 健康证管理_办证申请信息表(系统编号,体检号,姓名,性别,年龄,
	种类,申请编号,单位名称,体检日期,体检结论,检出病种,培训结论,
	处置,培训日期,体检系统编号,状态)
select newid() as 系统编号, a.体检单号, b.姓名, b.性别,datediff(yy,出生日期,getdate()) as 年龄,
	d.名称 as 种类,	b.单位申请编号,b.单位名称, a.体检日期,
	case when a.体检结论='正常' then '合格' else '不合格' end as 体检结论,
	case when a.体检结论='正常' then '无从业禁忌症' else 体检结论 end as 检出病种,
	培训结论='合格',处置=case when a.体检结论='正常' then '发健康证' else '调离' end,a.体检日期 as 培训日期, a.系统编号 as 体检系统编号,'未打印'
FROM 体检管理_体检基本信息表 a, 
	体检管理_体检人员基本信息表 b,
	体检管理_体检访问标志表 c,
	体检管理_卫生种类数据库 d
WHERE a.健康档案编号=b.健康档案编号 
and a.系统编号= c.系统编号 and b.卫生种类=d.编号
and a.体检状态=3 and isnull(复查体检表名,'')='' and c.健康证='1'
and a.系统编号 not in(select distinct isnull(体检系统编号,'') from 健康证管理_办证申请信息表)
--2008-08-01
update 健康证管理_办证申请信息表 set 处置='发健康证',检出病种='无从业禁忌症',体检结论='正常'
where 种类='公共卫生' and 检出病种 like '%乙肝%'

/*update 健康证管理_办证申请信息表
set 体检号=a.体检单号,姓名=b.姓名,性别=b.性别,年龄=datediff(yy,b.出生日期,getdate()),
种类=case b.卫生种类 when '1' then '食品卫生' when '2' then '公共卫生' end,
申请编号=b.单位申请编号,单位名称=b.单位名称,体检日期=a.体检日期,
体检结论=case when a.体检结论='正常' then '合格' else '不合格' end,
检出病种=case when a.体检结论='正常' then '无从业禁忌症' else a.体检结论 end
FROM 体检管理_体检基本信息表 a, 
	体检管理_体检人员基本信息表 b,
	体检管理_体检访问标志表 c,
	健康证管理_办证申请信息表 d
WHERE a.健康档案编号=b.健康档案编号 
and a.系统编号= c.系统编号
and a.体检状态=3 and isnull(复查体检表名,'')='' and c.健康证='1'
and a.系统编号=d.体检系统编号
*/
--导入图片数据。
--2006-1-12：不导图片，就用体检的图片。
--insert into 系统管理_系统图片管理表(图片编号,子系统名,图片)
--select 系统编号,'健康证管理',图片
--from 健康证管理_办证申请信息表 a,系统管理_系统图片管理表 b
--where a.体检系统编号=b.图片编号 and b.子系统名='体检管理'
--and  not exists(
--	select 图片编号 from 系统管理_系统图片管理表 
--	where 图片编号=a.系统编号 and 子系统名='健康证管理'
--	)

--设置已更新的标志。
update 体检管理_体检访问标志表 set 健康证='2' where 健康证='1'



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


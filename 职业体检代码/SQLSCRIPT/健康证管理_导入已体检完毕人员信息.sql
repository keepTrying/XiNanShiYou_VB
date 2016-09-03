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
	处置,培训日期,体检系统编号,状态,二代身份证编号)
select newid() as 系统编号, a.体检单号, b.姓名, b.性别,datediff(yy,出生日期,getdate()) as 年龄,
	d.名称 as 种类,	b.单位申请编号,b.单位名称, a.体检日期,
	case when a.体检结论='正常' then '合格' else '不合格' end as 体检结论,
	case when a.体检结论='正常' then '无从业禁忌症' else 体检结论 end as 检出病种,
	培训结论='合格',处置=case when a.体检结论='正常' then '发健康证' else '调离' end,
	a.体检日期 as 培训日期, a.系统编号 as 体检系统编号,'未打印',e.项目值
FROM 体检管理_体检基本信息表 a, 
	体检管理_体检人员基本信息表 b,
	体检管理_体检访问标志表 c,
	体检管理_卫生种类数据库 d,
	体检管理_体检附加信息表 e
WHERE a.健康档案编号=b.健康档案编号 
and a.系统编号= c.系统编号 and b.卫生种类=d.编号
and a.体检状态=3 and isnull(复查体检表名,'')='' and c.健康证='1'
and a.系统编号 not in(select distinct isnull(体检系统编号,'') from 健康证管理_办证申请信息表)
and a.系统编号=e.系统编号 and e.附加项目='身份证号'

update 健康证管理_办证申请信息表 set 处置='发健康证',检出病种='无从业禁忌症',体检结论='正常'
where 种类='公共卫生' and 检出病种 like '%乙肝%'



--设置已更新的标志。
update 体检管理_体检访问标志表 set 健康证='2' where 健康证='1'



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--select top  100 * from dbo.健康证管理_办证申请信息表  order by 二代身份证编号 desc
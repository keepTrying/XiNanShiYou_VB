create table 收费管理_票据交账记录表 (
	票据号 varchar(20) not null,
	已交账 tinyint null,
	交账日期 smalldatetime null,
	已收款 tinyint null,
	收款日期 smalldatetime null
)
go

create view 收费管理_票据交账记录视图 as
select 收据号,收费人,收费员,金额,convert(varchar(10),交费日期,120) 收费日期,已交账,convert(varchar(10),交账日期,120) 交账日期,已收款,convert(varchar(10),收款日期,120) 收款日期
from 收费管理_票据交账记录表 a,(select 收据号,收费人,姓名 收费员,交费日期,sum(金额) 金额 from 收费管理_费用信息表 a,系统管理_员工基本信息表 b where a.收费人=b.编号 and 收费状态=1 group by 收据号,收费人,姓名,交费日期) b
where a.票据号=*b.收据号
go

create proc 收费管理_获取票据交账信息 
	@pWhere1 varchar(100),
	@pWhere2 varchar(100)
as

set nocount on
declare @lstrSql varchar(1000)
select @lstrSql='select * into #temp from 收费管理_票据交账记录视图 where ' + @pWhere1

select @lstrSql=@lstrSql + ' select * from #temp where ' + @pWhere2
exec (@lstrSql)

go

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


ALTER   PROC 收费管理_返回收费信息 
(	
	@p收费批号 VARCHAR(40)='',
	@p收据号   VARCHAR(40)='',
	@p交费人   VARCHAR(100)='',
	@p交费单位 VARCHAR(100)='',	
	@p开始时间 VARCHAR(40)='2005/02/08',
	@p结束时间 VARCHAR(40)='2007/12/31',
--	@p业务科室 varchar(40)='',--名称
	@p对应业务  varchar(20)='', --一般、门诊
	@p收款人 varchar(10)='' --收费人编号
)
AS
SET NOCOUNT ON


--构造构临时表
CREATE  TABLE #TEMP_返回值表 (
	收费批号 VARCHAR(14) COLLATE database_default,
	票据号 VARCHAR(30) COLLATE database_default,
        交费人 VARCHAR(20) COLLATE database_default,
	交费单位 VARCHAR(80) COLLATE database_default,
	金额 dec(12,2),
        交费日期 varchar(10) COLLATE database_default,
	收费人 VARCHAR(20) COLLATE database_default,
	主管科室名称 VARCHAR(200) COLLATE database_default,
	打折比率 NUMERIC(5,2),
	交费方式 varchar(20) COLLATE database_default,
	收费编号 VARCHAR(14) COLLATE database_default,
	开户银行 VARCHAR(200) COLLATE database_default,
     	收费状态 CHAR(1) COLLATE database_default,
	标识 INT--标示：1收费；2退费
)
--未收费记录。
INSERT INTO #TEMP_返回值表 (收费批号,收费编号,交费单位,主管科室名称,交费人,
                           交费日期,收费人,打折比率,收费状态,交费方式,标识,开户银行,金额)
SELECT  收费批号,收费编号,交费单位名称,主管科室名称,交费人,
        convert(varchar(10),交费日期,120),收费人姓名,打折比率,收费状态,交费方式名称,收费状态,开户银行,sum(金额)        
FROM 收费管理_打印费用信息
WHERE  (@p收费批号='' OR 收费批号=@p收费批号) AND
	   (@p收据号='' OR 收据号=@p收据号	 ) AND
	   (@p交费人='' OR 交费人=@p交费人) AND
       (@p交费单位='' OR 交费单位名称=@p交费单位) AND
--	(主管科室名称=@p业务科室 or @p业务科室='') and
	(对应业务=@p对应业务 or @p对应业务='') and
--	(收费人=@p收款人 or @p收款人 ='') and
	   (收费状态 =0 ) 	
group by 收费批号,收费编号,交费单位名称,主管科室名称,交费人,
        convert(varchar(10),交费日期,120),收费人姓名,
        打折比率,收费状态,交费方式名称,开户银行

--收费记录（和报废记录）
INSERT INTO #TEMP_返回值表 (收费批号,收费编号,交费单位,主管科室名称,交费人,
                           交费日期,收费人,打折比率,收费状态,交费方式,标识,开户银行,金额)
SELECT  收费批号,收费编号,交费单位名称,主管科室名称,交费人,
        convert(varchar(10),交费日期,120),收费人姓名,
        打折比率,收费状态,交费方式名称,收费状态,开户银行,sum(金额)
FROM 收费管理_打印费用信息
WHERE  (@p收费批号='' OR 收费批号=@p收费批号) AND
	   (@p收据号='' OR 收据号=@p收据号	 ) AND
	   (@p交费人='' OR 交费人=@p交费人) AND
       (@p交费单位='' OR 交费单位名称=@p交费单位) AND
	   (@p开始时间='' OR @p结束时间='' OR ( 交费日期 between @p开始时间 AND @p结束时间 ))	AND
--	(主管科室名称=@p业务科室 or @p业务科室='') and
	(对应业务=@p对应业务 or @p对应业务='') and
	(收费人=@p收款人 or @p收款人 ='') and
	   (收费状态 =1 OR 收费状态 =3) 	
group by 收费批号,收费编号,交费单位名称,主管科室名称,交费人,
        convert(varchar(10),交费日期,120),收费人姓名,
        打折比率,收费状态,交费方式名称,开户银行

--退费记录
INSERT INTO #TEMP_返回值表 (收费批号,收费编号,交费单位,交费人,
                           交费日期,收费人,打折比率,收费状态,交费方式,标识,开户银行,金额)
SELECT   收费批号,收费编号,交费单位名称,交费人,
         convert(varchar(10),退费日期,111),收费人姓名,
         打折比率,收费状态,交费方式名称,收费状态,开户银行,sum(0-金额)
FROM dbo.收费管理_打印费用信息
WHERE  (@p收费批号='' OR 收费批号=@p收费批号) AND
	   (@p收据号='' OR 收据号=@p收据号	 ) AND
	   (@p交费人='' OR 交费人=@p交费人) AND
       (@p交费单位='' OR 交费单位名称=@p交费单位) AND
	   (@p开始时间='' OR @p结束时间='' OR ( 退费日期 between @p开始时间 AND @p结束时间 ))	AND
--	(主管科室名称=@p业务科室 or @p业务科室='') and
	(收费人=@p收款人  or 退费人=@p收款人 or @p收款人 ='') and
	(对应业务=@p对应业务 or @p对应业务='') and
	   (收费状态 =2) 
group by 收费批号,收费编号,交费单位名称,交费人,
         convert(varchar(10),退费日期,111),收费人姓名,
         打折比率,收费状态,交费方式名称,开户银行

update #TEMP_返回值表 set 票据号=isnull(a.收据号 ,'')
from (select 收费批号,min(收据号) as 收据号 from 收费管理_打印费用信息 group by 收费批号) a,#TEMP_返回值表 b
where a.收费批号=b.收费批号


update #TEMP_返回值表 set 票据号=票据号+ case when 票据号<>isnull(a.收据号,'') then '～'+a.收据号 else '' end
from (select 收费批号,max(收据号) as 收据号 from 收费管理_打印费用信息 group by 收费批号) a,
    #TEMP_返回值表 b
where a.收费批号=b.收费批号

--返回结果集
SELECT * FROM #TEMP_返回值表 ORDER BY 票据号 DESC,金额 DESC
DROP TABLE #TEMP_返回值表

--exec 收费管理_返回收费信息 '','','','','2008-07-16','2008-07-17','','',''


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


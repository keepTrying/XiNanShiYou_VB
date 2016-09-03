SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


ALTER  PROC 收费管理_返回费用次数
(	
	@p收费批号 VARCHAR(40)='',
	@p收据号   VARCHAR(40)='',
	@p交费人   VARCHAR(100)='',
	@p交费单位 VARCHAR(100)='',	
	@p开始时间 VARCHAR(40)='',
	@p结束时间 VARCHAR(40)='',
	@p对应业务 varchar(40)='',
	@p收款人 varchar(10)='' --收费人编号
)
AS
SET NOCOUNT ON
DECLARE 
	@l总次数		INT,
	@l退费次数   INT

SET @l总次数=0
SET @l退费次数=0

IF @p收费批号='' SET @p收费批号=NULL
IF @p收据号='' SET @p收据号=NULL
IF @p交费人 ='' SET @p交费人=NULL
IF @p交费单位='' SET @p交费单位=NULL
IF @p开始时间='' SET @p开始时间=NULL
IF @p结束时间='' SET @p结束时间=NULL


--构造临时表
CREATE TABLE #TEMP_返回值表 (项目 VARCHAR(100) COLLATE database_default,次数 INT)


SELECT  @l总次数=COUNT(DISTINCT(收据号))
FROM dbo.收费管理_打印费用信息
WHERE  (ISNULL(@p收费批号,'1')='1' OR 收费批号=@p收费批号) AND
	   (ISNULL(@p收据号,'1')='1' OR 收据号=@p收据号	 ) AND
	   (ISNULL(@p交费人,'1')='1' OR 交费人=@p交费人) AND
       (ISNULL(@p交费单位,'1')='1' OR 交费单位名称=@p交费单位) AND
	   ((ISNULL(@p开始时间,'1')='1') OR (ISNULL(@p结束时间,'1')='1') OR ( 交费日期>=@p开始时间 AND 交费日期 <=@p结束时间 ))	AND
	(对应业务=@p对应业务 or @p对应业务='') and
	(收费人=@p收款人 or @p收款人 ='') and
	   (收费状态 =1 OR 收费状态 =2) 	
--GROUP BY 收据号


--获取退费信息

SELECT @l退费次数=COUNT(DISTINCT(收据号))
FROM dbo.收费管理_打印费用信息
WHERE  (ISNULL(@p收费批号,'1')='1' OR 收费批号=@p收费批号) AND
	   (ISNULL(@p收据号,'1')='1' OR 收据号=@p收据号	 ) AND	
	   (ISNULL(@p交费人,'1')='1' OR 交费人=@p交费人) AND
       (ISNULL(@p交费单位,'1')='1' OR 交费单位名称=@p交费单位) AND
	   ((ISNULL(@p开始时间,'1')='1') OR (ISNULL(@p结束时间,'1')='1') OR ( 退费日期>=@p开始时间+' 00:00:01' AND 退费日期 <=@p结束时间+' 23:59:59' ))	AND
	(对应业务=@p对应业务 or @p对应业务='') and
	(收费人=@p收款人  or 退费人=@p收款人 or @p收款人 ='') and
	   (收费状态 =2) 
--GROUP BY 收据号

INSERT INTO #TEMP_返回值表 (项目,次数) VALUES ('总次数',@l总次数)
INSERT INTO #TEMP_返回值表 (项目,次数) VALUES ('退费次数',@l退费次数)

SELECT * FROM #TEMP_返回值表
--返回结果集
DROP TABLE #TEMP_返回值表


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


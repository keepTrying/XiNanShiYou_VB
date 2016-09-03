SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



ALTER   PROCEDURE 体检管理_各类体检人员的列表
  @p开始日期 VARCHAR ( 10 ) = NULL,
  @p结束日期 VARCHAR ( 10 ) = NULL,
  @p片区编号 VARCHAR ( 20 ) = NULL,
  @p行业类别编号 VARCHAR ( 20 ) = NULL,
  @p体检表名称 VARCHAR ( 50 ) = NULL
AS
--和存储过程 体检管理_已登记未下结论的体检人员列表 有些相似
--但增加了多个参数，并少了 体检结论 的判断条件
SET NOCOUNT ON

--将'','*',转换为NULL,以利标准和规范
IF @p开始日期 = '' OR @p开始日期 = '*' SELECT @p开始日期 = NULL
IF @p结束日期 = '' OR @p结束日期 = '*' SELECT @p结束日期 = NULL
IF @p体检表名称 = '' OR @p体检表名称 = '*' SELECT @p体检表名称 = NULL
IF @p片区编号 = '' OR @p片区编号 = '*' SELECT @p片区编号 = NULL
IF @p行业类别编号 = '' OR @p行业类别编号 = '*' SELECT @p行业类别编号 = NULL

SELECT a.系统编号,b.姓名,b.性别,DATEDIFF(year,b.出生日期,GETDATE()) AS 年龄,
       b.单位名称 AS 体检单位,a.体检表名称,a.体检日期,a.体检结论,a.诊断和处理意见,
       dbo.系统管理_获取字典名称(b.片区,'片区街道字典') AS 片区,
       dbo.系统管理_获取字典名称(b.行业类别,'行业属性字典') AS 行业类别
  FROM 体检管理_体检基本信息表 a,
       体检管理_体检人员基本信息表 b
 WHERE a.健康档案编号 = b.健康档案编号 AND
       (@p开始日期 IS NULL OR a.体检日期 >= @p开始日期) AND
       (@p结束日期 IS NULL OR a.体检日期 <= @p结束日期) AND
       (@p体检表名称 IS NULL OR @p体检表名称 = a.体检表名称) AND
       (@p片区编号 IS NULL OR b.片区 = @p片区编号) AND
       (@p行业类别编号 IS NULL OR b.行业类别 = @p行业类别编号) AND
	体检结论<>''

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
--exec 收费管理_产生收据号 '595'
ALTER    PROCEDURE dbo.收费管理_产生收据号 
	@userNo varchar(10)
AS
set nocount on
DECLARE @ls流水号 varchar(20)
declare @name varchar(20)

select @name='收费管理' + @userNo
--调用系统管理的存储过程
EXEC 系统管理_返回编号流水号 @name, '收据号',@ls流水号 output

--返回流水号。
SELECT  @ls流水号


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


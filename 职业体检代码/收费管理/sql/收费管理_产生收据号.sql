SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
--exec �շѹ���_�����վݺ� '595'
ALTER    PROCEDURE dbo.�շѹ���_�����վݺ� 
	@userNo varchar(10)
AS
set nocount on
DECLARE @ls��ˮ�� varchar(20)
declare @name varchar(20)

select @name='�շѹ���' + @userNo
--����ϵͳ����Ĵ洢����
EXEC ϵͳ����_���ر����ˮ�� @name, '�վݺ�',@ls��ˮ�� output

--������ˮ�š�
SELECT  @ls��ˮ��


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


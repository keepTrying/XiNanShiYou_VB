SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


ALTER  PROC �շѹ���_���ط��ô���
(	
	@p�շ����� VARCHAR(40)='',
	@p�վݺ�   VARCHAR(40)='',
	@p������   VARCHAR(100)='',
	@p���ѵ�λ VARCHAR(100)='',	
	@p��ʼʱ�� VARCHAR(40)='',
	@p����ʱ�� VARCHAR(40)='',
	@p��Ӧҵ�� varchar(40)='',
	@p�տ��� varchar(10)='' --�շ��˱��
)
AS
SET NOCOUNT ON
DECLARE 
	@l�ܴ���		INT,
	@l�˷Ѵ���   INT

SET @l�ܴ���=0
SET @l�˷Ѵ���=0

IF @p�շ�����='' SET @p�շ�����=NULL
IF @p�վݺ�='' SET @p�վݺ�=NULL
IF @p������ ='' SET @p������=NULL
IF @p���ѵ�λ='' SET @p���ѵ�λ=NULL
IF @p��ʼʱ��='' SET @p��ʼʱ��=NULL
IF @p����ʱ��='' SET @p����ʱ��=NULL


--������ʱ��
CREATE TABLE #TEMP_����ֵ�� (��Ŀ VARCHAR(100) COLLATE database_default,���� INT)


SELECT  @l�ܴ���=COUNT(DISTINCT(�վݺ�))
FROM dbo.�շѹ���_��ӡ������Ϣ
WHERE  (ISNULL(@p�շ�����,'1')='1' OR �շ�����=@p�շ�����) AND
	   (ISNULL(@p�վݺ�,'1')='1' OR �վݺ�=@p�վݺ�	 ) AND
	   (ISNULL(@p������,'1')='1' OR ������=@p������) AND
       (ISNULL(@p���ѵ�λ,'1')='1' OR ���ѵ�λ����=@p���ѵ�λ) AND
	   ((ISNULL(@p��ʼʱ��,'1')='1') OR (ISNULL(@p����ʱ��,'1')='1') OR ( ��������>=@p��ʼʱ�� AND �������� <=@p����ʱ�� ))	AND
	(��Ӧҵ��=@p��Ӧҵ�� or @p��Ӧҵ��='') and
	(�շ���=@p�տ��� or @p�տ��� ='') and
	   (�շ�״̬ =1 OR �շ�״̬ =2) 	
--GROUP BY �վݺ�


--��ȡ�˷���Ϣ

SELECT @l�˷Ѵ���=COUNT(DISTINCT(�վݺ�))
FROM dbo.�շѹ���_��ӡ������Ϣ
WHERE  (ISNULL(@p�շ�����,'1')='1' OR �շ�����=@p�շ�����) AND
	   (ISNULL(@p�վݺ�,'1')='1' OR �վݺ�=@p�վݺ�	 ) AND	
	   (ISNULL(@p������,'1')='1' OR ������=@p������) AND
       (ISNULL(@p���ѵ�λ,'1')='1' OR ���ѵ�λ����=@p���ѵ�λ) AND
	   ((ISNULL(@p��ʼʱ��,'1')='1') OR (ISNULL(@p����ʱ��,'1')='1') OR ( �˷�����>=@p��ʼʱ��+' 00:00:01' AND �˷����� <=@p����ʱ��+' 23:59:59' ))	AND
	(��Ӧҵ��=@p��Ӧҵ�� or @p��Ӧҵ��='') and
	(�շ���=@p�տ���  or �˷���=@p�տ��� or @p�տ��� ='') and
	   (�շ�״̬ =2) 
--GROUP BY �վݺ�

INSERT INTO #TEMP_����ֵ�� (��Ŀ,����) VALUES ('�ܴ���',@l�ܴ���)
INSERT INTO #TEMP_����ֵ�� (��Ŀ,����) VALUES ('�˷Ѵ���',@l�˷Ѵ���)

SELECT * FROM #TEMP_����ֵ��
--���ؽ����
DROP TABLE #TEMP_����ֵ��


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


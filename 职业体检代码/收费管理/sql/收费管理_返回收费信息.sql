SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


ALTER   PROC �շѹ���_�����շ���Ϣ 
(	
	@p�շ����� VARCHAR(40)='',
	@p�վݺ�   VARCHAR(40)='',
	@p������   VARCHAR(100)='',
	@p���ѵ�λ VARCHAR(100)='',	
	@p��ʼʱ�� VARCHAR(40)='2005/02/08',
	@p����ʱ�� VARCHAR(40)='2007/12/31',
--	@pҵ����� varchar(40)='',--����
	@p��Ӧҵ��  varchar(20)='', --һ�㡢����
	@p�տ��� varchar(10)='' --�շ��˱��
)
AS
SET NOCOUNT ON


--���칹��ʱ��
CREATE  TABLE #TEMP_����ֵ�� (
	�շ����� VARCHAR(14) COLLATE database_default,
	Ʊ�ݺ� VARCHAR(30) COLLATE database_default,
        ������ VARCHAR(20) COLLATE database_default,
	���ѵ�λ VARCHAR(80) COLLATE database_default,
	��� dec(12,2),
        �������� varchar(10) COLLATE database_default,
	�շ��� VARCHAR(20) COLLATE database_default,
	���ܿ������� VARCHAR(200) COLLATE database_default,
	���۱��� NUMERIC(5,2),
	���ѷ�ʽ varchar(20) COLLATE database_default,
	�շѱ�� VARCHAR(14) COLLATE database_default,
	�������� VARCHAR(200) COLLATE database_default,
     	�շ�״̬ CHAR(1) COLLATE database_default,
	��ʶ INT--��ʾ��1�շѣ�2�˷�
)
--δ�շѼ�¼��
INSERT INTO #TEMP_����ֵ�� (�շ�����,�շѱ��,���ѵ�λ,���ܿ�������,������,
                           ��������,�շ���,���۱���,�շ�״̬,���ѷ�ʽ,��ʶ,��������,���)
SELECT  �շ�����,�շѱ��,���ѵ�λ����,���ܿ�������,������,
        convert(varchar(10),��������,120),�շ�������,���۱���,�շ�״̬,���ѷ�ʽ����,�շ�״̬,��������,sum(���)        
FROM �շѹ���_��ӡ������Ϣ
WHERE  (@p�շ�����='' OR �շ�����=@p�շ�����) AND
	   (@p�վݺ�='' OR �վݺ�=@p�վݺ�	 ) AND
	   (@p������='' OR ������=@p������) AND
       (@p���ѵ�λ='' OR ���ѵ�λ����=@p���ѵ�λ) AND
--	(���ܿ�������=@pҵ����� or @pҵ�����='') and
	(��Ӧҵ��=@p��Ӧҵ�� or @p��Ӧҵ��='') and
--	(�շ���=@p�տ��� or @p�տ��� ='') and
	   (�շ�״̬ =0 ) 	
group by �շ�����,�շѱ��,���ѵ�λ����,���ܿ�������,������,
        convert(varchar(10),��������,120),�շ�������,
        ���۱���,�շ�״̬,���ѷ�ʽ����,��������

--�շѼ�¼���ͱ��ϼ�¼��
INSERT INTO #TEMP_����ֵ�� (�շ�����,�շѱ��,���ѵ�λ,���ܿ�������,������,
                           ��������,�շ���,���۱���,�շ�״̬,���ѷ�ʽ,��ʶ,��������,���)
SELECT  �շ�����,�շѱ��,���ѵ�λ����,���ܿ�������,������,
        convert(varchar(10),��������,120),�շ�������,
        ���۱���,�շ�״̬,���ѷ�ʽ����,�շ�״̬,��������,sum(���)
FROM �շѹ���_��ӡ������Ϣ
WHERE  (@p�շ�����='' OR �շ�����=@p�շ�����) AND
	   (@p�վݺ�='' OR �վݺ�=@p�վݺ�	 ) AND
	   (@p������='' OR ������=@p������) AND
       (@p���ѵ�λ='' OR ���ѵ�λ����=@p���ѵ�λ) AND
	   (@p��ʼʱ��='' OR @p����ʱ��='' OR ( �������� between @p��ʼʱ�� AND @p����ʱ�� ))	AND
--	(���ܿ�������=@pҵ����� or @pҵ�����='') and
	(��Ӧҵ��=@p��Ӧҵ�� or @p��Ӧҵ��='') and
	(�շ���=@p�տ��� or @p�տ��� ='') and
	   (�շ�״̬ =1 OR �շ�״̬ =3) 	
group by �շ�����,�շѱ��,���ѵ�λ����,���ܿ�������,������,
        convert(varchar(10),��������,120),�շ�������,
        ���۱���,�շ�״̬,���ѷ�ʽ����,��������

--�˷Ѽ�¼
INSERT INTO #TEMP_����ֵ�� (�շ�����,�շѱ��,���ѵ�λ,������,
                           ��������,�շ���,���۱���,�շ�״̬,���ѷ�ʽ,��ʶ,��������,���)
SELECT   �շ�����,�շѱ��,���ѵ�λ����,������,
         convert(varchar(10),�˷�����,111),�շ�������,
         ���۱���,�շ�״̬,���ѷ�ʽ����,�շ�״̬,��������,sum(0-���)
FROM dbo.�շѹ���_��ӡ������Ϣ
WHERE  (@p�շ�����='' OR �շ�����=@p�շ�����) AND
	   (@p�վݺ�='' OR �վݺ�=@p�վݺ�	 ) AND
	   (@p������='' OR ������=@p������) AND
       (@p���ѵ�λ='' OR ���ѵ�λ����=@p���ѵ�λ) AND
	   (@p��ʼʱ��='' OR @p����ʱ��='' OR ( �˷����� between @p��ʼʱ�� AND @p����ʱ�� ))	AND
--	(���ܿ�������=@pҵ����� or @pҵ�����='') and
	(�շ���=@p�տ���  or �˷���=@p�տ��� or @p�տ��� ='') and
	(��Ӧҵ��=@p��Ӧҵ�� or @p��Ӧҵ��='') and
	   (�շ�״̬ =2) 
group by �շ�����,�շѱ��,���ѵ�λ����,������,
         convert(varchar(10),�˷�����,111),�շ�������,
         ���۱���,�շ�״̬,���ѷ�ʽ����,��������

update #TEMP_����ֵ�� set Ʊ�ݺ�=isnull(a.�վݺ� ,'')
from (select �շ�����,min(�վݺ�) as �վݺ� from �շѹ���_��ӡ������Ϣ group by �շ�����) a,#TEMP_����ֵ�� b
where a.�շ�����=b.�շ�����


update #TEMP_����ֵ�� set Ʊ�ݺ�=Ʊ�ݺ�+ case when Ʊ�ݺ�<>isnull(a.�վݺ�,'') then '��'+a.�վݺ� else '' end
from (select �շ�����,max(�վݺ�) as �վݺ� from �շѹ���_��ӡ������Ϣ group by �շ�����) a,
    #TEMP_����ֵ�� b
where a.�շ�����=b.�շ�����

--���ؽ����
SELECT * FROM #TEMP_����ֵ�� ORDER BY Ʊ�ݺ� DESC,��� DESC
DROP TABLE #TEMP_����ֵ��

--exec �շѹ���_�����շ���Ϣ '','','','','2008-07-16','2008-07-17','','',''


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


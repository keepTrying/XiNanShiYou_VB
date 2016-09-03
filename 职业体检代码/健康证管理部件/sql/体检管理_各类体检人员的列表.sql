SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



ALTER   PROCEDURE ������_���������Ա���б�
  @p��ʼ���� VARCHAR ( 10 ) = NULL,
  @p�������� VARCHAR ( 10 ) = NULL,
  @pƬ����� VARCHAR ( 20 ) = NULL,
  @p��ҵ����� VARCHAR ( 20 ) = NULL,
  @p�������� VARCHAR ( 50 ) = NULL
AS
--�ʹ洢���� ������_�ѵǼ�δ�½��۵������Ա�б� ��Щ����
--�������˶�������������� ������ ���ж�����
SET NOCOUNT ON

--��'','*',ת��ΪNULL,������׼�͹淶
IF @p��ʼ���� = '' OR @p��ʼ���� = '*' SELECT @p��ʼ���� = NULL
IF @p�������� = '' OR @p�������� = '*' SELECT @p�������� = NULL
IF @p�������� = '' OR @p�������� = '*' SELECT @p�������� = NULL
IF @pƬ����� = '' OR @pƬ����� = '*' SELECT @pƬ����� = NULL
IF @p��ҵ����� = '' OR @p��ҵ����� = '*' SELECT @p��ҵ����� = NULL

SELECT a.ϵͳ���,b.����,b.�Ա�,DATEDIFF(year,b.��������,GETDATE()) AS ����,
       b.��λ���� AS ��쵥λ,a.��������,a.�������,a.������,a.��Ϻʹ������,
       dbo.ϵͳ����_��ȡ�ֵ�����(b.Ƭ��,'Ƭ���ֵ��ֵ�') AS Ƭ��,
       dbo.ϵͳ����_��ȡ�ֵ�����(b.��ҵ���,'��ҵ�����ֵ�') AS ��ҵ���
  FROM ������_��������Ϣ�� a,
       ������_�����Ա������Ϣ�� b
 WHERE a.����������� = b.����������� AND
       (@p��ʼ���� IS NULL OR a.������� >= @p��ʼ����) AND
       (@p�������� IS NULL OR a.������� <= @p��������) AND
       (@p�������� IS NULL OR @p�������� = a.��������) AND
       (@pƬ����� IS NULL OR b.Ƭ�� = @pƬ�����) AND
       (@p��ҵ����� IS NULL OR b.��ҵ��� = @p��ҵ�����) AND
	������<>''

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


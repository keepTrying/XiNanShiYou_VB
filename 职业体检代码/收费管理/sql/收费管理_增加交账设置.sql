create table �շѹ���_Ʊ�ݽ��˼�¼�� (
	Ʊ�ݺ� varchar(20) not null,
	�ѽ��� tinyint null,
	�������� smalldatetime null,
	���տ� tinyint null,
	�տ����� smalldatetime null
)
go

create view �շѹ���_Ʊ�ݽ��˼�¼��ͼ as
select �վݺ�,�շ���,�շ�Ա,���,convert(varchar(10),��������,120) �շ�����,�ѽ���,convert(varchar(10),��������,120) ��������,���տ�,convert(varchar(10),�տ�����,120) �տ�����
from �շѹ���_Ʊ�ݽ��˼�¼�� a,(select �վݺ�,�շ���,���� �շ�Ա,��������,sum(���) ��� from �շѹ���_������Ϣ�� a,ϵͳ����_Ա��������Ϣ�� b where a.�շ���=b.��� and �շ�״̬=1 group by �վݺ�,�շ���,����,��������) b
where a.Ʊ�ݺ�=*b.�վݺ�
go

create proc �շѹ���_��ȡƱ�ݽ�����Ϣ 
	@pWhere1 varchar(100),
	@pWhere2 varchar(100)
as

set nocount on
declare @lstrSql varchar(1000)
select @lstrSql='select * into #temp from �շѹ���_Ʊ�ݽ��˼�¼��ͼ where ' + @pWhere1

select @lstrSql=@lstrSql + ' select * from #temp where ' + @pWhere2
exec (@lstrSql)

go

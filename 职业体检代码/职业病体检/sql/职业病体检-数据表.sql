----------------ְҵ����� �������ݱ�------------------
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_ҵ��������Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_ҵ��������Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_ҵ��������Ϣ��] (
	[������Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[����ֵ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[ö����Դ] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[˵��] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_��������ʷ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_��������ʷ��]
GO

CREATE TABLE [dbo].[ְҵ�����_��������ʷ��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[ĩ���¾�] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[ͣ������] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[�Ƿ���] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[�������] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��ż�Ӵ�����] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[��żְҵ] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[��ż����״��] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[�д�] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��Ȼ����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��̥] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��̥] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��λ����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[���в���ԭ��] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[������Ů��Ŀ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��Ů����״��] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[����ʷ] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[���̶̳�] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[���Ƴ̶�] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����ʱ��] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����ʷ] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����ԱͼƬ�����]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����ԱͼƬ�����]
GO

CREATE TABLE [dbo].[ְҵ�����_�����ԱͼƬ�����] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[��Ŀ���] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ͼƬ] [image] NOT NULL ,
	[��дʱ��] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ա������Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ա������Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ա������Ϣ��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[������ݺ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[�Ա�] [varchar] (4) COLLATE Chinese_PRC_CI_AS NULL ,
	[��������] [datetime] NULL ,
	[������] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��λ������] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[��λ����] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[��������] [datetime] NULL ,
	[Σ������] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[ְҵ����] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[����Դ] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[�ֹ���] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[ְ���ְ��] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[�������] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ְҵΣ������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[�绰����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[סַ] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[�ʱ�] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[�Ļ��̶�] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[���] [varchar] (5) COLLATE Chinese_PRC_CI_AS NULL ,
	[У����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[У��ʱ��] [datetime] NULL ,
	[У�˺ϸ�] [varchar] (2) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ա������Ŀ���ñ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ա������Ŀ���ñ�]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ա������Ŀ���ñ�] (
	[������Ŀ] [varchar] (30) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[¼�����] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[��������] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ݳ���] [int] NULL ,
	[ö��ֵ] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_���ҽʦ��Ŀ���ñ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_���ҽʦ��Ŀ���ñ�]
GO

CREATE TABLE [dbo].[ְҵ�����_���ҽʦ��Ŀ���ñ�] (
	[ҽʦ���] [UDT_Ա�����] NOT NULL ,
	[�����Ŀ] [varchar] (6) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_��������Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_��������Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_��������Ϣ��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�Թܱ��] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[�������] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[������] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[�������] [datetime] NOT NULL ,
	[������] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[��Ϻʹ������] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[�½���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[����ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[���״̬] [varchar] (4) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�շ�����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[�½�������] [datetime] NULL ,
	[�շѽ��] [money] NULL ,
	[�������״̬] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�������Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�������Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_�������Ϣ��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (6) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (300) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[��д����] [datetime] NULL ,
	[�������] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�������ж�������]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�������ж�������]
GO

CREATE TABLE [dbo].[ְҵ�����_�������ж�������] (
	[���] [int] NOT NULL ,
	[������] [int] NOT NULL ,
	[����] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[���] [int] NOT NULL ,
	[�����Ŀ] [UDT_�����Ŀ���] NOT NULL ,
	[�ж�����] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[��׼ֵ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����ģ�������۱�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_����ģ�������۱�]
GO

CREATE TABLE [dbo].[ְҵ�����_����ģ�������۱�] (
	[��������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[������] [int] NOT NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����ģ�������Ŀ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_����ģ�������Ŀ��]
GO

CREATE TABLE [dbo].[ְҵ�����_����ģ�������Ŀ��] (
	[��������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����ģ�������Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_����ģ�������Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_����ģ�������Ϣ��] (
	[���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[��������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[�Թܱ����ĸ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��ϴ������] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[��쵥����] [varchar] (60) COLLATE Chinese_PRC_CI_AS NULL ,
	[�Ƿ񸴲�����] [smallint] NULL ,
	[����] [varchar] (4) COLLATE Chinese_PRC_CI_AS NULL ,
	[�շѱ�׼] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[�����Ա����] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_����ģ�帽����Ŀ��Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_����ģ�帽����Ŀ��Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_����ģ�帽����Ŀ��Ϣ��] (
	[��������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[������Ŀ] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[���] [smallint] NOT NULL ,
	[�Ƿ��¼] [varchar] (1) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_��츽����Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_��츽����Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_��츽����Ϣ��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[������Ŀ] [varchar] (30) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[��Ŀֵ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[��Ŀֵ���] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ŀ���ñ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ŀ���ñ�]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ŀ���ñ�] (
	[����] [varchar] (6) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[����] [varchar] (30) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ȱʡֵ] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[ö����Դ] [varchar] (60) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (4) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[������] [int] NOT NULL ,
	[�ȽϷ�ʽ] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[��׼ֵ] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[��λ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [money] NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ְҵ�����_�û�����Ȩ�ޱ�_ְҵ�����_���ò�����Ϣ��]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ְҵ�����_�û�����Ȩ�ޱ�] DROP CONSTRAINT FK_ְҵ�����_�û�����Ȩ�ޱ�_ְҵ�����_���ò�����Ϣ��
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_���ò�����Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_���ò�����Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_���ò�����Ϣ��] (
	[������] [UDT_������] NOT NULL ,
	[��������] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[�ϼ�������] [UDT_������] NULL ,
	[ҵ����] [UDT_ҵ����] NULL ,
	[������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ҵ��˳��] [int] NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_������ʷ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_������ʷ��]
GO

CREATE TABLE [dbo].[ְҵ�����_������ʷ��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[�������] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[��ϵ�λ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ƾ���] [varchar] (300) COLLATE Chinese_PRC_CI_AS NULL ,
	[ת��] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�û�����Ȩ�ޱ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�û�����Ȩ�ޱ�]
GO

CREATE TABLE [dbo].[ְҵ�����_�û�����Ȩ�ޱ�] (
	[�û����] [UDT_Ա�����] NOT NULL ,
	[Ȩ����] [UDT_������] NOT NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�û�����Ȩ�ޱ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�û�����Ȩ�ޱ�]
GO

CREATE TABLE [dbo].[ְҵ�����_�û�����Ȩ�ޱ�] (
	[�û����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[���ұ��] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_���ҽ��۱�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_���ҽ��۱�]
GO

CREATE TABLE [dbo].[ְҵ�����_���ҽ��۱�] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[����] [varchar] (5) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[���ֽ���] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[ҽ�����] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[��������] [datetime] NULL ,
	[�޸���ʼʱ��] [datetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_B��Ӱ���]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_B��Ӱ���]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_B��Ӱ���] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_X��Ӱ���]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_X��Ӱ���]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_X��Ӱ���] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_��ٿ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_��ٿ�]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_��ٿ�] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_�ڿ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_�ڿ�]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_�ڿ�] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_���]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_���]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_���] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_�򳣹滯���]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_�򳣹滯���]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_�򳣹滯���] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_�ĵ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_�ĵ��]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_�ĵ��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_Ⱦɫ�廯���]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_Ⱦɫ�廯���]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_Ⱦɫ�廯���] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_�������]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_�������]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_�������] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_�ι��ܻ����]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_�ι��ܻ����]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_�ι��ܻ����] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_�ι���Ӱ���]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_�ι���Ӱ���]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_�ι���Ӱ���] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_Ѫ���滯���]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_Ѫ���滯���]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_Ѫ���滯���] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_ְҵʷ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_ְҵʷ��]
GO

CREATE TABLE [dbo].[ְҵ�����_ְҵʷ��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[������λ] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[Σ������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[�Ӵ�ʱ��] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[������ʩ] [varchar] (80) COLLATE Chinese_PRC_CI_AS NULL ,
	[��ע] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[��������] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[ÿ�չ�����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[�ۻ�������] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��������ʷ] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[��ʼʱ��] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[����ʱ��] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[�Ƿ������] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�Ծ�֢״��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�Ծ�֢״��]
GO

CREATE TABLE [dbo].[ְҵ�����_�Ծ�֢״��] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[֢״] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[�̶�] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[����ʱ��] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼���]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼���]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼���] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


----------------ְҵ����� �������ݱ�������------------------


----------------ְҵ����� Ȩ�ޡ���Ŀ���ֵ������������------------------
-------------*****����������ְҵ�����-Ȩ�����á���������*****---------------

-----------������ݱ�3��,ר�ſ���ְҵ�����ģ��Ŀ��ҺͲ���Ȩ��
-----������ְҵ�����_�û�����Ȩ�ޱ������
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�û�����Ȩ�ޱ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�û�����Ȩ�ޱ�]
GO

CREATE TABLE [dbo].[ְҵ�����_�û�����Ȩ�ޱ�] (
	[�û����] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[���ұ��] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO
-----������ְҵ�����_�û�����Ȩ�ޱ������


-----������ְҵ�����_�û�����Ȩ�ޱ������
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�û�����Ȩ�ޱ�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�û�����Ȩ�ޱ�]
GO

CREATE TABLE [dbo].[ְҵ�����_�û�����Ȩ�ޱ�] (
	[�û����] [UDT_Ա�����] NOT NULL ,
	[Ȩ����] [UDT_������] NOT NULL 
) ON [PRIMARY]
GO
-----������ְҵ�����_�û�����Ȩ�ޱ������


-----������ְҵ�����_�û�����Ȩ�ޱ������
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ְҵ�����_�û�����Ȩ�ޱ�_ְҵ�����_���ò�����Ϣ��]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ְҵ�����_�û�����Ȩ�ޱ�] DROP CONSTRAINT FK_ְҵ�����_�û�����Ȩ�ޱ�_ְҵ�����_���ò�����Ϣ��
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_���ò�����Ϣ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_���ò�����Ϣ��]
GO

CREATE TABLE [dbo].[ְҵ�����_���ò�����Ϣ��] (
	[������] [UDT_������] NOT NULL ,
	[��������] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[�ϼ�������] [UDT_������] NULL ,
	[ҵ����] [UDT_ҵ����] NULL ,
	[������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[����] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ҵ��˳��] [int] NULL 
) ON [PRIMARY]
GO
-----������ְҵ�����_�û�����Ȩ�ޱ������

-----������ְҵ�����_ҵ��������Ϣ�������
insert into ְҵ�����_ҵ��������Ϣ�� values('��������',12,null,null)
insert into ְҵ�����_ҵ��������Ϣ�� values('�Ƿ�����','��',null,null)
insert into ְҵ�����_ҵ��������Ϣ�� values('�Ƿ���ٵǼ�','��',null,null)
insert into ְҵ�����_ҵ��������Ϣ�� values('���ٵǼ��Ƿ�Ʒ�','��',null,null)
insert into ְҵ�����_ҵ��������Ϣ�� values('�Ƿ��ӡ��쵥','��',null,null)
insert into ְҵ�����_ҵ��������Ϣ�� values('�Ƿ��ӡ����','��',null,null)
insert into ְҵ�����_ҵ��������Ϣ�� values('�Թܱ���Զ�����','��','�ǣ���',null)
insert into ְҵ�����_ҵ��������Ϣ�� values('�Ƿ��շ�','��',null,null)
insert into ְҵ�����_ҵ��������Ϣ�� values('��һ����λ��',0,15,'���ݴ�ӡ���봰��ȷ��')

--2012-06-06 �ڵ�� ��
--��¼���һ��ͳ��ʱ�Ļ�������
insert into ְҵ�����_ҵ��������Ϣ�� values('ͳ������_������','ͳ�����_�ϸ���','12','19')
insert into ְҵ�����_ҵ��������Ϣ�� values('ͳ������-��������','1','12','19')
--2012-06-06 �ڵ�� ��
go
-----������ְҵ�����_ҵ��������Ϣ�������


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '���������ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('���������ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '���������ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
go

--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�����ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('�����ֵ�','ϵͳ����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�����ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','�ѻ�','YH' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','δ��','WH' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'03','����','LY' ,'',0)
go

--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = 'ְҵ���������ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('ְҵ���������ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = 'ְҵ���������ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'11','��ȼ��ѭ��','HRL' ,'',0)
--�����ֵ䡰��ȼ��ѭ�����Ķ������ݡ�
select @llngParent=InnerID from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID and ���='11'
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1101','�˿󿪲�','YKKC' ,'',@llngParent)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1102','�˿�ˮұ','YKSY' ,'',@llngParent)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1103','ȼ������','RLZZ' ,'',@llngParent)
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'12','ҽѧӦ��','YXYY' ,'',0)
select @llngParent=InnerID from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID and ���='12'
--�����ֵ䡰ҽѧӦ�á��Ķ������ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1201','��Ϸ���ѧ','ZKFSX' ,'',@llngParent)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1202','���Ʒ���ѧ','' ,'',@llngParent)
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'13','��ҵӦ��','GYYY' ,'',0)
select @llngParent=InnerID from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID and ���='13'
--�����ֵ䡰��ҵӦ�á��Ķ������ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1301','��ҵ����','GYFZ' ,'',@llngParent)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1302','��ҵ����','GYKS' ,'',@llngParent)
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'14','��ȻԴ','TRY' ,'',0)
select @llngParent=InnerID from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID and ���='14'
--�����ֵ䡰��ȻԴ���Ķ������ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1401','���ú���','MYHK' ,'',@llngParent)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1402','ú�󿪲�','MKKC' ,'',@llngParent)
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'15','����','QT' ,'',0)
select @llngParent=InnerID from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID and ���='15'
--�����ֵ䡰�������Ķ������ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1501','����','JY' ,'',@llngParent)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'1502','��ѧ�о�','KXYJ' ,'',@llngParent)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '���������ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('���������ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '���������ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','��ͨ���','PTTJ' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','ְҵ����','ZYJK' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'03','���佡��','FSJK' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'04','��˲���','SHBD' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'05','8023����','8023' ,'',0)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '��������ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('��������ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '��������ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','�ϸ�ǰ','ZGQ' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','�ڸ��ڼ�','ZGQJ' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'03','���ʱ','LGS' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'04','Ӧ�����','YJJC' ,'',0)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�����ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('�����ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�����ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','������','SCB' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','����','GLB' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'03','�г���','SCB' ,'',0)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�����ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('�����ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�����ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','��ͨ','PU' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','����','TZ' ,'',0)
go

--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�����������ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('�����������ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�����������ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','X����','' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','Y����','' ,'',0)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '����̶��ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('����̶��ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '����̶��ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','��΢','QW' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','������','JMX' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'03','����','MX' ,'',0)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = 'Σ�������ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('Σ�������ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = 'Σ�������ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','����','ZS' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','ǿ��','QQ' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'03','������','FSX' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'04','�۳�','FC' ,'',0)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = 'ְҵ��ְ���ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('ְҵ��ְ���ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = 'ְҵ��ְ���ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','�쵼','LD' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','����ʦ','GCS' ,'',0)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�̶��ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('�̶��ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = '�̶��ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','�Ӳ�','CB' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','ż��','OE' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'03','����','CQ' ,'',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'04','����','YZ' ,'',0)
go


--��ȡ�ֵ�ID��
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = 'ְҵ���������ֵ�')
     insert into ϵͳ����_�ֵ�_�ֵ���б�(����,ҵ����,����) values('ְҵ���������ֵ�','ְҵ�����','������')
select @llngID = ID FROM ϵͳ����_�ֵ�_�ֵ���б� WHERE ���� = 'ְҵ���������ֵ�'
--ɾ���ֵ�����ݡ�
delete ϵͳ����_�ֵ�_�ֵ����ݱ� where ID=@llngID
--�����ֵ�һ�����ݡ�
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'01','��ٿ�','WG' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'02','�ڿ�','LK' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'03','���','WK' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'04','Ѫ���滯���','XCG' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'05','�ι��ܻ����','GGN' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'06','�򳣹滯���','LGN' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'07','Ⱦɫ�廯���','RST' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'08','�������','DCT' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'10','�ĵ��','XD' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'11','B��Ӱ���','BCYX' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'12','�ι���Ӱ���','FGN' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'09','X��Ӱ���','XGYX' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'14','���Ǽ�','TJDJ' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'15','ҵ������','YWSZ' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'13','�ܼ��߸�����Ϣ¼���','ZYBS' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'16','���ս���¼��','ZZJL' ,'ְҵ�����_����',0)
GO


-------------------------------------------
insert into ϵͳ����_ϵͳ��װ��Ϣ�� values('ְҵ�����',1)
go

-------------------------------------------
insert into ϵͳ����_ƽ̨���������� values('0001','ְҵ�����','ҵ����')
go


-----------�ⲿ����ÿ��ģ����ܽ�㲿�֡��ӽڵ��ٸ�������ֶ��Ӱɡ�
-----------��ϵͳ����_���ò�����Ϣ���У�Ҳ��Ҫ��������г��Ĵ����
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_���Ǽ�',null,null,'ְҵ�����','ְҵ������','clsmanagetestform',1)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_��ٿƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_ҵ������',null,null,'ְҵ�����','ְҵ������','clsmageconfform_zyb',3)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_�ܼ��߸�����Ϣ¼���',null,null,'ְҵ�����','ְҵ��ʷ¼��','clscareerhstmage',2)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_�ڿƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',5)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',6)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_Ѫ���滯��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',7)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_�ι��ܻ���ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',8)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_�򳣹滯��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',9)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_Ⱦɫ�廯��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',10)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_������ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',11)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_X��Ӱ��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',12)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_�ι���Ӱ��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',13)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_�ĵ�ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',14)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_B��Ӱ��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',15)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_���ս���¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',16)
go

-----------��ְҵ�����_���ò�����Ϣ���У���ȥ����ģ���Ҫ�����Ҫ���Ӳ�����
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���Ǽ�',null,null,'ְҵ�����','ְҵ������','clsmanagetestform',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_��ٿƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_ҵ������',null,null,'ְҵ�����','ְҵ������','clsmageconfform_zyb',3)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ܼ��߸�����Ϣ¼���',null,null,'ְҵ�����','ְҵ��ʷ¼��','clscareerhstmage',2)

insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ڿƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',5)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',6)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_Ѫ���滯��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',7)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ι��ܻ���ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',8)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�򳣹滯��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',9)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_Ⱦɫ�廯��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',10)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_������ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',11)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_X��Ӱ��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',12)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ι���Ӱ��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',13)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ĵ�ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',14)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_B��Ӱ��ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',15)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���ս���¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',16)

--2012-04-24 ��¶
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_��ٿƽ��¼��_�޸�',null,'ְҵ�����_��ٿƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_B��Ӱ��ƽ��¼��_�޸�',null,'ְҵ�����_B��Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_X��Ӱ��ƽ��¼��_�޸�',null,'ְҵ�����_X��Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_������ƽ��¼��_�޸�',null,'ְҵ�����_������ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ι���Ӱ��ƽ��¼��_�޸�',null,'ְҵ�����_�ι���Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ι��ܻ���ƽ��¼��_�޸�',null,'ְҵ�����_�ι��ܻ���ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ڿƽ��¼��_�޸�',null,'ְҵ�����_�ڿƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�򳣹滯��ƽ��¼��_�޸�',null,'ְҵ�����_�򳣹滯��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_Ⱦɫ�廯��ƽ��¼��_�޸�',null,'ְҵ�����_Ⱦɫ�廯��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_��ƽ��¼��_�޸�',null,'ְҵ�����_��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�ĵ�ƽ��¼��_�޸�',null,'ְҵ�����_�ĵ�ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_Ѫ���滯��ƽ��¼��_�޸�',null,'ְҵ�����_Ѫ���滯��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_B��Ӱ��ƽ��¼��_��������',null,'ְҵ�����_B��Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',3)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_B��Ӱ��ƽ��¼��_ɾ��',null,'ְҵ�����_B��Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',2)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_X��Ӱ��ƽ��¼��_ɾ��',null,'ְҵ�����_X��Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',2)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_X��Ӱ��ƽ��¼��_��������',null,'ְҵ�����_X��Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���ս���¼��_�������',null,'ְҵ�����_���ս���¼��','ְҵ�����','ְҵ��������','clsmanagetestform',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���ս���¼��_��ӡ����',null,'ְҵ�����_���ս���¼��','ְҵ�����','ְҵ��������','clsmanagetestform',2)
--2012-04-24 ��¶��ע�ͽ�����

--2012-05-23 ���� ��
--ְҵ�������������Ȩ����Ϣ
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_����Ǽ�','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',2)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_����Ǽ�','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',3)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_�޸�','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_ɾ��','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',5)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_����','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',6)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_��λ����','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',7)
--insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_��ӡ������','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',8)
--insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_���´�����','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',9)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_���Ǽ�_���Ǽ�','','ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',10)

update ְҵ�����_���ò�����Ϣ�� set ����='clsmanagetestform' where ������='ְҵ�����_ҵ������'

insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_��������','','','ְҵ�����','ְҵ������','clsMageConfForm_zyb',1)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_ҵ������_��������_����','','ְҵ�����_ҵ������_��������','ְҵ�����','ְҵ������','clsMageConfForm_zyb',2)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_ҵ������_��������_����','','ְҵ�����_ҵ������_��������','ְҵ�����','ְҵ������','clsMageConfForm_zyb',3)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_ҵ������_��������_����','','ְҵ�����_ҵ������_��������','ְҵ�����','ְҵ������','clsMageConfForm_zyb',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_ҵ������_��������_ɾ��','','ְҵ�����_ҵ������_��������','ְҵ�����','ְҵ������','clsMageConfForm_zyb',5)

insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ܼ��߸�����Ϣ¼���','','','ְҵ�����','ְҵ��ʷ¼��','clsCareerHstMage',1)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ܼ��߸�����Ϣ¼���_ְҵ��ʷ�Ǽ�','','ְҵ�����_�ܼ��߸�����Ϣ¼���','ְҵ�����','ְҵ��ʷ¼��','clsCareerHstMage',2)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ܼ��߸�����Ϣ¼���_�޸�','','ְҵ�����_�ܼ��߸�����Ϣ¼���','ְҵ�����','ְҵ��ʷ¼��','clsCareerHstMage',3)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ܼ��߸�����Ϣ¼���_ɾ��','','ְҵ�����_�ܼ��߸�����Ϣ¼���','ְҵ�����','ְҵ��ʷ¼��','clsCareerHstMage',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ܼ��߸�����Ϣ¼���_����','','ְҵ�����_�ܼ��߸�����Ϣ¼���','ְҵ�����','ְҵ��ʷ¼��','clsCareerHstMage',5)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ܼ��߸�����Ϣ¼���_��ӡ','','ְҵ�����_�ܼ��߸�����Ϣ¼���','ְҵ�����','ְҵ��ʷ¼��','clsCareerHstMage',6)

insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_ְҵ����ѯͳ��','','','ְҵ�����','ְҵ������','clsmanagetestform',6)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_ְҵ����ѯͳ��_����Excel','','ְҵ�����_ְҵ����ѯͳ��','ְҵ�����','ְҵ������','clsmanagetestform',6)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_ְҵ����ѯͳ��_����Excel','','ְҵ�����_ְҵ����ѯͳ��','ְҵ�����','ְҵ������','clsmanagetestform',6)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_ְҵ����ѯͳ��_��ӡ','','ְҵ�����_ְҵ����ѯͳ��','ְҵ�����','ְҵ������','clsmanagetestform',6)


--�����ҽ������¼��Ȩ����Ϣ
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_B��Ӱ��ƽ��¼��_�����޸�','','ְҵ�����_B��Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_X��Ӱ��ƽ��¼��_�����޸�','','ְҵ�����_X��Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_������ƽ��¼��_�����޸�','','ְҵ�����_������ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ι���Ӱ��ƽ��¼��_�����޸�','','ְҵ�����_�ι���Ӱ��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ι��ܻ���ƽ��¼��_�����޸�','','ְҵ�����_�ι��ܻ���ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ڿƽ��¼��_�����޸�','','ְҵ�����_�ڿƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�򳣹滯��ƽ��¼��_�����޸�','','ְҵ�����_�򳣹滯��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_Ⱦɫ�廯��ƽ��¼��_�����޸�','','ְҵ�����_Ⱦɫ�廯��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_��ƽ��¼��_�����޸�','','ְҵ�����_��ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_��ٿƽ��¼��_�����޸�','','ְҵ�����_��ٿƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_�ĵ�ƽ��¼��_�����޸�','','ְҵ�����_�ĵ�ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values('ְҵ�����_Ѫ���滯��ƽ��¼��_�����޸�','','','ְҵ�����','ְҵ�������¼��','clscommon',4)
--2012-05-23 ���� ��

--2012-06-15 �ڵ�� ��
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���Ǽ�_��ӡ�嵥',null,'ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',8)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���Ǽ�_��ӡ�Թܱ�ǩ',null,'ְҵ�����_���Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',9)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���Ǽ�_����Ǽ�_У��ͨ��',null,'ְҵ�����_���Ǽ�_����Ǽ�','ְҵ�����','ְҵ������','clsmanagetestform',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���ս���¼��_����ͨ��',null,'ְҵ�����_���ս���¼��','ְҵ�����','ְҵ������','clsmanagetestform',3)
--2012-06-15 �ڵ�� ��
go

-----------�����Ա�����еĿ��ҹ���(ֻ��"ְҵ������"��һ����)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID,���,����,���Ƿ�,����,Parent) values(1,'06','ְҵ������','','',0)
go

-------------*****����������ְҵ�����-Ȩ�����á�������*****---------------
----------------ְҵ����� Ȩ�ޡ���Ŀ���ֵ�����������ã�������------------------



----------------ְҵ����� ���������Ŀ---------------
-------------*****����������ְҵ�����-��ٿ�-�����Ŀ���á���������*****---------------
declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '��ٿ�' and ID = @intID)

insert ְҵ�����_�����Ŀ���ñ� values('01001','ɫ��-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01002','ɫ��-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01003','����Ӧ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01004','����Ӧ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01005','��Ұ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01006','��Ұ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01007','����-Զ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01008','����-Զ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01009','����-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01010','����-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01011','����-Զ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01012','����-Զ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01013','����-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01014','����-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01015','��ǰ��-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01016','��ǰ��-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01017','��Ĥ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01018','��Ĥ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01019','��Ĥ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01020','��Ĥ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01021','ǰ��-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01022','ǰ��-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01023','��Ĥ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01024','��Ĥ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01025','��״��-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01026','��״��-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01027','������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01028','������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01029','�۵�-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01030','�۵�-��','����','����,�쳣','����',@intInnerID,'','','',1)
--insert ְҵ�����_�����Ŀ���ñ� values('01031','��״�廷�漰����ͼ-��','����','����,�쳣','����',@intInnerID,'','','',1)
--insert ְҵ�����_�����Ŀ���ñ� values('01032','��״�廷�漰����ͼ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01033','����','����','����,�쳣','����',@intInnerID,'','','',1)


--------------------------��˲���
insert ְҵ�����_�����Ŀ���ñ� values('01034','����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01035','����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01036','����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01037','����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01038','�۳�״-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01039','�۳�״-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01040','�۳�״-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01041','�۳�״-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01042','�۳�״-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01043','�۳�״-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01044','��״-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01045','��״-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01046','��״-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01047','��״-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01048','��״-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01049','��״-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01050','Ƭ״-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01051','Ƭ״-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01052','Ƭ״-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01053','Ƭ״-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01054','Ƭ״-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01055','Ƭ״-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01056','����-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01057','����-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01058','����-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01059','����-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01060','����-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01061','����-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01062','����-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01063','����-������-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01064','����-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01065','����-ǰ����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01066','����-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01067','����-���-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01068','���','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01069','��״�廷�漰����ͼ','����','����,�쳣','����',@intInnerID,'','','',1)


-------------���Ǻ�Ƽ����
insert ְҵ�����_�����Ŀ���ñ� values('01070','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01071','���','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01072','��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01073','���','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01074','�ж�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01075','����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01076','����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01077','�����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01078','�����-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01079','��ͻ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01080','��ͻ-��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01081','��-ճĤ','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01082','��-��Ѫ','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01083','��ǻ-ճĤ','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01084','��ǻ-����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01085','�ʺ�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01086','��ǻ','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01087','���Ǻ������','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('01088','����','','','����',@intInnerID,'','','',1,)
go

-------------*****����������ְҵ�����-��ٿ�-�����Ŀ���á���������*****---------------


   
   
   --���ߣ�����
   --ʱ�䣺2012-04-11
-------------*****����������������������ְҵ�����-�ڿ�-�����Ŀ���á���������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '�ڿ�' and ID = @intID)

insert ְҵ�����_�����Ŀ���ñ� values('02001','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02002','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02003','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02004','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02005','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02006','Ƣ��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02007','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02008','ʹ��������','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02009','λ�þ�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02010','���ڷ���','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02011','���췴��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02012','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02013','������','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02014','�����˶�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02015','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02016','������','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('02018','����','','','����',@intInnerID,'','','',1,)
GO

-------------*****��������������������������ְҵ�����-�ڿ�-�����Ŀ���á�������������������������*****---------------



-------------*****������������������������ְҵ�����-���-�����Ŀ���á�����������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '���' and ID = @intID)

insert ְҵ�����_�����Ŀ���ñ� values('03001','Ƥ����ɫ','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03002','Ƥ��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03003','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03004','���','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03005','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03006','ǳ���ܰͽ�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03007','��״��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03008','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03009','��֫�ؽ�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03010','�ѷ�����ë','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03011','��Ѫ���','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03012','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03013','��м','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03014','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03015','ɫ�س���','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03016','ɫ�ؼ���','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03017','���Ƚǻ�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03018','�ູ','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03019','��״��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03020','Ƥ��ή��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03021','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03022','ָ��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03023','�ܰͽ�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('03024','����','','','����',@intInnerID,'','','',1)
GO

-------------*****��������������������������ְҵ�����-���-�����Ŀ���á�������������������������*****---------------

-------------*****������������������������ְҵ�����-Ѫ���滯���-�����Ŀ���á�����������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = 'Ѫ���滯���' and ID = @intID)

insert ְҵ�����_�����Ŀ���ñ� values('04001','��ϸ��*10E9/L','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04002','����%','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04003','�ܰ�%','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04004','����%','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04005','��ϸ��*10E12/L','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04006','Ѫ�쵰��g/L','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04007','ѪС��*10E9/L','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04008','���Ը�״����ϸ��%','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04009','���Է�Ҷ����ϸ��%','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04010','��������ϸ��%','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04011','�ȼ�����ϸ��%','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04012','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04013','���ص�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04014','пԭ߲��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04015','Ǧ','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04016','������ø(u)','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04017','Ѫ��(mmol/L)','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('04027','����','','','����',@intInnerID,'','','',1)
GO

-------------*****��������������������ְҵ�����-Ѫ���滯���-�����Ŀ���á�������������������*****---------------

-------------*****��������������������ְҵ�����-�ι��ܻ����-�����Ŀ���á�������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '�ι��ܻ����' and ID = @intID)

insert ְҵ�����_�����Ŀ���ñ� values('05001','SGPT','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('05002','HbsAg','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('05003','TTT','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('05008','����','','','����',@intInnerID,'','','',1)
GO

-------------*****��������������������ְҵ�����-�ι��ܻ����-�����Ŀ���á�������������������*****---------------

-------------*****��������������������ְҵ�����-�򳣹滯���-�����Ŀ���á�������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '�򳣹滯���' and ID = @intID)

insert ְҵ�����_�����Ŀ���ñ� values('06001','�򵰰�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06002','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06003','��ϸ��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06004','��ϸ��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06005','����','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06006','Ǧ','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06007','��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06008','��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06009','��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06010','��','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06011','��һ������������','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06012','��2һ΢�򵰰�','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('06013','����','','','����',@intInnerID,'','','',1)
GO


-------------*****��������������������ְҵ�����-�򳣹滯���-�����Ŀ���á�������������������*****---------------


-------------*****��������������������ְҵ�����-Ⱦɫ�廯���-�����Ŀ���á�������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = 'Ⱦɫ�廯���' and ID = @intID)

insert ְҵ�����_�����Ŀ���ñ� values('07001','�������ڷ���ϸ����(��)','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('07002','Ⱦɫ�������(%)','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('07003','��������','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('07004','����ϸ������','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('07005','΢���ܰ�ϸ����(��)','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('07006','�ܰ�ϸ��΢����(��)','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('07007','����','','','����',@intInnerID,'','','',1)
GO


-------------*****��������������������ְҵ�����-Ⱦɫ�廯���-�����Ŀ���á�������������������*****---------------

-------------*****��������������������ְҵ�����-�������-�����Ŀ���á�������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '�������' and ID = @intID)


insert ְҵ�����_�����Ŀ���ñ� values('08001','��������','����','����,�쳣','����',@intInnerID,'','','',100)
insert ְҵ�����_�����Ŀ���ñ� values('08002','����','','','����',@intInnerID,'','','',100)
go
-------------*****��������������������ְҵ�����-�������-�����Ŀ���á�������������������*****---------------


-------------*****��������������������ְҵ�����-X��Ӱ���-�����Ŀ���á�������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = 'X��Ӱ���' and ID = @intID)


insert ְҵ�����_�����Ŀ���ñ� values('09001','X��Ӱ����','����','����,�쳣','����',@intInnerID,'','','',100)
insert ְҵ�����_�����Ŀ���ñ� values('09002','����','','','����',@intInnerID,'','','',100)
go
-------------*****��������������������ְҵ�����-X��Ӱ���-�����Ŀ���á�������������������*****---------------


-------------*****��������������������ְҵ�����-�ĵ��-�����Ŀ���á�������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '�ĵ��' and ID = @intID)


insert ְҵ�����_�����Ŀ���ñ� values('10001','�ĵ���','����','����,�쳣','����',@intInnerID,'','','',100)
insert ְҵ�����_�����Ŀ���ñ� values('10002','����','','','����',@intInnerID,'','','',100)
go
-------------*****��������������������ְҵ�����-�ĵ��-�����Ŀ���á�������������������*****---------------


-------------*****��������������������ְҵ�����-B��Ӱ���-�����Ŀ���á�������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = 'B��Ӱ���' and ID = @intID)


insert ְҵ�����_�����Ŀ���ñ� values('11001','B��Ӱ����','����','����,�쳣','����',@intInnerID,'','','',100)
insert ְҵ�����_�����Ŀ���ñ� values('11002','����','','','����',@intInnerID,'','','',100)
go
-------------*****��������������������ְҵ�����-B��Ӱ���-�����Ŀ���á�������������������*****---------------


-------------*****��������������������ְҵ�����-�ι���Ӱ���-�����Ŀ���á�������������������*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '�ι���Ӱ���' and ID = @intID)


insert ְҵ�����_�����Ŀ���ñ� values('12001','�ι���Ӱ����','����','����,�쳣','����',@intInnerID,'','','',100)
insert ְҵ�����_�����Ŀ���ñ� values('12002','����','','','����',@intInnerID,'','','',100)
go
-------------*****��������������������ְҵ�����-�ι���Ӱ���-�����Ŀ���á�������������������*****---------------

-------------*****��������������������ְҵ�����-�ܼ��߸�����Ϣ¼���-�����Ŀ���á�������������������*****---------------
--2012-06-13 �ڵ�� ��
--ʡ����Ҫ�����������һ���������ְҵ���������ϣ�
declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '�ܼ��߸�����Ϣ¼���' and ID = @intID)

insert ְҵ�����_�����Ŀ���ñ� values('13001','Ӫ��','����','����,�е�,����','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('13002','���','170','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('13003','����','60','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('13004','����ѹ','120','����,ƫ��,ƫ��','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('13005','����ѹ','90','����,ƫ��,ƫ��','����',@intInnerID,'','','',1)
--2012-06-13 �ڵ�� ��
-------------*****��������������������ְҵ�����-�ܼ��߸�����Ϣ¼���-�����Ŀ���á�������������������*****---------------

--///////////////////////////////////////////////////////////////////////////////////--
--���ߣ�����
--ʱ�䣺2012-04-11��ע�ͽ�����

----------------ְҵ����� ���������Ŀ��������---------------


----------------ְҵ����� ģ�����շ����ʼ��������------------------
-----------------------------
insert into ְҵ�����_����ģ�������Ϣ�� values(1,'ְҵ���������','�ϸ�ǰ','B',' ','',0,1,'','ְҵ����')
insert into ְҵ�����_����ģ�������Ϣ�� values(2,'ְҵ�����ڸڼ���','�ڸ��ڼ�','B',' ','',0,1,'','ְҵ����')
insert into ְҵ�����_����ģ�������Ϣ�� values(3,'ְҵ������ڼ���','���ʱ','B',' ','',0,1,'','ְҵ����')
go

-----------------------------
insert into �շѹ���_�շ���Ŀ�ֵ�� values('005','ְҵ�����',0,'��',871,0,0,'')
go

----------------ְҵ����� ģ�����շ����ʼ�������ã�������------------------


----------------ְҵ����� �����Ŀ������� ��2012-08-21�����ǣ�------------------
�������ݱ�
--������
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_������]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_������]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_������] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO
--ɾ���ι��ܻ����
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_�ι��ܻ����]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_�ι��ܻ����]
GO
--���� ���߿�
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_�����Ϣ_���߿�]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ְҵ�����_�����Ϣ_���߿�]
GO

CREATE TABLE [dbo].[ְҵ�����_�����Ϣ_���߿�] (
	[ϵͳ���] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����Ŀ] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[�����] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[���ҽʦ] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[��дʱ��] [datetime] NULL ,
	[�������] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO




--��������	����
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'05','���߿�','MY' ,'ְҵ�����_����',0)
insert into ϵͳ����_�ֵ�_�ֵ����ݱ�(ID, ���, ����, ���Ƿ�, ����, Parent) values(@llngID ,'17','������','SH' ,'ְҵ�����_����',0)

--��������	��Ŀ

--���߿�
declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '���߿�' and ID = @intID)
insert ְҵ�����_�����Ŀ���ñ� values('05001','HbsAb','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('05003','HbeAg','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('05004','HbeAb','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('05005','HbcAb','����','����,�쳣','����',@intInnerID,'','','',1)
insert ְҵ�����_�����Ŀ���ñ� values('05008','����','','','����',@intInnerID,'','','',1)

--������
declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '������' and ID = @intID)
insert ְҵ�����_�����Ŀ���ñ� values('17001','�ܵ�����','','','����',@intInnerID,'=','11.60','umol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17002','ֱ�ӵ�����','','','����',@intInnerID,'=','4.00','umol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17003','��ӵ�����','','','����',@intInnerID,'=','7.60','umol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17004','�ȱ�ת��ø','','','����',@intInnerID,'=','52.00','U/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17005','�Ȳ�ת��ø','','','����',@intInnerID,'=','30.00','U/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17006','�Ȳݱȹȱ�','','','����',@intInnerID,'=','0.57','',1)
insert ְҵ�����_�����Ŀ���ñ� values('17007','�ܵ���','','','����',@intInnerID,'=','91.00','g/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17008','�׵���','','','����',@intInnerID,'=','53.00','g/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17009','�򵰰�','','','����',@intInnerID,'=','38.00','g/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17010','�׵���/�򵰰�','','','����',@intInnerID,'=','1.39','',1)
insert ְҵ�����_�����Ŀ���ñ� values('17011','�ܵ�֭��','','','����',@intInnerID,'=','2.00','umol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17012','r-�Ȱ���ת��ø','','','����',@intInnerID,'=','30.00','U/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17013','��������ø','','','����',@intInnerID,'=','68.00','U/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17014','����','','','����',@intInnerID,'=','3.55','mmol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17015','����','','','����',@intInnerID,'=','138.20','umol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17016','������','','','����',@intInnerID,'=','10','mmol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17017','���̴�','','','����',@intInnerID,'=','10','mmol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17018','��������','','','����',@intInnerID,'=','10','mmol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17019','����','','','����',@intInnerID,'=','10','mmol/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('17020','����','','','����',@intInnerID,'','','',1)


--����Ȩ�����
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_�����ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',19)
insert into ϵͳ����_���ò�����Ϣ�� values ('ְҵ�����_�����ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',20)
go

--�Ӽ�Ȩ��
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�����ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',17)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���߿ƽ��¼��',null,null,'ְҵ�����','ְҵ�������¼��','clscommon',18)

insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�����ƽ��¼��_�޸�',null,'ְҵ�����_�����ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���߿ƽ��¼��_�޸�',null,'ְҵ�����_���߿ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',1)
--����¼��
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_�����ƽ��¼��_�����޸�',null,'ְҵ�����_�����ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
insert into ְҵ�����_���ò�����Ϣ�� values ('ְҵ�����_���߿ƽ��¼��_�����޸�',null,'ְҵ�����_���߿ƽ��¼��','ְҵ�����','ְҵ�������¼��','clscommon',4)
go



--�޸� Ѫ���滯��Ƶ���Ŀ����
insert ְҵ�����_�����Ŀ���ñ� values('04001','��ϸ��ѹ��','','','����',@intInnerID,'=','45.80','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04002','ƽ��ѪС�����','','','����',@intInnerID,'=','11.02','fL',1)
insert ְҵ�����_�����Ŀ���ñ� values('04003','ѪС��ѹ��','','','����',@intInnerID,'=','0.32','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04004','��ϸ��ƽ�����','','','����',@intInnerID,'=','91.20','fL',1)
insert ְҵ�����_�����Ŀ���ñ� values('04005','ƽ��Ѫ�쵰����','','','����',@intInnerID,'=','28.30','pg',1)
insert ְҵ�����_�����Ŀ���ñ� values('04006','ƽ��Ѫ�쵰��Ũ��','','','����',@intInnerID,'=','310.00','g/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04007','����ϸ������','','','����',@intInnerID,'=','64.20','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04008','�ܰ�ϸ������','','','����',@intInnerID,'=','28.70','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04009','����ϸ������','','','����',@intInnerID,'=','5.60','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04010','��������ϸ������','','','����',@intInnerID,'=','1.20','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04011','�ȼ�����ϸ������','','','����',@intInnerID,'=','0.30','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04012','����ϸ����','','','����',@intInnerID,'=','3.77','10^9/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04013','�ܰ�ϸ����','','','����',@intInnerID,'=','1.69','10^9/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04014','����ϸ��','','','����',@intInnerID,'=','0.33','10^9/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04015','��������ϸ��','','','����',@intInnerID,'=','0.07','10^9/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04016','�ȼ�����ϸ��','','','����',@intInnerID,'=','0.02','10^9/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04017','��ϸ���ֲ����','','','����',@intInnerID,'=','44.00','fL',1)
insert ְҵ�����_�����Ŀ���ñ� values('04018','��ϸ���ֲ���ȱ���ϵ��','','','����',@intInnerID,'=','13.40','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04019','ѪС��ֲ����','','','����',@intInnerID,'=','11.02','fL',1)
insert ְҵ�����_�����Ŀ���ñ� values('04020','����ѪС�����','','','����',@intInnerID,'=','32.12','%',1)
insert ְҵ�����_�����Ŀ���ñ� values('04021','��ϸ��','','','����',@intInnerID,'=','5.88','10^9/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04022','��ϸ��','','','����',@intInnerID,'=','5.02','10^12/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04023','Ѫ�쵰��','','','����',@intInnerID,'=','142.00','g/L',1)
insert ְҵ�����_�����Ŀ���ñ� values('04024','ѪС��','','','����',@intInnerID,'=','143.00','10^9/L',1)

--�޸� ְҵ�����_�����Ŀ���ñ� �����ֶΣ����ţ�

alter table ְҵ�����_�����Ŀ���ñ� add ���� varchar(10) null

--���������Ŀ��Ϣ	�����ţ�
update ְҵ�����_�����Ŀ���ñ� set ���� = 'WBC' where ����='��ϸ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'RBC' where ����='��ϸ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'HGB' where ����='Ѫ�쵰��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'HCT' where ����='��ϸ��ѹ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'PCT' where ����='ѪС��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'MPV' where ����='ƽ��ѪС�����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'PCT' where ����='ѪС��ѹ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'MCV' where ����='��ϸ��ƽ�����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'MCH' where ����='ƽ��Ѫ�쵰����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'MCHC' where ����='ƽ��Ѫ�쵰��Ũ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'NEUT' where ����='����ϸ������'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'LYMPH' where ����='�ܰ�ϸ������'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'MONO' where ����='����ϸ������'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'EO' where ����='��������ϸ������'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'BASO' where ����='�ȼ�����ϸ������'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'NEUT#' where ����='����ϸ����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'LYMPH#' where ����='�ܰ�ϸ����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'MONO#' where ����='����ϸ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'EO#' where ����='��������ϸ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'BASO#' where ����='�ȼ�����ϸ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'RDW-SD' where ����='��ϸ���ֲ����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'RDW-CV' where ����='��ϸ���ֲ���ȱ���ϵ��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'PDW' where ����='ѪС��ֲ����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'P-LCR' where ����='����ѪС�����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'TBIL' where ����='�ܵ�����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'DBIL' where ����='ֱ�ӵ�����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'IBIL' where ����='��ӵ�����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'ACT' where ����='�ȱ�ת��ø'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'TP' where ����='�ܵ���'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'ALB' where ����='�׵���'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'GLO' where ����='�򵰰�'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'ALB/GLO' where ����='�׵���/�򵰰�'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'ACP' where ����='��������ø'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'GLU' where ����='������'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'Urea' where ����='����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'GRZ' where ����='����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'CHO' where ����='���̴�'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'TG' where ����='��������'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'UA' where ����='����'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'AST' where ����='�Ȳ�ת��ø'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'AST/ALT' where ����='�Ȳݱȹȱ�'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'TBA' where ����='�ܵ�֭��'
update ְҵ�����_�����Ŀ���ñ� set ���� = 'GGT' where ����='r-�Ȱ���ת��ø'



--�ڿ�
declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '�ڿ�' and ID = @intID)
insert ְҵ�����_�����Ŀ���ñ� values('02017','����','','','����',@intInnerID,'','','',1,'')


--���
declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '���' and ID = @intID)
insert ְҵ�����_�����Ŀ���ñ� values('03024','ȫ��Ƥ��','����','����,�쳣','����',@intInnerID,'','','',1,'')

--Ѫ���滯���
declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = 'Ѫ���滯���' and ID = @intID)
insert ְҵ�����_�����Ŀ���ñ� values('04025','����','','','����',@intInnerID,'','','',1,'')
insert ְҵ�����_�����Ŀ���ñ� values('04026','���ص�','','','����',@intInnerID,'','','',1,'')


--���߿�
declare @intID int,@intInnerID int
set @intID = (select ID from [ϵͳ����_�ֵ�_�ֵ���б�] where ���� = 'ְҵ���������ֵ�')
set @intInnerID = (select InnerID from [ϵͳ����_�ֵ�_�ֵ����ݱ�] where ���� = '���߿�' and ID = @intID)
insert ְҵ�����_�����Ŀ���ñ� values('05006','SGPT','','','����',@intInnerID,'','','',1,'')
insert ְҵ�����_�����Ŀ���ñ� values('05007','TTT','','','����',@intInnerID,'','','',1,'')





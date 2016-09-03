----------------职业病体检 创建数据表------------------
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_业务设置信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_业务设置信息表]
GO

CREATE TABLE [dbo].[职业病体检_业务设置信息表] (
	[设置项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[设置值] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[枚举来源] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[说明] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_个人生活史表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_个人生活史表]
GO

CREATE TABLE [dbo].[职业病体检_个人生活史表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[初潮] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[周期] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[经期] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[末次月经] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[停经年龄] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[是否结婚] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[结婚日期] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[配偶接触放射] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[配偶职业] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[配偶健康状况] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[孕次] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[活产] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[早产] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[死产] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[自然流产] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[畸胎] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[多胎] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[异位妊娠] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[不孕不育原因] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[现有子女数目] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[子女健康状况] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[过敏史] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[吸烟程度] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[饮酒程度] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[烟龄] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[酒龄] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[戒烟时长] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[家族史] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检人员图片结果表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检人员图片结果表]
GO

CREATE TABLE [dbo].[职业病体检_体检人员图片结果表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[项目编号] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[图片] [image] NOT NULL ,
	[填写时间] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检人员基本信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检人员基本信息表]
GO

CREATE TABLE [dbo].[职业病体检_体检人员基本信息表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[公民身份号码] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[姓名] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[性别] [varchar] (4) COLLATE Chinese_PRC_CI_AS NULL ,
	[出生日期] [datetime] NULL ,
	[出生地] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[年龄] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[单位申请编号] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[单位名称] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[建档日期] [datetime] NULL ,
	[危害因素] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[职业分类] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[照射源] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[现工种] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[职务或职称] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[放射剂量] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[工龄] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[职业危害工龄] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[电话号码] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[住址] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[邮编] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[文化程度] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[籍贯] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[民族] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[婚否] [varchar] (5) COLLATE Chinese_PRC_CI_AS NULL ,
	[校核人] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[校核时间] [datetime] NULL ,
	[校核合格] [varchar] (2) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检人员附加项目设置表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检人员附加项目设置表]
GO

CREATE TABLE [dbo].[职业病体检_体检人员附加项目设置表] (
	[附加项目] [varchar] (30) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[录入标题] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[数据类型] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[数据长度] [int] NULL ,
	[枚举值] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检医师项目设置表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检医师项目设置表]
GO

CREATE TABLE [dbo].[职业病体检_体检医师项目设置表] (
	[医师编号] [UDT_员工编号] NOT NULL ,
	[体检项目] [varchar] (6) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检基本信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检基本信息表]
GO

CREATE TABLE [dbo].[职业病体检_体检基本信息表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[试管编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检表编号] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检类型] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检类别] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检日期] [datetime] NOT NULL ,
	[体检结论] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[诊断和处理意见] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[下结论医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[复查体检表编号] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[复查系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检状态] [varchar] (4) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[收费批号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[下结论日期] [datetime] NULL ,
	[收费金额] [money] NULL ,
	[各科体检状态] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检结果信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检结果信息表]
GO

CREATE TABLE [dbo].[职业病体检_体检结果信息表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (6) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (300) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写日期] [datetime] NULL ,
	[单项结论] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检结论判断条件表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检结论判断条件表]
GO

CREATE TABLE [dbo].[职业病体检_体检结论判断条件表] (
	[编号] [int] NOT NULL ,
	[体检结论] [int] NOT NULL ,
	[描述] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[序号] [int] NOT NULL ,
	[体检项目] [UDT_体检项目编号] NOT NULL ,
	[判断条件] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[标准值] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检表模板体检结论表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检表模板体检结论表]
GO

CREATE TABLE [dbo].[职业病体检_体检表模板体检结论表] (
	[体检表名称] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结论] [int] NOT NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检表模板体检项目表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检表模板体检项目表]
GO

CREATE TABLE [dbo].[职业病体检_体检表模板体检项目表] (
	[体检表名称] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检表模板基本信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检表模板基本信息表]
GO

CREATE TABLE [dbo].[职业病体检_体检表模板基本信息表] (
	[编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检表名称] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检类别] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[试管编号字母] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[诊断处理意见] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检单名称] [varchar] (60) COLLATE Chinese_PRC_CI_AS NULL ,
	[是否复查体检表] [smallint] NULL ,
	[代号] [varchar] (4) COLLATE Chinese_PRC_CI_AS NULL ,
	[收费标准] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检人员类型] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检表模板附加项目信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检表模板附加项目信息表]
GO

CREATE TABLE [dbo].[职业病体检_体检表模板附加项目信息表] (
	[体检表名称] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[附加项目] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[序号] [smallint] NOT NULL ,
	[是否必录] [varchar] (1) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检附加信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检附加信息表]
GO

CREATE TABLE [dbo].[职业病体检_体检附加信息表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[附加项目] [varchar] (30) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[项目值] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[项目值编号] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_体检项目设置表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_体检项目设置表]
GO

CREATE TABLE [dbo].[职业病体检_体检项目设置表] (
	[编码] [varchar] (6) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[名称] [varchar] (30) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[缺省值] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[枚举来源] [varchar] (60) COLLATE Chinese_PRC_CI_AS NULL ,
	[属性] [varchar] (4) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检大类] [int] NOT NULL ,
	[比较方式] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[标准值] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[单位] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[单价] [money] NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_职业病体检_用户操作权限表_职业病体检_可用操作信息表]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[职业病体检_用户操作权限表] DROP CONSTRAINT FK_职业病体检_用户操作权限表_职业病体检_可用操作信息表
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_可用操作信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_可用操作信息表]
GO

CREATE TABLE [dbo].[职业病体检_可用操作信息表] (
	[操作名] [UDT_操作名] NOT NULL ,
	[操作描述] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[上级操作名] [UDT_操作名] NULL ,
	[业务名] [UDT_业务名] NULL ,
	[部件名] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[类名] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[业务顺序] [int] NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_既往病史表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_既往病史表]
GO

CREATE TABLE [dbo].[职业病体检_既往病史表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[疾病名称] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[诊断日期] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[诊断单位] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[治疗经过] [varchar] (300) COLLATE Chinese_PRC_CI_AS NULL ,
	[转归] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_用户操作权限表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_用户操作权限表]
GO

CREATE TABLE [dbo].[职业病体检_用户操作权限表] (
	[用户编号] [UDT_员工编号] NOT NULL ,
	[权限名] [UDT_操作名] NOT NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_用户科室权限表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_用户科室权限表]
GO

CREATE TABLE [dbo].[职业病体检_用户科室权限表] (
	[用户编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[科室编号] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_科室结论表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_科室结论表]
GO

CREATE TABLE [dbo].[职业病体检_科室结论表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[科室] [varchar] (5) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[文字结论] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[医生编号] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[结论日期] [datetime] NULL ,
	[修改起始时间] [datetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_B超影像科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_B超影像科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_B超影像科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_X光影像科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_X光影像科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_X光影像科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_五官科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_五官科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_五官科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_内科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_内科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_内科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_外科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_外科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_外科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_尿常规化验科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_尿常规化验科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_尿常规化验科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_心电科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_心电科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_心电科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_染色体化验科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_染色体化验科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_染色体化验科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_电测听科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_电测听科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_电测听科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_肝功能化验科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_肝功能化验科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_肝功能化验科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_肺功能影像科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_肺功能影像科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_肺功能影像科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_血常规化验科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_血常规化验科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_血常规化验科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_职业史表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_职业史表]
GO

CREATE TABLE [dbo].[职业病体检_职业史表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[工作单位] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[部门] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[工种] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[危害种类] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[接触时间] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[防护措施] [varchar] (80) COLLATE Chinese_PRC_CI_AS NULL ,
	[备注] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[放射种类] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[每日工作量] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[累积照射量] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[过量照射史] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[起始时间] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[结束时间] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[是否放射性] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_自觉症状表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_自觉症状表]
GO

CREATE TABLE [dbo].[职业病体检_自觉症状表] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[症状] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[程度] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[出现时间] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_受检者个人信息录入科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_受检者个人信息录入科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_受检者个人信息录入科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO


----------------职业病体检 创建数据表（结束）------------------


----------------职业病体检 权限、项目等字典表、操作表设置------------------
-------------*****↓↓↓↓↓职业病体检-权限设置↓↓↓↓↓*****---------------

-----------添加数据表3个,专门控制职业病体检模块的科室和操作权限
-----↓↓↓职业病体检_用户科室权限表↓↓↓
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_用户科室权限表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_用户科室权限表]
GO

CREATE TABLE [dbo].[职业病体检_用户科室权限表] (
	[用户编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[科室编号] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO
-----↑↑↑职业病体检_用户科室权限表↑↑↑


-----↓↓↓职业病体检_用户操作权限表↓↓↓
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_用户操作权限表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_用户操作权限表]
GO

CREATE TABLE [dbo].[职业病体检_用户操作权限表] (
	[用户编号] [UDT_员工编号] NOT NULL ,
	[权限名] [UDT_操作名] NOT NULL 
) ON [PRIMARY]
GO
-----↑↑↑职业病体检_用户操作权限表↑↑↑


-----↓↓↓职业病体检_用户操作权限表↓↓↓
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_职业病体检_用户操作权限表_职业病体检_可用操作信息表]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[职业病体检_用户操作权限表] DROP CONSTRAINT FK_职业病体检_用户操作权限表_职业病体检_可用操作信息表
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_可用操作信息表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_可用操作信息表]
GO

CREATE TABLE [dbo].[职业病体检_可用操作信息表] (
	[操作名] [UDT_操作名] NOT NULL ,
	[操作描述] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[上级操作名] [UDT_操作名] NULL ,
	[业务名] [UDT_业务名] NULL ,
	[部件名] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[类名] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[业务顺序] [int] NULL 
) ON [PRIMARY]
GO
-----↑↑↑职业病体检_用户操作权限表↑↑↑

-----↓↓↓职业病体检_业务设置信息表↓↓↓
insert into 职业病体检_业务设置信息表 values('复查周期',12,null,null)
insert into 职业病体检_业务设置信息表 values('是否照相','否',null,null)
insert into 职业病体检_业务设置信息表 values('是否快速登记','是',null,null)
insert into 职业病体检_业务设置信息表 values('快速登记是否计费','否',null,null)
insert into 职业病体检_业务设置信息表 values('是否打印体检单','否',null,null)
insert into 职业病体检_业务设置信息表 values('是否打印体检表','否',null,null)
insert into 职业病体检_业务设置信息表 values('试管编号自动生成','否','是，否',null)
insert into 职业病体检_业务设置信息表 values('是否收费','是',null,null)
insert into 职业病体检_业务设置信息表 values('下一条码位置',0,15,'根据打印条码窗体确定')

--2012-06-06 于登淼 ↓
--记录最后一次统计时的基本数据
insert into 职业病体检_业务设置信息表 values('统计内容_按工种','统计类别_合格率','12','19')
insert into 职业病体检_业务设置信息表 values('统计内容-按体检情况','1','12','19')
--2012-06-06 于登淼 ↑
go
-----↑↑↑职业病体检_业务设置信息表↑↑↑


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '照射种类字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('照射种类字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '照射种类字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
go

--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '婚姻字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('婚姻字典','系统管理','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '婚姻字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','已婚','YH' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','未婚','WH' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'03','离异','LY' ,'',0)
go

--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '职业照射种类字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('职业照射种类字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '职业照射种类字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'11','核燃料循环','HRL' ,'',0)
--生成字典“核燃料循环”的二级内容。
select @llngParent=InnerID from 系统管理_字典_字典内容表 where ID=@llngID and 编号='11'
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1101','铀矿开采','YKKC' ,'',@llngParent)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1102','铀矿水冶','YKSY' ,'',@llngParent)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1103','燃料制造','RLZZ' ,'',@llngParent)
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'12','医学应用','YXYY' ,'',0)
select @llngParent=InnerID from 系统管理_字典_字典内容表 where ID=@llngID and 编号='12'
--生成字典“医学应用”的二级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1201','诊断放射学','ZKFSX' ,'',@llngParent)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1202','牙科放射学','' ,'',@llngParent)
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'13','工业应用','GYYY' ,'',0)
select @llngParent=InnerID from 系统管理_字典_字典内容表 where ID=@llngID and 编号='13'
--生成字典“工业应用”的二级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1301','工业辐照','GYFZ' ,'',@llngParent)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1302','工业控伤','GYKS' ,'',@llngParent)
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'14','天然源','TRY' ,'',0)
select @llngParent=InnerID from 系统管理_字典_字典内容表 where ID=@llngID and 编号='14'
--生成字典“天然源”的二级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1401','民用航空','MYHK' ,'',@llngParent)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1402','煤矿开采','MKKC' ,'',@llngParent)
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'15','其它','QT' ,'',0)
select @llngParent=InnerID from 系统管理_字典_字典内容表 where ID=@llngID and 编号='15'
--生成字典“其它”的二级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1501','教育','JY' ,'',@llngParent)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'1502','科学研究','KXYJ' ,'',@llngParent)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '体检人类别字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('体检人类别字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '体检人类别字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','普通体检','PTTJ' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','职业健康','ZYJK' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'03','放射健康','FSJK' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'04','涉核部队','SHBD' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'05','8023部队','8023' ,'',0)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '体检类型字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('体检类型字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '体检类型字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','上岗前','ZGQ' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','在岗期间','ZGQJ' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'03','离岗时','LGS' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'04','应急检查','YJJC' ,'',0)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '部门字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('部门字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '部门字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','生产部','SCB' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','管理部','GLB' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'03','市场部','SCB' ,'',0)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '工种字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('工种字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '工种字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','普通','PU' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','特种','TZ' ,'',0)
go

--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '放射线种类字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('放射线种类字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '放射线种类字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','X射线','' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','Y射线','' ,'',0)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '病情程度字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('病情程度字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '病情程度字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','轻微','QW' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','较明显','JMX' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'03','明显','MX' ,'',0)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '危害种类字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('危害种类字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '危害种类字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','噪声','ZS' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','强光','QQ' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'03','放射线','FSX' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'04','粉尘','FC' ,'',0)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '职业或职称字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('职业或职称字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '职业或职称字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','领导','LD' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','工程师','GCS' ,'',0)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '程度字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('程度字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '程度字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','从不','CB' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','偶尔','OE' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'03','长期','CQ' ,'',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'04','严重','YZ' ,'',0)
go


--获取字典ID。
Declare @llngID int
Declare @llngParent int
if not exists(select ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '职业病体检科室字典')
     insert into 系统管理_字典_字典表列表(名称,业务名,级别) values('职业病体检科室字典','职业病体检','操作级')
select @llngID = ID FROM 系统管理_字典_字典表列表 WHERE 名称 = '职业病体检科室字典'
--删除字典旧数据。
delete 系统管理_字典_字典内容表 where ID=@llngID
--生成字典一级内容。
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'01','五官科','WG' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'02','内科','LK' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'03','外科','WK' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'04','血常规化验科','XCG' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'05','肝功能化验科','GGN' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'06','尿常规化验科','LGN' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'07','染色体化验科','RST' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'08','电测听科','DCT' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'10','心电科','XD' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'11','B超影像科','BCYX' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'12','肺功能影像科','FGN' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'09','X光影像科','XGYX' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'14','体检登记','TJDJ' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'15','业务设置','YWSZ' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'13','受检者个人信息录入科','ZYBS' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'16','最终结论录入','ZZJL' ,'职业病体检_科室',0)
GO


-------------------------------------------
insert into 系统管理_系统安装信息表 values('职业病体检',1)
go

-------------------------------------------
insert into 系统管理_平台操作组分类表 values('0001','职业病体检','业务类')
go


-----------这部分是每个模块的总结点部分。子节点再根据情况手动加吧。
-----------“系统管理_可用操作信息表”中，也需要添加下面列出的窗体项。
insert into 系统管理_可用操作信息表 values ('职业病体检_体检登记',null,null,'职业病体检','职业病界面','clsmanagetestform',1)
insert into 系统管理_可用操作信息表 values ('职业病体检_五官科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',4)
insert into 系统管理_可用操作信息表 values ('职业病体检_业务设置',null,null,'职业病体检','职业病设置','clsmageconfform_zyb',3)
insert into 系统管理_可用操作信息表 values ('职业病体检_受检者个人信息录入科',null,null,'职业病体检','职业病史录入','clscareerhstmage',2)
insert into 系统管理_可用操作信息表 values ('职业病体检_内科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',5)
insert into 系统管理_可用操作信息表 values ('职业病体检_外科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',6)
insert into 系统管理_可用操作信息表 values ('职业病体检_血常规化验科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',7)
insert into 系统管理_可用操作信息表 values ('职业病体检_肝功能化验科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',8)
insert into 系统管理_可用操作信息表 values ('职业病体检_尿常规化验科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',9)
insert into 系统管理_可用操作信息表 values ('职业病体检_染色体化验科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',10)
insert into 系统管理_可用操作信息表 values ('职业病体检_电测听科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',11)
insert into 系统管理_可用操作信息表 values ('职业病体检_X光影像科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',12)
insert into 系统管理_可用操作信息表 values ('职业病体检_肺功能影像科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',13)
insert into 系统管理_可用操作信息表 values ('职业病体检_心电科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',14)
insert into 系统管理_可用操作信息表 values ('职业病体检_B超影像科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',15)
insert into 系统管理_可用操作信息表 values ('职业病体检_最终结论录入',null,null,'职业病体检','职业病体检结果录入','clscommon',16)
go

-----------“职业病体检_可用操作信息表”中，除去下面的，还要添加需要的子操作。
insert into 职业病体检_可用操作信息表 values ('职业病体检_体检登记',null,null,'职业病体检','职业病界面','clsmanagetestform',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_五官科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values ('职业病体检_业务设置',null,null,'职业病体检','职业病设置','clsmageconfform_zyb',3)
insert into 职业病体检_可用操作信息表 values ('职业病体检_受检者个人信息录入科',null,null,'职业病体检','职业病史录入','clscareerhstmage',2)

insert into 职业病体检_可用操作信息表 values ('职业病体检_内科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',5)
insert into 职业病体检_可用操作信息表 values ('职业病体检_外科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',6)
insert into 职业病体检_可用操作信息表 values ('职业病体检_血常规化验科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',7)
insert into 职业病体检_可用操作信息表 values ('职业病体检_肝功能化验科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',8)
insert into 职业病体检_可用操作信息表 values ('职业病体检_尿常规化验科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',9)
insert into 职业病体检_可用操作信息表 values ('职业病体检_染色体化验科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',10)
insert into 职业病体检_可用操作信息表 values ('职业病体检_电测听科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',11)
insert into 职业病体检_可用操作信息表 values ('职业病体检_X光影像科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',12)
insert into 职业病体检_可用操作信息表 values ('职业病体检_肺功能影像科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',13)
insert into 职业病体检_可用操作信息表 values ('职业病体检_心电科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',14)
insert into 职业病体检_可用操作信息表 values ('职业病体检_B超影像科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',15)
insert into 职业病体检_可用操作信息表 values ('职业病体检_最终结论录入',null,null,'职业病体检','职业病体检结果录入','clscommon',16)

--2012-04-24 陶露
insert into 职业病体检_可用操作信息表 values ('职业病体检_五官科结果录入_修改',null,'职业病体检_五官科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_B超影像科结果录入_修改',null,'职业病体检_B超影像科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_X光影像科结果录入_修改',null,'职业病体检_X光影像科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_电测听科结果录入_修改',null,'职业病体检_电测听科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_肺功能影像科结果录入_修改',null,'职业病体检_肺功能影像科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_肝功能化验科结果录入_修改',null,'职业病体检_肝功能化验科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_内科结果录入_修改',null,'职业病体检_内科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_尿常规化验科结果录入_修改',null,'职业病体检_尿常规化验科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_染色体化验科结果录入_修改',null,'职业病体检_染色体化验科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_外科结果录入_修改',null,'职业病体检_外科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_心电科结果录入_修改',null,'职业病体检_心电科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_血常规化验科结果录入_修改',null,'职业病体检_血常规化验科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_B超影像科结果录入_网络配置',null,'职业病体检_B超影像科结果录入','职业病体检','职业病体检结果录入','clscommon',3)
insert into 职业病体检_可用操作信息表 values ('职业病体检_B超影像科结果录入_删除',null,'职业病体检_B超影像科结果录入','职业病体检','职业病体检结果录入','clscommon',2)
insert into 职业病体检_可用操作信息表 values ('职业病体检_X光影像科结果录入_删除',null,'职业病体检_X光影像科结果录入','职业病体检','职业病体检结果录入','clscommon',2)
insert into 职业病体检_可用操作信息表 values ('职业病体检_X光影像科结果录入_网络配置',null,'职业病体检_X光影像科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_最终结论录入_保存结论',null,'职业病体检_最终结论录入','职业病体检','职业病界面入','clsmanagetestform',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_最终结论录入_打印报告',null,'职业病体检_最终结论录入','职业病体检','职业病界面入','clsmanagetestform',2)
--2012-04-24 陶露（注释结束）

--2012-05-23 翁乔 ↓
--职业病体检其他操作权限信息
insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_初检登记','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',2)
insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_复查登记','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',3)
insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_修改','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_删除','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',5)
insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_导出','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',6)
insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_单位导入','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',7)
--insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_打印新条码','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',8)
--insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_重新打条码','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',9)
insert into 职业病体检_可用操作信息表 values('职业病体检_体检登记_年检登记','','职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',10)

update 职业病体检_可用操作信息表 set 类名='clsmanagetestform' where 操作名='职业病体检_业务设置'

insert into 职业病体检_可用操作信息表 values('职业病体检_体检表设置','','','职业病体检','职业病设置','clsMageConfForm_zyb',1)
insert into 职业病体检_可用操作信息表 values('职业病体检_业务设置_体检表设置_新增','','职业病体检_业务设置_体检表设置','职业病体检','职业病设置','clsMageConfForm_zyb',2)
insert into 职业病体检_可用操作信息表 values('职业病体检_业务设置_体检表设置_保存','','职业病体检_业务设置_体检表设置','职业病体检','职业病设置','clsMageConfForm_zyb',3)
insert into 职业病体检_可用操作信息表 values('职业病体检_业务设置_体检表设置_复制','','职业病体检_业务设置_体检表设置','职业病体检','职业病设置','clsMageConfForm_zyb',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_业务设置_体检表设置_删除','','职业病体检_业务设置_体检表设置','职业病体检','职业病设置','clsMageConfForm_zyb',5)

insert into 职业病体检_可用操作信息表 values('职业病体检_受检者个人信息录入科','','','职业病体检','职业病史录入','clsCareerHstMage',1)
insert into 职业病体检_可用操作信息表 values('职业病体检_受检者个人信息录入科_职业病史登记','','职业病体检_受检者个人信息录入科','职业病体检','职业病史录入','clsCareerHstMage',2)
insert into 职业病体检_可用操作信息表 values('职业病体检_受检者个人信息录入科_修改','','职业病体检_受检者个人信息录入科','职业病体检','职业病史录入','clsCareerHstMage',3)
insert into 职业病体检_可用操作信息表 values('职业病体检_受检者个人信息录入科_删除','','职业病体检_受检者个人信息录入科','职业病体检','职业病史录入','clsCareerHstMage',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_受检者个人信息录入科_导出','','职业病体检_受检者个人信息录入科','职业病体检','职业病史录入','clsCareerHstMage',5)
insert into 职业病体检_可用操作信息表 values('职业病体检_受检者个人信息录入科_打印','','职业病体检_受检者个人信息录入科','职业病体检','职业病史录入','clsCareerHstMage',6)

insert into 职业病体检_可用操作信息表 values('职业病体检_职业病查询统计','','','职业病体检','职业病界面','clsmanagetestform',6)
insert into 职业病体检_可用操作信息表 values('职业病体检_职业病查询统计_导入Excel','','职业病体检_职业病查询统计','职业病体检','职业病界面','clsmanagetestform',6)
insert into 职业病体检_可用操作信息表 values('职业病体检_职业病查询统计_导出Excel','','职业病体检_职业病查询统计','职业病体检','职业病界面','clsmanagetestform',6)
insert into 职业病体检_可用操作信息表 values('职业病体检_职业病查询统计_打印','','职业病体检_职业病查询统计','职业病体检','职业病界面','clsmanagetestform',6)


--各科室结果批量录入权限信息
insert into 职业病体检_可用操作信息表 values('职业病体检_B超影像科结果录入_批量修改','','职业病体检_B超影像科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_X光影像科结果录入_批量修改','','职业病体检_X光影像科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_电测听科结果录入_批量修改','','职业病体检_电测听科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_肺功能影像科结果录入_批量修改','','职业病体检_肺功能影像科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_肝功能化验科结果录入_批量修改','','职业病体检_肝功能化验科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_内科结果录入_批量修改','','职业病体检_内科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_尿常规化验科结果录入_批量修改','','职业病体检_尿常规化验科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_染色体化验科结果录入_批量修改','','职业病体检_染色体化验科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_外科结果录入_批量修改','','职业病体检_外科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_五官科结果录入_批量修改','','职业病体检_五官科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_心电科结果录入_批量修改','','职业病体检_心电科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values('职业病体检_血常规化验科结果录入_批量修改','','','职业病体检','职业病体检结果录入','clscommon',4)
--2012-05-23 翁乔 ↑

--2012-06-15 于登淼 ↓
insert into 职业病体检_可用操作信息表 values ('职业病体检_体检登记_打印清单',null,'职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',8)
insert into 职业病体检_可用操作信息表 values ('职业病体检_体检登记_打印试管标签',null,'职业病体检_体检登记','职业病体检','职业病界面','clsmanagetestform',9)
insert into 职业病体检_可用操作信息表 values ('职业病体检_体检登记_初检登记_校核通过',null,'职业病体检_体检登记_初检登记','职业病体检','职业病界面','clsmanagetestform',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_最终结论录入_复核通过',null,'职业病体检_最终结论录入','职业病体检','职业病界面','clsmanagetestform',3)
--2012-06-15 于登淼 ↑
go

-----------添加人员管理中的科室管理(只有"职业病体检科"这一大类)
insert into 系统管理_字典_字典内容表(ID,编号,名称,助记符,描述,Parent) values(1,'06','职业病体检科','','',0)
go

-------------*****↑↑↑↑↑职业病体检-权限设置↑↑↑↑*****---------------
----------------职业病体检 权限、项目等字典表、操作表设置（结束）------------------



----------------职业病体检 所有体检项目---------------
-------------*****↓↓↓↓↓职业病体检-五官科-体检项目设置↓↓↓↓↓*****---------------
declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '五官科' and ID = @intID)

insert 职业病体检_体检项目设置表 values('01001','色觉-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01002','色觉-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01003','暗适应-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01004','暗适应-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01005','视野-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01006','视野-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01007','裸眼-远视力-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01008','裸眼-远视力-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01009','裸眼-近视力-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01010','裸眼-近视力-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01011','矫正-远视力-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01012','矫正-远视力-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01013','矫正-近视力-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01014','矫正-近视力-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01015','眼前部-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01016','眼前部-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01017','角膜-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01018','角膜-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01019','结膜-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01020','结膜-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01021','前房-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01022','前房-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01023','虹膜-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01024','虹膜-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01025','晶状体-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01026','晶状体-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01027','玻璃体-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01028','玻璃体-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01029','眼底-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01030','眼底-左','正常','正常,异常','常规',@intInnerID,'','','',1)
--insert 职业病体检_体检项目设置表 values('01031','晶状体环面及正面图-右','正常','正常,异常','常规',@intInnerID,'','','',1)
--insert 职业病体检_体检项目设置表 values('01032','晶状体环面及正面图-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01033','其它','正常','正常,异常','常规',@intInnerID,'','','',1)


--------------------------涉核部队
insert 职业病体检_体检项目设置表 values('01034','裸眼-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01035','裸眼-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01036','矫正-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01037','矫正-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01038','粉尘状-后囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01039','粉尘状-后囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01040','粉尘状-前囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01041','粉尘状-前囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01042','粉尘状-赤道-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01043','粉尘状-赤道-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01044','点状-后囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01045','点状-后囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01046','点状-前囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01047','点状-前囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01048','点状-赤道-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01049','点状-赤道-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01050','片状-后囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01051','片状-后囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01052','片状-前囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01053','片状-前囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01054','片状-赤道-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01055','片状-赤道-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01056','空泡-后囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01057','空泡-后囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01058','空泡-前囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01059','空泡-前囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01060','空泡-赤道-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01061','空泡-赤道-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01062','其它-后囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01063','其它-后囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01064','其它-前囊下-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01065','其它-前囊下-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01066','其它-赤道-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01067','其它-赤道-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01068','诊断','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01069','晶状体环面及正面图','正常','正常,异常','常规',@intInnerID,'','','',1)


-------------耳鼻喉科检查结果
insert 职业病体检_体检项目设置表 values('01070','听力','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01071','嗅觉','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01072','鼻','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01073','外耳','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01074','中耳','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01075','听力-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01076','听力-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01077','外耳道-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01078','外耳道-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01079','乳突-左','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01080','乳突-右','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01081','鼻-粘膜','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01082','鼻-出血','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01083','口腔-粘膜','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01084','口腔-牙齿','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01085','咽喉','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01086','口腔','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01087','耳鼻喉科其它','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('01088','其它','','','常规',@intInnerID,'','','',1,)
go

-------------*****↑↑↑↑↑职业病体检-五官科-体检项目设置↑↑↑↑↑*****---------------


   
   
   --作者：翁乔
   --时间：2012-04-11
-------------*****↓↓↓↓↓↓↓↓↓↓↓职业病体检-内科-体检项目设置↓↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '内科' and ID = @intID)

insert 职业病体检_体检项目设置表 values('02001','肺脏','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02002','心率','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02003','心律','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02004','心音','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02005','肝脏','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02006','脾脏','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02007','肾脏','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02008','痛觉、触觉','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02009','位置觉','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02010','腹壁反射','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02011','跟腱反射','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02012','肌力','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02013','肌张力','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02014','共济运动','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02015','三颤','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02016','病理反射','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('02018','其他','','','常规',@intInnerID,'','','',1,)
GO

-------------*****↑↑↑↑↑↑↑↑↑↑↑↑↑职业病体检-内科-体检项目设置↑↑↑↑↑↑↑↑↑↑↑↑↑*****---------------



-------------*****↓↓↓↓↓↓↓↓↓↓↓↓职业病体检-外科-体检项目设置↓↓↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '外科' and ID = @intID)

insert 职业病体检_体检项目设置表 values('03001','皮肤颜色','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03002','皮疹','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03003','瘀斑','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03004','紫癜','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03005','瘀点','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03006','浅表淋巴结','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03007','甲状腺','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03008','脊柱','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03009','四肢关节','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03010','脱发、脱毛','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03011','出血紫癜','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03012','干燥','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03013','脱屑','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03014','皲裂','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03015','色素沉着','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03016','色素减退','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03017','过度角化','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03018','多汗','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03019','疣状物','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03020','皮肤萎缩','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03021','溃疡','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03022','指甲','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03023','淋巴结','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('03024','其他','','','常规',@intInnerID,'','','',1)
GO

-------------*****↑↑↑↑↑↑↑↑↑↑↑↑↑职业病体检-外科-体检项目设置↑↑↑↑↑↑↑↑↑↑↑↑↑*****---------------

-------------*****↓↓↓↓↓↓↓↓↓↓↓↓职业病体检-血常规化验科-体检项目设置↓↓↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '血常规化验科' and ID = @intID)

insert 职业病体检_体检项目设置表 values('04001','白细胞*10E9/L','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04002','中性%','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04003','淋巴%','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04004','单核%','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04005','红细胞*10E12/L','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04006','血红蛋白g/L','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04007','血小板*10E9/L','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04008','中性杆状核粒细胞%','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04009','中性分叶核粒细胞%','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04010','嗜酸性粒细胞%','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04011','嗜碱性粒细胞%','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04012','肌酐','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04013','尿素氮','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04014','锌原卟啉','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04015','铅','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04016','胆碱酯酶(u)','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04017','血糖(mmol/L)','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('04027','其他','','','常规',@intInnerID,'','','',1)
GO

-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-血常规化验科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------

-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-肝功能化验科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '肝功能化验科' and ID = @intID)

insert 职业病体检_体检项目设置表 values('05001','SGPT','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('05002','HbsAg','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('05003','TTT','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('05008','其他','','','常规',@intInnerID,'','','',1)
GO

-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-肝功能化验科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------

-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-尿常规化验科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '尿常规化验科' and ID = @intID)

insert 职业病体检_体检项目设置表 values('06001','尿蛋白','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06002','尿糖','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06003','红细胞','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06004','白细胞','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06005','管型','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06006','铅','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06007','砷','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06008','镉','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06009','锰','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06010','氟','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06011','δ一氨基乙酰丙酸','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06012','β2一微球蛋白','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('06013','其他','','','常规',@intInnerID,'','','',1)
GO


-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-尿常规化验科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------


-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-染色体化验科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '染色体化验科' and ID = @intID)

insert 职业病体检_体检项目设置表 values('07001','分析中期分裂细胞数(个)','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('07002','染色体畸变率(%)','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('07003','畸变类型','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('07004','分析细胞数量','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('07005','微核淋巴细胞率(‰)','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('07006','淋巴细胞微核率(‰)','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('07007','其他','','','常规',@intInnerID,'','','',1)
GO


-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-染色体化验科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------

-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-电测听科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '电测听科' and ID = @intID)


insert 职业病体检_体检项目设置表 values('08001','电测听结果','正常','正常,异常','常规',@intInnerID,'','','',100)
insert 职业病体检_体检项目设置表 values('08002','其他','','','常规',@intInnerID,'','','',100)
go
-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-电测听科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------


-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-X光影像科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = 'X光影像科' and ID = @intID)


insert 职业病体检_体检项目设置表 values('09001','X光影像结果','正常','正常,异常','常规',@intInnerID,'','','',100)
insert 职业病体检_体检项目设置表 values('09002','其他','','','常规',@intInnerID,'','','',100)
go
-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-X光影像科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------


-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-心电科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '心电科' and ID = @intID)


insert 职业病体检_体检项目设置表 values('10001','心电结果','正常','正常,异常','常规',@intInnerID,'','','',100)
insert 职业病体检_体检项目设置表 values('10002','其他','','','常规',@intInnerID,'','','',100)
go
-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-心电科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------


-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-B超影像科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = 'B超影像科' and ID = @intID)


insert 职业病体检_体检项目设置表 values('11001','B超影像结果','正常','正常,异常','常规',@intInnerID,'','','',100)
insert 职业病体检_体检项目设置表 values('11002','其他','','','常规',@intInnerID,'','','',100)
go
-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-B超影像科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------


-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-肺功能影像科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------

declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '肺功能影像科' and ID = @intID)


insert 职业病体检_体检项目设置表 values('12001','肺功能影像结果','正常','正常,异常','常规',@intInnerID,'','','',100)
insert 职业病体检_体检项目设置表 values('12002','其他','','','常规',@intInnerID,'','','',100)
go
-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-肺功能影像科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------

-------------*****↓↓↓↓↓↓↓↓↓↓职业病体检-受检者个人信息录入科-体检项目设置↓↓↓↓↓↓↓↓↓↓*****---------------
--2012-06-13 于登淼 ↓
--省疾控要求添加体检体格一般情况（在职业病体检界面上）
declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '受检者个人信息录入科' and ID = @intID)

insert 职业病体检_体检项目设置表 values('13001','营养','良好','良好,中等,不良','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('13002','身高','170','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('13003','体重','60','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('13004','收缩压','120','正常,偏高,偏低','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('13005','舒张压','90','正常,偏高,偏低','常规',@intInnerID,'','','',1)
--2012-06-13 于登淼 ↑
-------------*****↑↑↑↑↑↑↑↑↑↑职业病体检-受检者个人信息录入科-体检项目设置↑↑↑↑↑↑↑↑↑↑*****---------------

--///////////////////////////////////////////////////////////////////////////////////--
--作者：翁乔
--时间：2012-04-11（注释结束）

----------------职业病体检 所有体检项目（结束）---------------


----------------职业病体检 模板与收费项初始化与设置------------------
-----------------------------
insert into 职业病体检_体检表模板基本信息表 values(1,'职业健康初检表','上岗前','B',' ','',0,1,'','职业健康')
insert into 职业病体检_体检表模板基本信息表 values(2,'职业健康在岗检查表','在岗期间','B',' ','',0,1,'','职业健康')
insert into 职业病体检_体检表模板基本信息表 values(3,'职业健康离岗检查表','离岗时','B',' ','',0,1,'','职业健康')
go

-----------------------------
insert into 收费管理_收费项目字典表 values('005','职业病体检',0,'人',871,0,0,'')
go

----------------职业病体检 模板与收费项初始化与设置（结束）------------------


----------------职业病体检 体检项目表的增改 （2012-08-21；翁乔）------------------
新增数据表
--生化科
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_生化科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_生化科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_生化科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO
--删除肝功能化验科
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_肝功能化验科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_肝功能化验科]
GO
--新增 免疫科
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_结果信息_免疫科]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[职业病体检_结果信息_免疫科]
GO

CREATE TABLE [dbo].[职业病体检_结果信息_免疫科] (
	[系统编号] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检项目] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[体检结果] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[体检医师] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[填写时间] [datetime] NULL ,
	[单项结论] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO




--新增数据	科室
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'05','免疫科','MY' ,'职业病体检_科室',0)
insert into 系统管理_字典_字典内容表(ID, 编号, 名称, 助记符, 描述, Parent) values(@llngID ,'17','生化科','SH' ,'职业病体检_科室',0)

--新增数据	项目

--免疫科
declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '免疫科' and ID = @intID)
insert 职业病体检_体检项目设置表 values('05001','HbsAb','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('05003','HbeAg','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('05004','HbeAb','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('05005','HbcAb','正常','正常,异常','常规',@intInnerID,'','','',1)
insert 职业病体检_体检项目设置表 values('05008','其他','','','常规',@intInnerID,'','','',1)

--生化科
declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '生化科' and ID = @intID)
insert 职业病体检_体检项目设置表 values('17001','总胆红素','','','化验',@intInnerID,'=','11.60','umol/L',1)
insert 职业病体检_体检项目设置表 values('17002','直接胆红素','','','化验',@intInnerID,'=','4.00','umol/L',1)
insert 职业病体检_体检项目设置表 values('17003','间接胆红素','','','化验',@intInnerID,'=','7.60','umol/L',1)
insert 职业病体检_体检项目设置表 values('17004','谷丙转氨酶','','','化验',@intInnerID,'=','52.00','U/L',1)
insert 职业病体检_体检项目设置表 values('17005','谷草转氨酶','','','化验',@intInnerID,'=','30.00','U/L',1)
insert 职业病体检_体检项目设置表 values('17006','谷草比谷丙','','','化验',@intInnerID,'=','0.57','',1)
insert 职业病体检_体检项目设置表 values('17007','总蛋白','','','化验',@intInnerID,'=','91.00','g/L',1)
insert 职业病体检_体检项目设置表 values('17008','白蛋白','','','化验',@intInnerID,'=','53.00','g/L',1)
insert 职业病体检_体检项目设置表 values('17009','球蛋白','','','化验',@intInnerID,'=','38.00','g/L',1)
insert 职业病体检_体检项目设置表 values('17010','白蛋白/球蛋白','','','化验',@intInnerID,'=','1.39','',1)
insert 职业病体检_体检项目设置表 values('17011','总胆汁酸','','','化验',@intInnerID,'=','2.00','umol/L',1)
insert 职业病体检_体检项目设置表 values('17012','r-谷氨酰转肽酶','','','化验',@intInnerID,'=','30.00','U/L',1)
insert 职业病体检_体检项目设置表 values('17013','碱性磷酸酶','','','化验',@intInnerID,'=','68.00','U/L',1)
insert 职业病体检_体检项目设置表 values('17014','尿素','','','化验',@intInnerID,'=','3.55','mmol/L',1)
insert 职业病体检_体检项目设置表 values('17015','肌酐','','','化验',@intInnerID,'=','138.20','umol/L',1)
insert 职业病体检_体检项目设置表 values('17016','葡萄糖','','','化验',@intInnerID,'=','10','mmol/L',1)
insert 职业病体检_体检项目设置表 values('17017','胆固醇','','','化验',@intInnerID,'=','10','mmol/L',1)
insert 职业病体检_体检项目设置表 values('17018','甘油三酯','','','化验',@intInnerID,'=','10','mmol/L',1)
insert 职业病体检_体检项目设置表 values('17019','尿酸','','','化验',@intInnerID,'=','10','mmol/L',1)
insert 职业病体检_体检项目设置表 values('17020','其他','','','化验',@intInnerID,'','','',1)


--操作权限添加
insert into 系统管理_可用操作信息表 values ('职业病体检_生化科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',19)
insert into 系统管理_可用操作信息表 values ('职业病体检_生化科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',20)
go

--子级权限
insert into 职业病体检_可用操作信息表 values ('职业病体检_生化科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',17)
insert into 职业病体检_可用操作信息表 values ('职业病体检_免疫科结果录入',null,null,'职业病体检','职业病体检结果录入','clscommon',18)

insert into 职业病体检_可用操作信息表 values ('职业病体检_生化科结果录入_修改',null,'职业病体检_生化科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
insert into 职业病体检_可用操作信息表 values ('职业病体检_免疫科结果录入_修改',null,'职业病体检_免疫科结果录入','职业病体检','职业病体检结果录入','clscommon',1)
--批量录入
insert into 职业病体检_可用操作信息表 values ('职业病体检_生化科结果录入_批量修改',null,'职业病体检_生化科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
insert into 职业病体检_可用操作信息表 values ('职业病体检_免疫科结果录入_批量修改',null,'职业病体检_免疫科结果录入','职业病体检','职业病体检结果录入','clscommon',4)
go



--修改 血常规化验科的项目数据
insert 职业病体检_体检项目设置表 values('04001','红细胞压积','','','常规',@intInnerID,'=','45.80','%',1)
insert 职业病体检_体检项目设置表 values('04002','平均血小板体积','','','常规',@intInnerID,'=','11.02','fL',1)
insert 职业病体检_体检项目设置表 values('04003','血小板压积','','','常规',@intInnerID,'=','0.32','%',1)
insert 职业病体检_体检项目设置表 values('04004','红细胞平均体积','','','常规',@intInnerID,'=','91.20','fL',1)
insert 职业病体检_体检项目设置表 values('04005','平均血红蛋白量','','','常规',@intInnerID,'=','28.30','pg',1)
insert 职业病体检_体检项目设置表 values('04006','平均血红蛋白浓度','','','常规',@intInnerID,'=','310.00','g/L',1)
insert 职业病体检_体检项目设置表 values('04007','中性细胞比率','','','常规',@intInnerID,'=','64.20','%',1)
insert 职业病体检_体检项目设置表 values('04008','淋巴细胞比率','','','常规',@intInnerID,'=','28.70','%',1)
insert 职业病体检_体检项目设置表 values('04009','单核细胞比率','','','常规',@intInnerID,'=','5.60','%',1)
insert 职业病体检_体检项目设置表 values('04010','嗜酸性粒细胞比率','','','常规',@intInnerID,'=','1.20','%',1)
insert 职业病体检_体检项目设置表 values('04011','嗜碱性粒细胞比率','','','常规',@intInnerID,'=','0.30','%',1)
insert 职业病体检_体检项目设置表 values('04012','中性细胞数','','','常规',@intInnerID,'=','3.77','10^9/L',1)
insert 职业病体检_体检项目设置表 values('04013','淋巴细胞数','','','常规',@intInnerID,'=','1.69','10^9/L',1)
insert 职业病体检_体检项目设置表 values('04014','单核细胞','','','常规',@intInnerID,'=','0.33','10^9/L',1)
insert 职业病体检_体检项目设置表 values('04015','嗜酸性粒细胞','','','常规',@intInnerID,'=','0.07','10^9/L',1)
insert 职业病体检_体检项目设置表 values('04016','嗜碱性粒细胞','','','常规',@intInnerID,'=','0.02','10^9/L',1)
insert 职业病体检_体检项目设置表 values('04017','红细胞分布宽度','','','常规',@intInnerID,'=','44.00','fL',1)
insert 职业病体检_体检项目设置表 values('04018','红细胞分布宽度变异系数','','','常规',@intInnerID,'=','13.40','%',1)
insert 职业病体检_体检项目设置表 values('04019','血小板分布宽度','','','常规',@intInnerID,'=','11.02','fL',1)
insert 职业病体检_体检项目设置表 values('04020','大型血小板比率','','','常规',@intInnerID,'=','32.12','%',1)
insert 职业病体检_体检项目设置表 values('04021','白细胞','','','常规',@intInnerID,'=','5.88','10^9/L',1)
insert 职业病体检_体检项目设置表 values('04022','红细胞','','','常规',@intInnerID,'=','5.02','10^12/L',1)
insert 职业病体检_体检项目设置表 values('04023','血红蛋白','','','常规',@intInnerID,'=','142.00','g/L',1)
insert 职业病体检_体检项目设置表 values('04024','血小板','','','常规',@intInnerID,'=','143.00','10^9/L',1)

--修改 职业病体检_体检项目设置表 增加字段（代号）

alter table 职业病体检_体检项目设置表 add 代号 varchar(10) null

--更新体检项目信息	（代号）
update 职业病体检_体检项目设置表 set 代号 = 'WBC' where 名称='白细胞'
update 职业病体检_体检项目设置表 set 代号 = 'RBC' where 名称='红细胞'
update 职业病体检_体检项目设置表 set 代号 = 'HGB' where 名称='血红蛋白'
update 职业病体检_体检项目设置表 set 代号 = 'HCT' where 名称='红细胞压积'
update 职业病体检_体检项目设置表 set 代号 = 'PCT' where 名称='血小板'
update 职业病体检_体检项目设置表 set 代号 = 'MPV' where 名称='平均血小板体积'
update 职业病体检_体检项目设置表 set 代号 = 'PCT' where 名称='血小板压积'
update 职业病体检_体检项目设置表 set 代号 = 'MCV' where 名称='红细胞平均体积'
update 职业病体检_体检项目设置表 set 代号 = 'MCH' where 名称='平均血红蛋白量'
update 职业病体检_体检项目设置表 set 代号 = 'MCHC' where 名称='平均血红蛋白浓度'
update 职业病体检_体检项目设置表 set 代号 = 'NEUT' where 名称='中性细胞比率'
update 职业病体检_体检项目设置表 set 代号 = 'LYMPH' where 名称='淋巴细胞比率'
update 职业病体检_体检项目设置表 set 代号 = 'MONO' where 名称='单核细胞比率'
update 职业病体检_体检项目设置表 set 代号 = 'EO' where 名称='嗜酸性粒细胞比率'
update 职业病体检_体检项目设置表 set 代号 = 'BASO' where 名称='嗜碱性粒细胞比率'
update 职业病体检_体检项目设置表 set 代号 = 'NEUT#' where 名称='中性细胞数'
update 职业病体检_体检项目设置表 set 代号 = 'LYMPH#' where 名称='淋巴细胞数'
update 职业病体检_体检项目设置表 set 代号 = 'MONO#' where 名称='单核细胞'
update 职业病体检_体检项目设置表 set 代号 = 'EO#' where 名称='嗜酸性粒细胞'
update 职业病体检_体检项目设置表 set 代号 = 'BASO#' where 名称='嗜碱性粒细胞'
update 职业病体检_体检项目设置表 set 代号 = 'RDW-SD' where 名称='红细胞分布宽度'
update 职业病体检_体检项目设置表 set 代号 = 'RDW-CV' where 名称='红细胞分布宽度变异系数'
update 职业病体检_体检项目设置表 set 代号 = 'PDW' where 名称='血小板分布宽度'
update 职业病体检_体检项目设置表 set 代号 = 'P-LCR' where 名称='大型血小板比率'
update 职业病体检_体检项目设置表 set 代号 = 'TBIL' where 名称='总胆红素'
update 职业病体检_体检项目设置表 set 代号 = 'DBIL' where 名称='直接胆红素'
update 职业病体检_体检项目设置表 set 代号 = 'IBIL' where 名称='间接胆红素'
update 职业病体检_体检项目设置表 set 代号 = 'ACT' where 名称='谷丙转氨酶'
update 职业病体检_体检项目设置表 set 代号 = 'TP' where 名称='总蛋白'
update 职业病体检_体检项目设置表 set 代号 = 'ALB' where 名称='白蛋白'
update 职业病体检_体检项目设置表 set 代号 = 'GLO' where 名称='球蛋白'
update 职业病体检_体检项目设置表 set 代号 = 'ALB/GLO' where 名称='白蛋白/球蛋白'
update 职业病体检_体检项目设置表 set 代号 = 'ACP' where 名称='碱性磷酸酶'
update 职业病体检_体检项目设置表 set 代号 = 'GLU' where 名称='葡萄糖'
update 职业病体检_体检项目设置表 set 代号 = 'Urea' where 名称='尿素'
update 职业病体检_体检项目设置表 set 代号 = 'GRZ' where 名称='肝酐'
update 职业病体检_体检项目设置表 set 代号 = 'CHO' where 名称='胆固醇'
update 职业病体检_体检项目设置表 set 代号 = 'TG' where 名称='甘油三酯'
update 职业病体检_体检项目设置表 set 代号 = 'UA' where 名称='尿酸'
update 职业病体检_体检项目设置表 set 代号 = 'AST' where 名称='谷草转氨酶'
update 职业病体检_体检项目设置表 set 代号 = 'AST/ALT' where 名称='谷草比谷丙'
update 职业病体检_体检项目设置表 set 代号 = 'TBA' where 名称='总胆汁酸'
update 职业病体检_体检项目设置表 set 代号 = 'GGT' where 名称='r-谷氨酰转肽酶'



--内科
declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '内科' and ID = @intID)
insert 职业病体检_体检项目设置表 values('02017','体温','','','常规',@intInnerID,'','','',1,'')


--外科
declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '外科' and ID = @intID)
insert 职业病体检_体检项目设置表 values('03024','全身皮肤','正常','正常,异常','常规',@intInnerID,'','','',1,'')

--血常规化验科
declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '血常规化验科' and ID = @intID)
insert 职业病体检_体检项目设置表 values('04025','肌酐','','','常规',@intInnerID,'','','',1,'')
insert 职业病体检_体检项目设置表 values('04026','尿素氮','','','常规',@intInnerID,'','','',1,'')


--免疫科
declare @intID int,@intInnerID int
set @intID = (select ID from [系统管理_字典_字典表列表] where 名称 = '职业病体检科室字典')
set @intInnerID = (select InnerID from [系统管理_字典_字典内容表] where 名称 = '免疫科' and ID = @intID)
insert 职业病体检_体检项目设置表 values('05006','SGPT','','','常规',@intInnerID,'','','',1,'')
insert 职业病体检_体检项目设置表 values('05007','TTT','','','常规',@intInnerID,'','','',1,'')





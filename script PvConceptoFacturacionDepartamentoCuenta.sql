if exists (select * from sysobjects where id = object_id(N'[dbo].[PvConceptoFacturacionDepartamentoCuenta]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PvConceptoFacturacionDepartamentoCuenta]
GO

CREATE TABLE [dbo].[PvConceptoFacturacionDepartamentoCuenta] (
	[smiCveConcepto] [smallint] NOT NULL ,
	[smiCveDepartamento] [smallint] NOT NULL ,
	[intNumCuentaIngreso] [int] NOT NULL ,
	[intNumCuentaDescuento] [int] NOT NULL ,
	[intNumCuentaIVA] [int] NOT NULL 
) ON [PRIMARY]
GO


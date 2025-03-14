USE SIIAPP
GO

IF DB_NAME() <> N'SIIAPP' SET NOEXEC ON
GO

--
-- Create table [dbo].[SIIAPP_Des]
--
PRINT (N'Create table [dbo].[SIIAPP_Des]')
GO
CREATE TABLE dbo.SIIAPP_Des (
  UUID bigint IDENTITY,
  N_control varchar(50) NOT NULL,
  Fecha_Soli date NULL,
  N_cotizacion varchar(50) NULL,
  PT varchar(50) NULL,
  Producto varchar(50) NULL,
  Cliente varchar(50) NULL,
  Notif_Sanitaria varchar(50) NULL,
  Estado_M varchar(50) NULL,
  Dis_Des varchar(50) NULL,
  Obs varchar(50) NULL,
  CONSTRAINT PK_SIIAPP_Des PRIMARY KEY CLUSTERED (UUID, N_control)
)
ON [PRIMARY]
GO

--
-- Create index [KEY_SIIAPP_Des_N_control] on table [dbo].[SIIAPP_Des]
--
PRINT (N'Create index [KEY_SIIAPP_Des_N_control] on table [dbo].[SIIAPP_Des]')
GO
CREATE UNIQUE INDEX KEY_SIIAPP_Des_N_control
  ON dbo.SIIAPP_Des (N_control)
  ON [PRIMARY]
GO

--
-- Create table [dbo].[SIIAPP_Prod]
--
PRINT (N'Create table [dbo].[SIIAPP_Prod]')
GO
CREATE TABLE dbo.SIIAPP_Prod (
  UUID bigint IDENTITY,
  N_Control varchar(50) NOT NULL,
  Cod_Caja varchar(50) NULL,
  Caja_Embalaje varchar(50) NULL,
  FB_Dimensiones_Adec varchar(50) NULL,
  FB_Tunel_Temp varchar(50) NULL,
  FB_Tunel_Speed varchar(50) NULL,
  FB_Test_Tool varchar(50) NULL,
  FB_obs varchar(50) NULL,
  SCE_Temp varchar(50) NULL,
  SCE_Temp_Time varchar(50) NULL,
  SCE_Obs varchar(50) NULL,
  SCE_Hermt varchar(50) NULL,
  Ubi_Lote varchar(50) NULL,
  Etiqueta_Manual varchar(50) NULL,
  Dis_Prod varchar(50) NULL,
  Metod_Fab varchar(50) NULL,
  Diligenciado varchar(50) NULL,
  CONSTRAINT PK_SIIAPP_Prod_UUID PRIMARY KEY CLUSTERED (UUID)
)
ON [PRIMARY]
GO

--
-- Create foreign key [FK_SIIAPP_Prod_N_Control] on table [dbo].[SIIAPP_Prod]
--
PRINT (N'Create foreign key [FK_SIIAPP_Prod_N_Control] on table [dbo].[SIIAPP_Prod]')
GO
ALTER TABLE dbo.SIIAPP_Prod
  ADD CONSTRAINT FK_SIIAPP_Prod_N_Control FOREIGN KEY (N_Control) REFERENCES dbo.SIIAPP_Des (N_control)
GO

--
-- Create table [dbo].[SIIAPP_Cal]
--
PRINT (N'Create table [dbo].[SIIAPP_Cal]')
GO
CREATE TABLE dbo.SIIAPP_Cal (
  UUID bigint IDENTITY,
  N_Control varchar(50) NOT NULL,
  Mesofilos varchar(50) NULL,
  E_Coli varchar(50) NULL,
  S_Aureus varchar(50) NULL,
  P_Aureoginosa varchar(50) NULL,
  Moho_Levadura varchar(50) NULL,
  Micro_Obs varchar(50) NULL,
  Cumple_Microbiologia varchar(50) NULL,
  Densidad varchar(50) NULL,
  Ph varchar(50) NULL,
  Contenido varchar(50) NULL,
  Color varchar(50) NULL,
  Textura varchar(50) NULL,
  Olor varchar(50) NULL,
  Viscosidad varchar(50) NULL,
  Rpm varchar(50) NULL,
  Aguja varchar(50) NULL,
  Torque varchar(50) NULL,
  Apt_Container varchar(50) NULL,
  Cont_Obs varchar(50) NULL,
  Diligenciado varchar(50) NULL,
  CONSTRAINT PK_SIIAPP_Cal_UUID PRIMARY KEY CLUSTERED (UUID)
)
ON [PRIMARY]
GO

--
-- Create foreign key [FK_SIIAPP_Cal_N_Control] on table [dbo].[SIIAPP_Cal]
--
PRINT (N'Create foreign key [FK_SIIAPP_Cal_N_Control] on table [dbo].[SIIAPP_Cal]')
GO
ALTER TABLE dbo.SIIAPP_Cal
  ADD CONSTRAINT FK_SIIAPP_Cal_N_Control FOREIGN KEY (N_Control) REFERENCES dbo.SIIAPP_Des (N_control)
GO

--
-- Create table [dbo].[SIIAPP_Bod]
--
PRINT (N'Create table [dbo].[SIIAPP_Bod]')
GO
CREATE TABLE dbo.SIIAPP_Bod (
  UUID bigint IDENTITY,
  N_Control varchar(50) NOT NULL,
  Bottle_Cod varchar(50) NULL,
  Envase varchar(50) NULL,
  Bottle_Sum varchar(50) NULL,
  Bottle_Color varchar(50) NULL,
  Bottle_Material varchar(50) NULL,
  Bottle_ML varchar(50) NULL,
  Cap_Cod varchar(50) NULL,
  Tapa varchar(50) NULL,
  Cap_Sum varchar(50) NULL,
  Cap_Color varchar(50) NULL,
  Cap_Material varchar(50) NULL,
  Foil_Cod varchar(50) NULL,
  Foil varchar(50) NULL,
  Foil_Type varchar(50) NULL,
  FB_Cod varchar(50) NULL,
  Funda_Banda varchar(50) NULL,
  Ubicacion_Termoduc varchar(50) NULL,
  Etiq_Cod varchar(50) NULL,
  Etiqueta varchar(50) NULL,
  Box_Cod varchar(50) NULL,
  Box_Folding varchar(50) NULL,
  Box_Sum varchar(50) NULL,
  Dis_Bod varchar(50) NULL,
  CONSTRAINT PK_SIIAPP_Bod_UUID PRIMARY KEY CLUSTERED (UUID)
)
ON [PRIMARY]
GO

--
-- Create foreign key [FK_SIIAPP_Bod_N_Control] on table [dbo].[SIIAPP_Bod]
--
PRINT (N'Create foreign key [FK_SIIAPP_Bod_N_Control] on table [dbo].[SIIAPP_Bod]')
GO
ALTER TABLE dbo.SIIAPP_Bod
  ADD CONSTRAINT FK_SIIAPP_Bod_N_Control FOREIGN KEY (N_Control) REFERENCES dbo.SIIAPP_Des (N_control)
GO
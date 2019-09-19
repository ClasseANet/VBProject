USE G3R
Go
CREATE DEFAULT [dbo].[DF_Z] AS 'Z'
go
CREATE DEFAULT [dbo].[DF_Y] AS 'Y' 
go
CREATE DEFAULT [dbo].[DF_X] AS 'X' 
go
CREATE DEFAULT [dbo].[DF_W] AS 'W' 
go
CREATE DEFAULT [dbo].[DF_V] AS 'V' 
go
CREATE DEFAULT [dbo].[DF_U] AS 'U' 
go
CREATE DEFAULT [dbo].[DF_T] AS 'T' 
go
CREATE DEFAULT [dbo].[DF_S] AS 'S' 
go
CREATE DEFAULT [dbo].[DF_R] AS 'R' 
go
CREATE DEFAULT [dbo].[DF_Q] AS 'Q' 
go
CREATE DEFAULT [dbo].[DF_P] AS 'P' 
go
CREATE DEFAULT [dbo].[DF_O] AS 'O' 
go
CREATE DEFAULT [dbo].[DF_N] AS 'N' 
go
CREATE DEFAULT [dbo].[DF_M] AS 'M' 
go
CREATE DEFAULT [dbo].[DF_L] AS 'L' 
go
CREATE DEFAULT [dbo].[DF_K] AS 'K' 
go
CREATE DEFAULT [dbo].[DF_J] AS 'J' 
go
CREATE DEFAULT [dbo].[DF_I] AS 'I' 
go
CREATE DEFAULT [dbo].[DF_H] AS 'H' 
go
CREATE DEFAULT [dbo].[DF_G] AS 'G' 
go
CREATE DEFAULT [dbo].[DF_F] AS 'F' 
go
CREATE DEFAULT [dbo].[DF_E] AS 'E' 
go
CREATE DEFAULT [dbo].[DF_D] AS 'D' 
go
CREATE DEFAULT [dbo].[DF_C] AS 'C' 
go
CREATE DEFAULT [dbo].[DF_B] AS 'B' 
go
CREATE DEFAULT [dbo].[DF_A] AS 'A' 
go
CREATE DEFAULT [dbo].[DF_9] AS '9' 
go
CREATE DEFAULT [dbo].[DF_8] AS '8' 
go
CREATE DEFAULT [dbo].[DF_7] AS '7' 
go
CREATE DEFAULT [dbo].[DF_6] AS '6' 
go
CREATE DEFAULT [dbo].[DF_5] AS '5' 
go
CREATE DEFAULT [dbo].[DF_4] AS '4' 
go
CREATE DEFAULT [dbo].[DF_3] AS '3' 
go
CREATE DEFAULT [dbo].[DF_2] AS '2' 
go
CREATE DEFAULT [dbo].[DF_1] AS '1' 
go
CREATE DEFAULT [dbo].[DF_0] AS '0' 
go

CREATE DEFAULT [dbo].[DF_NowTIME] AS Convert(varchar(5), GetDate(),108) 
go
CREATE DEFAULT [dbo].[DF_NowDate] AS Cast(Convert(varchar, GetDate(),111) + ' 00:00:00' as datetime) 
go
CREATE DEFAULT [dbo].[DF_Now] AS GetDate() 
go
CREATE DEFAULT [dbo].[DF_NewID]	AS NewId() 
go

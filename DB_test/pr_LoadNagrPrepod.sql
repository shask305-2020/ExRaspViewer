﻿CREATE PROC [dbo].[pr_LoadNagrPrepod]
@IDP INT
AS
BEGIN
	SELECT GRUP.[NAIM] AS 'Группа', PRED.[NAIM] AS 'Дисциплина', [CHSEM] AS 'Час. в сем.' 
	FROM [dbo].[SPNAGR] AS NAGR
	INNER JOIN
	[dbo].[SPGRUP] AS GRUP
	ON NAGR.IDG = GRUP.IDG
	INNER JOIN
	[dbo].[SPPRED] AS PRED
	ON NAGR.IDD = PRED.IDD
	WHERE NAGR.[IDP] = @IDP
	ORDER BY GRUP.[NAIM];
END
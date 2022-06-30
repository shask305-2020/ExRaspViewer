SELECT
[DAT] AS 'Дата',
Prep.[FAMIO] AS 'Преподаватель',
Gr.[NAIM] AS 'Группа',
[IDD] AS 'Предмет'
FROM [dbo].[UROKI] AS Ur
INNER JOIN
[dbo].[SPPREP] AS Prep
ON Ur.IDP = Prep.IDP
INNER JOIN
[dbo].[SPGRUP] AS Gr
ON Ur.IDG = Gr.IDG
WHERE [DAT] = '2022-01-10'
AND Gr.[NAIM] = '111';
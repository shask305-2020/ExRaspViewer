SELECT
[DAT] AS 'Дата',
[IDP] AS 'Преподаватель',
[IDG] AS 'Группа',
[IDD] AS 'Предмет'
FROM [dbo].[UROKI]
WHERE [DAT] = '2022-01-10'
ORDER BY [IDP];
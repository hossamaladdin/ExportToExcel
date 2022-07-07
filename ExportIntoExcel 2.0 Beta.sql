EXEC sp_configure 'show advanced options',1 reconfigure
EXEC sp_configure 'xp_cmdshell',1 reconfigure

drop table if exists ##SQLCMD
SELECT 'SELECT @@SERVERNAME Server,HOST_NAME() Host' SQLStatement INTO ##SQLCMD

EXEC xp_cmdshell 'bcp "SELECT SQLStatement FROM ##SQLCMD" queryout %temp%\SQLCMD.sql -T -c',no_output

EXEC xp_cmdshell 'powershell "Invoke-Sqlcmd -InputFile %temp%\SQLCMD.sql -ServerInstance localhost | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel %temp%\report11.xlsx -autosize -tablestyle light7"'
EXEC xp_cmdshell 'del %temp%\SQLCMD.sql',no_output


exec dbo.UploadToSharePoint @username='h.alaa@iyelo.com',@password='H0$$@m#55481',@sharepointurl='https://wefaqdollar-my.sharepoint.com/personal/h_alaa_iyelo_com',@library='attachments',@filename='c:\users\sqlagsvc\appdata\local\temp\report11.xlsx'

USE [master]
GO

ALTER PROC [dbo].[ExportToExcel2]

/*Created by Hossam Alaa hossam.alaaddin@gmail.com https://github.com/hossamaladdin https://www.linkedin.com/in/hossamaladdin */
/*With the help of ImportExcel from https://www.powershellgallery.com/packages/ImportExcel written by Doug Finke https://github.com/dfinke */
/*And with using sharepoint PS scripts written by Salaudeen Rajack  https://www.sharepointdiary.com/ */
/*Introducing Export to Excel SQL 2019 Proc V1.0 */
/*for usage, please refer to ExporttoExcel_1.0_Example.sql*/

	@databasename varchar(20) = NULL /*optional, don't specify if temp table*/
	,@schemaname varchar(20) = NULL	/*optional, don't specify if temp table*/
	,@tablename varchar(20)	= NULL  /*specify table name only without schema name, db name, or an existing GLOBAL temp table, like ##SomeTable*/
	,@Query nvarchar(max)  = NULL	/*specify query to execute*/
	,@SQLSERVER varchar(100) = 'localhost' /*specify which server/instacne to get the results from*/
	,@filepath varchar(100) = NULL	/*optional, the default path is the temp folder %Temp%*/
	,@filename varchar(100)	= NULL	/*specify file name you want to export, default = tablename*/
	,@fullname varchar(100) = NULL output	/*an output parameter if you are going to use the file in sending an email or for any other purpose*/
	,@CSVOnly bit = 0					/*only use csv to attach or upload*/
	,@params varchar(100) = '-AutoSize -TableName table1'		/*optional, specify parameters of Export-Excel cmdlet, like -AutoSize or -TableName for extra formatting*/
	,@overwrite bit = 1		/*overwrite the current file, default is true*/
	,@deleteold bit = 0		/*clean old xlsx files in the temp folder, default is false*/
	,@uploadonly bit = 0	/*just upload the previously exported file without recreating it*/
	,@uploadfile bit = 0	/*upload the file to SharePpoint after exporting, you need to specify sharepoint parameters*/
	,@SharePointUrl varchar(100) = '' /*Sharepoint URL to upload to, it should be like this https://sharepoint.crescent.com/sites/operations */
	,@Library  varchar(100) = '' /*child folder name of your sharepoint library, leave empty if it is the root folder*/
	,@Username varchar(100) = '' /*username to access sharepoint library*/
	,@Password varchar(100) = '' /*pasword to access sharepoint library*/
	,@AttachToMail bit = 0		/*attach file to email*/
	,@MailList varchar(100) = '' /*mail recipients*/
	,@MailSubject	varchar(100) = 'Subject' /*mail subject*/
	,@MailBody	varchar(100) = 'PFA' /*mail body*/
AS
BEGIN
	/*enabling cmd and configureing primary settings*/
	EXEC sp_configure 'show advanced options',1;reconfigure;
	EXEC sp_configure 'xp_cmdshell',1;reconfigure;
	
	IF @tablename IS NULL AND @Query IS NULL
	 RAISERROR ('No table or query were defined, Please provide either @tablename or @query', 16, 1)
	ELSE
	IF @tablename IS NOT NULL AND @Query IS NOT NULL
	 RAISERROR ('Two conflicting parameters defined, Please provide one parameter only', 16, 1)

ELSE    
BEGIN
	declare @csvfile varchar(100),@exelfile varchar(100),@fourpartname varchar(100)

	IF LEFT(@params,1) <> ' '
	 SET @params = ' '+@params
	IF @filename IS NULL
		SET @filename = ISNULL(@tablename,'QueryFile')
	IF @uploadonly = 0
		SET @overwrite = 0
	IF @overwrite = 0
		SET @deleteold = 0
	IF @deleteold = 1
		EXEC xp_cmdshell 'del /f %temp%\*.csv && del /f %temp%\*.xlsx',no_output
	IF @CSVOnly = 0
		set @deleteold = 0
/*==========================================================*/

	/*setting filenames and table names*/
	SET @csvfile = IIF(@filepath IS NOT NULL,CONCAT(@filepath,@filename),'%temp%\'+@filename)
	IF RIGHT(@csvfile,4) <> '.csv'
		SET @csvfile += '.csv'
	SET @exelfile = @filename
	IF  RIGHT(@exelfile,5) <> '.xlsx'
		SET @exelfile +='.xlsx'
	IF SUBSTRING(@filename,2,2) <> ':\'
		SET @exelfile = '%temp%\'+@filename+'.xlsx'
	SET @fourpartname = IIF(@databasename IS NOT NULL,CONCAT(@databasename,'.',@schemaname,'.',@tablename),@tablename)

	/*getting table headers and data for csv*/
	DECLARE 
    @pivotcolumns NVARCHAR(MAX) = '', 
    @sql     VARCHAR(4000) = '';

IF @overwrite = 1
	SET @sql = 'del '+@exelfile+' && ' +@sql
EXEC xp_cmdshell @sql,no_output

IF (@uploadonly = 0 AND  @CSVOnly = 1)
		
begin
		
		SET @params = ' -Encoding utf8'
		DROP TABLE IF EXISTS ##SQLCMD
		SELECT ISNULL(@Query,'SELECT * FROM '+@fourpartname) SQLStatement INTO ##SQLCMD
		
		EXEC xp_cmdshell 'bcp "SELECT SQLStatement FROM ##SQLCMD" queryout %temp%\SQLCMD.sql -T -c',no_output
		
		SET @sql = 'powershell "Invoke-Sqlcmd -InputFile %temp%\SQLCMD.sql -ServerInstance "'+@SQLSERVER+'" | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Csv -NoTypeInformation '+@csvfile+@params+'"'

		EXEC xp_cmdshell @sql
		EXEC xp_cmdshell 'del %temp%\SQLCMD.sql',no_output	
end
/*==========================================================*/

/*Exporting CSV to Excel using Export-Excel cmdlet*/	
	IF (@uploadonly = 0 AND  @CSVOnly = 0)
	BEGIN
		
		DROP TABLE IF EXISTS ##ps
		SELECT 
		'#Add required references to Export-Excel  
		$ModuleName = "ImportExcel"  
		#$Module = Find-Module $ModuleName    
		if(!(Get-PackageProvider | Where-Object Name -eq NuGet)){Find-PackageProvider -Name NuGet | Install-PackageProvider -Force}  
		if(!(Get-Module -ListAvailable | Where-Object Name -Like $ModuleName)) {Install-Module $ModuleName -Force}' ps INTO ##ps
		
		EXEC xp_cmdshell 'bcp "SELECT ps FROM ##ps" queryout %temp%\InstallExcelModule.ps1 -T -c',no_output
		
		EXEC xp_cmdshell 'powershell %temp%\InstallExcelModule.ps1',no_output
		EXEC xp_cmdshell 'del %temp%\InstallExcelModule.ps1',no_output
		
		DROP TABLE IF EXISTS ##SQLEXEL
		SELECT ISNULL(@Query,'SELECT * FROM '+@fourpartname) SQLStatement INTO ##SQLEXEL
		
		EXEC xp_cmdshell 'bcp "SELECT SQLStatement FROM ##SQLEXEL" queryout %temp%\SQLCMD.sql -T -c',no_output
		
		SET @sql = 'powershell "Invoke-Sqlcmd -InputFile %temp%\SQLCMD.sql -ServerInstance "'+@SQLSERVER+'" | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel '+@exelfile+@params+'"'

		EXEC xp_cmdshell @sql
		EXEC xp_cmdshell 'del %temp%\SQLCMD.sql',no_output

	END
/*==========================================================*/

/*fixing folder names for attachment*/
	DECLARE @folder varchar(50)
	DECLARE @table table (folder varchar(50))
	
	INSERT @table
	EXEC xp_cmdshell 'echo %username%'--,no_output
	
	SELECT @folder = 'c:\users\'+folder+'\appdata\local\temp' FROM @table WHERE folder IS NOT NULL

	IF left(@exelfile,6)='%temp%'
	 SELECT @exelfile = replace(@exelfile,'%temp%',@folder)

	SET @fullname = @exelfile

	IF @CSVOnly = 1
	BEGIN
	DELETE @table
	INSERT @table
	EXEC xp_cmdshell 'echo %username%'--,no_output
	
	SELECT @folder = 'c:\users\'+folder+'\appdata\local\temp' FROM @table WHERE folder IS NOT NULL

	IF left(@csvfile,6)='%temp%'
	 SELECT @csvfile = replace(@csvfile,'%temp%',@folder)
	SET @fullname = @csvfile
	END
/*==========================================================*/

/*uploading file to sharepoint library using Microsoft.Online.SharePoint.PowerShell*/
	IF @uploadfile = 1
		BEGIN
		DROP TABLE IF EXISTS ##ps1
		SELECT 
			'#Add required references to SharePoint client assembly to use CSOM  
			$ModuleName = "Microsoft.Online.SharePoint.PowerShell"  
			#$Module = Find-Module $ModuleName    
			if(!(Get-PackageProvider | Where-Object Name -eq NuGet)){Find-PackageProvider -Name NuGet | Install-PackageProvider -Force}  
			if(!(Get-Module -ListAvailable | Where-Object Name -Like $ModuleName)) {Install-Module $ModuleName -Force}    
			Import-Module $ModuleName    
			##$ModuleFiles= "\Microsoft.SharePoint.Client.dll","\Microsoft.SharePoint.Client.Runtime.dll"  
			##foreach ($file in $ModuleFiles) {Add-Type -Path (Join-Path ($Module.Path | Split-Path) $file)}  
			
			#Add required references to SharePoint client assembly to use CSOM     
			
			#Create Local Folder, if it doesnt exist  
			$FolderName = ($OneDriveURL.ServerRelativeURL) -replace "/","\"  
			#$LocalFolder = ""  
			#If (!(Test-Path -Path $LocalFolder)) {New-Item -ItemType Directory -Path $LocalFolder | Out-Null}    
			
			#Parameters  
			$AdminAccount = "'+@Username+'"  
			$aesKey = (7,3,0,4,5,32,5,23,5,3,1,1,36,9,18,9,6,0,1,9,5,1,76,23)  #any random combination
			$Plain = "'+REPLACE(@Password,'$','`$')+'"
			$Secure = ConvertTo-SecureString -String $Plain -AsPlainText -Force
			$Encrypted = ConvertFrom-SecureString $Secure -Key $aesKey
			$AdminPass = ConvertTo-SecureString -String $Encrypted -key $aesKey    
			
			#Specify Users OneDrive Site URL and Folder name  
			$OneDriveURL = "'+@SharePointUrl+'"  
			$LibraryName ="Documents"  
			$SubFolder = "Documents'+ISNULL('\'+@Library,'')+'"  
			$FilePath =  "'+@fullname+'"  
			$UniqueFileName = [System.IO.Path]::GetFileName($FilePath)    
			
			#Setup the context  
			$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($OneDriveURL)  
			$Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminAccount,$AdminPass)    
			
			#Get the Library  
			$List = $Ctx.Web.Lists.GetByTitle($LibraryName)  
			$Ctx.Load($List)  
			$Ctx.ExecuteQuery()  
			$uploadFolder = $ctx.Web.GetFolderByServerRelativeUrl($SubFolder);  
			
			#Use regular approach.  
			$FileStream = New-Object IO.FileStream($FilePath,[System.IO.FileMode]::Open)  
			$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation  
			$FileCreationInfo.Overwrite = $true  
			$FileCreationInfo.ContentStream = $FileStream  
			$FileCreationInfo.URL = $UniqueFileName  
			$Upload = $uploadFolder.Files.Add($FileCreationInfo)  
			$ctx.Load($Upload)  
			$ctx.ExecuteQuery()  
			
			Write-host "File uploaded Successfully!" -f Green   
			
			#Read more: https://www.sharepointdiary.com/2020/05/upload-large-files-to-sharepoint-online-using-powershell.html#ixzz7IWnNDnEK' ps INTO ##ps1

			PRINT 'Uploading '+@filename
			SELECT @SQL = 'bcp "SELECT ps FROM ##ps1" queryout %temp%\UploadAttachments.ps1 -T -c'
			EXEC xp_cmdshell @SQL,no_output
			
			EXEC xp_cmdshell 'powershell %temp%\UploadAttachments.ps1'
			EXEC xp_cmdshell 'del %temp%\UploadAttachments.ps1',no_output
		END

/*==========================================================*/

/*sending email with attached file*/
	IF @AttachToMail = 1
	EXEC msdb.dbo.sp_send_dbmail
	@recipients = @MailList,
	@subject = @MailSubject,
	@file_attachments = @fullname,
	@body = @MailBody

	SELECT @fullname
	PRINT @fullname

--CleanUP
	--EXEC xp_cmdshell 'del /f %temp%\*.csv && del /f %temp%\*.xlsx',no_output

EXEC sp_configure 'xp_cmdshell',0 reconfigure
EXEC sp_configure 'show advanced options',0 reconfigure
END
END
GO
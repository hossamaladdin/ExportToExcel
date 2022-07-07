CREATE PROC UploadToSharePoint @Username VARCHAR(100),@Password VARCHAR(100),@SharePointUrl VARCHAR(100),@Library VARCHAR(100),@filename VARCHAR(100)
AS
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
$FilePath =  "'+@filename+'"  
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

EXEC sp_configure 'show advanced options',1 reconfigure
EXEC sp_configure 'xp_cmdshell',1 reconfigure

EXEC xp_cmdshell  'bcp "SELECT ps FROM ##ps1" queryout %temp%\UploadToSP.ps1 -T -c',no_output
EXEC xp_cmdshell 'powershell %temp%\UploadToSP.ps1'
EXEC xp_cmdshell 'del %temp%\UploadToSP.ps1',no_output

EXEC sp_configure 'xp_cmdshell',0 reconfigure
EXEC sp_configure 'show advanced options',0 reconfigure
END
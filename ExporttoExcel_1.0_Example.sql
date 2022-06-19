DROP TABLE IF EXISTS ##TBL
SELECT 'somedate' [Column1],
1231 [Column2],
'Just Another Column' [Column3]
INTO ##TBL

exec [ExportToExcel] 
 @tablename= '##mytable'
,@attachtomail=1
,@csvonly=1
,@MailList='mymail@domain.com;myfriendmail@domain.com'
,@uploadfile=1
,@SharePointUrl='https://domain.sharepoint.com/personal/myfolder/'
,@Library='MyFolder'
,@Username='mymail@domain.com'
,@Password='SomePassword'
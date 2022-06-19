# ExportToExcel
SQL 2019 Stored Procedure to Export any table to Excel and send it via email or upload it to OneDrive or SharePoint

I'm working as an SQL Server DBA and Database Developer, one of my main responsibilites is to analyze data and send it in Excel workbooks.

Sometimes it becomes a heavy burden when I'm querying data from a remote server without direct clipboard access or file share, which takes me several minutes to export the data, format it and send it to people requested it.
Using CSV mail or SQLCMD is pretty neat, but it makes heavy reports hard to send via mail since the file is too large, and also the formatting is not cool, on the other hand SSIS is a very good tool, but it also requires a lot of development and debugging.

After a lot of research and staying late at the office, and with the help of some genius guys (Doug Finke & Salaudeen Rajack, I'll share their profiles below) I've come up with a quick solution that combines very handy features that could help you overcome this problem.

Introducing ExportToExcel_V1.0, a very useful and also customizable stored procedure for SQL Server 2019, it utilizez T-SQL, ImportExcel PS Module and SharePoint Online Management Shell to quickly and automatically send SQL reports in CSV or Excel formatted file.
I'll add more details later, so hang on.

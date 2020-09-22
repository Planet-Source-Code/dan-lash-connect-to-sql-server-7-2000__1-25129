<div align="center">

## Connect to SQL Server 7/2000


</div>

### Description

This code/tutorial is a beginners guide to connecting to your own SQL Server database.
 
### More Info
 
You have, and installed Microsoft ActiveX Data Objects 2.x


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Lash](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-lash.md)
**Level**          |Beginner
**User Rating**    |4.3 (52 globes from 12 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-lash-connect-to-sql-server-7-2000__1-25129/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
'First, Project->References->Microsoft ActiveX Data Objects
'Now, declare the connection
Dim adoConn As New adodb.connection
'Declare the recordset
Dim adoRS As New adodb.Recordset
'Declare the querey
Dim sqlString As String
'Set the connection string:
'Driver tells it were using SQL Server
'Server says what it is named (click the properties of your server through Enterprise Manager. If you don't have E.M. you must reinstall)
'Database is the database within the SQL Server we want
'Also tell it our login/password (which you can setup by adding a user to your database, there is a button that says add user)
adoConn.ConnectionString = "Driver={SQL Server}; " & _
  "Server=MYSQLSERVER; " & _
  "Database=testdatabase; " & _
  "UID=admin; " & _
  "PWD=test"
'Set the querey
sqlString = "SELECT * FROM personal"
'Open the connection
adoConn.Open
'Execute the querey:
'Tell it what we want
'Tell it where to get it
'Allow the user to fully navigate the recordset
'Tell it that were going to lock the records right after we edit them
'Tell the server that the command is in text format
adoRS.Open sqlString, _
 adoConn, _
 adOpenKeyset, _
 adLockPessimistic, _
 adCmdText
'Loop through our recordset
While Not adoRS.EOF
 lstNames.AddItem Trim(adoRS("id")) & ". " & _
 Trim(adoRS("fname")) & " " & Trim(adoRS("lname"))
 'Get the next record
 adoRS.MoveNext
Wend
'Close the recordset
adoRS.Close
'Close the connection
adoConn.Close
'Set the objects to nothing
Set adoRS = Nothing
Set adoConn = Nothing
End Sub
```


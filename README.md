<div align="center">

## Upload File To An FTP Server \(Authorization Required For MOST Servers\)


</div>

### Description

Many people have been bugging me to give them the source for uploading a file to their website. Well, for everyone that wants it, here is the open source code. It is in text format, so you don't have to download anything. Don't forget to vote for me!
 
### More Info
 
INPUTS (ARGUMENTS/PARAMETERS)

¯¯¯¯¯¯ ¯¯¯¯¯¯¯¯¯ ¯¯¯¯¯¯¯¯¯¯

InetControl (Inet): The Inet control to use for the operation.

strURL (String): The server's URL that you want to upload to. MOST SERVERS REQUIRE USERNAMES AND PASSWORDS SO DON'T THINK THAT YOU CAN UPLOAD WITHOUT AUTHORIZATION.

strUserName (String): The username used to login to the server.

strPassword (String): The password used to login to the server.

strLocalFile (String): The LOCAL path AND file name of the file to upload.

strRemoteFile (String): The REMOTE path AND file name to save the file as on the server. NOTE: IT MUST NOT BE A FULL PATH! USE '/' FOR THE ROOT DIRECTORY!

REQUIRED KNOWLEDGE (KNOW-NEEDS)

¯¯¯¯¯¯¯¯ ¯¯¯¯¯¯¯¯¯ ¯¯¯¯ ¯¯¯¯¯

Know how to add a control/component to a form.

OUTPUTS (RETURN VALUES)

¯¯¯¯¯¯¯ ¯¯¯¯¯¯ ¯¯¯¯¯¯

This function returns TRUE if the upload WAS successful.

This function returns FALSE if the upload WAS NOT successful.

SIDE EFFECTS (BUGS/PROBLEMS)

¯¯¯¯ ¯¯¯¯¯¯¯ ¯¯¯¯ ¯¯¯¯¯¯¯¯

There as no side effects for this code that I know of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Synthesize](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/synthesize.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/synthesize-upload-file-to-an-ftp-server-authorization-required-for-most-servers__1-23454/archive/master.zip)

### API Declarations

```
API CALLS (DECLARATIONS)
¯¯¯ ¯¯¯¯¯ ¯¯¯¯¯¯¯¯¯¯¯¯
NONE
```


### Source Code

```
Public Function UploadFile(InetControl As Inet, ByVal strURL As String, ByVal strUserName As String, ByVal strPassword As String, ByVal strLocalFile As String, ByVal strRemoteFile As String) As Boolean
  ' INPUTS (ARGUMENTS/PARAMETERS)
  ' ¯¯¯¯¯¯ ¯¯¯¯¯¯¯¯¯ ¯¯¯¯¯¯¯¯¯¯
  '  InetControl: The Inet control to use for the operation.
  '  strURL: The server's URL that you want to upload to. MOST SERVERS
  '    REQUIRE USERNAMES AND PASSWORDS SO DON'T THINK THAT YOU CAN UPLOAD
  '    WITHOUT AUTHORIZATION.
  '  strUserName: The username used to login to the server.
  '  strPassword: The password used to login to the server.
  '  strLocalFile: The LOCAL path AND file name of the file to upload.
  '  strRemoteFile: The REMOTE path AND file name to save the file as on
  '    the server. NOTE: IT MUST NOT BE A FULL PATH! USE '/' FOR THE ROOT
  '    DIRECTORY!
  '
  ' OUTPUTS (RETURN VALUES)
  ' ¯¯¯¯¯¯¯ ¯¯¯¯¯¯ ¯¯¯¯¯¯
  '  This function returns TRUE if the upload WAS successful.
  '  This function returns FALSE if the upload WAS NOT successful.
  '
  ' EXAMPLE:
  ' ¯¯¯¯¯¯¯
  '  Example: Put the following commented line of code in a command button:
  '    blnUpload = UploadFile(Inet1, "the.url.DO.NOT.USE.HTTP://", "server_username", "server_password", "C:\The Local Path\To The Local File\The File.exe", "/public_html/the_remote_path/thefile.exe")
  '  blnUpload will return TRUE if the upload was successful and FALSE if not.
  '  NOTICE: YOU MAY NEED TO USE '/public_html' BECAUSE THAT IS THE HOME
  '    DIRECTORY OF MOST SERVERS!
  '
  ' NOW TO THE REAL CODE:
  ' ¯¯¯ ¯¯ ¯¯¯ ¯¯¯¯ ¯¯¯¯
  '
  ' If we run into an error, go to the label statement 'ErrHandle_UploadFile'
  On Error GoTo ErrHandle_UploadFile
  ' If the selected Inet control is still processing it's last operation,
  '  goto the label statement 'ErrHandle_UploadFile'
  If InetControl.StillExecuting Then GoTo ErrHandle_UploadFile
  ' Make the code simpler by using the With statement.
  With InetControl
    ' Cancel the last request if one as slipped in between the last line
    '  of code and this one.
    .Cancel
    ' Set the Protocol of the selected Inet control.
    .Protocol = icFTP
    ' Set the URL of the selected Inet control. YOU MUST SET THE URL BEFORE
    '  YOU SET THE USERNAME AND PASSWORD.
    .URL = strURL
    ' Set the UserName of the selected Inet control.
    .UserName = strUserName
    ' Set the Password of the selected Inet control.
    .Password = strPassword
  End With
  ' Execute the 'PUT' command using the selected Inet control. The first param
  '  of the PUT command is the LOCAL file path and name. The second (last) param
  '  of the PUT command is the REMOTE file path and name.
  InetControl.Execute , "PUT " & Chr(34) & strLocalFile & Chr(34) & " " & Chr(34) & strRemoteFile & Chr(34)
  ' Create a loop and kill it when the selected Inet control is FINISHED executing
  '  it's last command (in our case, the last command is 'PUT').
  Do While InetControl.StillExecuting
    ' Allow the processor to carry on other tasks
    DoEvents
  ' Continue the loop.
  Loop
  ' The upload WAS successful, no errors. Set 'UploadFile' to TRUE.
  UploadFile = True
  ' Exit the function so that we don't trip anymore events.
  Exit Function
  ' Finally, the 'ErrHandle_UploadFile' label statement. This label statement,
  '  when accessed, will trigger the code below it.
ErrHandle_UploadFile:
  ' In our case, if we had an error or something, we want to return a FALSE
  '  value telling the user that the upload WAS NOT successful.
  UploadFile = False
  ' Then we exit the function just incase an error triggered this label
  '  statement.
  Exit Function
End Function
```


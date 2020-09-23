<div align="center">

## deletefiles


</div>

### Description

This ASP page is a maintenance tool of sorts. It first lists all of the files in the directory.

The links when pressed pass the file name back to the same page which calls a delete function that delete the file which was just clicked on.
 
### More Info
 
passes a querystring delete=Yes to call the delete function and passes a querystring FN= the filename to be deleted.

This code uses the Filesystemobject, passes variables via a querystring, and use a for each loop.

Again, make sure you want to delete the file before clicking.

List of all files in the directory minus the deleted file.

Make sure you are clicking on the right file, because once you click on it, it is gone.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jason Buck](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-buck.md)
**Level**          |Beginner
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jason-buck-deletefiles__4-6787/archive/master.zip)

### API Declarations

This code was written by Jason Buck for free distribution.


### Source Code

```
<html><head>
<title>delte a file</title>
</head><body>
<center><H2>WARNING</H2></center>
This code is very dangerous. When loaded, clicking on any file on this page
will cause that file to be deleted. I have not put in a confirm delete page.
I recommend doing so if you are ever going to use this delete function in a production
environment.
<p>For help or comments, feel free to contact me at <a href="mailto:buck_jason@hotmail.com">Hotmail</a>
or visit my temporary site at <a href="http://www22.brinkster.com/jbuck">www22.brinkster.com/jbuck</a>
<p>
Click on any file below to delete it. I recommend adding a few junk files for demostration
purposes.<p>
<%
strdelete = request.querystring("delete")
strFN = request.querystring("FN")
if strdelete = "Yes" Then
call functionDF()
End if
Sub functionDF()
  Dim fso, f1
  Set fso = CreateObject("Scripting.FileSystemObject")
  Response.Write "Deleting file <b>" & strFN & "</b><br>"
  Set f1 = fso.GetFile(Server.MapPath(strFN))
  f1.Delete
  Response.Write "All done!<br>"
End Sub
' the dot . below represents the current directory that this file is in.
' this file will only delete in its current directory. to delete in
' directories, you would have to add that to the tobdel portion which is the FN
' that is being passed.
dirtowalk = "."
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(server.mappath(dirtowalk))
Set fc = f.Files
For Each tobdel in fc
    response.write "<a href='deletefiles.asp?delete=Yes&FN=" & tobdel.name & "'>"
    response.write tobdel.name & "</a></br>"
Next
%>
</body></html>
```


<div align="center">

## General File System Object Functions


</div>

### Description

These are few general functions i have written related to File System Object:

- Check if a SPECIFIC folder exists

- Check if A SPECIFIC file EXISTS in A SPECIFIC folder

- Create A SPECIFIC folder

- Delete A SPECIFIC folder

- Delete ALL FILES in a SPECIFIC folder

- Delete A SPECIFIC file in A SPECIFIC folder

Your comments and highly appreciated.

And finally thanks to all coders here at PSC for sharing their knowledge and their work!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vladimir Miho](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vladimir-miho.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vladimir-miho-general-file-system-object-functions__4-8315/archive/master.zip)





### Source Code

```
<%
'================================================================='
'Check if a SPECIFIC folder exists								 '
'Input:															 '
'		- Path: the path where the folder is suppose to be		 '
'		- FolderName: the name of the folder you want to check for'
'Output:														 '
'		- True: if the folder exists							 '
'		- False: if the folder does not exist					 '
'================================================================='
Function CheckIfFolderExists(byVal Path, byVal FolderName)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	dim FullPath
	FullPath = Path & FolderName
	if objFSO.FolderExists(FullPath) then
		CheckIfFolderExists = true
	else
		CheckIfFolderExists = false
	end if
	set objFSO = nothing
End Function
'-----------------------------------------------------------------'
'================================================================='
'Checks if A SPECIFIC file EXISTS in A SPECIFIC folder			 '
'Input:															 '
'		- FolderPath: the path of the folder					 '
'		- FileName: the name of the file you want to check for  '
'Output:														 '
'		- True: if the file exists								 '
'		- False: if the file does not exist						 '
'================================================================='
Function CheckIfFileExists(byVal FolderPath, byVal FileName)
	CheckIfFileExists = false
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set objFolder = objFSO.GetFolder(FolderPath)
  		for each file in objFolder.files
			if Lcase(file.name) = Trim(LCase(Filename)) then
				CheckIfFileExists = true
			end if
		next
	set objFSO = nothing
End Function
'-----------------------------------------------------------------'
'================================================================='
'Creates A SPECIFIC folder										 '
'Input:															 '
'		- Path: path where you want to create the folder		 '
'		- NewFolderName: the name of the new folder			   '
'Output:														 '
'		- True: if the folder is created successfully			 '
'		- False: if the folder creation failed					 '
'================================================================='
Function CreateNewFolder(byVal Path, byVal NewFolderName)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	dim FullPath
	FullPath = Path & NewFolderName
	if not objFSO.FolderExists(FullPath) then
		objFSO.CreateFolder(FullPath)
		CreateNewFolder = true
	else
		CreateNewFolder = false
	end if
	set objFSO = nothing
End Function
'-----------------------------------------------------------------'
'================================================================='
'Deletes A SPECIFIC folder										 '
'Input:															 '
'		- Path: path where the folder is						 '
'		- FolderName: the name of the folder you want to delete  '
'Output:														 '
'		- True: if the folder is deleted						 '
'		- False: if the folder could not be deleted				 '
'================================================================='
Function DeleteFolder(byVal Path, byVal FolderName)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	dim FullPath
	FullPath = Path & FolderName
	if objFSO.FolderExists(FullPath) then
		objFSO.DeleteFolder(FullPath)
		DeleteFolder = true
	else
		DeleteFolder = false
	end if
	set objFSO = nothing
End Function
'-----------------------------------------------------------------'
'================================================================='
'Deletes ALL FILES in a SPECIFIC folder							 '
'Input:															 '
'		- Path: path where the folder is						 '
'		- FolderName: the name of the folder where the files you '
'					 want to delete are						 '
'Output: none													 '
'================================================================='
Function DeleteFilesInFolder(byVal Path, byVal FolderName)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	dim FullPath
	FullPath = Path & FolderName
	set objFolder = objFSO.GetFolder(FullPath)
  		for each file in objFolder.files
			file.delete(true)
		next
	set objFSO = nothing
End Function
'-----------------------------------------------------------------'
'================================================================='
'Deletes A SPECIFIC file in A SPECIFIC folder					 '
'Input:															 '
'		- Path: path where the folder is						 '
'		- FolderName: the name of the folder where the file you  '
'					 want to delete is							 '
'		- FileName: name of the file you want to delete			 '
'Output:														 '
'		- True: if the file is deleted						   '
'		- False: if the file could not be deleted				 '
'================================================================='
Function DeleteFileInFolder(byVal Path, byVal FolderName, byVal FileName)
	DeleteFileInFolder = False
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	dim FullPath
	FullPath = Path & FolderName
	set objFolder = objFSO.GetFolder(FullPath)
		for each file in objFolder.files
			if Lcase(file.name) = Trim(LCase(Filename)) then
				file.delete(true)
				DeleteFileInFolder = True
			end if
		next
	set objFSO = nothing
End Function
'-----------------------------------------------------------------'
%>
```


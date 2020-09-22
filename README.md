<div align="center">

## CreatePath\(\) Function


</div>

### Description

CreatePath() uses API calls to create a directory tree. A simple application is included to demonstrate its use.
 
### More Info
 
sPath - A string containing the required directory structure.

Module modCreatePath contains all of the code and API declarations required to use the function in your project.

Boolean - True if the directory structure has been created or already exists, False on error.


<span>             |<span>
---                |---
**Submitted On**   |2001-01-09 01:31:30
**By**             |[Dave Lambert](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-lambert.md)
**Level**          |Advanced
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD13560182001\.zip](https://github.com/Planet-Source-Code/dave-lambert-createpath-function__1-14264/archive/master.zip)

### API Declarations

```
Public Type SECURITY_ATTRIBUTES
 nLength As Long
 lpSecurityDescriptor As Long
 bInheritHandle As Long
End Type
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
```






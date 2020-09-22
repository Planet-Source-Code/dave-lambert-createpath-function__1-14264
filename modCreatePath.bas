Attribute VB_Name = "modCreatePath"
Option Explicit

Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

' Function ensures that the specified Path exists, and if not, creates it.
'
' Arguments:
'         sPath
'             A string containing the required directory structure.
'
' Returns:
'         True on success, otherwise False.
'
' Notes:
'         Trailing \'s are ignored so can be included without causing problems.
'
' Examples:
'         bSuccess = CreatePath("C:\Dir1\Subdir1\Subdir2")
'         bSuccess = CreatePath("C:\Dir1\Subdir1\Subdir2\")
'         bSuccess = CreatePath("\\MyServer\Dir1\Subdir1\Subdir2")
'
' © D R Lambert - All Rights Reserved
' Any questions to vbcode@drldev.co.uk
' You may distribute this code in a compiled application.
'
Public Function CreatePath(ByVal sPath As String) As Boolean
  Dim RetArray() As String
  Dim intUbound As Integer
  Dim N As Integer
  Dim intDev As Integer
  Dim lngAttr As Long
  Dim apiSA As SECURITY_ATTRIBUTES
  Dim lngRetVal As Long
  
  Const CR2 = vbCrLf & vbCrLf         ' used for dialog box text formatting
  
  CreatePath = False                  ' it is of course false by default, but never
                                      ' assume anything
  
  If Len(sPath) = 0 Then              ' just in case the path is empty
    Exit Function
  End If
  
  lngAttr = GetFileAttributes(sPath)  ' check the requested directory path
  
  If lngAttr <> -1 Then               ' the complete path already exists
    CreatePath = True                 ' flagged as successful
    Exit Function                     ' so exit the function
  End If
  
  If Left$(sPath, 2) = "\\" Then      ' the root device is a server
    intDev = 2
  Else                                ' the root device is a disk or mapped drive
    intDev = 0
  End If
  
  RetArray() = Split(sPath, "\")      ' retrieve the substring components of the path
  
  intUbound = UBound(RetArray)        ' find the index of the last substring returned
                                      ' by split()
  
  lngAttr = 0                         ' reset our directory attributes variable
  
  For N = intDev To intUbound         ' loop through each element of the array
  
    If (intDev = N) And (intDev = 2) Then
      sPath = "\\" & RetArray(N)      ' this substring is the server name
      
    ElseIf (intDev = N) And (intDev = 0) Then
      sPath = RetArray(N)             ' this substring is the drive
      
    Else
      If Len(RetArray(N)) > 0 Then    ' these substrings are the directories
      
        sPath = sPath & "\" & RetArray(N) ' re-assemble the path one directory at a time
        
        If lngAttr <> -1 Then         ' once one directory doesn't exist, then anything after
                                      ' it obviously doesn't exist either so we don't need
                                      ' to check from that point onward
          lngAttr = GetFileAttributes(sPath)
        End If
        
        If lngAttr = -1 Then          ' the directory doesn't exist
          'Debug.Print "Create path: " & sPath
          If CreateDirectory(sPath, apiSA) <> 1 Then  ' create it
            MsgBox "An error occurred while creating directory" & CR2 & _
                   """" & sPath & """", vbCritical, "CreatePath() Error"
            Exit Function
          End If
        'Else
        '  Debug.Print "Path exists: " & sPath
        End If
                                                           
      End If
    End If
  Next
  CreatePath = True                   ' if we reached here then we succeeded
End Function


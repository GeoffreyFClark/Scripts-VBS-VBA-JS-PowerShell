' Force variable declaration
Option Explicit

' Declare variables
Dim objShell
Dim filePaths, filePath, i

' WScript Shell object
Set objShell = CreateObject("WScript.Shell")

' Array of filepaths
filePaths = Array( _
"""C:\filepath\file1.pdf""", _
"""C:\filepath\file2.pdf""", _
"""C:\filepath\file3.pdf""" _
) 

' Loop through each to open
For i = 0 To UBound(filePaths)
    filePath = filePaths(i)
    objShell.Run filePath
Next

' Free up memory, release reference to WScript Shell object
Set objShell = Nothing

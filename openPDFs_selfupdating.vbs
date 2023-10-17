' Force variable declaration
Option Explicit

' Declare variables
Dim objShell, objFSO, folderPath, file, closestMatch
Dim filePaths, filePath, i

' WScript Shell object
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Array of PDF file paths to open
filePaths = Array("C:\filepath\file1.pdf", _
                  "C:\filepath\file2.pdf", _
                  "C:\filepath\file3.pdf") 

' Loop through each to open
For i = 0 To UBound(filePaths)
    filePath = filePaths(i)
    
    ' Check if file exists
    If objFSO.FileExists(filePath) Then
        objShell.Run filePath
    Else
        ' If not, attempt to find a similar file
        folderPath = objFSO.GetParentFolderName(filePath)
        closestMatch = FindSimilarFile(folderPath, objFSO.GetFileName(filePath))
        If closestMatch <> "" Then
            objShell.Run folderPath & "\" & closestMatch
        End If
    End If
Next

' Free up memory, release references to objects
Set objShell = Nothing
Set objFSO = Nothing

Function FindSimilarFile(folderPath, originalFileName)
    Dim file, bestMatch
    Dim parts, part, matchedParts
    Dim baseName, fileBaseName
    baseName = Split(originalFileName, ".")(0)
    parts = Split(baseName, " ") ' Split the base name by spaces
    bestMatch = ""
    matchedParts = 0

    ' Iterate over each file in the folder
    For Each file In objFSO.GetFolder(folderPath).Files
        ' Extract the base name of the file (without extension)
        fileBaseName = Split(file.Name, ".")(0)
        
        matchedParts = 0
        For Each part In parts
            ' Increase the count for every part of original file name found in current file
            If InStr(1, fileBaseName, part, vbTextCompare) > 0 Then
                matchedParts = matchedParts + 1
            End If
        Next

        ' If the majority of the parts match, assume this is the best match
        ' You can adjust the logic here to better fit your needs
        If matchedParts >= (UBound(parts) + 1) / 2 Then
            bestMatch = file.Name
            Exit For
        End If
    Next

    FindSimilarFile = bestMatch
End Function

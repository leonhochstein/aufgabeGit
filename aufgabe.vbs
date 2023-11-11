Const ForReading = 1 
Const ForWriting = 2
Const ForAppending = 8
Dim lastline

Set oWSH = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim strFilePath
strFilePath = "pyramid.txt"

If objFSO.FileExists(strFilePath) Then
    Set objFile = objFSO.OpenTextFile(strFilePath)

    Do Until objFile.AtEndOfStream
       lastline = objFile.ReadLine
    Loop

    objFile.Close

    If lastline = "XXXXXXXXXXXXXXX" Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        objFile.WriteLine "       X"
        objFile.Close
    Else
        removedSpace = Replace(lastline, " ", "",1,1)
        line = removedSpace & "XX"

        Set objFile = objFSO.OpenTextFile(strFilePath, ForAppending)
        objFile.WriteLine line
        objFile.Close
    End If
Else
    Set objFile = objFSO.CreateTextFile(strFilePath)
    objFile.WriteLine "       X"
    objFile.Close
End If

Set objFile = objFSO.OpenTextFile(strFilePath)
text = objFile.ReadAll
WScript.StdOut.Write text
objFile.Close
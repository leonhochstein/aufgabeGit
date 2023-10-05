Const ForReading = 1 
Const ForWriting = 2

Set oWSH = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim strFilePath
strFilePath = "counter.txt"

If objFSO.FileExists(strFilePath) Then
    Set objFile = objFSO.OpenTextFile(strFilePath)
    Dim content
    content = objFile.ReadAll
    objFile.Close

    If content = 0 Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        WScript.StdOut.WriteLine "       X"
        objFile.WriteLine "1"
        objFile.Close

    ElseIf content = 1 Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        WScript.StdOut.WriteLine "       X"
        WScript.StdOut.WriteLine "      XXX"
        objFile.WriteLine "2"
        objFile.Close

    ElseIf content = 2 Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        WScript.StdOut.WriteLine "       X"
        WScript.StdOut.WriteLine "      XXX"
        WScript.StdOut.WriteLine "     XXXXX"
        objFile.WriteLine "3"
        objFile.Close

    ElseIf content = 3 Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        WScript.StdOut.WriteLine "       X"
        WScript.StdOut.WriteLine "      XXX"
        WScript.StdOut.WriteLine "     XXXXX"
        WScript.StdOut.WriteLine "    XXXXXXX"
        objFile.WriteLine "4"
        objFile.Close

    ElseIf content = 4 Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        WScript.StdOut.WriteLine "       X"
        WScript.StdOut.WriteLine "      XXX"
        WScript.StdOut.WriteLine "     XXXXX"
        WScript.StdOut.WriteLine "    XXXXXXX"
        WScript.StdOut.WriteLine "   XXXXXXXXX"
        objFile.WriteLine "5"
        objFile.Close

    ElseIf content = 5 Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        WScript.StdOut.WriteLine "       X"
        WScript.StdOut.WriteLine "      XXX"
        WScript.StdOut.WriteLine "     XXXXX"
        WScript.StdOut.WriteLine "    XXXXXXX"
        WScript.StdOut.WriteLine "   XXXXXXXXX"
        WScript.StdOut.WriteLine "  XXXXXXXXXXX"
        objFile.WriteLine "6"
        objFile.Close

    ElseIf content = 6 Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        WScript.StdOut.WriteLine "       X"
        WScript.StdOut.WriteLine "      XXX"
        WScript.StdOut.WriteLine "     XXXXX"
        WScript.StdOut.WriteLine "    XXXXXXX"
        WScript.StdOut.WriteLine "   XXXXXXXXX"
        WScript.StdOut.WriteLine "  XXXXXXXXXXX"
        WScript.StdOut.WriteLine " XXXXXXXXXXXXX"
        objFile.WriteLine "7"
        objFile.Close

    ElseIf content = 7 Then
        Set objFile = objFSO.OpenTextFile(strFilePath, ForWriting)
        WScript.StdOut.WriteLine "       X"
        WScript.StdOut.WriteLine "      XXX"
        WScript.StdOut.WriteLine "     XXXXX"
        WScript.StdOut.WriteLine "    XXXXXXX"
        WScript.StdOut.WriteLine "   XXXXXXXXX"
        WScript.StdOut.WriteLine "  XXXXXXXXXXX"
        WScript.StdOut.WriteLine " XXXXXXXXXXXXX"
        WScript.StdOut.WriteLine "XXXXXXXXXXXXXXX"
        objFile.WriteLine "0"
        objFile.Close 
    End If
Else
    Set objFile = objFSO.CreateTextFile(strFilePath)
    objFile.WriteLine "1"
    WScript.StdOut.WriteLine "       X"
    objFile.Close
End If

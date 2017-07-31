Set objFSO=CreateObject("Scripting.FileSystemObject")

' How to write file
outFile="c:\test\file.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write "test string" & vbCrLf
objFile.Close

'How to read a file
strFile = "c:\test\file.txt"
Set objFile = objFS.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine= objFile.ReadLine
    Wscript.Echo strLine
Loop
objFile.Close

'to get file path without drive letter, assuming drive letters are c:, d:, etc
strFile="c:\test\file"
s = Split(strFile,":")
WScript.Echo s(1)
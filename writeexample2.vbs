'~ Create a FileSystemObject
Set objFSO=CreateObject("Scripting.FileSystemObject")

'~ Provide file path
outFile="c:\code\Results.txt"

'~ Setting up file to write
Set objFile = objFSO.CreateTextFile(outFile,True)


strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_DataFile Where Extension = 'mdb' OR Extension = 'ldb'")

For Each obj_File in colFiles
    'Wscript.Echo objFile.Name  'Commented out

    '~ Write to file
    objFile.WriteLine obj_File.Name
Next

'~ Close the file
objFile.Close
